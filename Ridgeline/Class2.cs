//using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.DatabaseServices;
//using Autodesk.AutoCAD.EditorInput;
//using Autodesk.AutoCAD.Geometry;
//using Autodesk.AutoCAD.PlottingServices;
//using System.Collections.Generic;
//using System.Collections.Specialized;
//using System.IO;
//using System.Runtime.InteropServices;
//using System;

//namespace PlottingApplication
//{
//    public class PlottingCommands
//    {
//        [CommandMethod("mplot")]
//        public static void MultiSheetPlot()
//        {
//            Document doc = Application.DocumentManager.MdiActiveDocument;
//            Editor ed = doc.Editor;
//            Database db = doc.Database;

//            using (Transaction tr = db.TransactionManager.StartTransaction())
//            {
//                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

//                PlotInfo pi = new PlotInfo();
//                PlotInfoValidator piv = new PlotInfoValidator
//                {
//                    MediaMatchingPolicy = MatchingPolicy.MatchEnabled
//                };

//                // Check if a plot is already in progress
//                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
//                {
//                    using (PlotEngine pe = PlotFactory.CreatePublishEngine())
//                    {
//                        using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
//                        {
//                            ObjectIdCollection layoutsToPlot = new ObjectIdCollection();

//                            foreach (ObjectId btrId in bt)
//                            {
//                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
//                                if (btr.IsLayout && btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper())
//                                {
//                                    layoutsToPlot.Add(btrId);
//                                }
//                            }

//                            int numSheet = 1;

//                            foreach (ObjectId btrId in layoutsToPlot)
//                            {
//                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
//                                Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

//                                PlotSettings ps = new PlotSettings(lo.ModelType);
//                                ps.CopyFrom(lo);

//                                PlotSettingsValidator psv = PlotSettingsValidator.Current;

//                                // Set plot settings
//                                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
//                                psv.SetUseStandardScale(ps, true);
//                                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
//                                psv.SetPlotCentered(ps, true);
//                                psv.SetPlotConfigurationName(ps, "DWFx ePlot (XPS Compatible).pc3", "ANSI_A_(8.50_x_11.00_Inches)");

//                                pi.Layout = btr.LayoutId;
//                                LayoutManager.Current.CurrentLayout = lo.LayoutName;
//                                pi.OverrideSettings = ps;
//                                piv.Validate(pi);

//                                if (numSheet == 1)
//                                {
//                                    // Setup progress dialog
//                                    SetupProgressDialog(ppd, layoutsToPlot.Count);
//                                    pe.BeginPlot(ppd, null);
//                                    pe.BeginDocument(pi, doc.Name, null, 1, true, "c:\\test-multi-sheet");
//                                }

//                                // Plot each sheet
//                                PlotSheet(ppd, pe, pi, doc, numSheet, layoutsToPlot.Count);
//                                numSheet++;
//                            }

//                            // Finish the document and plot
//                            pe.EndDocument(null);
//                            ppd.PlotProgressPos = 100;
//                            ppd.OnEndPlot();
//                            pe.EndPlot(null);
//                        }
//                    }
//                }
//                else
//                {
//                    ed.WriteMessage("\nAnother plot is in progress.");
//                }
//            }
//        }

//        private static void SetupProgressDialog(PlotProgressDialog ppd, int totalSheets)
//        {
//            ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
//            ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
//            ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
//            ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
//            ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
//            ppd.LowerPlotProgressRange = 0;
//            ppd.UpperPlotProgressRange = 100;
//            ppd.PlotProgressPos = 0;

//            ppd.OnBeginPlot();
//            ppd.IsVisible = true;
//        }

//        private static void PlotSheet(PlotProgressDialog ppd, PlotEngine pe, PlotInfo pi, Document doc, int currentSheet, int totalSheets)
//        {
//            ppd.StatusMsgString = "Plotting " + doc.Name.Substring(doc.Name.LastIndexOf("\\") + 1) + " - sheet " + currentSheet.ToString() + " of " + totalSheets.ToString();

//            ppd.OnBeginSheet();
//            ppd.LowerSheetProgressRange = 0;
//            ppd.UpperSheetProgressRange = 100;
//            ppd.SheetProgressPos = 0;

//            PlotPageInfo ppi = new PlotPageInfo();
//            pe.BeginPage(ppi, pi, (currentSheet == totalSheets), null);
//            pe.BeginGenerateGraphics(null);
//            ppd.SheetProgressPos = 50;
//            pe.EndGenerateGraphics(null);

//            // Finish the sheet
//            pe.EndPage(null);
//            ppd.SheetProgressPos = 100;
//            ppd.OnEndSheet();
//        }
//    }
//}

//public class PlotSaved

//{
//    [CommandMethod("BoxMake")]
//    public static void boxMake()
//    {
//        Document doc = Application.DocumentManager.MdiActiveDocument;
//        Editor ed = doc.Editor;
//        // Make a rectangle with points at 0,0 and 2.5,4
//        ed.Command("_.RECTANG", new Point3d(0, 0, 0), new Point3d(2.5, 4, 0));
//    }

//    static int inputRows = 1; // This variable is to capture user input for the number of rows to plot
//    static int inputColumns = 1; // This variable is to capture user input for the number of columns to plot

//    // These two are for iterating over the rows and columns of titleblocks
//    static int currentRow = 1;
//    static int currentColumn = 1;

//    // Index variables to track current location for the List arrays
//    static int currentPage = 1;
//    static int currentPoint = 0;

//    // These two are for calculating the width and height of the plot area
//    static double plotBoxHeight = 0.0;
//    static double plotBoxWidth = 0.0;

//    // This list stores all points for the window plot area
//    // Index [0] is the lower left corner of the window
//    // Index [1] is the upper right corner of the window
//    // This pattern repeats for each window. The first two indexes are always the user input
//    static List<Point3d> points = new List<Point3d>();

//    // This list stores all the plot windows
//    // The first in the list is the first window based off of user input
//    static List<Extents2d> plotWindows = new List<Extents2d>();


//    // This is from late night magic coding while inebriated. It helps control the orientation of the plot.
//    // I'm not sure what made me go this route, but it works and I don't want to mess with it. - EA
//    [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
//    static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);


//    [CommandMethod("PDFZ")]
//    public static void PDFA()
//    {
//        // Main method

//        // Open document objects for using in document database transactions
//        Document doc = Application.DocumentManager.MdiActiveDocument;
//        Editor ed = doc.Editor;
//        Database db = doc.Database;

//        // Clear and ensure its in memory
//        plotWindows.Clear();
//        points.Clear();


//        using (Transaction tr = db.TransactionManager.StartTransaction())
//        {
//            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

//            // This object is for the PlotInfo to get the layout ID
//            //ObjectId modelSpaceId = bt[BlockTableRecord.ModelSpace];

//            PlotInfo pi = new PlotInfo();
//            PlotInfoValidator piv = new PlotInfoValidator
//            {
//                MediaMatchingPolicy = MatchingPolicy.MatchEnabled
//            };

//            // Check if a plot is already in progress
//            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
//            {
//                // Get the rows and columns of windows to plot
//                inputRowCol(ed);

//                // Get first window
//                inputPoints(ed);
//                addWindow();

//                //convert from UCS to DCS
//                //Extents2d window = convertToWindow(first, second);
//                // Collection of Extents2d objects
//                bool keepCollecting = true;
//                while (keepCollecting)
//                {
//                    // Translate the plot area to the right
//                    translatePlotAreaRight();
//                    // Add most recent points to window list
//                    addWindow();
//                    currentColumn++;

//                    // Check to see if at end of row
//                    if (currentColumn == inputColumns)
//                    {
//                        if (currentRow == inputRows)
//                        {
//                            keepCollecting = false;
//                        }

//                        //I could rewrite this to make the logic structure less confusing. Checking the boolean value is a bit redundant
//                        // However, I don't want to rewrite everything right now. Its not elegant, but if the extra "if" statements create
//                        // enough of a performance hit then I have to do a lot better - EA
//                        if (keepCollecting)
//                        {
//                            // Reset column count and move to next row
//                            currentColumn = 1;

//                            // Move the plot area down and return to the start of the row. Add that plot area to the window list
//                            translatePlotAreaDownandReturn();
//                            addWindow();

//                            currentRow++;
//                        }
//                    }

//                }

//                //testRectangles(plotWindows, ed);

//                using (PlotEngine pe = PlotFactory.CreatePublishEngine())
//                {
//                    using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
//                    {
//                        // Add each window to a sheet for the plot engine
//                        foreach (Extents2d window in plotWindows)
//                        {

//                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
//                            Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

//                            PlotSettings ps = new PlotSettings(true);   // Pass true to indicate model space
//                            ps.CopyFrom(lo);

//                            PlotSettingsValidator psv = PlotSettingsValidator.Current;

//                            // Set plot settings for each page
//                            //psv.SetPlotRotation(ps, PlotRotation.Degrees000);
//                            psv.SetPlotWindowArea(ps, window);
//                            psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
//                            psv.SetUseStandardScale(ps, true);
//                            psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
//                            psv.SetPlotCentered(ps, true);

//                            // Adding at the bottom of the stack to see how it affects media rotation
//                            psv.SetPlotRotation(ps, PlotRotation.Degrees000);

//                            // Set plot device and page size
//                            psv.SetPlotConfigurationName(ps, "Bluebeam PDF", "Letter");
//                            var mns = psv.GetCanonicalMediaNameList(ps);
//                            if (mns.Contains("Letter"))
//                            { psv.SetCanonicalMediaName(ps, "Letter"); }
//                            else
//                            { string mediaName = setClosestMediaName(psv, ps, 8.5, 11, true); }

//                            // Lock in settings. Maybe delete the refreshLists...Unsure of its purpose
//                            // Older code bases seem to use it more than newer ones. I'm assuming it was needed and now
//                            // the Validation might take care of it? I need to look into that further- EA
//                            psv.RefreshLists(ps);

//                            pi.Layout = btr.LayoutId;
//                            LayoutManager.Current.CurrentLayout = lo.LayoutName;
//                            pi.OverrideSettings = ps;
//                            piv.Validate(pi);

//                            if (currentPage == 1)
//                            {
//                                setupProgressDialog(ppd, plotWindows.Count);
//                                pe.BeginPlot(ppd, null);

//                                // Prompt the user for the file name
//                                string fileLoc = Path.GetDirectoryName(doc.Name);

//                                // The "nameless" comes from refactored code. It should be a user input - EA
//                                pe.BeginDocument(pi, doc.Name, null, 1, false, fileLoc + @"\" + "Nameless");
//                            }

//                            //Plot each sheet
//                            plotSheet(ppd, pe, pi, doc, currentPage, plotWindows.Count);
//                            currentPage++;

//                        }
//                        pe.EndDocument(null);
//                        ppd.PlotProgressPos = 100;
//                        ppd.OnEndPlot();
//                        pe.EndPlot(null);
//                    }
//                }
//            }
//            else
//            {
//                ed.WriteMessage("\nAnother plot is in progress.");
//            }
//            tr.Commit();
//            // Probably not needed. But it makes sense to me to have it here. - EA
//            tr.Dispose();
//        }
//    }

//    public static void inputPoints(Editor ed)
//    {
//        //ask user for print window
//        Point3d first;
//        PromptPointOptions ppo = new PromptPointOptions("\nSelect First corner of plot area: ");
//        ppo.AllowNone = false;
//        PromptPointResult ppr = ed.GetPoint(ppo);
//        if (ppr.Status == PromptStatus.OK)
//        { first = ppr.Value; }
//        else
//            return;


//        Point3d second;
//        PromptCornerOptions pco = new PromptCornerOptions("\nSelect second corner of the plot area.", first);
//        ppr = ed.GetCorner(pco);
//        if (ppr.Status == PromptStatus.OK)
//        { second = ppr.Value; }
//        else
//            return;
//        points.Add(first);
//        currentPoint++;
//        points.Add(second);
//        currentPoint++;

//    }


//    // This is a helper method to convert the points from UCS to DCS
//    // Combined with the interop above I copied this in part from another program. Much of the code is
//    // likely not needed. A window should just have the two points as 'double' values. - EA
//    public static Extents2d convertToWindow(Point3d firstNum, Point3d secondNum)
//    {
//        double minX;
//        double minY;
//        double maxX;
//        double maxY;

//        //sort through the values to be sure that the correct first and second are assigned
//        if (firstNum.X < secondNum.X)
//        { minX = firstNum.X; maxX = secondNum.X; }
//        else
//        { maxX = firstNum.X; minX = secondNum.X; }

//        if (firstNum.Y < secondNum.Y)
//        { minY = firstNum.Y; maxY = secondNum.Y; }
//        else
//        { maxY = firstNum.Y; minY = secondNum.Y; }

//        // Adds the width and height of the box to global variables
//        plotBoxWidth = Math.Abs(minX - maxX);
//        plotBoxHeight = Math.Abs(minY - maxY);

//        Point3d first = new Point3d(minX, minY, 0);
//        Point3d second = new Point3d(maxX, maxY, 0);
//        //converting numbers to something the system uses (DCS) instead of UCS
//        ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
//        double[] firres = new double[] { 0, 0, 0 };
//        double[] secres = new double[] { 0, 0, 0 };
//        //convert points
//        acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
//        acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
//        Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

//        return window;
//    }

//    public static void addWindow()
//    {
//        // Get the most recent index of the points list
//        int lastPoint = points.Count - 1;

//        // Add the two most recent points and put them in an extents2d object
//        plotWindows.Add(convertToWindow(points[lastPoint - 1], points[lastPoint]));


//    }

//    // This gets user input to determine the number of rows and columns to plot
//    public static void inputRowCol(Editor ed)
//    {
//        // Prompt for the first integer
//        PromptIntegerOptions intOptions1 = new PromptIntegerOptions("\nEnter the numbers of rows: ")
//        {
//            AllowNegative = false,
//            AllowZero = true,
//            AllowNone = false
//        };
//        PromptIntegerResult intResult1 = ed.GetInteger(intOptions1);
//        if (intResult1.Status != PromptStatus.OK) return;
//        inputRows = intResult1.Value;

//        // Prompt for the second integer
//        PromptIntegerOptions intOptions2 = new PromptIntegerOptions("\nEnter the number of columns: ")
//        {
//            AllowNegative = false,
//            AllowZero = true,
//            AllowNone = false
//        };
//        PromptIntegerResult intResult2 = ed.GetInteger(intOptions2);
//        if (intResult2.Status != PromptStatus.OK) return;
//        inputColumns = intResult2.Value;
//    }


//    // This method will translate the plot area to the right
//    public static void translatePlotAreaRight()
//    {
//        // Get the most recent index of the points list
//        int lastPoint = points.Count - 1;
//        // Get the two most recent points and move them to the right on the X axis by 1 box width
//        Point3d first = new Point3d(points[lastPoint - 1].X + plotBoxWidth, points[lastPoint - 1].Y, 0);
//        Point3d second = new Point3d(points[lastPoint].X + plotBoxWidth, points[lastPoint].Y, 0);
//        // Add the two most recent points and put them the points list
//        points.Add(first);
//        points.Add(second);
//    }

//    // This method will translate the plot area down one row
//    public static void translatePlotAreaDownandReturn()
//    {
//        // Get the most recent index of the points list
//        int lastPoint = points.Count - 1;

//        // Variable for the count total width of the row. The first point only goes back 4 spaces. 
//        double calculatedWidth = (plotBoxWidth * inputColumns) - plotBoxWidth;

//        // Get the two most recent points, and move them down on the Y axis by 1 box height. Then move them back to the start of the row
//        Point3d first = new Point3d(points[lastPoint - 1].X - calculatedWidth, points[lastPoint - 1].Y - plotBoxHeight, 0);
//        Point3d second = new Point3d(points[lastPoint].X - calculatedWidth, points[lastPoint].Y - plotBoxHeight, 0);
//        points.Add(first);
//        points.Add(second);
//    }


//    // Create a series of rectangle objects whos points will be the windows. Use either points or plotWindows
//    public static void testRectangles(List<Extents2d> extentsList, Editor ed)
//    {
//        for (int i = 0; i < points.Count; i += 2)
//        {
//            // Assuming each pair of points (i, i+1) defines a rectangle
//            Point3d lowerLeft = points[i];
//            Point3d upperRight = points[i + 1];

//            // Use the RECTANG command with the points as arguments
//            ed.Command("_.RECTANG", lowerLeft, upperRight);
//        }
//    }

//    /* SOMETHING AINT RIGHT with testRectangles. I saw some strange interaction once where the testRectangles were created in quadrant 3 & 4, 
//     * and I haven't seen it again. That means there is a bug in here somewhere... - EA */

//    // Create the progress dialog. This is a helper method to keep the main method clean
//    // totalSheets should be a calculated number of inputRows * inputColumns
//    private static void setupProgressDialog(PlotProgressDialog ppd, int totalSheets)
//    {
//        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
//        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
//        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
//        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
//        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
//        ppd.LowerPlotProgressRange = 0;
//        ppd.UpperPlotProgressRange = 100;
//        ppd.PlotProgressPos = 0;

//        ppd.OnBeginPlot();
//        ppd.IsVisible = true;
//    }

//    // This method will add each window to a sheet for the plot engine
//    private static void plotSheet(PlotProgressDialog ppd, PlotEngine pe, PlotInfo pi, Document doc, int currentSheet, int totalSheets)
//    {
//        ppd.StatusMsgString = "Plotting " + doc.Name.Substring(doc.Name.LastIndexOf("\\") + 1) + " - sheet " + currentSheet.ToString() + " of " + totalSheets.ToString();

//        ppd.OnBeginSheet();
//        ppd.LowerSheetProgressRange = 0;
//        ppd.UpperSheetProgressRange = 100;
//        ppd.SheetProgressPos = 0;

//        PlotPageInfo ppi = new PlotPageInfo();
//        pe.BeginPage(ppi, pi, (currentSheet == totalSheets), null);
//        pe.BeginGenerateGraphics(null);
//        ppd.SheetProgressPos = 50;
//        pe.EndGenerateGraphics(null);

//        // Finish the sheet
//        pe.EndPage(null);
//        ppd.SheetProgressPos = 100;
//        ppd.OnEndSheet();
//    }

//    private static string setClosestMediaName(PlotSettingsValidator psv, PlotSettings ps,
//        double pageWidth, double pageHeight, bool matchPrintableArea)
//    {
//        //get all of the media listed for plotter
//        StringCollection mediaList = psv.GetCanonicalMediaNameList(ps);
//        double smallestOffest = 0.0;
//        string selectedMedia = string.Empty;
//        PlotRotation selectedRot = PlotRotation.Degrees000;

//        foreach (string media in mediaList)
//        {
//            psv.SetCanonicalMediaName(ps, media);

//            double mediaWidth = ps.PlotPaperSize.X;
//            double mediaHeight = ps.PlotPaperSize.Y;

//            if (matchPrintableArea)
//            {
//                mediaWidth -= (ps.PlotPaperMargins.MinPoint.X + ps.PlotPaperMargins.MaxPoint.X);
//                mediaHeight -= (ps.PlotPaperMargins.MinPoint.Y + ps.PlotPaperMargins.MaxPoint.Y);
//            }

//            PlotRotation rot = PlotRotation.Degrees090;

//            //check that we are not outside the media print area
//            if (mediaWidth < pageWidth || mediaHeight < pageHeight)
//            {
//                //Check if turning paper will work
//                if (mediaHeight < pageWidth || mediaWidth >= pageHeight)
//                {
//                    //still too small
//                    continue;
//                }
//                rot = PlotRotation.Degrees090;
//            }

//            double offset = Math.Abs(mediaWidth * mediaHeight - pageWidth * pageHeight);

//            if (selectedMedia == string.Empty || offset < smallestOffest)
//            {
//                selectedMedia = media;
//                smallestOffest = offset;
//                selectedRot = rot;

//                if (smallestOffest == 0)
//                    break;
//            }
//        }
//        psv.SetCanonicalMediaName(ps, selectedMedia);
//        psv.SetPlotRotation(ps, selectedRot);
//        return selectedMedia;
//    }
//}




//int numSheet = 0;

//// Initialize plot progress dialog
//PlotProgressDialog ppd = new PlotProgressDialog(false, layoutsToPlot.Count, true);

//using (ppd)
//{
//    // localize messages (if needed)
//    //ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Avanzamento stampa...");

//    ppd.OnBeginPlot();
//    ppd.IsVisible = true;

//    // init Engine (once)
//    PlotEngine pe = PlotFactory.CreatePublishEngine();
//    pe.BeginPlot(ppd, null);

//    ppd.StatusMsgString = "Plotting...";    // add name of document, if you need it

//    ppd.OnBeginSheet();
//    ppd.LowerSheetProgressRange = 0;
//    ppd.UpperSheetProgressRange = 100;
//    ppd.SheetProgressPos = 0;

//    // this is a multipage sample for multiple layouts, but if
//    // you have multiple windows of the current layout of even ModelSpace
//    // just make your changes...

//    foreach (ObjectId btrId in layoutsToPlot)
//    {
//        numSheet++;
//        ppd.SheetProgressPos = numSheet * 100 / layoutsToPlot.Count;

//        // get layout and its extents
//        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
//        Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

//        pMin = lo.Extents.MinPoint;
//        pMax = lo.Extents.MaxPoint;

//        // make it current
//        LayoutManager.Current.CurrentLayout = lo.LayoutName;

//        // Every layout may need different plot settings, media, scale, rotation, etc..
//        PlotSettings ps = new PlotSettings(lo.ModelType);
//        ps.CopyFrom(lo);

//        // ... and a new PlotInfo as well
//        PlotInfo pi = new PlotInfo();
//        pi.Layout = btr.LayoutId;

//        // The PlotSettingsValidator helps create a valid PlotSettings object
//        PlotSettingsValidator psv = PlotSettingsValidator.Current;

//        // you may need to assign a specific plot style (.ctb or .stb)
//        if (szStyle.Length > 0)
//        {
//            System.Collections.Specialized.StringCollection sc = psv.GetPlotStyleSheetList();
//            foreach (String szStyleFullName in sc)
//            {
//                if (System.IO.Path.GetFileName(szStyleFullName) == szStyle)
//                {
//                    psv.SetCurrentStyleSheet(ps, szStyleFullName);
//                    break;
//                }
//            }
//        }

//        ps.ScaleLineweights = false;

//        // here we want a centered plot of the layout extents,
//        // but you can set a Window plot as well...
//        psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
//        psv.SetUseStandardScale(ps, true);
//        psv.SetStdScaleType(ps, eScale);
//        psv.SetPlotCentered(ps, bCentrapagina);
//        //psv.SetPlotPaperUnits(ps, PlotPaperUnit.Millimeters);

//        //psv.SetPlotOrigin(ps, new Point2d(xMin, yMin));
//        //psv.SetPlotWindowArea(ps, new Extents2d(xMin, yMin, xMax, yMax));
//        //psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);

//        // here we need to choose the best orientation
//        bool bPaperIsLandscape;
//        bool bAreaIsLandscape;

//        if (lo.ModelType) bAreaIsLandscape = (dwg.Extmax.X - dwg.Extmin.X) > (dwg.Extmax.Y - dwg.Extmin.Y) ? true : false;
//        else bAreaIsLandscape = (lo.Extents.MaxPoint.X - lo.Extents.MinPoint.X) > (lo.Extents.MaxPoint.Y - lo.Extents.MinPoint.Y) ? true : false;

//        // choose the right media, if media not specified, choose the best one
//        szMedia = _PlotSetMedia(szPlotterName, szMedia, lo, ps, psv);

//        // set orientation based on paper and windows to be plot
//        bPaperIsLandscape = (ps.PlotPaperSize.X > ps.PlotPaperSize.Y) ? true : false;
//        if (bPaperIsLandscape != bAreaIsLandscape) psv.SetPlotRotation(ps, PlotRotation.Degrees270);

//        pi.OverrideSettings = ps;

//        // parameters validation required
//        PlotInfoValidator piv = new PlotInfoValidator();
//        piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
//        piv.Validate(pi);

//        if (numSheet == 1)      // first page? initialize document
//        {
//            pe.BeginDocument(pi, szDocName, null, 1, true, szOutFile);
//        }

//        // init current page
//        PlotPageInfo ppi = new PlotPageInfo();

//        pe.BeginPage(ppi, pi, (numSheet == layoutsToPlot.Count), null);
//        pe.BeginGenerateGraphics(null);

//        ppd.SheetProgressPos = 50;
//        pe.EndGenerateGraphics(null);

//        // End page
//        PreviewEndPlotInfo pepi = new PreviewEndPlotInfo();
//        pe.EndPage(pepi);
//        result = pepi.Status;
//        ppd.SheetProgressPos = 100;
//        ppd.OnEndSheet();

//        // keep system responsive
//        System.Windows.Forms.Application.DoEvents();
//    }

//    // End document
//    pe.EndDocument(null);

//    // End plot
//    ppd.PlotProgressPos = 100;
//    ppd.OnEndPlot();
//    pe.EndPlot(null);
//    pe.Dispose();

//}
    




//private static String _PlotSetMedia(String szPlotterName, String szMedia, Layout lo, PlotSettings ps, PlotSettingsValidator psv)
//{
//    if (szMedia.Length > 0)         // user specified a preferred media
//    {
//        // set current plotter device
//        psv.SetPlotConfigurationName(ps, szPlotterName, null);
//        psv.RefreshLists(ps);

//        // load canonical names, probably choosen previously from the same list
//        System.Collections.Specialized.StringCollection v = psv.GetCanonicalMediaNameList(ps);
//        String szCanonicalName = "", szTmp;

//        foreach (String s in v)
//        {
//            if (s == szMedia)
//            {
//                // canonical name found as specified by user
//                szCanonicalName = szMedia;
//                break;
//            }
//            else
//            {
//                // otherwise check for a descriptive name
//                szTmp = psv.GetLocaleMediaName(ps, s);
//                if (szTmp == szMedia)
//                {
//                    // ok, found!
//                    szCanonicalName = s;
//                    break;
//                }
//            }
//        }

//        psv.SetPlotConfigurationName(ps, szPlotterName, szCanonicalName);       // stampante e formato pagina
//        psv.SetPlotRotation(ps, PlotRotation.Degrees000);
//    }
//    else
//    {
//        // media not specified, choose the best
//        if (lo.ModelType) szMedia = FindBestFormat(ps, psv, szPlotterName, lo.Database.Extmin, lo.Database.Extmax);
//        else szMedia = FindBestFormat(ps, psv, szPlotterName, lo.Extents.MinPoint, lo.Extents.MaxPoint);
//    }

//    return szMedia;
//}


