//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.DatabaseServices;
//using Autodesk.AutoCAD.EditorInput;
//using Autodesk.AutoCAD.Geometry;
//using Autodesk.AutoCAD.PlottingServices;
//using Autodesk.AutoCAD.Runtime;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Runtime.InteropServices;


//namespace Ridgeline
//{
//    public class PrintPDF
//    {
//        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
//        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

//        string plotterName = "Bluebeam PDF"; // Plotter name for the PDF plotter
//        string mediaName = "Letter"; // Media name for the PDF plotter
//        string plotStyle = "acad.ctb"; // Plot style for the PDF plotter
//        static string filename = "default";   // String that captures the user input for the file name and location
//        string printVar = "";   // String that captures the user input for the layer visibility for printing
//        Point3d lowerLeft;   // Starting boundary point for the print window. For conversion from 3d to 2d
//        Point3d upperRight;    // Ending boundary point for the print window. For conversion from 3d to 2d
//        Extents2d window;   // The print window location
//        int rownum = 0; // Used for calculation how many times to translate the print window over the X-Axis
//        int colnum = 0; // Used for calculation how many times to translate the print window over the Y-Axis


//        [CommandMethod("PDFA")]
//        public void printPDFA()
//        {
//            // Active document objects
//            Document doc = Application.DocumentManager.MdiActiveDocument;
//            Database db = doc.Database;
//            Editor ed = doc.Editor;
//            PlotInfo pi = null;  // Putting this in the method scope to avoid the "unassigned variable" and avoid data clobbering
//            using (Transaction tr = db.TransactionManager.StartTransaction())
//            {






//                /* Main method calls */
//                getDrawingType(ed, ref printVar);   // Get user input for the printVar

//                getRowsCols(ed, ref rownum, ref colnum);    // Get user input for the number of rows and columns

//                getPrintBox(ed);    // Get user input for the initial print window location

//                //PromptForSavePath(ed, ref filename);   // Get user input for the file name and location

//                // Switch case for the printVar settings
//                switch (printVar)
//                {
//                    case "A":
//                        printVar = "A";
//                        break;
//                    case "C":
//                        printVar = "C";
//                        break;
//                    case "B":
//                        printVar = "B";
//                        break;
//                    case "N":
//                        printVar = "N";
//                        break;
//                    default:
//                        printVar = "B";
//                        break;
//                }


//                window = convertCoordinates(lowerLeft, upperRight);  // Returns an updated Extents2d object for every print window location
//                pi = plotSetup(window, tr, db, ed);  // Set plot settings with new window location
//                plotEngine(pi, ref filename, doc, ed, (rownum * colnum));      // Plot the window location

//                // Loop through each of the print window locations
//                for (int row = 1; row <= rownum; row++)
//                {
//                    for (int col = 1; col <= colnum; col++)
//                    {

                        
//                                                                                       //Translate window location by column

//                    }
//                    //Translate window location by row
//                }
//                ed.WriteMessage("\nHave a nice day!");
//            }

//        }

//        // Get user input for the printVar
//        public void getDrawingType(Editor ed, ref string printVar)
//        {
//            PromptStringOptions printVarOptions = new PromptStringOptions("\nAssembly(A), Cut(C) or Both(B), None(N)? Default (B): ")
//            {
//                AllowSpaces = false,
//                DefaultValue = "B",
//                UseDefaultValue = true
//            };
//            printVar = ed.GetString(printVarOptions).StringResult.ToUpper();

//            ed.WriteMessage("\nYou selected: " + printVar);
//        }

//        // Get user input for the number of rows and columns
//        public void getRowsCols(Editor ed, ref int rownum, ref int colnum)
//        {
//            rownum = PromptForInteger(ed, "\nHow many ROWS?: ");
//            colnum = PromptForInteger(ed, "\nHow many COLUMNS?: ");
//        }

//        // Get user input for the initial print window location
//        public void getPrintBox(Editor ed)
//        {

//            // Get user input for the print window location
//            PromptPointOptions ppo = new PromptPointOptions("\nSelect lower left corner of print window: ");
//            ppo.AllowNone = false;
//            PromptPointResult ppr = ed.GetPoint(ppo);
//            if (ppr.Status != PromptStatus.OK) return;  // Check for Error, return if error
//            lowerLeft = ppr.Value;

//            PromptCornerOptions pco = new PromptCornerOptions("\nSelect upper right corner of print window: ", lowerLeft);
//            ppr = ed.GetCorner(pco);
//            if (ppr.Status != PromptStatus.OK) return;  // Check for Error, return if error
//            upperRight = ppr.Value;

//        }

//        // Service function for 'getRowsCols'
//        private static int PromptForInteger(Editor ed, string message)
//        {

//            PromptIntegerOptions options = new PromptIntegerOptions(message)
//            {
//                AllowNone = false,
//                AllowZero = false,
//                AllowNegative = false
//            };
//            PromptIntegerResult result = ed.GetInteger(options);
//            return result.Value;
//        }

//        // Get user input for the file name and location
//        // public void PromptForSavePath(Editor ed, ref string filename)
//        //{

//        //    PromptSaveFileOptions options = new PromptSaveFileOptions("\nSelect Save Directory & Name")
//        //    {
//        //        Filter = "PDF files (*.pdf)|*.pdf",
//        //        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // TODO: Set to the desired default directory to be the document working directory
//        //    };
//        //    PromptFileNameResult result = ed.GetFileNameForSave(options);
//        //    if (result.Status == PromptStatus.OK) // Check if the user clicked OK
//        //    {
//        //        filename = result.StringResult; // Assign the selected file name to filename
//        //    }
//        //    else
//        //    {
//        //        filename = "defaultFileName.pdf"; // Example default value assignment TODO: Evalute the need for this case checking
//        //    }
//        //}

//        // Convert 3d coordinates to 2d coordinates
//        public Extents2d convertCoordinates(Point3d lowerLeft, Point3d upperRight)
//        {
//            ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
//            double[] firres = new double[] { 0, 0, 0 };
//            double[] secres = new double[] { 0, 0, 0 };
//            //convert points
//            acedTrans(lowerLeft.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
//            acedTrans(upperRight.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
//            Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

//            return window;

//        }

//        // Set plot settings. This method is called for each print window location
//        public PlotInfo plotSetup(Extents2d window, Transaction tr, Database db, Editor ed)
//        {
//            using (tr)
//            {
//                // New block table record to manage layout copy
//                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
//                // New PlotInfo to be linked to the layout
//                PlotInfo pi = new PlotInfo();
//                pi.Layout = btr.LayoutId;

//                // Get the current layout and assign to object
//                Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

//                // Create plot settings object to pass to validator
//                PlotSettings ps = new PlotSettings(lo.ModelType);
//                ps.CopyFrom(lo);

//                // Create new plot settings validator object
//                PlotSettingsValidator psv = PlotSettingsValidator.Current;

//                // Set area rotation
//                psv.SetPlotRotation(ps, PlotRotation.Degrees000);

//                // Set window area, scale and offset
//                psv.SetPlotWindowArea(ps, window);
//                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
//                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
//                psv.SetPlotCentered(ps, true);

//                psv.RefreshLists(ps);   // Appears in other libraries before setting configuration name. Not sure if necessary

//                // Set plot device to PDF
//                psv.SetPlotConfigurationName(ps, plotterName, mediaName);

//                psv.RefreshLists(ps);   // Has to be set before setting the plot style

//                // Set plot style
//                psv.SetCurrentStyleSheet(ps, plotStyle);

//                pi.OverrideSettings = ps;
//                PlotInfoValidator piv = new PlotInfoValidator();
//                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
//                piv.Validate(pi);

//                return pi;
//            }
//        }

//        static void plotEngine(PlotInfo pi, ref string name, Document doc, Editor ed, int progressNum)
//        {
//            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
//            {
//                PlotEngine pe = PlotFactory.CreatePublishEngine();
//                using (pe)
//                {
//                    // Create a Progress Dialog to provide info or allow the user to cancel
//                    PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true);
//                    using (ppd)
//                    {
//                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Elward Plot Progress");
//                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
//                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
//                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Elward Block Progress");
//                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Block Progress");
//                        ppd.LowerPlotProgressRange = 0;
//                        ppd.UpperPlotProgressRange = progressNum;   // Default 100 
//                        ppd.PlotProgressPos = 0;

//                        // Let's start the plot, at last
//                        ppd.OnBeginPlot();
//                        ppd.IsVisible = true;
//                        pe.BeginPlot(ppd, null);

//                        // We'll be plotting a single document
//                        //name should be file location + prompeted answer
//                        string fileLoc = Path.GetDirectoryName(doc.Name);
//                        pe.BeginDocument(pi, doc.Name, null, 1, true, fileLoc + @"\" + filename);

//                        // Which contains a single sheet
//                        ppd.OnBeginSheet();
//                        ppd.LowerSheetProgressRange = 0;
//                        ppd.UpperSheetProgressRange = progressNum;
//                        ppd.SheetProgressPos = 0;

//                        PlotPageInfo ppi = new PlotPageInfo();
//                        pe.BeginPage(ppi, pi, true, null);
//                        pe.BeginGenerateGraphics(null);
//                        pe.EndGenerateGraphics(null);

//                        // Finish the sheet
//                        pe.EndPage(null);
//                        ppd.SheetProgressPos = progressNum;
//                        ppd.OnEndSheet();

//                        // Finish the document
//                        pe.EndDocument(null);

//                        // And finish the plot
//                        ppd.PlotProgressPos = progressNum;
//                        ppd.OnEndPlot();
//                        pe.EndPlot(null);
//                    }
//                }
//            }

//            else
//            {
//                ed.WriteMessage("\nAnother plot is in progress.");
//            }
//        }
//    }
//}