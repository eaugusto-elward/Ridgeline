using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Specialized;
using System.IO;
using System.Runtime.InteropServices;
using WinForms = System.Windows.Forms;


namespace Ridgeline
{
    public class PrintPDF
    {
        static string plotterName = "Bluebeam PDF";
        static string mediaName = "ANSI_expand_A_(8.50_x_11.00_Inches)";
        static string plotStyle = "acad.ctb";


        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);

        [CommandMethod("PDFX")]
        static public void plotterscaletofit()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //ask user for print window
                Point3d first;
                PromptPointOptions ppo = new PromptPointOptions("\nSelect First corner of plot area: ");
                ppo.AllowNone = false;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status == PromptStatus.OK)
                { first = ppr.Value; }
                else
                    return;

                Point3d second;
                PromptCornerOptions pco = new PromptCornerOptions("\nSelect second corner of the plot area.", first);
                ppr = ed.GetCorner(pco);
                if (ppr.Status == PromptStatus.OK)
                { second = ppr.Value; }
                else
                    return;

                //convert from UCS to DCS
                Extents2d window = coordinates(first, second);

                //if the current view is paperspace then need to set up a viewport first
                if (LayoutManager.Current.CurrentLayout != "Model")
                { }

                //set up the plotter
                PlotInfo pi = plotSetUp(window, tr, db, ed);

                //call plotter engine to run
                plotEngine(pi, "Nameless", doc, ed);

                tr.Dispose();
            }
        }

        // A PlotEngine does the actual plotting
        // (can also create one for Preview)
        //***NOTE- always be sure that back ground plotting is off, in code and the users computer.
        static void plotEngine(PlotInfo pi, string name, Document doc, Editor ed)
        {
            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
            {
                PlotEngine pe = PlotFactory.CreatePublishEngine();
                using (pe)
                {
                    // Create a Progress Dialog to provide info or allow the user to cancel
                    PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true);
                    using (ppd)
                    {
                        
                        ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                        ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                        ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                        ppd.LowerPlotProgressRange = 0;
                        ppd.UpperPlotProgressRange = 100;
                        ppd.PlotProgressPos = 0;

                        // Let's start the plot, at last
                        ppd.OnBeginPlot();
                        ppd.IsVisible = true;
                        pe.BeginPlot(ppd, null);

                        // We'll be plotting a single document
                        //name should be file location + prompeted answer
                        string fileLoc = Path.GetDirectoryName(doc.Name);
                        pe.BeginDocument(pi, doc.Name, null, 1, false, fileLoc + @"\" + name);

                        // Which contains a single sheet
                        ppd.OnBeginSheet();
                        ppd.LowerSheetProgressRange = 0;
                        ppd.UpperSheetProgressRange = 100;
                        ppd.SheetProgressPos = 0;

                        PlotPageInfo ppi = new PlotPageInfo();
                        pe.BeginPage(ppi, pi, true, null);
                        
                        pe.BeginGenerateGraphics(null);
                        pe.EndGenerateGraphics(null);

                        // Finish the sheet
                        pe.EndPage(null);
                        ppd.SheetProgressPos = 100;
                        ppd.OnEndSheet();

                        // Finish the document
                        pe.EndDocument(null);

                        // And finish the plot
                        ppd.PlotProgressPos = 100;
                        ppd.OnEndPlot();
                        pe.EndPlot(null);
                    }
                }
            }

            else
            {
                ed.WriteMessage("\nAnother plot is in progress.");
            }
        }

        //acquire the extents of the frame and convert them from UCS to DCS, in case of view rotation
        static public Extents2d coordinates(Point3d firstInput, Point3d secondInput)
        {
            double lowerLeftX = firstInput.X;
            double lowerLeftY = firstInput.Y;
            double upperRightX = secondInput.X;
            double upperRightY = secondInput.Y;

            //sort through the values to be sure that the correct first and second are assigned
            //if (firstInput.X < secondInput.X)
            //{ minX = firstInput.X; maxX = secondInput.X; }
            //else
            //{ maxX = firstInput.X; minX = secondInput.X; }

            //if (firstInput.Y < secondInput.Y)
            //{ minY = firstInput.Y; maxY = secondInput.Y; }
            //else
            //{ maxY = firstInput.Y; minY = secondInput.Y; }


            //Point3d first = new Point3d(lowerLeftX, lowerLeftY, 0);
            //Point3d second = new Point3d(upperRightX, upperRightY, 0);
            ////converting numbers to something the system uses (DCS) instead of UCS
            //ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
            //double[] firres = new double[] { 0, 0, 0 };
            //double[] secres = new double[] { 0, 0, 0 };
            ////convert points
            //acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
            //acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
            //Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);
            Extents2d window = new Extents2d(lowerLeftX, lowerLeftY, upperRightX, upperRightY);

            return window;
        }

        static public PlotRotation orientation(Extents2d ext)
        {
            PlotRotation portrait = PlotRotation.Degrees180;
            PlotRotation landscape = PlotRotation.Degrees270;
            double width = ext.MinPoint.X - ext.MaxPoint.X;
            double height = ext.MinPoint.Y - ext.MaxPoint.Y;
            if (Math.Abs(width) > Math.Abs(height))
            { return landscape; }
            else
            { return portrait; }
        }

        //set up plotinfo
        static public PlotInfo plotSetUp(Extents2d window, Transaction tr, Database db, Editor ed)
        {
            using (tr)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                // We need a PlotInfo object linked to the layout
                PlotInfo pi = new PlotInfo();
                pi.Layout = btr.LayoutId;

                //current layout
                Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);
                

                // We need a PlotSettings object based on the layout settings which we then customize
                PlotSettings ps = new PlotSettings(lo.ModelType);
                ps.CopyFrom(lo);
                
                //The PlotSettingsValidator helps create a valid PlotSettings object
                PlotSettingsValidator psv = PlotSettingsValidator.Current;

                //set rotation
                psv.SetPlotRotation(ps, orientation(window));
                


                // We'll plot the window, centered, scaled, landscape rotation
                psv.SetPlotWindowArea(ps, window);
                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);//breaks here on some drawings                
                
                // Set the plot scale
                psv.SetUseStandardScale(ps, true);

                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);


                // Center the plot
                psv.SetPlotCentered(ps, true);//finding best location

                //get printerName from system settings
                //PrinterSettings settings = new PrinterSettings();
                //string defaultPrinterName = settings.PrinterName;

                psv.RefreshLists(ps);
                // Set Plot device & page size 
                psv.SetPlotConfigurationName(ps, "Bluebeam PDF", "Letter");


                //rebuilts plotter, plot style, and canonical media lists
                //(must be called before setting the plot style)
                psv.RefreshLists(ps);
                //psv.SetPlotRotation(ps, PlotRotation.Degrees090);
                //ps.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
                //ps.ShadePlotResLevel = ShadePlotResLevel.Normal;

                //plot options
                //ps.PrintLineweights = true;
                //ps.PlotTransparency = false;
                //ps.PlotPlotStyles = true;
                //ps.DrawViewportsFirst = true;
                //ps.CurrentStyleSheet

                // Use only on named layouts - Hide paperspace objects option
                // ps.PlotHidden = true;

                psv.SetPlotRotation(ps, PlotRotation.Degrees000);


                //plot table needs to be the custom heavy lineweight for the Uphol specs 
                psv.SetCurrentStyleSheet(ps, "acad.ctb");

                // We need to link the PlotInfo to the  PlotSettings and then validate it
                pi.OverrideSettings = ps;
                PlotInfoValidator piv = new PlotInfoValidator();
                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                piv.Validate(pi);

                return pi;
            }
        }
    }
}