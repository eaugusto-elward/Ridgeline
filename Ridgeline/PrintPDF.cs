using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;

namespace Ridgeline
{
    public class PrintPDF
    {
        string plotterName = "Bluebeam PDF"; // Plotter name for the PDF plotter
        string mediaName = "Letter"; // Media name for the PDF plotter
        Point2d lowerLeft;   // Starting boundary point for the print window. For conversion from 3d to 2d
        Point2d upperRight;    // Ending boundary point for the print window. For conversion from 3d to 2d

        private PromptFileNameResult filename;

        [CommandMethod("PDFA")]
        public void printPDFA()
        {
            //Algorithm Outline
            /*
             *Get Input from user for what layers to use
             *Get Input from user for what rows
             *Get Input from user for what columns
             *Get Input from user for the lower left and upper right corner
             *Set the file and name for savings
             *Switch case for the print var settings
             *
             *Cycle through each of the print window locations
             *Print with settings definited by the switch case
             */



            // Active document objects
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            string printVar = "";   // String that captures the user input for the layer visibility for printing
            string filename = "";   // String that captures the user input for the file name and location


            int rownum = 0; // Used for calculation how many times to translate the print window over the X-Axis
            int colnum = 0; // Used for calculation how many times to translate the print window over the Y-Axis

            double blockx = 0;  // Used for calculation of the print window width
            double blocky = 0;  // Used for calculation of the print window height

            Point3d pt1 = new Point3d();    // Starting boundary input point for the print window
            Point3d pt2 = new Point3d();    // Ending boundary input point for the print window


            /* Main method calls */
            getDrawingType(ed, ref printVar);   // Get user input for the printVar

            getRowsCols(ed, ref rownum, ref colnum);    // Get user input for the number of rows and columns

            getPrintBox(ed, ref blockx, ref blocky, ref pt1, ref pt2);    // Get user input for the initial print window location

            PromptForSavePath(ed, ref filename);   // Get user input for the file name and location

            // Switch case for the printVar settings
            switch (printVar)
            {
                case "A":
                    printVar = "A";
                    break;
                case "C":
                    printVar = "C";
                    break;
                case "B":
                    printVar = "B";
                    break;
                case "N":
                    printVar = "N";
                    break;
                default:
                    printVar = "B";
                    break;
            }

            // Loop through each of the print window locations
            for (int row = 1; row <= rownum; row++)
            {
                for (int col = 1; col <= colnum; col++)
                {
                    PlotCurrentView(doc, db, filename, row, col, pt1, blockx, blocky);
                }
            }
            ed.WriteMessage("\nHave a nice day!");

        }

        // Get user input for the printVar
        public void getDrawingType(Editor ed, ref string printVar)
        {
            PromptStringOptions printVarOptions = new PromptStringOptions("\nAssembly(A), Cut(C) or Both(B), None(N)? Default (B): ")
            {
                AllowSpaces = false,
                DefaultValue = "B",
                UseDefaultValue = true
            };
            printVar = ed.GetString(printVarOptions).StringResult.ToUpper();

            ed.WriteMessage("\nYou selected: " + printVar);
        }

        // Get user input for the number of rows and columns
        public void getRowsCols(Editor ed, ref int rownum, ref int colnum)
        {
            rownum = PromptForInteger(ed, "\nHow many ROWS(-)?: ");
            colnum = PromptForInteger(ed, "\nHow many COLUMNS(||)?: ");
        }

        // Get user input for the initial print window location
        public void getPrintBox(Editor ed, ref double blockx, ref double blocky, ref Point3d pt1, ref Point3d pt2)
        {
            pt1 = PromptForPoint(ed, "\nSelect the LOWER LEFT corner of the 1st cell: ");
            pt2 = PromptForPoint(ed, "\nSelect the UPPER RIGHT corner of the 1st cell: ");

            blockx = Math.Abs(pt1.X - pt2.X);
            blocky = Math.Abs(pt1.Y - pt2.Y);

            lowerLeft = new Point2d(pt1.X, pt1.Y);
            upperRight = new Point2d(pt2.X, pt2.Y);
        }

        // Service function for 'getRowsCols'
        private static int PromptForInteger(Editor ed, string message)
        {

            PromptIntegerOptions options = new PromptIntegerOptions(message)
            {
                AllowNone = false,
                AllowZero = false,
                AllowNegative = false
            };
            PromptIntegerResult result = ed.GetInteger(options);
            return result.Value;
        }

        // Service function for 'getPrintBox'
        private static Point3d PromptForPoint(Editor ed, string message)
        {

            PromptPointOptions options = new PromptPointOptions(message);
            PromptPointResult result = ed.GetPoint(options);
            return result.Value;
        }

        // Get user input for the file name and location
        private void PromptForSavePath(Editor ed, ref string filename)
        {

            PromptSaveFileOptions options = new PromptSaveFileOptions("\nSelect Save Directory & Name")
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // TODO: Set to the desired default directory to be the document working directory
            };
            PromptFileNameResult result = ed.GetFileNameForSave(options);
            if (result.Status == PromptStatus.OK) // Check if the user clicked OK
            {
                filename = result.StringResult; // Assign the selected file name to filename
            }
            else
            {
                filename = "defaultFileName.pdf"; // Example default value assignment TODO: Evalute the need for this case checking
            }
        }

        // Plot the current view
        private void PlotCurrentView(Document doc, Database db, string fileName, int row, int col, Point3d basePoint, double width, double height)
        {
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {

                LayoutManager lm = LayoutManager.Current;             
                string currentLayoutName = lm.CurrentLayout;
                ObjectId layoutId = lm.GetLayoutId(currentLayoutName);
                Layout layoutObj = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);

                PlotInfo pi = new PlotInfo();
                pi.Layout = layoutId; // Use the ObjectId of the layout directly

                PlotSettings ps = new PlotSettings(layoutObj.ModelType);
                ps.CopyFrom(layoutObj);

                PlotSettingsValidator psv = PlotSettingsValidator.Current;


                psv.SetPlotConfigurationName(ps, "Bluebeam PDF",
                                                    "Letter");


                // Set necessary plot settings here
                // This is an example and might need adjustment
                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                
                psv.SetUseStandardScale(ps, true);
                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(ps, true);

                // Set the plot area               
                psv.SetPlotWindowArea(ps, new Extents2d(lowerLeft, upperRight));

                pi.OverrideSettings = ps;

                PlotInfoValidator piv = new PlotInfoValidator();
                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;



                // Method to push all the plot settings to the PlotInfo object
                piv.Validate(pi);

                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                {
                    using (PlotEngine pe = PlotFactory.CreatePublishEngine())
                    {
                        using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
                        {
                            ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");

                            ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                            ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                            ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                            ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");

                            ppd.LowerPlotProgressRange = 0;
                            ppd.UpperPlotProgressRange = 100;
                            ppd.PlotProgressPos = 0;

                            ppd.OnBeginPlot();
                            ppd.IsVisible = true;

                            pe.BeginPlot(ppd, null);

                            pe.BeginDocument(pi, doc.Name, null, 1, true, fileName + "_" + row + "_" + col);

                            PlotPageInfo ppi = new PlotPageInfo();
                            pe.BeginPage(ppi, pi, true, null);
                            pe.BeginGenerateGraphics(null);
                            pe.EndGenerateGraphics(null);

                            pe.EndPage(null);
                            pe.EndDocument(null);

                            pe.EndPlot(null);
                            ppd.OnEndPlot();
                        }
                    }
                }

                tr.Commit();
            }
        }

        //TODO: Make the point selection style a bit different
        /*
          private Point3d? PromptForStartPoint(string message)
        {
            PromptPointOptions options = new PromptPointOptions(message);
            PromptPointResult result = Editor.GetPoint(options);

            if (result.Status == PromptStatus.OK)
            {
                return result.Value;
            }
            else
            {
                return null;
            }
        }

        private Point3d? PromptForCorner(Point3d start, string message)
        {
            PromptCornerOptions options = new PromptCornerOptions(message, start);
            PromptPointResult result = Editor.GetCorner(options);

            if (result.Status == PromptStatus.OK)
            {
                return result.Value;
            }
            else
            {
                return null;
            }
        }
        */

    }
}


