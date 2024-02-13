//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;


//using MSREG = Microsoft.Win32;
//using System.Reflection;

//using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.DatabaseServices;
//using Autodesk.AutoCAD.Colors;
//using Autodesk.AutoCAD.Windows.Data;
//using System.Security.AccessControl;
//using Autodesk.AutoCAD.EditorInput;
//using System.IO;
//using Autodesk.AutoCAD.PlottingServices;
//using Autodesk.AutoCAD.Geometry;



//namespace Ridgeline
//{
//    public class Commands
//    {
//        // Global values to quickly change plugin settings
//        private const string PlotStyleTable = "acad.ctb";
//        private double offsetx = 0.0;
//        private double offsety = 0.0;

//        [CommandMethod("PDFX")]

//        // Target .NET Framework version as used by AutoCAD
//        public void BatchPlotToPdf()
//        {
//            Document acDoc = Application.DocumentManager.MdiActiveDocument;
//            Editor ed = acDoc.Editor;

//            // Prompt for number of columns
//            PromptIntegerOptions columnOptions = new PromptIntegerOptions("\nEnter number of columns: ")
//            {
//                AllowNegative = false,
//                AllowZero = false
//            };
//            PromptIntegerResult columnResult = ed.GetInteger(columnOptions);
//            if (columnResult.Status != PromptStatus.OK) return;

//            int columns = columnResult.Value;

//            // Prompt for number of rows
//            PromptIntegerOptions rowOptions = new PromptIntegerOptions("\nEnter number of rows: ")
//            {
//                AllowNegative = false,
//                AllowZero = false
//            };
//            PromptIntegerResult rowResult = ed.GetInteger(rowOptions);
//            if (rowResult.Status != PromptStatus.OK) return;

//            int rows = rowResult.Value;

//            // Ensure the "PDFs" folder exists
//            string pdfFolderPath = EnsurePdfFolderExists(acDoc);

//            var rectangles = GetUserDefinedRectangles(ed);
//            if (rectangles == null || rectangles.Count == 0)
//            {
//                ed.WriteMessage("No rectangles defined or operation cancelled by the user.");
//                return; // Exit if no rectangles or user cancelled
//            }

//            foreach (var rectangle in rectangles)
//            {
//                PlotRectangleToPdf(rectangle, pdfFolderPath, acDoc);
//            }

//            // If AutoCAD or a third-party library supports combining PDFs, invoke that here
//        }

//        private string EnsurePdfFolderExists(Document acDoc)
//        {
//            string dwgFilePath = acDoc.Name;
//            string dwgDirectory = Path.GetDirectoryName(dwgFilePath);
//            string pdfFolderPath = Path.Combine(dwgDirectory, "PDFs");
//            if (!Directory.Exists(pdfFolderPath))
//            {
//                Directory.CreateDirectory(pdfFolderPath);
//            }
//            return pdfFolderPath;
//        }

//        private void PlotRectangleToPdf(Point3d[] rectangleCorners, string pdfFolderPath, Document acDoc)
//        {
//            if (rectangleCorners == null || rectangleCorners.Length != 2)
//            {
//                throw new ArgumentException("Rectangle corners must be an array of two Point3d objects.");
//            }

//            // Extract lower-left and upper-right points
//            Point3d lowerLeft = rectangleCorners[0];
//            Point3d upperRight = rectangleCorners[1];

//            LayoutManager lm = LayoutManager.Current;
//            string currentLayout = lm.CurrentLayout;
//            ObjectId layoutId = lm.GetLayoutId(currentLayout);

//            using (Transaction tr = acDoc.Database.TransactionManager.StartTransaction())
//            {
//                var plotInfo = new PlotInfo();
//                plotInfo.Layout = layoutId;

//                var plotSettings = new PlotSettings(true);
//                var plotSettingsValidator = PlotSettingsValidator.Current;

//                // Configure plot settings as before
//                plotSettingsValidator.SetPlotType(plotSettings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);

//                // Use standard scale 
//                plotSettingsValidator.SetUseStandardScale(plotSettings, true);

//                // Scale the plot to fit the rectangle created by the user
//                plotSettingsValidator.SetStdScaleType(plotSettings, StdScaleType.ScaleToFit);

//                // Set the plot units for custom layouts
//                plotSettingsValidator.SetPlotPaperUnits(plotSettings, PlotPaperUnit.Millimeters);

//                // Plot Rotation
//                plotSettingsValidator.SetPlotRotation(plotSettings, PlotRotation.Degrees000);

//                // Set the plot device to the desired printer
//                plotSettingsValidator.SetPlotConfigurationName(plotSettings, "DWG To PDF.pc3", "ISO_A4_(210.00_x_297.00_MM)");

//                // Centered flag to see if the plis considered centered
//                plotSettingsValidator.SetPlotCentered(plotSettings, true);

//                // Set the plot style to const global string
//                plotSettingsValidator.SetCurrentStyleSheet(plotSettings, PlotStyleTable);

//                // Set the plot window using the lower-left and upper-right points
//                plotSettingsValidator.SetPlotWindowArea(plotSettings, new Extents2d(lowerLeft.X, lowerLeft.Y, upperRight.X, upperRight.Y));
                



//                plotInfo.OverrideSettings = plotSettings;

//                using (var pe = PlotFactory.CreatePublishEngine())
//                {
//                    using (var plotProgress = new PlotProgressDialog(false, 1, true))
//                    {
//                        plotProgress.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");

//                        pe.BeginPlot(plotProgress, null);

//                        // Define the plot output
//                        var plotFileName = Path.Combine(pdfFolderPath, $"{Guid.NewGuid()}.pdf");
//                        pe.BeginDocument(plotInfo, acDoc.Name, null, 1, true, plotFileName);

//                        var pageSetup = new PlotPageInfo();
//                        pe.BeginPage(pageSetup, plotInfo, true, null);
//                        pe.BeginGenerateGraphics(null);
//                        pe.EndGenerateGraphics(null);

//                        // End the plot
//                        pe.EndPage(null);
//                        pe.EndDocument(null);
//                        pe.EndPlot(null);
//                    }
//                }

//                tr.Commit();
//            }
//        }


//        // Placeholder for the method to get rectangles based on user input
//        private List<Point3d[]> GetUserDefinedRectangles(Editor ed)
//        {
//            List<Point3d[]> rectangles = new List<Point3d[]>();

//            PromptPointResult pprLowerLeft = ed.GetPoint("\nSelect lower-left corner: ");
//            if (pprLowerLeft.Status != PromptStatus.OK) return null;

//            PromptPointOptions ppoUpperRight = new PromptPointOptions("\nSelect upper-right corner: ")
//            {
//                UseBasePoint = true,
//                BasePoint = pprLowerLeft.Value
//            };
//            PromptPointResult pprUpperRight = ed.GetPoint(ppoUpperRight);
//            if (pprUpperRight.Status != PromptStatus.OK) return null;

//            // Adjusted to include offsets
//            Point3d lowerLeft = new Point3d(pprLowerLeft.Value.X - offsetx, pprLowerLeft.Value.Y - offsety, pprLowerLeft.Value.Z);
//            Point3d upperRight = new Point3d(pprUpperRight.Value.X + offsetx, pprUpperRight.Value.Y + offsety, pprUpperRight.Value.Z);

//            rectangles.Add(new Point3d[] { lowerLeft, upperRight });

//            return rectangles;
//        }

//    }
//}
