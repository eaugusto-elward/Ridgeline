using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;

namespace PlottingApplication
{
    public class PlottingCommands
    {
        [CommandMethod("mplot")]
        public static void MultiSheetPlot()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                PlotInfo pi = new PlotInfo();
                PlotInfoValidator piv = new PlotInfoValidator
                {
                    MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                };

                // Check if a plot is already in progress
                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                {
                    using (PlotEngine pe = PlotFactory.CreatePublishEngine())
                    {
                        using (PlotProgressDialog ppd = new PlotProgressDialog(false, 1, true))
                        {
                            ObjectIdCollection layoutsToPlot = new ObjectIdCollection();

                            foreach (ObjectId btrId in bt)
                            {
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                if (btr.IsLayout && btr.Name.ToUpper() != BlockTableRecord.ModelSpace.ToUpper())
                                {
                                    layoutsToPlot.Add(btrId);
                                }
                            }

                            int numSheet = 1;

                            foreach (ObjectId btrId in layoutsToPlot)
                            {
                                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                                Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead);

                                PlotSettings ps = new PlotSettings(lo.ModelType);
                                ps.CopyFrom(lo);

                                PlotSettingsValidator psv = PlotSettingsValidator.Current;

                                // Set plot settings
                                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                                psv.SetUseStandardScale(ps, true);
                                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                                psv.SetPlotCentered(ps, true);
                                psv.SetPlotConfigurationName(ps, "DWFx ePlot (XPS Compatible).pc3", "ANSI_A_(8.50_x_11.00_Inches)");

                                pi.Layout = btr.LayoutId;
                                LayoutManager.Current.CurrentLayout = lo.LayoutName;
                                pi.OverrideSettings = ps;
                                piv.Validate(pi);

                                if (numSheet == 1)
                                {
                                    // Setup progress dialog
                                    SetupProgressDialog(ppd, layoutsToPlot.Count);
                                    pe.BeginPlot(ppd, null);
                                    pe.BeginDocument(pi, doc.Name, null, 1, true, "c:\\test-multi-sheet");
                                }

                                // Plot each sheet
                                PlotSheet(ppd, pe, pi, doc, numSheet, layoutsToPlot.Count);
                                numSheet++;
                            }

                            // Finish the document and plot
                            pe.EndDocument(null);
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
        }

        private static void SetupProgressDialog(PlotProgressDialog ppd, int totalSheets)
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
        }

        private static void PlotSheet(PlotProgressDialog ppd, PlotEngine pe, PlotInfo pi, Document doc, int currentSheet, int totalSheets)
        {
            ppd.StatusMsgString = "Plotting " + doc.Name.Substring(doc.Name.LastIndexOf("\\") + 1) + " - sheet " + currentSheet.ToString() + " of " + totalSheets.ToString();

            ppd.OnBeginSheet();
            ppd.LowerSheetProgressRange = 0;
            ppd.UpperSheetProgressRange = 100;
            ppd.SheetProgressPos = 0;

            PlotPageInfo ppi = new PlotPageInfo();
            pe.BeginPage(ppi, pi, (currentSheet == totalSheets), null);
            pe.BeginGenerateGraphics(null);
            ppd.SheetProgressPos = 50;
            pe.EndGenerateGraphics(null);

            // Finish the sheet
            pe.EndPage(null);
            ppd.SheetProgressPos = 100;
            ppd.OnEndSheet();
        }
    }
}
