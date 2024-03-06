using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Internal.Calculator;
using System.IO;
//using System.Windows.Forms;


namespace Ridgeline
{
    internal class Class3
    {



        // Creates a new page setup or edits the page set if it exists
        [CommandMethod("CreateOrEditPageSetup")]
        public static void CreateOrEditPageSetup()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {

                DBDictionary plSets = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                        OpenMode.ForRead) as DBDictionary;
                DBDictionary vStyles = acTrans.GetObject(acCurDb.VisualStyleDictionaryId,
                                                         OpenMode.ForRead) as DBDictionary;

                PlotSettings acPlSet = default(PlotSettings);
                bool createNew = false;

                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Get the current layout and output its name in the Command Line window
                Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                    OpenMode.ForRead) as Layout;

                // Check to see if the page setup exists
                if (plSets.Contains("MyPageSetup") == false)
                {
                    createNew = true;

                    // Create a new PlotSettings object: 
                    //    True - model space, False - named layout
                    acPlSet = new PlotSettings(acLayout.ModelType);
                    acPlSet.CopyFrom(acLayout);

                    acPlSet.PlotSettingsName = "MyPageSetup";
                    acPlSet.AddToPlotSettingsDictionary(acCurDb);
                    acTrans.AddNewlyCreatedDBObject(acPlSet, true);
                }
                else
                {
                    acPlSet = plSets.GetAt("MyPageSetup").GetObject(OpenMode.ForWrite) as PlotSettings;
                }

                // Update the PlotSettings object
                try
                {
                    PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                    // Set the Plotter and page size
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWF6 ePlot.pc3", "ANSI_B_(17.00_x_11.00_Inches)");

                    // Set to plot to the current display
                    if (acLayout.ModelType == false)
                    {
                        acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout);
                    }
                    else
                    {
                        acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);

                        acPlSetVdr.SetPlotCentered(acPlSet, true);
                    }

                    // Use SetPlotWindowArea with PlotType.Window
                    //acPlSetVdr.SetPlotWindowArea(plSet,
                    //                             new Extents2d(New Point2d(0.0, 0.0),
                    //                             new Point2d(9.0, 12.0)));

                    // Use SetPlotViewName with PlotType.View
                    //acPlSetVdr.SetPlotViewName(plSet, "MyView");

                    // Set the plot offset
                    acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(0, 0));

                    // Set the plot scale
                    acPlSetVdr.SetUseStandardScale(acPlSet, true);
                    acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                    acPlSetVdr.SetPlotPaperUnits(acPlSet, PlotPaperUnit.Inches);
                    acPlSet.ScaleLineweights = true;

                    // Specify if plot styles should be displayed on the layout
                    acPlSet.ShowPlotStyles = true;

                    // Rebuild plotter, plot style, and canonical media lists 
                    // (must be called before setting the plot style)
                    acPlSetVdr.RefreshLists(acPlSet);

                    // Specify the shaded viewport options
                    acPlSet.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;

                    acPlSet.ShadePlotResLevel = ShadePlotResLevel.Normal;

                    // Specify the plot options
                    acPlSet.PrintLineweights = true;
                    acPlSet.PlotTransparency = false;
                    acPlSet.PlotPlotStyles = true;
                    acPlSet.DrawViewportsFirst = true;

                    // Use only on named layouts - Hide paperspace objects option
                    // plSet.PlotHidden = true;

                    // Specify the plot orientation
                    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees000);

                    // Set the plot style
                    if (acCurDb.PlotStyleMode == true)
                    {
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.ctb");
                    }
                    else
                    {
                        acPlSetVdr.SetCurrentStyleSheet(acPlSet, "acad.stb");
                    }

                    // Zoom to show the whole paper
                    acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, true);
                }
                catch (Autodesk.AutoCAD.Runtime.Exception es)
                {
                    System.Windows.Forms.MessageBox.Show(es.Message);
                }

                // Save the changes made
                acTrans.Commit();

                if (createNew == true)
                {
                    acPlSet.Dispose();
                }
            }
        }
    }
}
