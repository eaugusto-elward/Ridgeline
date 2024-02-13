using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.PlottingServices;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.IO;

//[assembly: CommandClass(typeof(MyFirstCadPlugin.Class1))]
//namespace Ridgeline
//{
//    internal class Main
//    {

//        // Get application state and set plugin state
//        // TODO

//        // Populate ribbon with buttons
//        // TODO

//        // Begin runtime loop
//        // TODO

//        // Clean up and close
//    }
//}


namespace MyFirstCadPlugin
{
    public class Class1
    {
        [CommandMethod("hello")]
        public void HelloCommand()
        {
            // Here we connect to the active AutoCAD Document and Database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Here we create a Transaction in the current Database
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Access to the Model Blocktable for write
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForWrite) as BlockTable;
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                // Create a Circle with its center in the coordinates origin and radius = 50.
                Circle myCircle = new Circle();
                myCircle.Center = new Point3d(0, 0, 0);
                myCircle.Radius = 50;

                // Create text with "Hello!" in the middle of the circle
                DBText myText = new DBText();
                myText.SetDatabaseDefaults();
                myText.Height = 20;
                myText.TextString = "Hello!";
                myText.Justify = AttachmentPoint.MiddleCenter;


                // Append Circle and Text to the Blocktable record and Database
                acBlkTblRec.AppendEntity(myCircle);
                acTrans.AddNewlyCreatedDBObject(myCircle, true);
                acBlkTblRec.AppendEntity(myText);
                acTrans.AddNewlyCreatedDBObject(myText, true);

                // Finish Transaction
                acTrans.Commit();
            }
        }

        [CommandMethod("ListEntities")]
        public static void ListEntities()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for read
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForRead) as BlockTableRecord;

                int nCnt = 0;
                acDoc.Editor.WriteMessage("\nModel space objects: ");

                // Step through each object in Model space and
                // display the type of object found
                foreach (ObjectId acObjId in acBlkTblRec)
                {
                    acDoc.Editor.WriteMessage("\n" + acObjId.ObjectClass.DxfName);

                    nCnt = nCnt + 1;
                }

                // If no objects are found then display a message
                if (nCnt == 0)
                {
                    acDoc.Editor.WriteMessage("\n  No objects found");
                }

                // Dispose of the transaction
            }
        }

        [CommandMethod("CalculateDefinedArea")]
        public static void CalculateDefinedArea()
        {
            // Prompt the user for 5 points
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            PromptPointResult pPtRes;
            Point2dCollection colPt = new Point2dCollection();
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            // Prompt for the first point
            pPtOpts.Message = "\nSpecify first point: ";
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));

            // Exit if the user presses ESC or cancels the command
            if (pPtRes.Status == PromptStatus.Cancel) return;

            int nCounter = 1;

            while (nCounter <= 4)
            {
                // Prompt for the next points
                switch (nCounter)
                {
                    case 1:
                        pPtOpts.Message = "\nSpecify second point: ";
                        break;
                    case 2:
                        pPtOpts.Message = "\nSpecify third point: ";
                        break;
                    case 3:
                        pPtOpts.Message = "\nSpecify fourth point: ";
                        break;
                    case 4:
                        pPtOpts.Message = "\nSpecify fifth point: ";
                        break;
                }

                // Use the previous point as the base point
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = pPtRes.Value;

                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));

                if (pPtRes.Status == PromptStatus.Cancel) return;

                // Increment the counter
                nCounter = nCounter + 1;
            }

            // Create a polyline with 5 points
            using (Polyline acPoly = new Polyline())
            {
                acPoly.AddVertexAt(0, colPt[0], 0, 0, 0);
                acPoly.AddVertexAt(1, colPt[1], 0, 0, 0);
                acPoly.AddVertexAt(2, colPt[2], 0, 0, 0);
                acPoly.AddVertexAt(3, colPt[3], 0, 0, 0);
                acPoly.AddVertexAt(4, colPt[4], 0, 0, 0);

                // Close the polyline
                acPoly.Closed = true;

                // Query the area of the polyline
                Application.ShowAlertDialog("Area of polyline: " +
                                            acPoly.Area.ToString());

                // Dispose of the polyline
            }
        }

        // List all the layouts in the current drawing
        [CommandMethod("ListLayouts")]
        public void ListLayouts()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Get the layout dictionary of the current database
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                DBDictionary lays =
                    acTrans.GetObject(acCurDb.LayoutDictionaryId,
                        OpenMode.ForRead) as DBDictionary;

                acDoc.Editor.WriteMessage("\nLayouts:");

                // Step through and list each named layout and Model
                foreach (DBDictionaryEntry item in lays)
                {
                    acDoc.Editor.WriteMessage("\n  " + item.Key);
                }

                // Abort the changes to the database
                acTrans.Abort();
            }
        }

        // Create a new layout with the LayoutManager
        [CommandMethod("CreateLayout")]
        public void CreateLayout()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Get the layout and plot settings of the named pagesetup
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Create the new layout with default settings
                ObjectId objID = acLayoutMgr.CreateLayout("newLayout");

                // Open the layout
                Layout acLayout = acTrans.GetObject(objID,
                                                    OpenMode.ForRead) as Layout;

                // Set the layout current if it is not already
                if (acLayout.TabSelected == false)
                {
                    acLayoutMgr.CurrentLayout = acLayout.LayoutName;
                }

                // Output some information related to the layout object
                acDoc.Editor.WriteMessage("\nTab Order: " + acLayout.TabOrder +
                                          "\nTab Selected: " + acLayout.TabSelected +
                                          "\nBlock Table Record ID: " +
                                          acLayout.BlockTableRecordId.ToString());

                // Save the changes made
                acTrans.Commit();
            }
        }

        // Lists the available local media names for a specified plot configuration (PC3) file
        [CommandMethod("PlotterLocalMediaNameList")]
        public static void PlotterLocalMediaNameList()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            using (PlotSettings plSet = new PlotSettings(true))
            {
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                // Set the Plotter and page size
                acPlSetVdr.SetPlotConfigurationName(plSet, "DWF6 ePlot.pc3",
                                                    "ANSI_A_(8.50_x_11.00_Inches)");

                acDoc.Editor.WriteMessage("\nCanonical and Local media names: ");

                int cnt = 0;

                foreach (string mediaName in acPlSetVdr.GetCanonicalMediaNameList(plSet))
                {
                    // Output the names of the available media for the specified device
                    acDoc.Editor.WriteMessage("\n  " + mediaName + " | " +
                                              acPlSetVdr.GetLocaleMediaName(plSet, cnt));

                    cnt = cnt + 1;
                }
            }
        }

        // Lists the available plotters (plot configuration [PC3] files)
        [CommandMethod("PlotterList")]
        public static void PlotterList()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;

            acDoc.Editor.WriteMessage("\nPlot devices: ");

            foreach (string plotDevice in PlotSettingsValidator.Current.GetPlotDeviceList())
            {
                // Output the names of the available plotter devices
                acDoc.Editor.WriteMessage("\n  " + plotDevice);
            }
        }

        // Changes the plot settings for a layout directly
        [CommandMethod("ChangeLayoutPlotSettings")]
        public static void ChangeLayoutPlotSettings()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Get the current layout and output its name in the Command Line window
                Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                    OpenMode.ForRead) as Layout;

                // Output the name of the current layout and its device
                acDoc.Editor.WriteMessage("\nCurrent layout: " + acLayout.LayoutName);

                acDoc.Editor.WriteMessage("\nCurrent device name: " + acLayout.PlotConfigurationName);

                // Get a copy of the PlotSettings from the layout
                using (PlotSettings acPlSet = new PlotSettings(acLayout.ModelType))
                {
                    acPlSet.CopyFrom(acLayout);

                    // Update the PlotConfigurationName property of the PlotSettings object
                    PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3", "ANSI_B_(11.00_x_17.00_Inches)");

                    // Zoom to show the whole paper
                    acPlSetVdr.SetZoomToPaperOnUpdate(acPlSet, true);

                    // Update the layout
                    acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForWrite);
                    acLayout.CopyFrom(acPlSet);
                }

                // Output the name of the new device assigned to the layout
                acDoc.Editor.WriteMessage("\nNew device name: " + acLayout.PlotConfigurationName);

                // Save the new objects to the database
                acTrans.Commit();
            }

            // Update the display
            acDoc.Editor.Regen();
        }

        // Lists the available page setups
        [CommandMethod("ListPageSetup")]
        public static void ListPageSetup()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                DBDictionary plSettings = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                            OpenMode.ForRead) as DBDictionary;

                acDoc.Editor.WriteMessage("\nPage Setups: ");

                // List each named page setup
                foreach (DBDictionaryEntry item in plSettings)
                {
                    acDoc.Editor.WriteMessage("\n  " + item.Key);
                }

                // Abort the changes to the database
                acTrans.Abort();
            }
        }

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

        // Assigns a page setup to a layout
        [CommandMethod("AssignPageSetupToLayout")]
        public static void AssignPageSetupToLayout()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Get the current layout and output its name in the Command Line window
                Layout acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                    OpenMode.ForRead) as Layout;

                DBDictionary acPlSet = acTrans.GetObject(acCurDb.PlotSettingsDictionaryId,
                                                         OpenMode.ForRead) as DBDictionary;

                // Check to see if the page setup exists
                if (acPlSet.Contains("MyPageSetup") == true)
                {
                    PlotSettings plSet = acPlSet.GetAt("MyPageSetup").GetObject(OpenMode.ForRead) as PlotSettings;

                    // Update the layout
                    acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForWrite);
                    acLayout.CopyFrom(plSet);

                    // Save the new objects to the database
                    acTrans.Commit();
                }
                else
                {
                    // Ignore the changes made
                    acTrans.Abort();
                }
            }

            // Update the display
            acDoc.Editor.Regen();
        }
        
    }
}







