using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Ridgeline
{
    public class PrintPDF
    {
        string pageName = "Page"; // File name base value. Will be apprended with sequential numbers
        string pageCounter = "001"; // Counter for file name
        double plotLowerLeftX = 0.0; // Lower left corner of the plot area
        double plotLowerLeftY = 0.0; // Lower left corner of the plot area
        double plotUpperRightX = 0.0; // Upper right corner of the plot area
        double plotUpperRightY = 0.0; // Upper right corner of the plot area

        [DllImport("PrintPDF_DLL.dll")]
        private static extern void PrintPDFConsole(double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY);

        [CommandMethod("SelectCorners")]
        public void SelectCorners()
        {

            UpdateActiveDocument(); // Update the active document and editor

            // Prompt the user for the number of columns
            double colnum = PromptForReal("\nHow many COLUMNS(||)?: ");
            if (double.IsNaN(colnum))
            {
                return;
            }
            Editor.WriteMessage("\nColumns Entered: " + colnum.ToString());

            // Prompt the user for the number of rows
            double rownum = PromptForReal("\nHow many ROWS(-)?: ");
            if (double.IsNaN(rownum))
            {
                return;
            }
            Editor.WriteMessage("\nRows Entered: " + rownum.ToString());




            // Prompt the user to select the lower-left corner of the first cell
            Point3d? Point1 = PromptForStartPoint("\nSelect the LOWER LEFT corner of the 1st cell: ");
            if (Point1 == null)
            {
                return;
            }
            // Print the lower left corner point to console
            Editor.WriteMessage("\nLower Left Corner: " + Point1.Value.ToString());

            // Prompt the user to select the upper-right corner of the first cell
            Point3d? Point2 = PromptForCorner(Point1.Value, "\nSelect the UPPER RIGHT corner of the 1st cell: ");
            if (Point2 == null)
                return;
            // Print the upper right corner point to console
            Editor.WriteMessage("\nUpper Right Corner: " + Point2.Value.ToString());

            // Calculate block dimensions
            double blockx = Math.Abs(Point1.Value.X - Point2.Value.X);
            double blocky = Math.Abs(Point1.Value.Y - Point2.Value.Y);

            // Lower left corner values
            double lowerLeftX = Point1.Value.X;
            double lowerLeftY = Point1.Value.Y;

            // Set the plot area
            SetPlotArea(Point1.Value, Point2.Value);


            // Now we bring in the C++ code to print the PDF and set the plot area.
            // I would like to pass the plot area values to the C++ code, but I am not sure how to do that.
            //Then I want to set the plot area settings in the c++ code and print the PDF.

            PrintPDFConsole(plotLowerLeftX, plotLowerLeftY, plotUpperRightX, plotUpperRightY);
        }

        private double PromptForReal(string message)
        {
            PromptDoubleOptions options = new PromptDoubleOptions(message);
            PromptDoubleResult result = Editor.GetDouble(options);

            if (result.Status == PromptStatus.OK)
            {
                return result.Value;
            }
            else
            {
                return double.NaN;
            }
        }


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

        // Objects required for active document and editor
        // Keeping at bottom of class for readability
        private Document _document; // backing field for Document property
        public Document Document // property to access active document
        {
            get => _document;
            set => _document = value;
        }

        private Editor _editor; // backing field for Editor property
        public Editor Editor // property to access active editor
        {
            get => _editor;
            set => _editor = value;
        }

        // Update the active document and editor. To be called at the start of every [Commandmethod]
        private void UpdateActiveDocument()
        {
            Document = Application.DocumentManager.MdiActiveDocument;
            Editor = Document.Editor;
        }

        private void SetPlotArea(Point3d lowerLeft, Point3d upperRight)
        {
            plotLowerLeftX = lowerLeft.X;
            plotLowerLeftY = lowerLeft.Y;
            plotUpperRightX = upperRight.X;
            plotUpperRightY = upperRight.Y;
        }
    }
}