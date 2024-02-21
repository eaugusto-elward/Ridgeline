using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Ridgeline
{
    public class PrintPDF
    {

        [CommandMethod("PDFA")]
        public void printPDFA()
        {
            //Algorithm
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
            
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            
            string printVar = "";

            int rownum = 0;
            int colnum = 0;

        }

        public void getDrawingType(Editor ed, string printVar)
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

        public void getRowsCols(Editor ed, int rownum, int colnum)
        {
            rownum = PromptForInteger("\nHow many ROWS(-)?: ");
            colnum = PromptForInteger("\nHow many COLUMNS(||)?: ");
        }

        public void getPrintBox()
        {
            Point3d pt1 = PromptForPoint("\nSelect the LOWER LEFT corner of the 1st cell: ");
            Point3d pt2 = PromptForPoint("\nSelect the UPPER RIGHT corner of the 1st cell: ");
        }
    }
}


