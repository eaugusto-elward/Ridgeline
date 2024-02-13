using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;


namespace Ridgeline
{

    // Helper class that keeps object abstractions of the active document
    public static class Active
    {
        // The active document stored in an object for use
        public static Document Document => Application.DocumentManager.MdiActiveDocument;

        // Editor object for the active document
        public static Editor Editor => Document.Editor;

        // Database object for the active document
        public static Database Database => Document.Database;

        // Helper method to start a transaction
        public static void UsingTransaction(Action<Transaction> action)
        {
            using (Transaction transaction = Active.Database.TransactionManager.StartTransaction())
            {
                action(transaction);

                transaction.Commit();
            }
        }

        // Helper method to update the drawing
        public static void UpdateDrawing()
        {
            Document.Editor.UpdateScreen();
        }
    }
}
