using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using MSREG = Microsoft.Win32;
using System.Reflection;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Windows.Data;
using System.Security.AccessControl;



namespace Ridgeline
{
    internal class Commands
    {
        [CommandMethod("RegisterMyApp")]
        public void RegisterMyApp()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.MachineRegistryProductRootKey;
            string sAppName = "Ridgeline";

            RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
            RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Check to see if the "MyApp" key exists
            string[] subKeys = regAcadAppKey.GetSubKeyNames();
            foreach (string subKey in subKeys)
            {
                // If the application is already registered, exit
                if (subKey.Equals(sAppName))
                {
                    regAcadAppKey.Close();
                    return;
                }
            }

            // Get the location of this module
            string sAssemblyPath = Assembly.GetExecutingAssembly().Location;


            //            DESCRIPTION
            //Description of the .NET assembly and is optional.
            //            LOADCTRLS
            //Controls how and when the .NET assembly is loaded.

            //1 - Load application upon detection of proxy object
            //2 - Load the application at startup
            //4 - Load the application at start of a command
            //8 - Load the application at the request of a user or another application
            //16 - Do not load the application
            //32 - Load the application transparently
            //LOADER
            //Specifies which .NET assembly file to load.

            //MANAGED
            //Specifies the file that should be loaded is a.NET assembly or ObjectARX file. Set to 1 for .NET assembly files.


            // Register the application
            RegistryKey regAppAddInKey = regAcadAppKey.CreateSubKey(sAppName);
            regAppAddInKey.SetValue("DESCRIPTION", sAppName, MSREG.RegistryValueKind.String);
            regAppAddInKey.SetValue("LOADCTRLS", 2, MSREG.RegistryValueKind.DWord);
            regAppAddInKey.SetValue("LOADER", sAssemblyPath, MSREG.RegistryValueKind.String);
            regAppAddInKey.SetValue("MANAGED", 1, MSREG.RegistryValueKind.DWord);

            regAcadAppKey.Close();
        }

        [CommandMethod("UnregisterMyApp")]
        public void UnregisterMyApp()
        {
            // Get the AutoCAD Applications key
            string sProdKey = HostApplicationServices.Current.MachineRegistryProductRootKey;
            string sAppName = "MyApp";

            RegistryKey regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey);
            RegistryKey regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true);

            // Delete the key for the application
            regAcadAppKey.DeleteSubKeyTree(sAppName);
            regAcadAppKey.Close();
        }

    }
}
