#include "pch.h"
//#include "rxregsvc.h"
//#include "acutads.h"
//#include <tchar.h>
//
//// Simple acrxEntryPoint code. Normally initialization and cleanup
//// (such as registering and removing commands) should be done here.
////
//extern "C" AcRx::AppRetCode
//acrxEntryPoint(AcRx::AppMsgCode msg, void* appId)
//{
//    switch (msg) {
//    case AcRx::kInitAppMsg:
//        // Allow application to be unloaded
//        // Without this statement, AutoCAD will
//        // not allow the application to be unloaded
//        // except on AutoCAD exit.
//        //
//        acrxUnlockApplication(appId);
//        // Register application as MDI aware. 
//        // Without this statement, AutoCAD will
//        // switch to SDI mode when loading the
//        // application.
//        //
//        acrxRegisterAppMDIAware(appId);
//        acutPrintf(_T("\nExample Application Loaded")); // Use _T() macro for wide-character strings
//        break;
//    case AcRx::kUnloadAppMsg:
//        acutPrintf(_T("\nExample Application Unloaded")); // Use _T() macro for wide-character strings
//        break;
//    }
//    return AcRx::kRetOK;
//}
