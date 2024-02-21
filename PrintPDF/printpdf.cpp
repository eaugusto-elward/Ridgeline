#include "pch.h"
#include "printpdf.h"
#include "acutads.h"
#include "acedads.h" // Include the header for acedAlert
#include "dbmain.h" // Include the header for AcDbEntity

#include <iostream>
#include <string>

// Function to print PDF and set plot area
extern "C" __declspec(dllexport) void PrintPDFConsole(double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY)
{
	acutPrintf(L"Hello, World!\n");
}
