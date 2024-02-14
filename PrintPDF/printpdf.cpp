#include "pch.h"
#include "printpdf.h"
#include <iostream>
#include <string>

using namespace std;

// Function to print PDF and set plot area
extern "C" __declspec(dllexport) void PrintPDFConsole(double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY)
{
    // Implement your printing to PDF logic here
    cout << "Printing PDF with plot area: (" << lowerLeftX << ", " << lowerLeftY << ") to (" << upperRightX << ", " << upperRightY << ")" << endl;
}
