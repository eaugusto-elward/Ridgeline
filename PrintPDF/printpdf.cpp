#include "pch.h"
#include "printpdf.h"
#include <iostream>
#include <string>



void SetPlotArea(double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY) {
    // Here you would set the plot area using the provided coordinates
    // For demonstration purposes, let's just print the received coordinates
    std::cout << "Lower Left: (" << lowerLeftX << ", " << lowerLeftY << ")" << std::endl;
    std::cout << "Upper Right: (" << upperRightX << ", " << upperRightY << ")" << std::endl;
}

void PlotAndPrintPDF() {
    // Here you would set up the plot settings and print the PDF
    // For demonstration purposes, let's just print a message
    std::cout << "Plotting and printing PDF..." << std::endl;
}

// Example of calling the C++ functions from C#
extern "C" {
    __declspec(dllexport) void SetPlotAreaWrapper(double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY) {
        SetPlotArea(lowerLeftX, lowerLeftY, upperRightX, upperRightY);
    }

    __declspec(dllexport) void PlotAndPrintPDFWrapper() {
        PlotAndPrintPDF();
    }
}
