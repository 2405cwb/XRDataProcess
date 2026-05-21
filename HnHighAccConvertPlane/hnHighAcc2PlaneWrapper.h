#pragma once
#include "hnHighAcc2Plane.h"
namespace hnHighAcc2PlaneWrapper
{
 	extern "C"  _declspec(dllexport) void _stdcall initialParam(POS_CONVERT_INFO * paramInfo); 
    extern "C" _declspec(dllexport) bool _stdcall convertBLHToProjection(double dL, double dB, double dH, double& dEast, double& dNorth, double& dHeight);

}