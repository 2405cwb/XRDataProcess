#pragma once
#include<iostream>
#include<vector>
#include "..\hnDxfIO\dl_dxf.h"
typedef struct
{
	/// <summary>
   /// ¾­¶È
   /// </summary>
	double DiseaseLon;
	/// <summary>
	/// Î³¶È
	/// </summary>
	double DiseaseLat;
	double DiseaseHeight;
}HighAccureacyGps;

typedef struct
{
	char name[64];
	int mile;
	HighAccureacyGps p0;
	HighAccureacyGps p1;
	HighAccureacyGps p2;
	HighAccureacyGps p3;
	HighAccureacyGps center;

}HighAccDisease;

extern "C" _declspec(dllexport)bool _stdcall HighAccOut(const char* path, int diseaseNum, HighAccDisease * disease);
bool initCAD(DL_Dxf* dxf, DL_WriterA* dw);