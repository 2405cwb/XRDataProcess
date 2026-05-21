#pragma  once
#include "hnCalcuIRIMethodApi.h"
#include <math.h>
#include <string>
//// 定义椭球参数,椭球长半轴; 
#define   EARTH_WGS84_EA	6378137
#define   EARTH_WGS84_EF	298.2572236
#define PI_M      3.141592653589793238462643383279
namespace  hnCalcuIRIMethodApiWrapper
{
	extern "C" _declspec(dllexport) void _stdcall setParam(const char* strDaqPath, int nImuHz);
	extern "C" _declspec(dllexport) void _stdcall setSaveResamplePath(const char* strSaveResamplePath);
	extern "C" _declspec(dllexport) void _stdcall setSaveIRIPath10(const char* strSavePath, bool bSaveIRI10);
	extern "C" _declspec(dllexport) void _stdcall setSaveIRIPath100(const char* strSavePath, bool bSaveIRI100);
	extern "C" _declspec(dllexport) void _stdcall setSaveIRIPath1000(const char* strSavePath, bool bSaveIRI1000);
	extern "C" _declspec(dllexport) void _stdcall setIsOnRight(int onRight);
	extern "C" _declspec(dllexport) bool _stdcall calcuIRI();
	extern "C" _declspec(dllexport) void _stdcall setCallBack(void(*func)(float, const char*));
	extern "C" _declspec(dllexport) bool _stdcall calcuCelerator(const char * savePath);

	extern "C" _declspec(dllexport) bool _stdcall calcLatToPicCenter(double dCurLon, double dCurLat, double dCurHeight, double dLastLon, double dLastLat, double dLastHeight,
		double dOffsetX,double dOffsetY, double dOffsetZ, OUT double& returnLon, OUT double& returnLat, OUT double& returnHeight, bool bInverse);

	extern "C" _declspec(dllexport) bool _stdcall calcLatToPicPos(bool showGps,double dCurPicLon, double dCurPicLat, double dCurPicH,
		double dLastPicLon, double dLastPicLat, double dLastPicH,
		int picX, int picY, int picWidth, int picHeight,
		double& returnLon, double& returnLat, double& returnHeight, int equip,
		bool bInverse , double dWidth, double dHeight);
}			   