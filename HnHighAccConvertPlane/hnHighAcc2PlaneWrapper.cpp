#include "hnHighAcc2PlaneWrapper.h"

namespace  hnHighAcc2PlaneWrapper
{
	hnHighAcc2Plane* api = new hnHighAcc2Plane();
	void _stdcall initialParam(POS_CONVERT_INFO* paramInfo)
	{
		api->initialParam(paramInfo);
	}
	bool _stdcall convertBLHToProjection(double dL, double dB, double dH, double& dEast, double& dNorth,
		double& dHeight)
	{
		return api->convertBLHToProjection(dL, dB, dH, dEast, dNorth, dHeight);
	}
}