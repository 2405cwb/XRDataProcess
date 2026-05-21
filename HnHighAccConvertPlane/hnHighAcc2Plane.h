#pragma once
#include <stdio.h>
#include <stdlib.h>
#include <iostream>
#include "IHdPJTranslator.h"
#include "HnHighAccConvertPlane_global.h"
struct POS_CONVERT_INFO
{
	POS_CONVERT_INFO()
	{
		dCenterL = 114.0;
		nSphereType = 3;
		nProjectType = 0;
		dProjectHeight = 0.0;
		dEastAdd = 500000.0;
		dAverageLat = 0.0;
		dProjectScale = 1.0;
		nUseConvertModel = 0;
		dOffsetX = 0.0;
		dOffsetY = 0.0;
		dOffsetZ = 0.0;
		dRotateX = 0.0;
		dRotateY = 0.0;
		dRotateZ = 0.0;
		dK = 1.0;

		dFourX = 0.0;
		dFourY = 0.0;
		dFourR = 0.0;
		dFourK = 1.0;
	}

	// 1 设置中央经线,单位度点度;
	double dCenterL;

	// 2 设置椭球体，0表示北京54,1表示西安80,2表示WGS84,3表示CGCS2000;
	int nSphereType;

	//3 投影方法设置，0表示高斯三度带投影，1表示高斯6度带，2表示墨卡托投影，3表示横轴墨卡托投影;
	int nProjectType;

	//4 设置投影面高程;
	double dProjectHeight;

	// 5 东向加常数;
	double dEastAdd;

	//6 平均纬度;
	double dAverageLat;

	//7 尺度因子;
	double dProjectScale;

	//8 是否使用七参数或四参数转换，1表示使用四参数，2表示使用七参数，0表示不使用;
	int nUseConvertModel;

	// 四参数设置,不考虑高程;
	double dFourX;
	double dFourY;
	double dFourR;
	double dFourK;

	//9 七参数值设置;
	double dOffsetX;
	double dOffsetY;
	double dOffsetZ;
	double dRotateX;
	double dRotateY;
	double dRotateZ;
	double dK;
};

const double PI64 = 3.1415926535897932384626433832795028841971693993751;
 class HnHighAccConvertPlane_API hnHighAcc2Plane
{
	// 定义成员变量;
// 坐标投影参数设置对象;
	IHdPJTranslator* m_ptr_convert_translator = NULL;
	Spatial_Ref_t    m_src_param;				// 原始数据转换参数，默认即为WGS84;
	Spatial_Ref_t    m_dst_param;				// 目标数据转换参数;
public:
	void initialParam(POS_CONVERT_INFO* paramInfo);
	bool convertBLHToProjection(double dL, double dB, double dH, double& dEast, double& dNorth, double& dHeight);


};

