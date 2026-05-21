#include "stdafx.h"
#include "hnCalcuIRIMethodApiWrapper.h"
namespace hnCalcuIRIMethodApiWrapper
{
	hn::hnCalcuIRIMethodApi* calcuIRIApi = new hn::hnCalcuIRIMethodApi();
	void _stdcall setParam(const char* strDaqPath, int nImuHz)
	{
		calcuIRIApi->setParam(strDaqPath,nImuHz);
	}
	void _stdcall setSaveResamplePath(const char* strSaveResamplePath)
	{
		calcuIRIApi->setSaveResamplePath(strSaveResamplePath);
	}
	void _stdcall setSaveIRIPath10(const char* strSavePath, bool bSaveIRI10)
	{

		calcuIRIApi->setSaveIRIPath10(strSavePath,bSaveIRI10);
	}
	void _stdcall setSaveIRIPath100(const char* strSavePath, bool bSaveIRI100)
	{

		calcuIRIApi->setSaveIRIPath100(strSavePath,bSaveIRI100);
	}
	void _stdcall setSaveIRIPath1000(const char* strSavePath, bool bSaveIRI1000)
	{

		calcuIRIApi->setSaveIRIPath1000(strSavePath,bSaveIRI1000);
	}
	void _stdcall setIsOnRight(int onRight)
	{
		calcuIRIApi->setIsOnRight(onRight);
	}
	bool _stdcall calcuIRI()
	{
		return calcuIRIApi->calcuIRI();
	}

	void _stdcall setCallBack( void(*func)(float, const char*) )
	{
		calcuIRIApi->loadCallback = func;
	}

	bool  _stdcall calcuCelerator(const char * savePath)
	{
		return calcuIRIApi->calcuCelerator( savePath);
	}
	

#pragma region gps 朱旭波
	//将经纬度通过高斯投影成为平面坐标 (弧度);
	void Gauss_Btox(double a, double f, double h, double Bm, double B, double L, double L0, double& m_x, double& m_y)
	{
		double b, e1, e2, l, t, m0, n2, N;
		double A0, A2, A4, A6, A8, X0;
		double da, dN, dB, dH;

		b = a * (1 - 1 / f);
		e2 = (a * a - b * b) / (a * a);

		// 更换计算方法;
		da = h * (1.0) / sqrt(1.0 - e2);

		a += da;
		double W = 1 - e2 * sin(B) * sin(B);
		double M = a * (1 - e2) / sqrt(pow(1 - e2 * sin(B) * sin(B), 3));
		double Ni = a / W;
		dB = (e2 * sin(B) * cos(B)) / (M + h) / W;

		dN = h * sqrt((1 - e2 * pow(sin(B), 2)) / (1 - e2));
		N = a / sqrt(1 - e2 * pow(sin(B), 2));

		//N+=dN;
		B += dB;


		A0 = 1 + 3 / 4.0 * e2 + 45 / 64.0 * pow(e2, 2) + 350 / 512.0 * pow(e2, 3) + 11025 / 16384.0 * pow(e2, 4);
		A2 = (-1 / 2.0) * (3 / 4.0 * e2 + 60 / 64.0 * pow(e2, 2) + 525 / 512.0 * pow(e2, 3) + 17640 / 16384.0 * pow(e2, 4));
		A4 = 1 / 4.0 * (15 / 64.0 * pow(e2, 2) + 210 / 512.0 * pow(e2, 3) + 8820 / 16384.0 * pow(e2, 4));
		A6 = (-1 / 6.0) * (35 / 512.0 * pow(e2, 3) + 2520 / 16384.0 * pow(e2, 4));
		A8 = (1 / 8.0) * 315 / 16384.0 * pow(e2, 4);
		l = L - L0;
		t = tan(B);
		m0 = l * cos(B);

		n2 = e2 / (1 - e2) * pow(cos(B), 2);

		X0 = a * (1 - e2) * (A0 * B + A2 * sin(2 * B) + A4 * sin(4 * B) + A6 * sin(6 * B) + A8 * sin(8 * B));

		double M0 = a * (1 - e2) * (A0 * Bm + A2 * sin(2 * Bm) + A4 * sin(4 * Bm) + A6 * sin(6 * Bm) + A8 * sin(8 * Bm));

		m_x = X0 - M0 + 0.5 * N * t * pow(m0, 2.0) + 1.0 / 24.0 * (5 - pow(t, 2) + 9 * n2 + 4 * pow(n2, 2.0)) * N * t * pow(m0, 4.0)
			+ 1.0 * (61.0 - 58.0 * pow(t, 2.0) + pow(t, 4.0)) * N * t * pow(m0, 6.0) / 720.0;
		m_y = N * m0 + 1.0 / 6.0 * (1.0 - pow(t, 2.0) + n2) * N * pow(m0, 3.0)
			+ 1.0 / 120.0 * (5.0 - 18 * pow(t, 2.0) + pow(t, 4.0) + 14.0 * n2 - 58.0 * n2 * pow(t, 2.0)) * N * pow(m0, 5.0);
	}

	////将平面坐标通过高斯投影成为经纬度;
	void Gauss_xtoB(double a, double f, double h, double Bm, double x, double y, double L0, double& m_B, double& m_L)//弧度
	{
		double da, e2, ep2;
		e2 = 1 - pow((1 - (1.0 / f)), 2);
		ep2 = e2 / (1 - e2);
		double v1, p1, w1, e1, u1;
		double T1, C1, D;

		da = h * (1. - e2 * pow(sin(Bm), 2)) / sqrt(1. - e2);
		a += da;

		double M, M0, M1, M2, M3, M4;
		M1 = 1 - e2 / 4 - pow(e2, 2) * 3 / 64 - pow(e2, 3) * 5 / 256;
		M2 = e2 * 3 / 8 + pow(e2, 2) * 3 / 32 + pow(e2, 3) * 45 / 1024;
		M3 = pow(e2, 2) * 15 / 256 + pow(e2, 3) * 45 / 1024;
		M4 = pow(e2, 3) * 35 / 3072;

		M0 = a * (M1 * Bm - M2 * sin(2 * Bm) + M3 * sin(4 * Bm) - M4 * sin(6 * Bm));

		M = M0 + x;
		u1 = M / (a * (1 - e2 / 4 - 3 * pow(e2, 2) / 64 - 5 * pow(e2, 3) / 256));
		e1 = (1 - sqrt(1 - e2)) / (1 + sqrt(1 - e2));

		w1 = u1 + (e1 * 3 / 2 - pow(e1, 3) * 27 / 32) * sin(2 * u1);
		w1 += (pow(e1, 2) * 21 / 16 - pow(e1, 4) * 55 / 32) * sin(4 * u1);
		w1 += (pow(e1, 3) * 151 / 96) * sin(6 * u1);
		w1 += (pow(e1, 4) * 1097 / 512) * sin(8 * u1);

		p1 = a * (1 - e2) / pow(1 - e2 * pow(sin(w1), 2), 3.0 / 2.0);
		v1 = a / sqrt(1 - e2 * pow(sin(w1), 2));
		T1 = pow(tan(w1), 2);
		C1 = ep2 * pow(cos(w1), 2);
		D = y / v1;

		m_B = pow(D, 2) / 2 - (5 + 3 * T1 + 10 * C1 - 4 * pow(C1, 2) - 9 * ep2) * pow(D, 4) / 24
			+ (61 + 90 * T1 + 298 * C1 + 45 * pow(T1, 2) - 252 * ep2 - 3 * pow(C1, 2)) * pow(D, 6) / 720;
		m_B = w1 - v1 * tan(w1) / p1 * m_B;

		m_L = D - (1 + 2 * T1 + C1) * pow(D, 3) / 6
			+ (5 - 2 * C1 + 28 * T1 - 3 * pow(C1, 2) + 8 * ep2 + 24 * pow(T1, 2)) * pow(D, 5) / 120;
		m_L = L0 + m_L / cos(w1);
	}

	// dCurLon当前路面拍照时刻对应的经度，单位为度；
	// dCurLat当前路面拍照时刻对应的纬度，单位为度；
	// dLastLon前一张路面拍照时刻对应的经度(bInverse为false)，单位为度（当当前拍照时刻为第一张时，dLastLon可传入第二张的经纬度，同时标记bInverse 为true）；
	// dLastLat前一张路面拍照时刻对应的纬度，单位为度；
	// returnLon返回的当前路面中心位置的经度;
	// returnLat返回的当前路面中心位置的纬度;
	// 传入经纬度均为度点度值，dOffsetX为天线中心定义的车体坐标系（X轴指向前进方向车右方，应设置为0，
	//即硬件上天线中心与路面破损中心在前进方向一条直线上，减少计算量，Y轴指向前进方向车前方，Z轴指天）;
	__declspec(dllexport)  bool calcLatToPicCenter(double dCurLon, double dCurLat, double dCurHeight, double dLastLon, double dLastLat, double dLastHeight,
		double dXOffset,	double dOffsetY, double dOffsetZ, OUT double& returnLon, OUT double& returnLat, OUT double& returnHeight, bool bInverse)
	{
		// 转换高斯三度带投影;
		double dCurEast, dCurNorth;
		dCurEast = dCurNorth = 0.0;
		int iNo = (int)((dCurLon + 1.5) / 3.0);
		 
		double dL0 = iNo * 3.0 * PI_M / 180.0;
		Gauss_Btox(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dCurLat * PI_M / 180.0, dCurLon * PI_M / 180.0, dL0, dCurNorth, dCurEast);

		// 转平面坐标;
		double dLastEast, dLastNorth;
		dLastEast = dLastNorth = 0.0;
		Gauss_Btox(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dLastLat * PI_M / 180.0, dLastLon * PI_M / 180.0, dL0, dLastNorth, dLastEast);

		// 由两点构成三维线;
		double dNormalNorth, dNormalEast, dNormalH;
		if (bInverse)
		{
			// 计算第一张照片，dLastLon实际为第二张照片的经纬度信息;
			dNormalEast = dCurEast - dLastEast;
			dNormalNorth = dCurNorth - dLastNorth;
			dNormalH = dCurHeight - dLastHeight;
		}
		else
		{
			// 其他张照片，last为当前上一帧;
			dNormalEast = dLastEast - dCurEast;
			dNormalNorth = dLastNorth - dCurNorth;
			dNormalH = dLastHeight - dCurHeight;
		}

		// 向量单位化;
		double length = dNormalEast * dNormalEast + dNormalNorth * dNormalNorth + dNormalH * dNormalH;
		if (length == 0)
		{
			return false;
		}
		length = 1.0 / sqrt(length);
		dNormalEast = dNormalEast * length;
		dNormalNorth = dNormalNorth * length;
		dNormalH = dNormalH * length;

		// 得到图片中心点坐标;
		double dTmpEast, dTmpNorth, dTmpH;
		dTmpEast = dCurEast + dOffsetY * dNormalEast;
		dTmpNorth = dCurNorth + dOffsetY * dNormalNorth;
		dTmpH = dCurHeight + dOffsetY * dNormalH;
		dTmpH = dTmpH - dOffsetZ;

		// 当存在x方向偏移量时，需要将向量旋转;
		if (dXOffset != 0.0)
		{
			double dTmpEast_addX, dTmpNorth_addX, dTmpH_addX;
			dTmpEast_addX = dXOffset * dNormalEast;
			dTmpNorth_addX = dXOffset * dNormalNorth;
			dTmpH_addX = dTmpH + dXOffset * dNormalH;

			// 旋转90度;
			double dRotateAngle = 90.0 * PI_M / 180.0;
			double cs = cos(dRotateAngle);
			double sn = sin(dRotateAngle);
			double dtmpX, dtmpY, dtmpX2, dtmpY2;
			dtmpX2 = dTmpEast_addX * cs - dTmpNorth_addX * sn;
			dtmpY2 = dTmpEast_addX * sn + dTmpNorth_addX * cs;
			dtmpX2 += dTmpEast;
			dtmpY2 += dTmpNorth;

			dTmpEast = dtmpX2;
			dTmpNorth = dtmpY2;
			dTmpH = dTmpH_addX;


		}

		// 转换经纬度传出;
		 Gauss_xtoB(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dTmpNorth, dTmpEast, dL0, returnLat, returnLon);
		 returnLat = returnLat * 180.0 / PI_M;
		 returnLon = returnLon * 180.0 / PI_M;

		//平面坐标
		//returnLat = dTmpNorth;
		//returnLon = dTmpEast;

		returnHeight = dTmpH;
		return true;
	}

	//// 传入当前路面图片中心点的经纬高;
	//// 前一个（下一个）图片中心点的经纬高;
	//// 图片目标点像素x（x对应宽） y（y对应高）;
	//// 图片像素宽高;
	//// bInverse标记是否反向;
	//// 图片拍照实际宽高，宽为2米，高默认3.75米;
	bool _stdcall calcLatToPicPos(bool showGps,double dCurPicLon, double dCurPicLat, double dCurPicH,
		double dLastPicLon, double dLastPicLat, double dLastPicH,
		int picX, int picY, int picWidth, int picHeight,
		double& returnLon, double& returnLat, double& returnHeight,  int equip,
		bool bInverse, double dWidth, double dHeight)
	{
		// 转换高斯三度带投影;
		double dCurEast, dCurNorth;
		dCurEast = dCurNorth = 0.0;
		int iNo = (int)((dCurPicLon+1.5) / 3.0);



		//int iNo =  (int)(dCurPicLon / 3.0);
		double dL0 = iNo * 3.0 * PI_M / 180.0 ;
		Gauss_Btox(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dCurPicLat * PI_M / 180.0, dCurPicLon * PI_M / 180.0, dL0, dCurNorth, dCurEast);

		// 转平面坐标;
		double dLastEast, dLastNorth;
		dLastEast = dLastNorth = 0.0;
		Gauss_Btox(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dLastPicLat * PI_M / 180.0, dLastPicLon * PI_M / 180.0, dL0, dLastNorth, dLastEast);

		// 由两点构成三维线;
		double dNormalNorth, dNormalEast, dNormalH;
		if (bInverse)
		{
			// 计算第一张照片，dLastLon实际为第二张照片的经纬度信息;
			dNormalEast = dCurEast - dLastEast;
			dNormalNorth = dCurNorth - dLastNorth;
			dNormalH = dCurPicH - dLastPicH;
		}
		else
		{
			// 其他张照片，last为当前上一帧;
			dNormalEast = dLastEast - dCurEast;
			dNormalNorth = dLastNorth - dCurNorth;
			dNormalH = dLastPicH - dCurPicH;
		}

		// 向量单位化;
		double length = dNormalEast * dNormalEast + dNormalNorth * dNormalNorth + dNormalH * dNormalH;
		if (length == 0)
		{
			return false;
		}
		length = 1.0 / sqrt(length);
		dNormalEast = dNormalEast * length;
		dNormalNorth = dNormalNorth * length;
		dNormalH = dNormalH * length;

		// 根据像素比例换算;
		//20250725修改  前进方向左上方为像素零点 宽为x 高为y
		float dHeightOffset = 0;
			float dWidthOffset = 0;
		if(equip == 1)
		{
			//二三维设备
			  dHeightOffset = 1.0 * (picHeight / 2.0 -picY) * dHeight / picHeight;
			  dWidthOffset = 1.0 * (picX - picWidth / 2.0) * dWidth / picWidth;
		}
		else
		{
			dHeightOffset = 1.0 * (picHeight / 2.0 - picY) * dHeight / picHeight;
		    dWidthOffset = 1.0 * (picWidth / 2.0 - picX) * dWidth / picWidth;
		}
		
		
		

		// 得到width方向上的坐标;
		double dTmpEast, dTmpNorth, dTmpH;
		dTmpEast = dCurEast + dHeightOffset * dNormalEast;
		dTmpNorth = dCurNorth + dHeightOffset * dNormalNorth;
		dTmpH = dCurPicH + dHeightOffset * dNormalH;

		double dRotateAngle = 90.0;
		if (picX == picWidth / 2)
		{
			// 转换经纬度传出;
			Gauss_xtoB(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dTmpNorth, dTmpEast, dL0, returnLat, returnLon);
			returnLat = returnLat * 180.0 / PI_M;
			returnLon = returnLon * 180.0 / PI_M;
			returnHeight = dTmpH;
			return true;
		} 
		else
		{
			if (picX < picWidth / 2)
				dRotateAngle = 90.0;
			else
			{
				dRotateAngle = -90.0;

			}  
		}

		// 旋转90度;
		double dTmpEast1, dTmpNorth1, dTmpH1;
		dTmpEast1 = dTmpEast + fabs(dWidthOffset) * dNormalEast;
		dTmpNorth1 = dTmpNorth + fabs(dWidthOffset) * dNormalNorth;
		dTmpH1 = dTmpH + fabs(dWidthOffset) * dNormalH;

		dRotateAngle = dRotateAngle * PI_M / 180.0;
		double cs = cos(dRotateAngle);
		double sn = sin(dRotateAngle);

		dTmpEast1 -= dTmpEast;
		dTmpNorth1 -= dTmpNorth;

		double dRotatedE = dTmpEast1 * cs - dTmpNorth1 * sn;
		double dRotatedN = dTmpEast1 * sn + dTmpNorth1 * cs;
		dRotatedE += dTmpEast;
		dRotatedN += dTmpNorth;

		// 转换经纬度传出;
	
		
		if (showGps)
		{
			Gauss_xtoB(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dRotatedN, dRotatedE, dL0, returnLat, returnLon);
			returnLat = returnLat * 180.0 / PI_M;
			returnLon = returnLon * 180.0 / PI_M;
			returnHeight = dTmpH;
		}
		else
		{
			Gauss_xtoB(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dRotatedN, dRotatedE, dL0, returnLat, returnLon);
			returnLat = returnLat * 180.0 / PI_M;
			returnLon = returnLon * 180.0 / PI_M;
			returnHeight = dTmpH;
				//平面坐标
			returnLat = dRotatedN;
			returnLon = dRotatedE + 500000;
			returnHeight = dTmpH; 
		}


		
		
		//std::string strPath = "J:\\ROAD_DATA\\gps\\gps数据\\gps数据\\_动态数据_上行_02_湖北省_武汉市_汉阳区_20240306_154533\\检核点-001.txt";
		//FILE* ptrF = fopen(strPath.data(), "at+");
		//fprintf_s(ptrF, "%.10lf,%.10lf,%.3lf\n",
		//	returnLon, returnLat, returnHeight);
		//fprintf_s(ptrF, "%.4lf,%.4lf,%.4lf\n",
		//	dRotatedE, dRotatedN, returnHeight);
		//fclose(ptrF);
		 
		return true;
	}

	void convertBlh()
	{
		std::string strPath = "J:\\ROAD_DATA\\gps\\gps数据\\gps数据\\检核点.txt";
		std::string strPathSave = "J:\\ROAD_DATA\\gps\\gps数据\\gps数据\\检核点-save.csv";
		FILE* ptrFile = fopen(strPath.data(), "rt");
		FILE* ptrSave = fopen(strPathSave.data(), "wt+");
		char strline[1024];
		int iIdex, nL, nB;
		double dMinL, dMinB, dH;
		while (!feof(ptrFile))
		{
			// 1,30,27.3008029,114,23.9988389,12.113
			memset(strline, 0, 1024);
			fgets(strline, 1024, ptrFile);
			int ret = sscanf(strline, "%d,%d,%lf,%d,%lf,%lf\n",
				&iIdex, &nB, &dMinB, &nL, &dMinL, &dH);
			if (ret < 6)
			{
				continue;
			}

			double dL = nL + dMinL / 60.0;
			double dB = nB + dMinB / 60.0;

			// 转换高斯三度带投影;
			double dCurEast, dCurNorth;
			dCurEast = dCurNorth = 0.0;
			int iNo = (int)((dL + 1.5) / 3.0);
			//int iNo = (int)(dL / 3.0);
			double dL0 = iNo * 3.0 * PI_M / 180.0;
			Gauss_Btox(EARTH_WGS84_EA, EARTH_WGS84_EF, 0.0, 0.0, dB * PI_M / 180.0, dL * PI_M / 180.0, dL0, dCurNorth, dCurEast);

			fprintf_s(ptrSave, "%d,%.10lf,%.10lf,%.3lf,%.3lf,%.3lf,%.3lf\n",
				iIdex, dL, dB, dH, dCurEast, dCurNorth, dH);
		}
		fclose(ptrFile);
		fclose(ptrSave);
	}

	void calcCenterPos()
	{
		double dCurLon, dCurLat, dCurH, dLastLon, dLastLat, dLastH;
		dLastLon = 114 + 23.9986248 / 60.0;
		dLastLat = 30 + 27.3015866 / 60.0;
		dLastH = 13.943;

		dCurLon = 114 + 23.9997468 / 60.0;
		dCurLat = 30 + 27.3008938 / 60.0;
		dCurH = 13.932;

		double dYOffsetLength = 2.61;
		double dZOffsetLength = 1.8;
		double dCenterLon, dCenterLat, dCenterH;
		calcLatToPicCenter(dCurLon, dCurLat, dCurH, dLastLon, dLastLat, dLastH, 0,dYOffsetLength, dZOffsetLength, dCenterLon, dCenterLat, dCenterH, false);

		std::string strPath = "J:\\ROAD_DATA\\gps\\gps数据\\gps数据\\检核点-centerPt.txt";
		FILE* ptrF = fopen(strPath.data(), "wt+");
		fprintf_s(ptrF, "%.10lf,%.10lf,%.3lf\n",
			dCenterLon, dCenterLat, dCenterH);

		dLastLon = 114 + 23.9997468 / 60.0;
		dLastLat = 30 + 27.3008938 / 60.0;
		dLastH = 13.932;

		dCurLon = 114 + 24.0006456 / 60.0;
		dCurLat = 30 + 27.3003212 / 60.0;
		dCurH = 13.917;

		calcLatToPicCenter(dCurLon, dCurLat, dCurH, dLastLon, dLastLat, dLastH,0, dYOffsetLength, dZOffsetLength, dCenterLon, dCenterLat, dCenterH, false);
		fprintf_s(ptrF, "%.10lf,%.10lf,%.3lf\n",
			dCenterLon, dCenterLat, dCenterH);

		fclose(ptrF);
	}

	/*void testPicPos()
	{
		double dCurCenterL, dCurCenterB, dCurH;
		double dLastCenterL, dLastCenterB, dLastH;

		dLastCenterL = 114.3999736509;
		dLastCenterB = 30.4550285864;
		dLastH = 12.145;

		dCurCenterL = 114.3999888680;
		dCurCenterB = 30.4550193259;
		dCurH = 12.139;

		double dRltL, dRltB, dRh;

		int picX, picY;
		picX = 1087;
		picY = 428;

		int picWidth = 4096;
		int picHeight = 2168;
		calcLatToPicPos(dCurCenterL, dCurCenterB, dCurH, dLastCenterL, dLastCenterB, dLastH, picX, picY, picWidth, picHeight, dRltL, dRltB, dRh);

		picX = 2679;
		picY = 327;
		calcLatToPicPos(dCurCenterL, dCurCenterB, dCurH, dLastCenterL, dLastCenterB, dLastH, picX, picY, picWidth, picHeight, dRltL, dRltB, dRh);
	}*/
#pragma endregion

}