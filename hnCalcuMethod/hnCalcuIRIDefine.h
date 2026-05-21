#ifndef HNCALCU_IRI_DEFINE_HEADER_H
#define HNCALCU_IRI_DEFINE_HEADER_H
#include <iosfwd>

// 将字符串转换为整型 包括 (unsigned) int/short/long;
template <class T>
T str2INT(const std::string& str)
{
	T num;
	std::istringstream iss(str);

	iss >> num;

	return num;
}

namespace hn
{
		//// 将字符串转换为浮点(float)
		//float str2float(const std::string& str)
		//{
		//	short int hg1120Result = 0;
		//	float fResult = 0.0f;

		//	// 定义初始变量，确定字符串长度;
		//	int i = (int)strlen(str.data());
		//	int j = 0;
		//	int counter = 0;
		//	char zc[2];
		//	unsigned int bytes[2];
		//	unsigned char strDest[128] = { 0 };
		//	for (j = 0; j < i; j += 2)
		//	{
		//		if (0 == j % 2)
		//		{
		//			zc[0] = str[j];
		//			zc[1] = str[j + 1];
		//			sscanf_s(zc, "%02x", &bytes[0]);
		//			strDest[counter] = bytes[0];
		//			counter++;
		//		}
		//	}

		//	fResult = 0.0f;
		//	memcpy(&hg1120Result, strDest, 2);
		//	fResult = hg1120Result;

		//	return fResult;
		//}

	struct HNTIME
	{
		HNTIME()
		{
			year = 2000;
			mon = 1;
			day = 1;
			hour = 0;
			min = 0;
			sec = 0;
			microSec = 0;
			milliSec = 0;
		}
		short year;
		short mon;
		short day;
		short hour;
		short min;
		short sec;
		short microSec;
		short milliSec;

	};

	struct DAQ_STRUCT_INFO
	{
		DAQ_STRUCT_INFO()
		{
			gpsWeek = 0;
			gpsSecond = 0.0;
			dim = 0;
			xVelocity = yVelocity = zVelocity = 0.0;
			xAccelerate = yAccelerate = zAccelerate = 0.0;
		}

		short gpsWeek;
		double gpsSecond;
		unsigned long dim;

		double xVelocity;
		double yVelocity;
		double zVelocity;
		double xAccelerate;
		double yAccelerate;//加速度
		double zAccelerate;

		HNTIME utc;
	};
	struct ACC_STRUCT_INFO
	{
		ACC_STRUCT_INFO()
		{
			mile = 0; 
			time = 0;
			yAccelerate = 0; 
			speed = 0;
			nowTime = 0;
		}
		//桩号 km
		double mile;
		//时长 s
		double time;
		//加速度
		double yAccelerate;//加速度 m/s2
		//速度 m/s
		double speed;
		  
	 double nowTime;
	};
	struct dmi_speed_type
	{
		double gpsSecond;
		double speed;
	};

	// 编码器单条数据结构 LONG型;
	struct dmi_lrec_type
	{
		short    sSync;  // set to 0xffee
		short    sWeek;  // set to -1 if not known
		double   dTime;  // GPS time of week,in seconds
		unsigned long lValue[1];  // values(counts) should be equal to sDim
	};

	// 编码器单条数据结构 DOUBLE型;
	struct dmi_drec_type
	{
		short    sSync;      // set to 0xffee
		short    sWeek;      // set to -1 if not known
		double   dTime;      // GPS time of week,in seconds
		double   dValue[1];  // values(double precision) should be equal to sDim
	};

	struct POSD_STRUCT_INFO 
	{
		POSD_STRUCT_INFO()
		{
			dGpsTime = 0.0;
			dL = dB = dH = 0.0;
			dYaw = dPitch = dRoll = 0.0;
			dDist = 0.0;
		}

		double dGpsTime;
		double dL;
		double dB;
		double dH;
		double dYaw;
		double dPitch;
		double dRoll;
		double dDist;
	};

	struct POSD_RESAMPLE250_INFO 
	{
		POSD_RESAMPLE250_INFO()
		{
			dim = 0;
			dHeight = 0;
			dGpsTime = 0.0;
		}

		HNTIME time;
		unsigned long dim;
		double dHeight;
		double dGpsTime;
	};
}



#endif // HNCALCU_IRI_METHOD_API_H
