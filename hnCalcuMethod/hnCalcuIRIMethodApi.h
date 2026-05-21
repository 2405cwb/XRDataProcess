#ifndef HNCALCU_IRI_METHOD_API_H
#define HNCALCU_IRI_METHOD_API_H
#include "hnCalcuMethod_global.h"
#include <stdio.h>
#include <stdlib.h>
#include <string>
#include <vector>
#include "hnCalcuIRIDefine.h"

using namespace std;

namespace hn
{
	// 不同速度标定系数
	typedef struct _KB_PARAM
	{
		_KB_PARAM()
		{
			dSpeed = 0;
			dK = 1.0;
			dB = 0.0;
		}

		// 车轮速度
		double dSpeed;

		// 标定K值
		double dK;

		// 标定B值
		double dB;

	}hnKbParam;

	class HNCALCUIRIMETHOD_API hnCalcuIRIMethodApi
	{
	public:
		hnCalcuIRIMethodApi();
		~hnCalcuIRIMethodApi();

		// 设置进度条回调函数;
		void(*loadCallback)(float, const char*);

		// 设置参数;
		void setParam(const char* strDaqPath, int nImuHz, int nDmiHz, double dDmiWheel);

		// 设置参数
		void setParam(const char* strDaqPath, int nImuHz);

		//// 设置保存文件路径;
		//void setSaveResultPath(const char* strSaveDir,const char* strSavePreName);

		// 设置保存250mm抽样的数据文件位置;
		void setSaveResamplePath(const char* strSaveResamplePath);

		// 设置保存10m IRI文件路径以及是否保存该文件;
		void setSaveIRIPath10(const char* strSavePath, bool bSaveIRI10);

		// 设置保存100m IRI文件路径以及是否保存该文件;
		void setSaveIRIPath100(const char* strSavePath, bool bSaveIRI100);

		// 设置保存1000m IRI文件路径以及是否保存该文件;
		void setSaveIRIPath1000(const char* strSavePath, bool bSaveIRI1000);

		// 设置车轮编码器安装在左侧还是右侧，为0表示在左侧，为1表示在右侧，安装不同惯导坐标系指向不一样;
		void setIsOnRight(int onRight);

		// 计算处理;
		bool calcuIRI();

		// 测试代码，读取resample.txt文件，计算100m值对比结果;
		void testResample100(const char* strResamplePath);

		bool calcuCelerator(const char * savePath);
	private:
		// DAQ数据解析;
		bool loadDaqData(std::vector<DAQ_STRUCT_INFO>& vecDaqInfo,std::vector<dmi_speed_type>& vecDmiSpeedInfo);
		void outAccLPData(std::vector<DAQ_STRUCT_INFO>& vecDaq, std::vector<dmi_speed_type>& speed,int dist);
		void  outAccLPDataUser(const char* path, std::vector<DAQ_STRUCT_INFO>& vecDaq, std::vector<dmi_speed_type>& speed, int dist);
		// 解析I300原始惯导数据;
		bool parseHGI300Data(const std::string& str, DAQ_STRUCT_INFO& daqInfo);

		// utc时间转换gps时间;
		bool utc2gps(const HNTIME& stTime, short& nGPSWeek, double& dGPSSeconds, double dGPSSubUTC = 0.0);

		// gps周秒转换utc;
		bool gps2utc(short nGpsWeek, double dGpsSeconds, HNTIME& stTime, double dGPSSubUTC = 0.0);

		// 航位推算位置计算;
		bool calcuPosBySinsDr(std::vector<DAQ_STRUCT_INFO>& vecDaqInfo, std::vector<dmi_speed_type>& vecDmiSpeedInfo, std::vector<POSD_STRUCT_INFO>& vecMatPos);

		// 计算平整度值;
		void calcuIRIMethod(double dIntervel, double DeltLen,std::vector<POSD_RESAMPLE250_INFO>& vecResample250, std::vector<double>& vecIRIResult);

		// 重采样250mm数据;
		bool resampleIRI250(std::vector<POSD_STRUCT_INFO>& vecMatPos,std::vector<POSD_RESAMPLE250_INFO>& vecResample250);

		// 给定间隔数据计算IRI值;
		double CalculateIRI(double dIntervel, std::vector<POSD_RESAMPLE250_INFO>& vecIRIInfo,
			double DeltLen = 0.25);

		// 保存给定数据IRI结果值;
		void saveIRIs(const char* strSavePath, std::vector<double>& vecIRIResult);
		//保存给定数据速度结果值 
		void saveSpeed10m(const char* strSavePath, std::vector<double>& vecSpeed);
		// 保存250mm重采样信息数据结果值;
		void saveReSample250(const char* strSavePath,std::vector<POSD_RESAMPLE250_INFO>& vecResample250);

		// 添加一个速度计算变量;
		double addSpeedBlong(double iriResult, double speedval);

		// 单轴加速度计计算速度和高程信息-测试代码;
		void calcAccDisIntegrate(std::vector<DAQ_STRUCT_INFO>& vecDaqInfo, std::vector<POSD_STRUCT_INFO>& vecMatPos);

		// 多段线拟合;
		void polyfit(int n, double *y, int poly_n, double *a, double *tempx, double *tempy, double *sumxx, double *sumxy, double *ata);//拟合
		void gauss_solve(int n, double *A, double *x, double *b);
		void polyfitVal(int n, double *y, int poly_n, double *a);
		void Acc2Vel(int n, double *acc, double *vel, double t,int fitnum);
		void Vel2Dis(int n, double *vel, double *dis, double t,int fitnum);

		void Acc2Vel2(int ptcount,double* ptAcc,double*& ptVel,double dTimeScale,int fitnum);

		// 读取车速标定系数值
		bool readKbParam();

		// 写入车速标定系数值
		bool writeKbParam(const char* strPath);

		// 根据车速获取标定系数值
		bool getKbParam(double dSpeed, double& dK, double& dB);
		
		// 获取分割数目
		int getSplitCnt(const char* strData);
	private:
		// 记录原始DAQ文件路径;
		std::string m_strDaqPath;

		// 记录保存结果值文件夹路径;
		std::string m_strSaveResultDir;

		// 记录保存结果值文件名前缀信息;
		std::string m_strSavePreName;

		// 记录保存结果值文件-Resample250.txt文件路径;
		std::string m_strSaveResample250Path;

		std::string m_strSaveIRIPath10;
		bool m_bUseSave10;

		std::string m_strSaveIRIPath100;
		bool m_bUseSave100;

		std::string m_strSaveIRIPath1000;
		bool m_bUseSave1000;

		// 惯导频率;
		int m_nImuHz;

		// 编码器频率;
		int m_nDmiHz;

		// 车轮周长;
		double m_dDmiWheel;

		// 记录计算的GPS周;
		int m_nCurGpsWeek;

		double* m_SZU;
		double* m_PZU;
		double* m_ZSU;
		double* m_oldZSU;

		// 标记设备安装在左侧还是右侧;
		int m_isOnRight;

		// 设备标定系数
		double m_dK;
		double m_dB;

		// 不同车速标定系数值
		vector<hnKbParam> m_vecKbParam;

	private:
		double *m_tempx, *m_tempy;
		double *m_sumxx_0, *m_sumxy_0, *m_ata_0;
		double *m_sumxx_1, *m_sumxy_1, *m_ata_1;
		double *m_sumxx_2, *m_sumxy_2, *m_ata_2;
		double m_poly_a[6];
		std::vector<double> vecSpeed10m;
	};
}



#endif // HNCALCU_IRI_METHOD_API_H
