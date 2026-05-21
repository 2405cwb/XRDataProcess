#pragma once
#include "hnOutDiseaseDxf.h"

typedef struct _HN_MILEAGE_PILE_
{
	_HN_MILEAGE_PILE_()
	{
		dMileage = 0.0;
		dRealMileage = 0.0;
	}

	// 相对里程
	double dMileage;

	// 真实里程
	double dRealMileage;
}hnMileagePile;

class DXFLIB_EXPORT hnOutputRoadDisease
{
public:
	hnOutputRoadDisease(void);
	~hnOutputRoadDisease(void);

public:
	bool outDisease(const char* strDisease, const char* strCenter, const char* strMileagePile,
		const char* outDir,void (*processCallback)(float, const char*) = NULL);

	// 设置里程信息
	void setMileageInfo(double dBegMil, int nMilDir, int nCenterType);
	
private:

	// 删除数据
    void clearData();

	// 读取病害
	bool readDisease(const char* strDisease, const char* strCenter, const char* strMileagePile);

	// 计算病害属性
	bool calDiseaseAttr();

	// 划分病害
	bool divideDiseaseGroup(vector<hnGridDiseaseInfo>& vecGrid);

	// 分割病害
	bool segGroupDisease(hnGridDiseaseInfo gridDisease);

	// 检查病害
	bool checkDisease(hnGridDiseaseInfo gridDisease);

	// 计算某点到中心线的投影距离以及相对里程
	bool getMileageAndDist(hn2dPt pt, double& dDist, double& dMileage);

	// 里程
	bool getRealMileage(double& dMileage);

	// 点坐标
	void getPoint(hn2dPt ptBeg, hn2dPt ptEnd, hn2dPt& oriPt);

	// 分割病害
	bool segDisease(double dBegMile, double dEndMile, hnRoadDiseaseInfo& roadDisease);

private:
	// 起始里程
	double m_dBegMileage;

	// 里程方向 0- 大里程， 1- 小里程
	int m_nMileageDir;

	// 道路病害信息
	vector<hnRoadDiseaseInfo> m_vecDiseaseInfo;

	// 中心线数据
	vector<hn2dPt> m_vecCenter;

	// 里程桩数据
	vector<hnMileagePile> m_vecMileagePile;

	// 最大最小里程
	double m_dMaxMile;
	double m_dMinMile;

	// 道路
	int m_nCenterType;
};

