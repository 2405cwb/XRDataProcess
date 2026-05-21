#include "StdAfx.h"
#include "hnOutputRoadDisease.h"

// 解析字符
bool analysisStr(char* strSource, hnRoadDiseaseInfo& roadDisease)
{
	const char *d = ",";
	char *p;
	p = strtok(strSource, d);

	vector<double> vecValue;

	int nCnt = 0;
	while (p)
	{
		if (nCnt == 0)
		{
			// 病害类型
			roadDisease.nDiseaseType = (DISEASETYPE)(atoi(p));
		}
		else if (nCnt == 1)
		{
			// 病害程度
			roadDisease.nLevel = atoi(p);
		}
		else if (nCnt == 2)
		{
			// 病害长度
			roadDisease.dLength = atof(p);
		}
		else if (nCnt == 3)
		{
			// 病害宽度
			roadDisease.dWidth = atof(p);
		}
		else if (nCnt == 4)
		{
			// 病害面积
			roadDisease.dArea = atof(p);
		}
		else if (nCnt == 5)
		{
			// 病害深度
			roadDisease.dDepth = atof(p);
		}
		else if (nCnt == 6)
		{
			// 病害几何形态
			roadDisease.nGeometry = atoi(p);
		}
		else
		{
			// 病害顶点
			vecValue.push_back(atof(p));
		}
		
		p = strtok(NULL, d);

		++nCnt;
	}

	if (vecValue.size() <= 0)
	{
		return false;
	}

	if (vecValue.size() % 2 != 0)
	{
		return false;
	}

	int nPtCnt= vecValue.size() / 2;
	roadDisease.vecPt.resize(nPtCnt);
	
	for (int i = 0; i < nPtCnt; i++)
	{
		roadDisease.vecPt[i].dx = vecValue[i*2];
		roadDisease.vecPt[i].dy = vecValue[i*2 + 1];
	}

	return true;
}

// 排序
bool sortMileagePile(hnMileagePile begMilPile, hnMileagePile endMilPile)
{
	return begMilPile.dMileage < endMilPile.dMileage;
}

hnOutputRoadDisease::hnOutputRoadDisease(void):m_dBegMileage(0.0),m_nMileageDir(0),m_nCenterType(0)
{
	// 最大最小里程
	m_dMaxMile = -100000000.0;
	m_dMinMile = 100000000.0;
}


hnOutputRoadDisease::~hnOutputRoadDisease(void)
{
}

bool hnOutputRoadDisease::outDisease(const char* strDisease, const char* strCenter, 
	const char* strMileagePile, const char* outDir,void (*processCallback)(float, const char*)/* = NULL*/)
{
	// 导出dxf
	hnOutDiseaseDxf outDxf;
	outDxf.setCenterType(m_nCenterType);

	// 清空数据
	clearData();

	if (processCallback)
	{
		processCallback(0.2, "正在读取数据...");
	}

	// 读取数据
	if (!readDisease(strDisease, strCenter, strMileagePile))
	{
		return false;
	}

	if (m_vecCenter.size() <= 0 || m_vecDiseaseInfo.size() <= 0)
	{
		return false;
	}

	if (processCallback)
	{
		processCallback(0.4, "计算病害的属性信息...");
	}

	// 计算病害的属性信息
	calDiseaseAttr();

	vector<hnGridDiseaseInfo> vecGroupDisease;

	if (processCallback)
	{
		processCallback(0.6, "划分病害...");
	}

	// 划分病害
	divideDiseaseGroup(vecGroupDisease);

	if (vecGroupDisease.size() <= 0)
	{
		return false;
	}

	string strDir = outDir;
	strDir = strDir.substr(0, strDir.find_last_of("\\"));
	string strOutPath = "";

	if (processCallback)
	{
		processCallback(0.8, "输出病害...");
	}

	// 输出
	for (int i = 0; i < vecGroupDisease.size(); i++)
	{
		// 分割病害
		segGroupDisease(vecGroupDisease[i]);

		// 检查病害
		checkDisease(vecGroupDisease[i]);

		// 输出病害
		strOutPath = strDir + "\\" + vecGroupDisease[i].strName + ".dxf";

		outDxf.outDiseaseDxf(strOutPath.c_str() ,vecGroupDisease[i]);
	}

	if (processCallback)
	{
		processCallback(1.0, "输出完成...");
	}

	return true;
}

// 设置里程信息
void hnOutputRoadDisease::setMileageInfo(double dBegMil, int nMilDir, int nCenterType)
{
	m_dBegMileage = dBegMil;
	m_nMileageDir = nMilDir;
	m_nCenterType = nCenterType;
}

// 删除数据
void hnOutputRoadDisease::clearData()
{
	m_vecDiseaseInfo.clear();
	m_vecCenter.clear();
	m_vecMileagePile.clear();

	m_dMaxMile = -100000000.0;
	m_dMinMile = 100000000.0;
}

// 读取病害
bool hnOutputRoadDisease::readDisease(const char* strDisease, const char* strCenter, const char* strMileagePile)
{
	// 读取病害信息
	FILE* pf = fopen(strDisease, "r");
	if (!pf)
	{
		return false;
	}

	// 将文件指针放在初始位置
	fseek(pf, 0, SEEK_SET);
	char str[1024] = {0};
	int ncnt = 0;
	hnRoadDiseaseInfo roadDiseaseInfo;

	// 读取第一行数据
	while (!feof(pf))
	{
		memset(str, 0, 1024);
		fgets(str, 1024, pf);

		// 解析字符
		if (!analysisStr(str, roadDiseaseInfo))
		{
			continue;
		}

		if (roadDiseaseInfo.vecPt.size() < 2)
		{
			continue;
		}

		m_vecDiseaseInfo.push_back(roadDiseaseInfo);
	}

	fclose(pf);

	if (m_vecDiseaseInfo.size() <= 0)
	{
		return false;
	}

	// 读取中心点
	pf = fopen(strCenter, "r");
	const char *d = ",";
	char *p = NULL;
	int nCnt = 0;
	hn2dPt pt;
	if (pf)
	{
		fseek(pf, 0, SEEK_SET);
		while (!feof(pf))
		{
			memset(str, 0, 1024);
			fgets(str, 1024, pf);
			p = strtok(str, d);

			nCnt = 0;
			
			while (p)
			{
				if (nCnt == 0)
				{
					pt.dx = atof(p);
				}
				else
				{
					pt.dy = atof(p);
				}
				p = strtok(NULL, d);

				++nCnt;
			}

			if (nCnt == 0)
			{
				continue;
			}

			m_vecCenter.push_back(pt);
		}

		fclose(pf);
	}

	if (m_vecCenter.size() <= 0)
	{
		return false;
	}

	// 读取里程桩
	pf = fopen(strMileagePile, "r");
	hnMileagePile mileagePile;
	if (pf)
	{
		fseek(pf, 0, SEEK_SET);
		while (!feof(pf))
		{
			memset(str, 0, 1024);
			fgets(str, 1024, pf);
			p = strtok(str, d);

			nCnt = 0;

			while (p)
			{
				if (nCnt == 0)
				{
					mileagePile.dMileage = atof(p);
				}
				else
				{
					mileagePile.dRealMileage = atof(p);
				}
				p = strtok(NULL, d);

				++nCnt;
			}

			if (nCnt == 0)
			{
				continue;
			}

			m_vecMileagePile.push_back(mileagePile);
		}

		fclose(pf);
	}

	return true;
}

// 计算病害属性
bool hnOutputRoadDisease::calDiseaseAttr()
{
	// 获取车辆行驶的里程方向
	double nDir = 1.0;

	if (m_vecMileagePile.size() > 1)
	{
		// 里程桩排序
		sort(m_vecMileagePile.begin(), m_vecMileagePile.end(),sortMileagePile);

		if (m_vecMileagePile[0].dRealMileage > m_vecMileagePile[1].dRealMileage)
		{
			nDir = -1.0;
		}
	}
	else
	{
		if (m_nMileageDir == 1) // 小里程
		{
			nDir = -1.0;
		}
	}

	// 
	double dDist, dMileage;
	dDist = dMileage = 0.0;

	double dMaxMileage, dMinMileage;
	dMaxMileage = -10000000.0;
	dMinMileage = 10000000.0;

	for (int i = 0; i < m_vecDiseaseInfo.size(); i++)
	{
		// 计算每个点的里程和距离
		for (int j = 0; j < m_vecDiseaseInfo[i].vecPt.size(); j++)
		{
			// 计算里程和相对距离
			getMileageAndDist(m_vecDiseaseInfo[i].vecPt[j], dDist, dMileage);

			dDist = nDir*dDist;

			// 相对里程桩绝对里程
			getRealMileage(dMileage);

			m_vecDiseaseInfo[i].vecPt[j].dx = dMileage;
			m_vecDiseaseInfo[i].vecPt[j].dy = dDist;

			if (dMileage < dMinMileage)
			{
				dMinMileage = dMileage;
			}

			if (dMileage > dMaxMileage)
			{
				dMaxMileage = dMileage;
			}
		}

		if (dMaxMileage > m_dMaxMile)
		{
			m_dMaxMile = dMaxMileage;
		}

		if (dMinMileage < m_dMinMile)
		{
			m_dMinMile = dMinMileage;
		}

		m_vecDiseaseInfo[i].dBegMileage = dMinMileage;
		m_vecDiseaseInfo[i].dEndMileage = dMaxMileage;
	}

	return true;
}

// 划分病害
bool hnOutputRoadDisease::divideDiseaseGroup(vector<hnGridDiseaseInfo>& vecGrid)
{
	 int nBegMile = m_dMinMile / 1000.0;
	 double dBeg = m_dMinMile - nBegMile*1000.0;
     dBeg = nBegMile*1000.0 + (int)(dBeg / 500.0)*500.0;

	 int nEndMile = m_dMaxMile / 1000.0;
	 double dEnd = m_dMaxMile - nEndMile*1000.0;
	 dEnd = nEndMile*1000.0 + ((int)(dEnd / 500.0) + 1)*500.0;

	 char strName[1000] = {0};

	 // 划分
	 hnGridDiseaseInfo gridDisease;
	 int nGridCnt = (dEnd - dBeg) / 500.0;
	 for (int i = 0; i < nGridCnt; i ++)
	 {
		 gridDisease.dBegMileage = dBeg + i*500.0;
		 gridDisease.dEndMileage = dBeg + (i+1)*500.0;

		 // 起始位置
		 memset(strName, 0, 1000);
		 nBegMile = gridDisease.dBegMileage / 1000.0;
		 sprintf(strName, "K%03d+%03d", nBegMile, (int)(gridDisease.dBegMileage - nBegMile*1000.0));
		 gridDisease.strBegMile = strName;

		 // 终止位置
		 memset(strName, 0, 1000);
		 nEndMile = gridDisease.dEndMileage / 1000.0;
		 sprintf(strName, "K%03d+%03d", nEndMile, (int)(gridDisease.dEndMileage - nEndMile*1000.0));
		 gridDisease.strEndMile = strName;

		 gridDisease.strName = gridDisease.strBegMile + "-" + gridDisease.strEndMile;

		 // 划分病害
		 for (int j = 0; j < m_vecDiseaseInfo.size(); j++)
		 {
			 if ((m_vecDiseaseInfo[j].dBegMileage >= gridDisease.dBegMileage &&
				 m_vecDiseaseInfo[j].dBegMileage <= gridDisease.dEndMileage) ||
				 (m_vecDiseaseInfo[j].dEndMileage >= gridDisease.dBegMileage &&
				 m_vecDiseaseInfo[j].dEndMileage <= gridDisease.dEndMileage) )
			 {
				 gridDisease.vecDiseaseInfo.push_back(m_vecDiseaseInfo[j]);
				 continue;
			 }

			 if ((gridDisease.dBegMileage >= m_vecDiseaseInfo[j].dBegMileage &&
				 gridDisease.dBegMileage <= m_vecDiseaseInfo[j].dEndMileage) ||
				 (gridDisease.dEndMileage >= m_vecDiseaseInfo[j].dBegMileage &&
				 gridDisease.dEndMileage <= m_vecDiseaseInfo[j].dEndMileage))
			 {
				 gridDisease.vecDiseaseInfo.push_back(m_vecDiseaseInfo[j]);
			 }
		 }

		 vecGrid.push_back(gridDisease);
	 }

	return true;
}

// 分割病害
bool hnOutputRoadDisease::segGroupDisease(hnGridDiseaseInfo gridDisease)
{
	// 分割跨500m里程病害
	for (int i = 0; i < gridDisease.vecDiseaseInfo.size(); i++)
	{
		if (gridDisease.vecDiseaseInfo[i].vecPt.size() < 2)
		{
			continue;
		}

		// 分割病害
		segDisease(gridDisease.dBegMileage, gridDisease.dEndMileage, gridDisease.vecDiseaseInfo[i]);
	}

	vector<hnRoadDiseaseInfo> vecAllDisease;

	double dBegMile, dEndMile;

	for (int i = 0; i < 5; i++)
	{
		dBegMile = gridDisease.dBegMileage + i*100.0;
		dEndMile = gridDisease.dBegMileage + (i+1)*100.0;

		for (int j = 0; j < gridDisease.vecDiseaseInfo.size(); j++)
		{
			if ((gridDisease.vecDiseaseInfo[j].dBegMileage >= dBegMile &&
				gridDisease.vecDiseaseInfo[j].dBegMileage <= dEndMile) ||
				(gridDisease.vecDiseaseInfo[j].dEndMileage >= dBegMile &&
				gridDisease.vecDiseaseInfo[j].dEndMileage <= dEndMile) )
			{
				// 分割病害
				if (!segDisease(dBegMile, dEndMile, gridDisease.vecDiseaseInfo[j]))
				{
					continue;
				}

				vecAllDisease.push_back(gridDisease.vecDiseaseInfo[j]);
				continue;
			}

			if ((dBegMile >= gridDisease.vecDiseaseInfo[j].dBegMileage &&
				dBegMile <= gridDisease.vecDiseaseInfo[j].dEndMileage) ||
				(dEndMile >= gridDisease.vecDiseaseInfo[j].dBegMileage &&
				dEndMile <= gridDisease.vecDiseaseInfo[j].dEndMileage))
			{
				// 分割病害
				if (!segDisease(dBegMile, dEndMile, gridDisease.vecDiseaseInfo[j]))
				{
					continue;
				}

				vecAllDisease.push_back(gridDisease.vecDiseaseInfo[j]);
			}
		}
	}

	gridDisease.vecDiseaseInfo = vecAllDisease;

	return true;
}

// 检查病害
bool hnOutputRoadDisease::checkDisease(hnGridDiseaseInfo gridDisease)
{
	return true;

}

// 计算某点到中心线的投影距离以及相对里程
bool hnOutputRoadDisease::getMileageAndDist(hn2dPt pt, double& dDist, double& dMileage)
{
	dDist = 0.0;
	dMileage = 0.0;

	if (m_vecCenter.size() <= 0)
	{
		return false;
	}

	hn2dPt pBegPoint, pEndPoint;

	double dA = 0.0;
	double dB = 0.0;

	vector<int> vecIndex;

	// 遍历线段的所有点
	for (int j = 0; j < m_vecCenter.size() - 1; j++)
	{
		pBegPoint = m_vecCenter[j];
		pEndPoint = m_vecCenter[j + 1];

		if ((pBegPoint.dx == pEndPoint.dx) && (pBegPoint.dy == pEndPoint.dy))
		{
			continue;
		}

		dA = (pEndPoint.dx - pBegPoint.dx) * (pt.dx - pBegPoint.dx) +
			(pEndPoint.dy - pBegPoint.dy) * (pt.dy - pBegPoint.dy);

		dB = (pBegPoint.dx - pEndPoint.dx) * (pt.dx - pEndPoint.dx) +
			(pBegPoint.dy - pEndPoint.dy) * (pt.dy - pEndPoint.dy);


		if (dA * dB < 0)
		{
			continue;
		}

		vecIndex.push_back(j);
	}

	if (vecIndex.size() <= 0)
	{
		return false;
	}

	double dMin = 100000;
	double dTempMileage = 0.0;
	int nIndex = -1;

	for (int i = 0; i < vecIndex.size(); i++ )
	{
		// 获取垂足以及方向向量
		pBegPoint = m_vecCenter[vecIndex[i]];
		pEndPoint = m_vecCenter[vecIndex[i] + 1];

		// 两线的距离
		dDist = sqrt((pEndPoint.dx - pBegPoint.dx) * (pEndPoint.dx - pBegPoint.dx) +
			(pEndPoint.dy - pBegPoint.dy) * (pEndPoint.dy - pBegPoint.dy));

		dA = (pEndPoint.dx - pBegPoint.dx) * (pt.dx - pBegPoint.dx) +
			(pEndPoint.dy - pBegPoint.dy) * (pt.dy - pBegPoint.dy);

		dB = sqrt((pt.dx - pBegPoint.dx) * (pt.dx - pBegPoint.dx) +
			(pt.dy - pBegPoint.dy) * (pt.dy - pBegPoint.dy));

		dA = dA / dB / dDist;

		dTempMileage = dA * dB;

		dA = sqrt(1 - dA * dA) * dB;

		if (dA > 200)
		{
			continue;
		}

		if (dA < dMin)
		{
			dMin = dA;
			nIndex = vecIndex[i];
		}
	}

	if (nIndex == -1)
	{
		dDist = 0.0;
		dMileage = 0.0;
		return false;
	}

	// 距离
	dDist = dMin;
	dA = 0.0;

	// 计算里程
	for (int i = 0; i < nIndex; i++)
	{
		pBegPoint = m_vecCenter[i];
		pEndPoint = m_vecCenter[i + 1];

		// 两线的距离
		dA += sqrt((pEndPoint.dx - pBegPoint.dx) * (pEndPoint.dx - pBegPoint.dx) +
			(pEndPoint.dy - pBegPoint.dy) * (pEndPoint.dy - pBegPoint.dy));
	}

	dMileage = dA + dTempMileage;

	// 计算方向
	double dC;
	dA = dB = dC = 0.0;

	pBegPoint = m_vecCenter[nIndex];
	pEndPoint = m_vecCenter[nIndex + 1];

	// 求直线的参数
	if (pBegPoint.dx == pEndPoint.dx)
	{
		dA = 1;
		dB = 0;
		dC = -pBegPoint.dx;
	}
	else
	{
		dA = (pEndPoint.dy - pBegPoint.dy) / (pEndPoint.dx - pBegPoint.dx);
		dB = 1;
		//dC = (pLineBegPt.X * pLineEndPt.Y - pLineEndPt.X * pLineBegPt.Y) / (pLineEndPt.X - pLineBegPt.X);
		dC = pBegPoint.dy - dA * pBegPoint.dx;
	}

	// 方向 0 - 左边 1-右边
	int nDir = 0;

	dTempMileage = pt.dy - pt.dx * dA - dC;

	if (dA >= 0)
	{
		if (dTempMileage < 0)
		{
			nDir = 1;
		}
		else
		{
			nDir = 0;
		}
	}
	else
	{
		if (dTempMileage >= 0)
		{
			nDir = 1;
		}
		else
		{
			nDir = 0;
		}
	}

	if (nDir == 0)
	{
		dDist = - dDist;
	}

	return true;
}

// 里程
bool hnOutputRoadDisease::getRealMileage(double& dMileage)
{
	if (m_vecMileagePile.size() >=2)
	{
		// 判断里程桩真实里程方向
		bool bBigMileage = true;
		if (m_vecMileagePile[0].dRealMileage > m_vecMileagePile[1].dRealMileage)
		{
			bBigMileage = false;
		}

		double dRealMileage = 0.0;

		// 选择的里程桩索引
		int nPileIndex = 0;

		if (dMileage < m_vecMileagePile[0].dMileage)
		{
			if (bBigMileage)
			{
				dRealMileage = m_vecMileagePile[0].dRealMileage - (m_vecMileagePile[0].dMileage - dMileage);
			}
			else
			{
				dRealMileage = m_vecMileagePile[0].dRealMileage + (m_vecMileagePile[0].dMileage - dMileage);
			}
		}
		else if (dMileage > m_vecMileagePile[m_vecMileagePile.size() - 1].dMileage)
		{
			if (bBigMileage)
			{
				dRealMileage = m_vecMileagePile[m_vecMileagePile.size() - 1].dRealMileage + (dMileage - m_vecMileagePile[m_vecMileagePile.size() - 1].dMileage);
			}
			else
			{
				dRealMileage = m_vecMileagePile[m_vecMileagePile.size() - 1].dRealMileage - (dMileage - m_vecMileagePile[m_vecMileagePile.size() - 1].dMileage);
			}

		}
		else
		{
			for (int i = 0; i < m_vecMileagePile.size() - 1; i++)
			{
				if (dMileage > m_vecMileagePile[i].dMileage &&
					dMileage < m_vecMileagePile[i + 1].dMileage)
				{
					nPileIndex = i;
					break;
				}
			}

			dRealMileage = m_vecMileagePile[nPileIndex].dRealMileage + (m_vecMileagePile[nPileIndex + 1].dRealMileage -
				m_vecMileagePile[nPileIndex].dRealMileage) / (m_vecMileagePile[nPileIndex + 1].dMileage -
				m_vecMileagePile[nPileIndex].dMileage) * (dMileage - m_vecMileagePile[nPileIndex].dMileage);

		}

		dMileage = dRealMileage;
	}
	else if (m_vecMileagePile.size() == 1)
	{
		dMileage = m_vecMileagePile[0].dRealMileage + (dMileage - m_vecMileagePile[0].dMileage);
	}
	else
	{
		if (m_nMileageDir == 0)
		{
			dMileage = dMileage + m_dBegMileage;
		}
		else
		{
			dMileage = m_dBegMileage - dMileage;
		}
	}

	return true;
}

// 点坐标
void hnOutputRoadDisease::getPoint(hn2dPt ptBeg, hn2dPt ptEnd, hn2dPt& oriPt)
{
	double dA,dB,dC;
	dA = dB = dC = 0.0;

	// 求直线的参数
	if (ptBeg.dx == ptEnd.dx)
	{
		dA = 1;
		dB = 0;
		dC = -ptBeg.dx;
	}
	else
	{
		dA = (ptEnd.dy - ptBeg.dy) / (ptEnd.dx - ptBeg.dx);
		dB = 1;
		dC = ptBeg.dy - dA * ptBeg.dx;
	}

	if (dB = 0.0)
	{
		oriPt.dy = (ptBeg.dy + ptEnd.dy) / 2.0;
	}
	else
	{
		oriPt.dy = (-1.0*dA*oriPt.dx - dC) / dB;
	}

}

// 分割病害
bool hnOutputRoadDisease::segDisease(double dBegMile, double dEndMile, 
	hnRoadDiseaseInfo& roadDisease)
{
	vector<hn2dPt> vecPt;

	// 起始点过界 
	if (roadDisease.dBegMileage - dBegMile < -0.000001)
	{
		// 线状病害
		if (roadDisease.nGeometry == 0)
		{
			int nIndex = -1;
			hn2dPt ptOri;

			vecPt = roadDisease.vecPt;

			for (int j = 0; j < vecPt.size() - 1; j++)
			{
				if (dBegMile >= vecPt[j].dx &&
					dBegMile <= vecPt[j+1].dx)
				{
					nIndex = j;
					break;
				}
			}

			if (nIndex == -1)
			{
				return false;
			}

			// 起始点过界
			if (vecPt[0].dx < dBegMile)
			{
				ptOri.dx = dBegMile;
				getPoint(vecPt[nIndex], vecPt[nIndex+1], ptOri);

				ptOri.dx = ptOri.dx + 0.0001;
				vecPt[nIndex + 1] = ptOri;
				vecPt.erase(vecPt.begin() + nIndex + 2, vecPt.end());
				roadDisease.vecPt = vecPt;
			}

			// 终止点过界
			if (vecPt[vecPt.size() - 1].dx < dBegMile)
			{
				ptOri.dx = dBegMile;
				getPoint(vecPt[nIndex], vecPt[nIndex+1], ptOri);

				ptOri.dx = ptOri.dx + 0.0001;
				vecPt[nIndex + 1] = ptOri;
				vecPt.erase(vecPt.begin(), vecPt.begin() + nIndex + 2);
				roadDisease.vecPt = vecPt;
			}

		}
		else // 面状病害
		{
			if (roadDisease.vecPt.size() < 4)
			{
				return false;
			}

			if (roadDisease.vecPt[0].dx < dBegMile)
			{
				roadDisease.vecPt[0].dx = dBegMile + 0.0001;
				roadDisease.vecPt[3].dx = dBegMile + 0.0001;
			}

			if (roadDisease.vecPt[1].dx < dBegMile)
			{
				roadDisease.vecPt[1].dx = dBegMile + 0.0001;
				roadDisease.vecPt[2].dx = dBegMile + 0.0001;
			}
		}
	}

	// 终止点过界
	if (roadDisease.dEndMileage - dEndMile > 0.000001)
	{
		// 线状病害
		if (roadDisease.nGeometry == 0)
		{
			int nIndex = -1;
			hn2dPt ptOri;

			vecPt = roadDisease.vecPt;

			for (int j = 0; j < vecPt.size() - 1; j++)
			{
				if (dEndMile >= vecPt[j].dx &&
					dEndMile <= vecPt[j+1].dx)
				{
					nIndex = j;
					break;
				}
			}

			if (nIndex == -1)
			{
				return false;
			}

			// 起始点过界
			if (vecPt[0].dx > dEndMile)
			{
				ptOri.dx = dEndMile;
				getPoint(vecPt[nIndex], vecPt[nIndex+1], ptOri);

				ptOri.dx = ptOri.dx - 0.0001;
				vecPt[nIndex + 1] = ptOri;
				vecPt.erase(vecPt.begin(), vecPt.begin() + nIndex + 2);
				roadDisease.vecPt = vecPt;
			}

			// 终止点过界
			if (vecPt[vecPt.size() - 1].dx > dEndMile)
			{
				ptOri.dx = dEndMile;
				getPoint(vecPt[nIndex], vecPt[nIndex+1], ptOri);

				ptOri.dx = ptOri.dx - 0.0001;
				vecPt[nIndex + 1] = ptOri;
				vecPt.erase(vecPt.begin() + nIndex + 2, vecPt.end());
				roadDisease.vecPt = vecPt;
			}

		}
		else // 面状病害
		{
			if (roadDisease.vecPt.size() < 4)
			{
				return false;
			}

			if (roadDisease.vecPt[0].dx > dEndMile)
			{
				roadDisease.vecPt[0].dx = dEndMile - 0.0001;
				roadDisease.vecPt[3].dx = dEndMile - 0.0001;
			}

			if (roadDisease.vecPt[1].dx > dEndMile)
			{
				roadDisease.vecPt[1].dx = dEndMile - 0.0001;
				roadDisease.vecPt[2].dx = dEndMile - 0.0001;
			}
		}
	}

	return true;
}