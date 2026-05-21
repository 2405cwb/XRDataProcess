#include "StdAfx.h"
#include "ConvertImpl.h"
#include"Dir.h"
#include<iostream>
#include<fstream>
#include <iomanip>
#define IRIExcelSide 2
#define RQIJudgeType 0
std::string UTF8ToGB(const char* str)
{
	std::string result;
	wchar_t *wstr;
	char* szRes;

	int len = MultiByteToWideChar(CP_UTF8,0,str,-1,NULL,0);
	wstr = new wchar_t[len+1];
	memset(wstr,0,len+1);
	MultiByteToWideChar(CP_UTF8,0,str,-1,wstr,len);

	len=WideCharToMultiByte(CP_ACP,0,wstr,-1,NULL,0,NULL,NULL);
	szRes = new char[len+1];
	WideCharToMultiByte(CP_ACP,0,wstr,-1,szRes,len,NULL,NULL);
	szRes[len] = '\0';
	result = szRes;
	delete[] wstr;
	delete[] szRes;

	return result;
}
ConvertImpl::ConvertImpl(void)
{
}


ConvertImpl::~ConvertImpl(void)
{
}

void ConvertImpl::LoadWorkspace(string projPath)
{
	vecProjectPath.clear();
	if(Dir::Exist(projPath))
	{
		vector<string> subPath = Dir::GetDirectories(projPath);
		
		for(int i=0;i<subPath.size();i++)
		{
			if (Dir::Exist(subPath[i] + "\\ProjectInfo.txt"))
			{
				vecProjectPath.push_back(subPath[i]);
			}
			else
			{
				/*vector<string> subsub =*/ LoadWorkspace(subPath[i]);
				/*if(subsub.size()>0)
					vecProjectPath.insert(subPath.end(),subsub.begin(),subsub.end());*/
			}
		}
		
	}
	if(Dir::Exist(projPath + "\\ProjectInfo.txt"))
	{
		vecProjectPath.push_back(projPath);
	}

	//return projects;
}

void ConvertImpl::LoadConfig()
{
	for(int i=0;i<vecProjectPath.size();i++)
	{
		ProjectSetting setting;
		string projectInfoPath = vecProjectPath[i] + "\\ProjectInfo.txt";
		std::ifstream ff(projectInfoPath.c_str());
		string strLine;
		getline(ff,strLine);
		while(getline(ff,strLine))
		{
			strLine = UTF8ToGB(strLine.c_str());
			string header = strLine.substr(0,strLine.find("："));
			string tailer = strLine.substr(strLine.find("：")+2);
			if(header == "省")	setting.province = tailer;
			else if(header == "市")	setting.city = tailer;
			else if(header == "县")	setting.region = tailer;
			else if(header == "工程起点道路编号")	setting.roadCode = tailer;
			else if(header == "工程起点道路名称")	setting.roadName = tailer;
			else if(header == "工程起点桩号")	setting.startMile = tailer;
			else if(header == "行车方向")	setting.direction = tailer;
			else if(header == "公路等级")	setting.roadLevel = tailer;
			else if(header == "车道")	setting.roadNum = tailer;
			else if(header == "采集日期")	setting.dataDate = tailer;
			else if(header == "工程开始时刻")	setting.dataTime = tailer;
			else if(header == "检测员")	setting.dataPerson = tailer;
			else if(header == "检测天气")	setting.dataWeather = tailer;
			else if(header == "路面材质")	setting.roadType = tailer;
			else if(header == "工程终点道路标识桩号")	setting.endMile = tailer;
			else if(header == "工程总里程数")	setting.endDmi = tailer;
		}
		vecProjectSetting.push_back(setting);
		ff.close();
	}
}

void ConvertImpl::CreateCICS()
{
	vector<double>LIRIVal;
	vector<double>RIRIVal;
	double irival;
	for(int i=0;i<vecProjectPath.size();i++)
	{
		vector<double> IRIData;
		string DAQ0 = vecProjectPath[i]+"\\IRIMTD\\DAQ0";
		string DAQ1 = vecProjectPath[i]+"\\IRIMTD\\DAQ1";
		bool bDAQ0 = Dir::Exist(DAQ0);
		bool bDAQ1 = Dir::Exist(DAQ1);
		if(bDAQ0)
		{
			LoadIRIFile(DAQ0+"\\IRI_10m.txt",LIRIVal);
		}
		if(bDAQ1)
		{
			LoadIRIFile(DAQ1+"\\IRI_10m.txt",RIRIVal);
		}
		int IRISum = max(LIRIVal.size(),RIRIVal.size());
		for(int j =0; j<IRISum;j++)
		{
			if(bDAQ0 && bDAQ1)
			{
				if (RQIJudgeType == 0)
				{
					irival = (LIRIVal[j] + RIRIVal[j]) * 0.5;
				}
				else if (RQIJudgeType == 1)
				{
					irival = max(LIRIVal[j], RIRIVal[j]);
				}
			}
			else if (bDAQ0)
			{
				irival = LIRIVal[j];
			}
			else if (bDAQ1)
			{
				irival = RIRIVal[j];
			}
			IRIData.push_back(irival);
		}
		string strDirection;
		if(vecProjectSetting[i].direction == "上行")
			strDirection = "A";
		else if (vecProjectSetting[i].direction == "下行")
			strDirection = "B";

		string strData = vecProjectSetting[i].dataDate;
		strData.insert(6,"-",1);
		strData.insert(4,"-",1);
		string strTime = vecProjectSetting[i].dataTime;
		strTime.insert(4,"-",1);
		strTime.insert(2,"-",1);

		string strMile = vecProjectSetting[i].startMile;
		strMile = strMile.substr(1);
		string strKMile = strMile.substr(0,strMile.find("+"));
		strMile = strMile.substr(strMile.find("+")+1);
		double dMile = atoi(strKMile.c_str())*1000 + atoi(strMile.c_str());
		dMile = dMile / 1000.0;
		char tmp[64];
		sprintf(tmp,"%.3lf",dMile);
		string strTmp = tmp;
		string outputName = vecProjectSetting[i].roadCode + vecProjectSetting[i].roadName+strDirection +"-"+strTmp+"-"+strData+"-"+strTime+".IRI";

		ofstream of(vecProjectPath[i]+"\\"+outputName);
		of<<"CICS平整度检测结果"<<endl;
		of<<endl;
		if(vecProjectSetting[i].direction == "上行")
		{
			for(int m=0;m<IRIData.size();m++)
			{
				sprintf(tmp,"%.3lf",dMile+(m+1)*0.01);
				string tmp1 = tmp;
				sprintf(tmp,"%.2lf",IRIData[m]);
				string tmp2 = tmp;
				of<<"L"<<tmp1<<"\t\t"<<tmp2<<endl;
			}
		}else if(vecProjectSetting[i].direction == "下行")
		{
			for(int m=0;m<IRIData.size();m++)
			{
				sprintf(tmp,"%.3lf",dMile-(m)*0.01);
				string tmp1 = tmp;
				sprintf(tmp,"%.2lf",IRIData[m]);
				string tmp2 = tmp;
				of<<"L"<<tmp1<<"\t\t"<<tmp2<<endl;
			}
		}

		of.close();
	}
}

void ConvertImpl::LoadIRIFile(string filePath,vector<double> &data)
{
	std::ifstream ff(filePath);
	string strLine;
	data.clear();
	while(getline(ff,strLine))
	{
		strLine = strLine.substr(strLine.find_first_of(" ")+1);
		data.push_back(atof(strLine.c_str()));
	}
	ff.close();
}