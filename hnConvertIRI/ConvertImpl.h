#pragma once
#include<string>
#include<vector>

using namespace std;
#ifndef _PROJ_SETTING
#define _PROJ_SETTING
typedef struct _Proj_Setting
{
	string province;
	string city;
	string region;
	string roadCode;
	string roadName;
	string startMile;
	string direction;
	string roadLevel;
	string roadNum;
	string dataDate;
	string dataTime;
	string dataPerson;
	string dataWeather;
	string roadType;
	string endMile;
	string endDmi;
}ProjectSetting;
#endif
class ConvertImpl
{
public:
	ConvertImpl(void);
	~ConvertImpl(void);

	void LoadWorkspace(string projPath);
	void LoadConfig();
	void CreateCICS();

private:
	void LoadIRIFile(string filePath,vector<double> &data);
private:
	vector<string> vecProjectPath;
	vector<ProjectSetting> vecProjectSetting;
};

