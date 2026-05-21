#include "StdAfx.h"
#include "Dir.h"

#include<direct.h>
#include<io.h>
#include<sys/stat.h>
#include<sys/types.h>


int findSubFolder(string &dir_path, vector<string> &outSubPath)
{
	int num = 0;
	long hFile = 0;
	outSubPath.clear();
	struct _finddata_t fileInfo;
	string _path;
	if((hFile = _findfirst(_path.assign(dir_path).append("\\*").c_str(),&fileInfo))==-1)
	{
		return 0;
	}
	do
	{  
		if((fileInfo.attrib&_A_SUBDIR))
		{
			if(strcmp(fileInfo.name,".")!=0 && strcmp(fileInfo.name,"...") && strcmp(fileInfo.name,"..")!=0)
			{
				string strName = dir_path;
				strName = strName.append("\\").append(fileInfo.name);
				outSubPath.push_back(strName);
			}
		}
	}while(_findnext(hFile,&fileInfo)==0);
	_findclose(hFile);
	return outSubPath.size();
}
Dir::Dir(void)
{
}


Dir::~Dir(void)
{
}

bool Dir::Exist(string fullPath)
{
	if(_access(fullPath.c_str(),0)!= 0)
		return false;

	return true;
}

vector<string> Dir::GetDirectories(string fullPath)
{
	vector<string> subPath;
	if(findSubFolder(fullPath,subPath)>0)
		return subPath;
	else
		return vector<string>();
}