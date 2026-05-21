// hnConvertIRI.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "hnConvertIRI.h"
#include "ConvertImpl.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 唯一的应用程序对象

CWinApp theApp;

using namespace std;

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
	int nRetCode = 0;

	HMODULE hModule = ::GetModuleHandle(NULL);

	if (hModule != NULL)
	{
		// 初始化 MFC 并在失败时显示错误
		if (!AfxWinInit(hModule, NULL, ::GetCommandLine(), 0))
		{
			//  更改错误代码以符合您的需要
			_tprintf(_T("错误: MFC 初始化失败\n"));
			nRetCode = 1;
		}
		else
		{
			// 在此处为应用程序的行为编写代码。
		}
	}
	else
	{
		//  更改错误代码以符合您的需要
		_tprintf(_T("错误: GetModuleHandle 失败\n"));
		nRetCode = 1;
	}

	ConvertImpl _impl;
	CString strFile = "vzm";
	CString strFileName = "File";
	CFolderPickerDialog dlg(NULL,0,theApp.GetMainWnd(),0);
	if(dlg.DoModal() == IDOK)
	{
		string strFilePath = dlg.GetPathName();
		_impl.LoadWorkspace(strFilePath);
		_impl.LoadConfig();
		_impl.CreateCICS();
	}
	cout<<"导出完成"<<endl;
	Sleep(3000);
	return 1;
}
