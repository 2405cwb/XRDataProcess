// IMUDecodeTest.cpp: 主项目文件。

#include "stdafx.h"
#include <stdio.h>
#include "IMUDecode.h"

using namespace System;

#pragma comment( lib, "IMUDecode.lib ")

int main(array<System::String ^> ^args)
{
	char srcfpath[] = "G:\\06_模块化设备数据\\2021年\\几何线形测试数据\\2021-08-04-藏龙岛\\__上行__河北省_邢台市_威县_20210804_115242\\camera0\\imu.hon";
	DecodeIMUBin(srcfpath);
	
	return 0;
}
