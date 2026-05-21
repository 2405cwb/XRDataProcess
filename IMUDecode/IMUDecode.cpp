// 这是主 DLL 文件。

#include "stdafx.h"
#include "NavDataParser.h"
#include <stdio.h>
#include <stdlib.h>
#include <string>
#include "IMUDecode.h"

using namespace NavDataParser;

#define BUF_SIZE 4096 //4kB
unsigned int BigLittleSwap32(unsigned int value)
{
	// 32位大小端转换
	return ((value & 0x000000FF) << 24 |
		(value & 0x0000FF00) << 8  |
		(value & 0x00FF0000) >> 8  |
		(value & 0xFF000000) >> 24 );
}

unsigned long long BigLittleSwap64(unsigned long long value)
{
	// 64位大小端转换
	return ((value & 0x00000000000000FF) << 56 |
		(value & 0x000000000000FF00) << 40 |
		(value & 0x0000000000FF0000) << 24 |
		(value & 0x00000000FF000000) << 8  |
		(value & 0x000000FF00000000) >> 8  |
		(value & 0x0000FF0000000000) >> 24 |
		(value & 0x00FF000000000000) >> 40 |
		(value & 0xFF00000000000000) >> 56);
}

unsigned short BigLittleSwap16(unsigned short value)
{
	// 16位大小端转换
	return ((value & 0xff00) >> 8 | 
		(value & 0x00ff) << 8);
}

//解码惯导原始数据
extern "C" __declspec(dllexport) void
	DecodeIMUBin(char* binfpath )
{
	UINT8 * RxBuffer;
	//Allocate memory for buffers
	RxBuffer = (UINT8 *)malloc(BUF_SIZE*2);

	char *outfpath;
	outfpath = (char*)malloc(1024);

	UINT8 *RxBufferAct = RxBuffer;
	UINT8 *RxBufferEnd = RxBuffer;
	float initLat;
	float initLon;
	float initAlt;

	//Create container structure for HG navigation set
	NavDataParser::HgNavMessageSet Message;
	Message.Init();

	int pointer;
	int readCount = 0;
	int endOffset = 0;
	int status = 2;

	int readbyte = 0;
	char msgstr[4096];

	UINT32 msgid = 0;

	sprintf(outfpath, "%s.csv", binfpath);
	FILE *fpout = fopen(outfpath, "wt");
	if(fpout != NULL)
	{
		FILE *fp = fopen(binfpath, "rb");
		int len = 0;
		if(fp != NULL)
		{
			while (!feof(fp))
			{
				//Copy the remainder of previous reading before the newone
				memcpy(RxBuffer, RxBufferAct, (RxBufferEnd - RxBufferAct));
				//Reset the end of the buffer
				RxBufferEnd = RxBuffer + (RxBufferEnd - RxBufferAct);
				//Set Reading pointer to beginning
				RxBufferAct = RxBuffer;

				readCount = fread(RxBufferEnd, sizeof(UINT8), BUF_SIZE, fp);
				readbyte = readbyte + readCount;
				//printf("%d\t", readbyte);

				endOffset = 0;
				if (readCount>0)
				{
					// Move the pointer to the end of the Read buffer
					RxBufferEnd += readCount;
					while (RxBufferAct <= RxBufferEnd - (MSG_LEN_MAX*4))
					{
						if (*(UINT32*)(RxBufferAct) == 0xA5C381FF)
						{
							msgid = *(UINT32*)(RxBufferAct + 4);
							switch (msgid)
							{
							case 0x2401: case 0x2402:
								{
									status = NavDataParser::Get0x2402Message(RxBufferAct, 0, &Message.NavOutput, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.NavOutput.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6108:
								{
									status = NavDataParser::Get0x6108Message(RxBufferAct, 0, &Message.GnssPosition, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.GnssPosition.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6109:
								{
									status = NavDataParser::Get0x6109Message(RxBufferAct, 0, &Message.GnssAttitude, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.GnssAttitude.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x2311:
								{
									status = NavDataParser::Get0x2311Message(RxBufferAct, 0, &Message.InertialData, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.InertialData.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x2201:
								{
									status = NavDataParser::Get0x2201Message(RxBufferAct, 0, &Message.TimeMark, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.TimeMark.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6504:
								{
									status = NavDataParser::Get0x6504Message(RxBufferAct, 0, &Message.NEDVelocity, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.NEDVelocity.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6405:
								{
									status = NavDataParser::Get0x6405Message(RxBufferAct, 0, &Message.EulerAttitudes, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.EulerAttitudes.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6403:
								{
									status = NavDataParser::Get0x6403Message(RxBufferAct, 0, &Message.GeoPosition, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.GeoPosition.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);

										if (Message.GeoPosition.InsGnssBIT.InsMode == 4 && initLat == -999.0)
										{
											initLat = Message.GeoPosition.Latitude;
											initLon = Message.GeoPosition.Longitude;
											initAlt = Message.GeoPosition.AltitudeAboveEllipsoid;
										}
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x20ff:
								{
									status = NavDataParser::Get0x20ffMessage(RxBufferAct, 0, &Message.Ack, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.Ack.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x2011:
								{
									status = NavDataParser::Get0x2011Message(RxBufferAct, 0, &Message.Status, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.Status.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x2001:
								{
									status = NavDataParser::Get0x2001Message(RxBufferAct, 0, &Message.Configuration, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.Configuration.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6110:
								{
									status = NavDataParser::Get0x6110Message(RxBufferAct, 0, &Message.DistanceTraveled, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.DistanceTraveled.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6111:
								{
									status = NavDataParser::Get0x6111Message(RxBufferAct, 0, &Message.MotionDetection, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.MotionDetection.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{	
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6424:
								{
									status = NavDataParser::Get0x6424Message(RxBufferAct, 0, &Message.AntennaLeverArmEstimates, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.AntennaLeverArmEstimates.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							case 0x6438:
								{
									status = NavDataParser::Get0x6438Message(RxBufferAct, 0, &Message.OdometerCalibration, &endOffset);
									if (!status)
									{
										RxBufferAct += endOffset-1;
										Message.OdometerCalibration.printDataToCsv(msgstr, sizeof(msgstr));
										fprintf(fpout, "0x%04X,%s\n", msgid, msgstr);
									}
									else if (status == 1) // Status == 1 Indicates error in check sum
									{	
										fprintf(fpout, "0x%04X,ChecksumError\n", msgid);
									}
									break;
								}
							default:
								//ROS_INFO("found: %.4x", *(UINT32*)(RxBufferAct + 4));
								fprintf(fpout, "0x%04X,NotFound\n", msgid);					
								break;

							} // SWITCH (RxBufferAct+4)

						} // if *RxBufferAct == 0x0E
						//Move the read buffer pointer
						RxBufferAct++;
					} //while RxAct < RxEnd
				}
			}
			fclose(fp);
			fp = NULL;
		}

		fclose(fpout);
		fpout = NULL;

		//free(RxBufferAct);
		//free(outfpath);
	}
}


