// HgDataParser.cpp : Defines the exported functions for the DLL application.
// Compile by using: cl /EHsc /HGDATAPARSER_EXPORTS /LD HgDataParser.cpp /doc 

/*
HONEYWELL hereby grants to you a, perpetual, free of charge, worldwide, irrevocable, non-exclusive license to use, copy, modify, merge,
publish, distribute, sublicense the software and associated documentation (the “Software?, subject to the following conditions:

YOU AGREE THAT YOU ASSUME ALL THE RESPONSIBILITY AND RISK FOR YOUR USE OF THE SOFTWARE AND THE RESULTS AND PERFORMANCE THEREOF.
THE SOFTWARE IS PROVIDED TO YOU ON AN “AS IS?AND “AS AVAILABLE?BASIS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.

IN NO EVENT WILL HONEYWELL BE LIABLE TO YOU FOR ANY DIRECT, SPECIAL, INDIRECT, INCIDENTAL, CONSEQUENTIAL, EXEMPLARY OR PUNITIVE DAMAGES,
INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOST DATA, LOST PROFITS, LOSS OF GOODWILL, LOST REVENUE, SERVICE INTERRUPTION, DEVICE DAMAGE
OR SYSTEM FAILURE, UNDER ANY THEORY OF LIABILITY, INCLUDING, WITHOUT LIMITATION CONTRACT OR TORT, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE.

All technology from the United States is subject to export regulations. This software is related to a device that has a United States
Export Commodity Classification of ECCN 7A994 with associated country chart control code of AT1. This generally will not require a license
to be exported or re-exported. However, if you plan to export this item to an embargoed or sanctioned country, to a party of concern,
or in support of a prohibited end-use, you may be required to obtain a license.
*/

#include "Stdafx.h"
#include "NavDataParser.h"


namespace NavDataParser
{

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	//Supproted HW: HG4940
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	// Deserialize Data Overloaded functions definition
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2402NavigationOutput *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2001Configuration *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2011Status *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x20ffAck *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2201TimeMark *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6109GnssAttitude *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6108GnssPosition *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2311InertialData *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6504NEDVelocity *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6405EulerAttitudes *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6403GeodeticPosition *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6110DistanceTraveled *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6111MotionDetection *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6424AntennaLeverArmEstimates *Message);
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6438OdometerCalibration *Message);


	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-OUTPUT-MESSAGES-------------------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/

	/// <summary>Calculate HG INS Checksum</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="wordLength">Word length of the message</param>  
	/// <returns>0 - OK, 1 - Incorrect, -9 - Incorrect wordLength</returns> 
	int HgInsChecksum(UINT8 *buffer, int startOffset, int wordLength)
	{

		UINT32 u32sum = 0;
		UINT8* pb = &buffer[startOffset];

		if (wordLength < 0)
			return -9;

		for (int i = 0; i < wordLength * 4; i = i + 4)
		{
			if (i != 12)
				u32sum += *(UINT32 *)(pb + i);
		}

		//calculate 2's complement
		u32sum = 0 - (INT32)u32sum;

		UINT32 Checksum = *(UINT32 *)(pb + 12);

		if (Checksum == u32sum)
			return 0;
		else
			return 1;


	}

	/// <summary>Return the INS / GNSS BIT Summary structure from byte buffer</summary>
	/// <param name="byteOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <returns>Filled structure containing the results of BIT</returns> 
	struct InsGnssBIT getInsGnssBIT(int byteOffset, UINT8*buffer)
	{
		struct InsGnssBIT bitValue;
		bitValue.InsMode = (buffer[byteOffset] & 0x07); //0-2 bits
		bitValue.InsStatus = (buffer[byteOffset] & 0x10) != 0;
		bitValue.ImuStatus = (buffer[byteOffset] & 0x20) != 0;
		bitValue.GnssStatus = (buffer[byteOffset] & 0x40) != 0;

		bitValue.MotionDetectActive = (buffer[byteOffset] & 0x80) != 0;
		bitValue.ZeroVelocity = (buffer[byteOffset+1] & 0x01) != 0;
		bitValue.ZeroVelocityPending= (buffer[byteOffset+2] & 0x01) != 0; // All Motion detection test are passed

		bitValue.GpsMode = (buffer[byteOffset+3] & 0xF0) >> 4; //28-31 bits

		/*StatusOut |= StatusMode.NavigationMode;     // 3 bits
		StatusOut |= StatusMode.Reserved1 << 3;     // 1 bit
		StatusOut |= StatusMode.INSStatusTest << 4; // 1 bit
		StatusOut |= StatusMode.IMUStatusTest << 5; // 1 bit
		StatusOut |= StatusMode.GNSSStatusTest << 6;// 1 bit
		StatusOut |= StatusMode.MotionDetectActive << 7; // 1 bit
		StatusOut |= StatusMode.StationaryMeasOn << 8;
		StatusOut |= StatusMode.StationaryMeasPending << 9;
		StatusOut |= StatusMode.GNSSmode << 28;     // 4 bits*/


		return bitValue;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	
	/// <summary>Fill Navigation Message structure from raw byte array. Selection of appropriate message is based on the message ID (2nd word)</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - Incorrect start offset
	///				-2 - Incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - No defined message present</returns>  
	int GetNavMessage(UINT8 *buffer, int startOffset, struct HgNavMessageSet *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		switch (*(UINT32*)(pb + 4))
		{
		//case 0x2401: //Navigation Output Message ||LEGACY||
		//	*endOffset = startOffset + 57 * 4;
		//	return Deserialize(buffer, startOffset, &Message->NavOutput);
		case 0x2402: //Smoothed Navigation Output Message
			*endOffset = startOffset + MSG_LEN_X2402 * 4;
			return Deserialize(buffer, startOffset, &Message->NavOutput);
		case 0x2001: //Configuration Output Message
			*endOffset = startOffset + MSG_LEN_X2001 * 4;
			return Deserialize(buffer, startOffset, &Message->Configuration);
		case 0x2011: //Status Output Message
			*endOffset = startOffset + MSG_LEN_X2011 * 4;
			return Deserialize(buffer, startOffset, &Message->Status);
		case 0x20ff: //ACK/NAK Output Message
			*endOffset = startOffset + MSG_LEN_X20FF * 4;
			return Deserialize(buffer, startOffset, &Message->Ack);
		case 0x2201: //Time Mark / PPS Output Message
			*endOffset = startOffset + MSG_LEN_X2201 * 4;
			return Deserialize(buffer, startOffset, &Message->TimeMark);
		case 0x2311: // Inertial Data Output Message
			*endOffset = startOffset + MSG_LEN_X2311 * 4;
			return Deserialize(buffer, startOffset, &Message->InertialData);
		case 0x6403: // INS Geodetic Position Output Message
			*endOffset = startOffset + MSG_LEN_X6403 * 4;
			return Deserialize(buffer, startOffset, &Message->GeoPosition);
		case 0x6504: // INS NED Velocity Output Message
			*endOffset = startOffset + MSG_LEN_X6504 * 4;
			return Deserialize(buffer, startOffset, &Message->NEDVelocity);
		case 0x6405: // INS Euler Attitudes Output Message
			*endOffset = startOffset + MSG_LEN_X6405 * 4;
			return Deserialize(buffer, startOffset, &Message->EulerAttitudes);
		case 0x6108: // GNSS position From Receiver Output Message
			*endOffset = startOffset + MSG_LEN_X6108 * 4;
			return Deserialize(buffer, startOffset, &Message->GnssPosition);
		case 0x6109: // GNSS Attitude From Receiver Output Message
			*endOffset = startOffset + MSG_LEN_X6109 * 4;
			return Deserialize(buffer, startOffset, &Message->GnssAttitude);
		case 0x6110: // Distance Traveled Based On Speed Aiding Output Message
			*endOffset = startOffset + MSG_LEN_X6110 * 4;
			return Deserialize(buffer, startOffset, &Message->DistanceTraveled);
		case 0x6111: // Motion Detection Output Message
			*endOffset = startOffset + MSG_LEN_X6111 * 4;
			return Deserialize(buffer, startOffset, &Message->MotionDetection);
		case 0x6424: // Antenna Lever Arms Estimation Output Message
			*endOffset = startOffset + MSG_LEN_X6424 * 4;
			return Deserialize(buffer, startOffset, &Message->AntennaLeverArmEstimates);
		case 0x6438: // Odometer Calibration Output Message
			*endOffset = startOffset + MSG_LEN_X6438 * 4;
			return Deserialize(buffer, startOffset, &Message->OdometerCalibration);
			/*Add additional messages*/

		default:
			*endOffset = startOffset;
			return -10;
		}
	}

	/// <summary>Retrieve the 0x2401 Navigation Output Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - Incorrect start offset
	///				-2 - Incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength</returns>  
	//int Get0x2401Message(UINT8 *buffer, int startOffset, struct Hg0x2402NavigationOutput *Message, int *endOffset)
	//{
	//	UINT8 *pb = &buffer[startOffset];
	//
	//	if (*(UINT32*)pb != 0xA5C381FF)
	//		return -3;
	//
	//	if ((*(UINT32*)(pb + 4)) == 0x2401)
	//	{
	//		*endOffset = startOffset + MSG_LEN_X2402 * 4;
	//		return Deserialize(buffer, startOffset, Message);
	//	}
	//	return -10;
	//}
	
	/// <summary>Retrieve the 0x2402 Smoothed Navigation Output Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns>  
	int Get0x2402Message(UINT8 *buffer, int startOffset, struct Hg0x2402NavigationOutput *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x2402) //Smoothed Navigation Output Message
		{
			*endOffset = startOffset + MSG_LEN_X2402 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x2001 Configuration Output Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns>  
	int Get0x2001Message(UINT8 *buffer, int startOffset, struct Hg0x2001Configuration *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x2001) //Configuration Output Message
		{
			*endOffset = startOffset + MSG_LEN_X2001 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;

	}

	/// <summary>Retrieve the 0x2011 Status Output Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns>  
	int Get0x2011Message(UINT8 *buffer, int startOffset, struct Hg0x2011Status *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x2011) //Status Output Message
		{
			*endOffset = startOffset + MSG_LEN_X2011 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;

	}

	/// <summary>Retrieve the 0x20FF ACK / NACK Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns>  
	int Get0x20ffMessage(UINT8 *buffer, int startOffset, struct Hg0x20ffAck *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x20ff) //ACK/NAK Output Message
		{
			*endOffset = startOffset + MSG_LEN_X20FF * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x2201 Time Mark / PPS Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x2201Message(UINT8 *buffer, int startOffset, struct Hg0x2201TimeMark *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x2201) //Time Mark / PPS Output Message
		{
			*endOffset = startOffset + MSG_LEN_X2201 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x2311 Inertial Data Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x2311Message(UINT8 *buffer, int startOffset, struct Hg0x2311InertialData *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x2311) // Inertial Data Output Message
		{
			*endOffset = startOffset + MSG_LEN_X2311 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6403 Geodetic Position Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6403Message(UINT8 *buffer, int startOffset, struct Hg0x6403GeodeticPosition *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6403) // INS Geodetic Position Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6403 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6504 NED Velocity Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6504Message(UINT8 *buffer, int startOffset, struct Hg0x6504NEDVelocity *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6504) // INS NED Velocity Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6504 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6405 Euler Attitudes Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6405Message(UINT8 *buffer, int startOffset, struct Hg0x6405EulerAttitudes *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6405) // INS Euler Attitudes Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6405 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6108 GNSS Position Message from byte array. Prepared for both legacy 26 word and current 32 word long messages</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6108Message(UINT8 *buffer, int startOffset, struct Hg0x6108GnssPosition *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6108) // GNSS position From Receiver Output Message
		{
			if ((*(UINT32*)(pb + 8))==26)
				*endOffset = startOffset - 1 + (*(UINT32*)(pb + 8)) * 4; //legacy message has 26 words // (-1) length is one based
			else
				*endOffset = startOffset + MSG_LEN_X6108 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;

	}

	/// <summary>Retrieve the 0x6109 GNSS Attitude Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6109Message(UINT8 *buffer, int startOffset, struct Hg0x6109GnssAttitude *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6109) // GNSS Attitude From Receiver Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6109 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6110 Distance from speed aiding from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6110Message(UINT8 *buffer, int startOffset, struct Hg0x6110DistanceTraveled *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6110) // GNSS Attitude From Receiver Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6110 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6111 Results of Motion detection algorithm in kalman filter Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6111Message(UINT8 *buffer, int startOffset, struct Hg0x6111MotionDetection *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6110) // GNSS Attitude From Receiver Output Message
		{
			*endOffset = startOffset + MSG_LEN_X6111 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6424 Message containing the Kalman filter estimates of Antenna Lever Amrs</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6424Message(UINT8 *buffer, int startOffset, struct Hg0x6424AntennaLeverArmEstimates *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6424) // Antenna Lever Arms Estimates
		{
			*endOffset = startOffset + MSG_LEN_X6424 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}

	/// <summary>Retrieve the 0x6438 Message containing the Kalman filter calibration Odometer input</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>(inherited from deserialize function)
	///				0 - OK
	///				-1 - incorrect start offset
	///				-2 - incorrect message type
	///				1 - Incorrect Checksum
	///				-3 - Incorrect start of Buffer
	///				-9 - Incorrect wordLength
	///				-10 - Incorrect Message Type</returns> 
	int Get0x6438Message(UINT8 *buffer, int startOffset, struct Hg0x6438OdometerCalibration *Message, int *endOffset)
	{
		UINT8 *pb = &buffer[startOffset];

		if (*(UINT32*)pb != 0xA5C381FF)
			return -3;

		if ((*(UINT32*)(pb + 4)) == 0x6438) // Odometer Calibration
		{
			*endOffset = startOffset + MSG_LEN_X6438 * 4;
			return Deserialize(buffer, startOffset, Message);
		}
		return -10;
	}
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x2402-Navigation-Output----------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2402NavigationOutput *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->InsGnssBIT = getInsGnssBIT(4 * 4, pb);

		// 5 Reserved 
		//Message->INSMode = pb[5 * 4]; // ||LEGACY||

		Message->GpsTov = *(double*)(pb + 6 * 4);
		Message->SystemTov = *(double*)(pb + 8 * 4);

		Message->GpsWeek = *(INT16*)(pb + 40);
		//Message->UtcTimeFom = *(INT16*)(pb + 42); // ||LEGACY||
		Message->GnssTimeFom = *(INT16*)(pb + 44);
		//Message->InsBlendedFom = *(INT16*)(pb + 46); // ||LEGACY||
		Message->INSMode = *(INT16*)(pb + 46);

		Message->PositionSystemTov = *(double*)(pb + 12 * 4);
		Message->Latitude = *(INT32*)(pb + 14 * 4) * 1.462918E-09f; // pi radians 
		Message->Longitude = *(INT32*)(pb + 15 * 4) * 1.462918E-09f; // pi radians
		Message->AltitudeElips = *(INT32*)(pb + 16 * 4) * 6.103516E-05f;
		Message->AltitudeGeoid = *(INT32*)(pb + 17 * 4) * 6.103516E-05f;

		Message->ECEFPosition[0] = *(INT32*)(pb + 18 * 4) * 0.0078125f;
		Message->ECEFPosition[1] = *(INT32*)(pb + 19 * 4) * 0.0078125f;
		Message->ECEFPosition[2] = *(INT32*)(pb + 20 * 4) * 0.0078125f;

		// 21 Reserved

		Message->VelocitySystemTov = *(double*)(pb + 22 * 4);
		Message->NEDVelocity[0] = *(INT32*)(pb + 24 * 4) * 7.629395E-06f;
		Message->NEDVelocity[1] = *(INT32*)(pb + 25 * 4) * 7.629395E-06f;
		Message->NEDVelocity[2] = *(INT32*)(pb + 26 * 4) * 7.629395E-06f;
		Message->ECEFVelocity[0] = *(INT32*)(pb + 27 * 4) * 7.629395E-06f;
		Message->ECEFVelocity[1] = *(INT32*)(pb + 28 * 4) * 7.629395E-06f;
		Message->ECEFVelocity[2] = *(INT32*)(pb + 29 * 4) * 7.629395E-06f;

		//Message->AttitudeTov = *(double*)(pb + 30 * 4); // ||LEGACY||
		// 30-31 Reserved 

		Message->VehicleEulerAngles[0] = *(float*)(pb + 32 * 4);
		Message->VehicleEulerAngles[1] = *(float*)(pb + 33 * 4);
		Message->VehicleEulerAngles[2] = *(float*)(pb + 34 * 4);

		Message->WanderAngle = *(float*)(pb + 35 * 4);

		Message->DCM[0][0] = *(INT32*)(pb + 36 * 4) * 4.656613E-10f;
		Message->DCM[0][1] = *(INT32*)(pb + 37 * 4) * 4.656613E-10f;
		Message->DCM[0][2] = *(INT32*)(pb + 38 * 4) * 4.656613E-10f;

		Message->DCM[1][0] = *(INT32*)(pb + 39 * 4) * 4.656613E-10f;
		Message->DCM[1][1] = *(INT32*)(pb + 40 * 4) * 4.656613E-10f;
		Message->DCM[1][2] = *(INT32*)(pb + 41 * 4) * 4.656613E-10f;

		Message->DCM[2][0] = *(INT32*)(pb + 42 * 4) * 4.656613E-10f;
		Message->DCM[2][1] = *(INT32*)(pb + 43 * 4) * 4.656613E-10f;
		Message->DCM[2][2] = *(INT32*)(pb + 44 * 4) * 4.656613E-10f;

		Message->VehicleBodyAngularRate[0] = *(float*)(pb + 45 * 4);
		Message->VehicleBodyAngularRate[1] = *(float*)(pb + 46 * 4);
		Message->VehicleBodyAngularRate[2] = *(float*)(pb + 47 * 4);

		Message->VehicleBodyAcceleration[0] = *(float*)(pb + 48 * 4);
		Message->VehicleBodyAcceleration[1] = *(float*)(pb + 49 * 4);
		Message->VehicleBodyAcceleration[2] = *(float*)(pb + 50 * 4);

		Message->AttitudeFom = *(INT32*)(pb + 51 * 4);

		Message->Quaternion[0] = *(float*)(pb + 52 * 4);
		Message->Quaternion[1] = *(float*)(pb + 53 * 4);
		Message->Quaternion[2] = *(float*)(pb + 54 * 4);
		Message->Quaternion[3] = *(float*)(pb + 55 * 4);

		// 56 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X2402)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x2001-INS-Configuration----------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2001Configuration *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);


		//Device Serial Number
		for (int i = 0; i < 8; i++)
		{
			Message->ImuSerialNumber[i] = pb[16 + i];
		}

		// 6-7 Reserved 
		//IMU SW Version
		for (int i = 0; i < 16; i++)
		{
			Message->ImuSwVersion[i] = pb[32 + i];
		}

		// 12-23 Reserved
		//HGuide SW Version
		for (int i = 0; i < 16; i++)
		{
			Message->HGuideSwVersion[i] = pb[24*4 + i];
		}
		//HGuide SW Build Date
		for (int i = 0; i < 16; i++)
		{
			Message->HGuideSwBuildDate[i] = pb[28 * 4 + i];
		}

		// 32-38 Reserved
		//HGuide Serial Number
		for (int i = 0; i < 8; i++)
		{
			Message->HGuideSerialNumber[i] = pb[39*4 + i];
		}
		//HGuide Part Number
		for (int i = 0; i < 8; i++)
		{
			Message->HGuidePartNumber[i] = pb[41*4 + i];
		}

		// 43-51 Reserved
		Message->VehicleLeverArms[0] = *(float*)(pb + 52 * 4);
		Message->VehicleLeverArms[1] = *(float*)(pb + 53 * 4);
		Message->VehicleLeverArms[2] = *(float*)(pb + 54 * 4);
		// 55-59 Reserved
		Message->MainAntennaLeverArms[0] = *(float*)(pb + 60 * 4);
		Message->MainAntennaLeverArms[1] = *(float*)(pb + 61 * 4);
		Message->MainAntennaLeverArms[2] = *(float*)(pb + 62 * 4);
		// 63-72 Reserved
		Message->AuxAntennaLeverArms[0] = *(float*)(pb + 73 * 4);
		Message->AuxAntennaLeverArms[1] = *(float*)(pb + 74 * 4);
		Message->AuxAntennaLeverArms[2] = *(float*)(pb + 75 * 4);
		// 76-79 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X2001)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x2011-INS-Status-----------------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2011Status *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		// 4-5 Reserved
		Message->NavigationFlag = *(UINT32*)(pb + 6 * 4);
		Message->InsGnssBIT = getInsGnssBIT(7 * 4, pb);
		Message->SystemTov = *(double*)(pb + 8 * 4);
		Message->GpsTov = *(double*)(pb + 10 * 4);
		Message->GpsWeek = *(INT32*)(pb + 12 * 4);
		Message->PowerCycleCount = *(INT32*)(pb + 13 * 4);
		Message->DeviceElapsedTime = *(double*)(pb + 14 * 4);
		Message->InsDeviceTemperature = *(float*)(pb + 16 * 4);

		// 17 Reserved

		Message->InsFom = *(INT32*)(pb + 18 * 4);
		Message->GnssFom = *(INT32*)(pb + 19 * 4);
		Message->UtcFom = *(INT32*)(pb + 20 * 4);
		// 21-22 Reserved
		//Enabled Messages
		//Output Group 1
		Message->EnabledMessages.MessageWord1 = *(INT32*)(pb + 23 * 4);
		Message->EnabledMessages.EnableX2001INSConfigurationOneShot = (pb[92] & 0x01) != 0;
		Message->EnabledMessages.EnableX2011INSModeStatusBIT1Hz = (pb[92] & 0x02) != 0;
		Message->EnabledMessages.EnableX2201TimeMark = (pb[92] & 0x80) != 0;
		Message->EnabledMessages.EnableX2311InertialDataOutputMessage100Hz = (pb[93] & 0x40) != 0;
		Message->EnabledMessages.SaveToFlash = (pb[93] & 0x80) != 0;
		Message->EnabledMessages.EnableX2402NavigationOutputMessage50Hz = (pb[94] & 0x10) != 0;
		Message->EnabledMessages.EnableDebugMessages = (pb[94] & 0x20) != 0; //can be extended to all messages
		//Output Group 2
		Message->EnabledMessages.MessageWord2 = *(INT32*)(pb + 24 * 4);
		Message->EnabledMessages.EnableX6108GNSSPositionFromReceiver = (pb[98] & 0x04) != 0;
		Message->EnabledMessages.EnableX6403GeodeticPosition = (pb[98] & 0x08) != 0;
		Message->EnabledMessages.EnableX6405EulerAttitude = (pb[98] & 0x10) != 0;
		Message->EnabledMessages.EnableX6504NEDVelocity = (pb[98] & 0x20) != 0;
		Message->EnabledMessages.EnableX6109GNSSAttitudeFromReceiver = (pb[98] & 0x40) != 0;
		Message->EnabledMessages.EnableX6110DistanceTraveled = (pb[98] & 0x80) != 0;

		Message->ImuBitStatus = (pb[100] & 0x01) != 0;

		// 26-27 Reserved for GNSS Status

		// 28-30 Reserved (TBD)? ||LEGACY||
		Message->NumberOfSatellitesUsed = *(INT32*)(pb + 28 * 4);
		Message->PseudoRangeValidity = *(UINT32*)(pb + 29 * 4);
		Message->DeltaRangeValidity = *(UINT32*)(pb + 30 * 4);
		// 31 Reserved
		Message->SolutionConvergence = *(UINT32*)(pb + 32 * 4); // 0 = Not Converged | 1 = Converged
		Message->AttitudeFom = *(INT32*)(pb + 33 * 4);
		//GNSS Aiding Status
		Message->GnssAidingStatus.BaroAidingValid = (pb[136] & 0x20) != 0;
		Message->GnssAidingStatus.BaroAidingUse = (pb[136] & 0x40) != 0;
		Message->GnssAidingStatus.MagAidingValid = (pb[137] & 0x02) != 0;
		Message->GnssAidingStatus.MagAidingUse = (pb[137] & 0x04) != 0;
		// 35-38 Reserved
		// 35-47 Reserved (TBD)? ||LEGACY||
		//INS Word 2 BIT Status 
		Message->InsWord2BitStatus.FirstStageBootLoader = (pb[156] & 0x01) != 0;
		Message->InsWord2BitStatus.FlashLoaderTable = (pb[156] & 0x02) != 0;
		Message->InsWord2BitStatus.RegisterInitializationTable = (pb[156] & 0x04) != 0;
		//GNSS Word 1 BIT Status
		Message->GnssWord1BitStatus.GNSSFunction = (pb[160] & 0x01) != 0;
		Message->GnssWord1BitStatus.GNSSCommunication = (pb[160] & 0x02) != 0;
		Message->GnssWord1BitStatus.GNSSTimeMark = (pb[160] & 0x04) != 0;
		Message->GnssWord1BitStatus.GNSST20Synchronization = (pb[160] & 0x08) != 0;
		// 41-47 Reserved
		Message->AccelTemperature = *(float*)(pb + 48 * 4);
		// 49-50 Reserved
		Message->GyroTemperature = *(float*)(pb + 51 * 4);
		// 52-54 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X2011)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x20FF-ACK------------------------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x20ffAck *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->Ack = *(UINT32*)(pb + 4 * 4);
		Message->InputMessageID = *(UINT32*)(pb + 5 * 4);
		Message->NoOfValidMessagesSinceLast = *(UINT32*)(pb + 6 * 4);
		Message->NoOfValidMessagesSincePowerUp = *(UINT32*)(pb + 7 * 4);
		Message->MessageTimeOfReception = *(double*)(pb + 8 * 4);
		// 10 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X20FF)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x2201-Time-Mark-/-PPS-Data-------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2201TimeMark *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 8);
		UINT32 Checksum = *(UINT32*)(pb + 12);

		Message->TimeValidityBits.GnssTime = (pb[16] & 0x01) != 0;
		Message->TimeValidityBits.UtcTime = (pb[16] & 0x02) != 0;

		Message->TimeValidityBits.HardwarePulse = (pb[16] & 0x70) >> 4;

		Message->UtcTimeFom = *(UINT32*)(pb + 20);
		Message->EventInSystemTov = *(double*)(pb + 24);
		Message->EventInGpsTov = *(double*)(pb + 32);
		Message->GpsWeek = *(INT32*)(pb + 40);
		// 11-14 Reserved
		Message->PpsSystemTov = *(double*)(pb + 15*4);
		Message->PpsGpsTov = *(double*)(pb + 17*4);
		Message->EventInCount = *(INT32*)(pb + 19*4);
		// 20-23 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X2201)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x2311-Inertial-Measurement-Unit-Data---------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x2311InertialData *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);

		//Message->InsGnssBIT = getInsGnssBIT(24, pb);

		Message->DeltaTheta[0] = *(INT32*)(pb + 7 * 4)* 1.164153E-10f;
		Message->DeltaTheta[1] = *(INT32*)(pb + 8 * 4)* 1.164153E-10f;
		Message->DeltaTheta[2] = *(INT32*)(pb + 9 * 4)* 1.164153E-10f;
		// 10-11 Reserved
		Message->DeltaVelocity[0] = *(INT32*)(pb + 12 * 4)* 1.862645E-09f;
		Message->DeltaVelocity[1] = *(INT32*)(pb + 13 * 4) * 1.862645E-09f;
		Message->DeltaVelocity[2] = *(INT32*)(pb + 14 * 4)* 1.862645E-09f;
		// 15-16 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X2311)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6403-INS-Geodetic-Position-Only-------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6403GeodeticPosition *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->GpsTov = *(double*)(pb + 6 * 4);

		Message->Latitude = *(double*)(pb + 8 * 4);
		Message->Longitude = *(double*)(pb + 10 * 4);
		Message->AltitudeAboveEllipsoid = *(double*)(pb + 12 * 4);

		Message->InsGnssBIT = getInsGnssBIT(56, pb);

		Message->StdvNED[0] = *(float*)(pb + 15 * 4);
		Message->StdvNED[1] = *(float*)(pb + 16 * 4);
		Message->StdvNED[2] = *(float*)(pb + 17 * 4);

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6403)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6504-INS-NED-Velocity-----------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6504NEDVelocity *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->GpsTov = *(double*)(pb + 6 * 4);

		Message->VelocityNED[0] = *(float*)(pb + 8 * 4);
		Message->VelocityNED[1] = *(float*)(pb + 9 * 4);
		Message->VelocityNED[2] = *(float*)(pb + 10 * 4);

		Message->InsGnssBIT = getInsGnssBIT(44, pb);

		Message->VelocityStdvNED[0] = *(float*)(pb + 12 * 4);
		Message->VelocityStdvNED[1] = *(float*)(pb + 13 * 4);
		Message->VelocityStdvNED[2] = *(float*)(pb + 14 * 4);

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6504)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6405-INS-Euler-Attitudes--------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6405EulerAttitudes *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->GpsTov = *(double*)(pb + 6 * 4);

		Message->EulerAttitude[0] = *(float*)(pb + 8 * 4);
		Message->EulerAttitude[1] = *(float*)(pb + 9 * 4);
		Message->EulerAttitude[2] = *(float*)(pb + 10 * 4);

		Message->InsGnssBIT = getInsGnssBIT(44, pb);

		Message->EulerAttitudeStdv[0] = *(float*)(pb + 12 * 4);
		Message->EulerAttitudeStdv[1] = *(float*)(pb + 13 * 4);
		Message->EulerAttitudeStdv[2] = *(float*)(pb + 14 * 4);

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6405)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6108-GNSS-Position-From-Receiver------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6108GnssPosition *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		if (WordLength == MSG_LEN_X6108)
		{
			Message->GpsTov = *(double*)(pb + 4 * 4);
			Message->GpsWeek = *(INT32*)(pb + 6 * 4);

			//Message->InsGnssBIT = getInsGnssBIT(28, pb);

			Message->Latitude = *(double*)(pb + 8 * 4);
			Message->Longitude = *(double*)(pb + 10 * 4);
			Message->AltitudeAboveEllipsoid = *(double*)(pb + 12 * 4);

			Message->VelocityNED[0] = *(float*)(pb + 14 * 4);
			Message->VelocityNED[1] = *(float*)(pb + 15 * 4);
			Message->VelocityNED[2] = *(float*)(pb + 16 * 4);

			Message->RxClkBias = *(float*)(pb + 17 * 4);
			Message->PVT_comp = *(UINT32*)(pb + 18 * 4);
			Message->corr_info = *(UINT32*)(pb + 19 * 4);
			Message->signal_information = *(UINT32*)(pb + 20 * 4);
			Message->PPP_info = *(UINT32*)(pb + 21 * 4);

			Message->GnssStdvLat = *(float*)(pb + 22 * 4);
			Message->GnssStdvLon = *(float*)(pb + 23 * 4);
			Message->GnssStdvAlt = *(float*)(pb + 24 * 4);
			Message->GnssStdvNEDVelocity[0] = *(float*)(pb + 25 * 4);
			Message->GnssStdvNEDVelocity[1] = *(float*)(pb + 26 * 4);
			Message->GnssStdvNEDVelocity[2] = *(float*)(pb + 27 * 4);
			Message->SystemTov = *(double*)(pb + 28 * 4);
			// 30-31 Reserved
		}
		else // ||LEGACY|| 26 word message length
		{
			Message->GpsTov = *(double*)(pb + 4 * 4);
			Message->GpsWeek = *(INT32*)(pb + 6 * 4);

			//Message->InsGnssBIT.GpsMode = *(INT32*)(pb + 7 * 4);

			Message->Latitude = *(double*)(pb + 8 * 4);
			Message->Longitude = *(double*)(pb + 10 * 4);
			Message->AltitudeAboveEllipsoid = *(double*)(pb + 12 * 4);

			Message->VelocityNED[0] = *(float*)(pb + 14 * 4);
			Message->VelocityNED[1] = *(float*)(pb + 15 * 4);
			Message->VelocityNED[2] = *(float*)(pb + 16 * 4);

			Message->RxClkBias = *(float*)(pb + 17 * 4);
			Message->PVT_comp = *(UINT32*)(pb + 18 * 4);
			Message->corr_info = *(UINT32*)(pb + 19 * 4);
			Message->signal_information = *(UINT32*)(pb + 20 * 4);
			Message->PPP_info = *(UINT32*)(pb + 21 * 4);

			//Not present in legacy message
			//Message->InsGnssBIT.InsMode = 0;
			//Message->InsGnssBIT.InsStatus = 0;
			//Message->InsGnssBIT.ImuStatus = 0;
			//Message->InsGnssBIT.GnssStatus = 0;
			Message->GnssStdvLat =0;
			Message->GnssStdvLon =0;
			Message->GnssStdvAlt = 0;
			Message->GnssStdvNEDVelocity[0] = 0;
			Message->GnssStdvNEDVelocity[1] = 0;
			Message->GnssStdvNEDVelocity[2] = 0;
			Message->SystemTov =0;
		}

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6108 || WordLength == 26)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6109-GNSS-Attitude-From-Receiver------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6109GnssAttitude *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->GpsTov = *(double*)(pb + 6 * 4);

		// 8 Reserved
		Message->GnssAttitude[0] = *(float*)(pb + 9 * 4);
		Message->GnssAttitude[1] = *(float*)(pb + 10 * 4);
		// 11-12 Reserved
		Message->GnssAttitudeStdv[0] = *(float*)(pb + 13 * 4);
		Message->GnssAttitudeStdv[1] = *(float*)(pb + 14 * 4);
		Message->HeadingValid = *(UINT32*)(pb + 15 * 4);

		//Message->InsGnssBIT = getInsGnssBIT(64, pb);

		// 17-18 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6109)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}	

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6110-Vehicle-Velocity-as-inputted-by-Speed-aiding--------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6110DistanceTraveled *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->GpsTov = *(double*)(pb + 6 * 4);

		Message->DistanceTraveled[0] = *(float*)(pb + 8 * 4);
		Message->DistanceTraveled[1] = *(float*)(pb + 9 * 4);
		Message->DistanceTraveled[2] = *(float*)(pb + 10 * 4);
		// 11-13 Reserved
		Message->OdometerCumulativePulses = *(UINT32*)(pb + 14 * 4);

		Message->OdometerSettings.VehicleVelocityValid = *(UINT8*)(pb + 15 * 4) & 0x01;
		Message->OdometerSettings.OdometerValid = (*(UINT8*)(pb + 15 * 4) & 0x02) >> 1;
		Message->OdometerSettings.TovMode = (*(UINT8*)(pb + 15 * 4) & 0x04) >> 2;
		Message->OdometerSettings.AidingStatus = (*(UINT8*)(pb + 15 * 4) & 0x08) >> 3;
		Message->OdometerSettings.ZeroVelocityDetected= (*(UINT8*)(pb + 15 * 4) & 0x10) >> 4;
		
		Message->Counter = *(UINT32*)(pb + 16 * 4);
		// 17 - 23 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6110)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6111-Motion-Detection-Algorithm-Results-From-Kalman-Filter-----------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6111MotionDetection *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->SystemTov = *(double*)(pb + 4 * 4);
		Message->TriggerTime = *(double*)(pb + 6 * 4);

		//Message->InsGnssBIT = getInsGnssBIT(8 * 4, pb);

		Message->Latitude = *(double*)(pb + 9 * 4);
		Message->Longitude = *(double*)(pb + 11 * 4);

		Message->MD1Rotation[0] = *(float*)(pb + 13 * 4);
		Message->MD1Rotation[1] = *(float*)(pb + 14 * 4);
		Message->MD1Rotation[2] = *(float*)(pb + 15 * 4);
		Message->MDD1AngularRateTotal = *(float*)(pb + 16 * 4);
		Message->MDT1Rotation = *(float*)(pb + 17 * 4);

		Message->MD2NavigationValid = *(float*)(pb + 18 * 4);

		Message->MDD2SpeedStdv = *(float*)(pb + 19 * 4);
		Message->MDT2SpeedStdv = *(float*)(pb + 20 * 4);

		Message->MD3AngularRateInstant[0] = *(float*)(pb + 21 * 4);
		Message->MD3AngularRateInstant[1] = *(float*)(pb + 22 * 4);
		Message->MD3AngularRateInstant[2] = *(float*)(pb + 23 * 4);
		Message->MD3InstantFN3dB = *(float*)(pb + 24 * 4);
		Message->MD3AngularRateNominal[0] = *(float*)(pb + 25 * 4);
		Message->MD3AngularRateNominal[1] = *(float*)(pb + 26 * 4);
		Message->MD3AngularRateNominal[2] = *(float*)(pb + 27 * 4);
		Message->MD3NominalFN3dB = *(float*)(pb + 28 * 4);
		Message->MDD3AngularRate[0] = *(float*)(pb + 29 * 4);
		Message->MDD3AngularRate[1] = *(float*)(pb + 30 * 4);
		Message->MDD3AngularRate[2] = *(float*)(pb + 31 * 4);
		Message->MDT3AngularRateInstant = *(float*)(pb + 32 * 4);

		Message->MDD4LinearAcceleration = *(float*)(pb + 33 * 4);
		Message->MDT4LinearAcceleration = *(float*)(pb + 34 * 4);

		Message->MDD5OdometerDeltaDistance = *(float*)(pb + 35 * 4);
		Message->MDT5Odometer = *(float*)(pb + 36 * 4);
		Message->MDOdometerTimeAtRest = *(float*)(pb + 37 * 4);

		Message->MDTimeStationary = *(float*)(pb + 38 * 4);
		Message->MDTSettlingTime = *(float*)(pb + 39 * 4);

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6111)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6424-Antenna-Lever-Arms-Estimates-As-Reported-By-Kalman-Filter-------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6424AntennaLeverArmEstimates *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		// 4 - 5 Reserved
		Message->SystemTov = *(double*)(pb + 6 * 4);
		Message->GpsTov = *(double*)(pb + 8 * 4);
		Message->GpsWeek = *(UINT32*)(pb + 10 * 4);
		// 11 - 21 Reserved
		Message->MainAntennaLeverArm[0] = *(float*)(pb + 22 * 4);
		Message->MainAntennaLeverArm[1] = *(float*)(pb + 23 * 4);
		Message->MainAntennaLeverArm[2] = *(float*)(pb + 24 * 4);

		Message->AntennaBoresight[0] = *(float*)(pb + 25 * 4);
		Message->AntennaBoresight[1] = *(float*)(pb + 26 * 4);
		Message->AntennaBoresight[2] = *(float*)(pb + 27 * 4);
		Message->AntennaBoresightStdv[0] = *(float*)(pb + 28 * 4);
		Message->AntennaBoresightStdv[1] = *(float*)(pb + 29 * 4);
		Message->AntennaBoresightStdv[2] = *(float*)(pb + 30 * 4);
		// 31 - 35 Reserved
		Message->MainAntennaLeverArmStdv[0] = *(float*)(pb + 36 * 4);
		Message->MainAntennaLeverArmStdv[1] = *(float*)(pb + 37 * 4);
		Message->MainAntennaLeverArmStdv[2] = *(float*)(pb + 38 * 4);
		// 39 - 41 Reserved

		//Calculate Checksum
		if (WordLength == MSG_LEN_X6424)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-0x6438-Kalman-Filter-Calibration-Of-Odometer-Inputs--------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/
	int Deserialize(UINT8 *buffer, int startOffset, struct Hg0x6438OdometerCalibration *Message)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);
		// 4 Reserved
		// 5 - Only INS Mode is filled

		Message->SystemTov = *(double*)(pb + 6 * 4);
		Message->GpsTov = *(double*)(pb + 8 * 4);
		Message->GpsWeek = *(UINT32*)(pb + 10 * 4);
		Message->DistanceTraveled = *(float*)(pb + 11 * 4);
		// 12 - 13 Reserved
		Message->ScaleFactorCorrection = *(float*)(pb + 14 * 4);
		// 15 - 16 Reserved
		Message->Boresight[0] = *(float*)(pb + 17 * 4);
		// 18 - 19 Reserved
		Message->Boresight[1] = *(float*)(pb + 20 * 4);
		// 21 - 22 Reserved
		Message->LeverArms[0] = *(float*)(pb + 23 * 4);
		Message->LeverArms[1] = *(float*)(pb + 24 * 4);
		Message->LeverArms[2] = *(float*)(pb + 25 * 4);
		Message->ScaleFactorStdv = *(float*)(pb + 26 * 4);
		// 27 - 28 Reserved
		Message->BoresightStdv[0] = *(float*)(pb + 29 * 4);
		// 30 - 31 Reserved
		Message->BoresightStdv[1] = *(float*)(pb + 32 * 4);
		// 33 - 34 Reserved
		Message->LeverArmsStdv[0] = *(float*)(pb + 35 * 4);
		Message->LeverArmsStdv[1] = *(float*)(pb + 36 * 4);
		Message->LeverArmsStdv[2] = *(float*)(pb + 37 * 4);
		Message->StoredLeverArms[0] = *(float*)(pb + 38 * 4);
		Message->StoredLeverArms[1] = *(float*)(pb + 39 * 4);
		Message->StoredLeverArms[2] = *(float*)(pb + 40 * 4);
		// 41 - 47 Reserved


		//Calculate Checksum
		if (WordLength == MSG_LEN_X6438)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*--INPUT-MESSAGES-------------------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/


	/// <summary>Calculate HG INS Checksum</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containing the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="byteLength">Message Word Length (byte length / 4)</param>  
	/// <returns>Calculated Checksum</returns>  
	UINT32 CalcHgInsChecksum(UINT8 *buffer, int startOffset, int wordLength)
	{

		UINT32 u32sum = 0;
		UINT8* pb = &buffer[startOffset];

		if (wordLength < 0)
			return 1;

		for (int i = 0; i < wordLength * 4; i = i + 4)
		{
			if (i != 12)
				u32sum += *(UINT32 *)(pb + i);
		}

		//calculate 2's complement
		u32sum = 0 - (INT32)u32sum;

		return u32sum;
	}


	/// <summary>Create 0x1001 Enable Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1001Message(UINT8 *buffer, int startOffset, struct Hg0x1001EnableInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X1001;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1001; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum
		// Output Group 1
		*(UINT32*)(pb + 16) = Message.Messages.MessageWord1|((int)Message.Messages.EnableX2001INSConfigurationOneShot |
			(int)Message.Messages.EnableX2011INSModeStatusBIT1Hz << 1 |
			(int)Message.Messages.EnableX2201TimeMark << 8 |
			(int)Message.Messages.EnableX2311InertialDataOutputMessage100Hz << 14 |
			(int)Message.Messages.SaveToFlash << 15 |
			(int)Message.Messages.EnableX2402NavigationOutputMessage50Hz << 20);
		if (Message.Messages.EnableDebugMessages) // Enable 0x2108, 0x2422
			*(UINT32*)(pb + 16) |= 0x01 << 5 | 0x01 << 26;
		// Output Group 2
		*(UINT32*)(pb + 20) = Message.Messages.MessageWord2 | (Message.Messages.EnableX6108GNSSPositionFromReceiver<<18|
			Message.Messages.EnableX6403GeodeticPosition<<19|
			Message.Messages.EnableX6405EulerAttitude<<20|
			Message.Messages.EnableX6504NEDVelocity<<21|
			Message.Messages.EnableX6109GNSSAttitudeFromReceiver<<22|
			Message.Messages.EnableX6110DistanceTraveled<<23);
		if (Message.Messages.EnableDebugMessages) // Enable 0x6438, 0x6428, 0x6424, 0x2501
		*(UINT32*)(pb + 20) |= 0x01 << 15 | 0x01 << 14 | 0x01 << 13 | 0x01 << 8;
		
		// 6 - 15 Reserved
		
		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	} 

	/// <summary>Create 0x1001 Select Navigation Mode Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1002Message(UINT8 *buffer, int startOffset, struct Hg0x1002NavigationModeInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X1002;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1002; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum
		// Output Group 1
		*(UINT32*)(pb + 16) = Message.INSMode;
		// 5 - 8 Reserved
		*(UINT32*)(pb + 36) = Message.TimeMarkPpsMode;
		// 10 Reserved
		*(float*)(pb + 44) = Message.CoarseLevelDuration;
		*(UINT32*)(pb + 48) = (Message.NavAidingSourcesEn.GnssPvtVelocity << 5 |
			Message.NavAidingSourcesEn.GnssPvtPosition <<6 |
			Message.NavAidingSourcesEn.ZeroVelocity << 11 |
			Message.NavAidingSourcesEn.ZeroHeadingChange << 12 |
			Message.NavAidingSourcesEn.BarometricAltitude << 13 |
			Message.NavAidingSourcesEn.MagneticHeading << 16 |
			Message.NavAidingSourcesEn.AidingSourcesEnable << 31);
		// 13 - 29 Reserved
		*(float*)(pb + 120) = Message.ZeroVelocityStdv;
		*(float*)(pb + 124) = Message.ZeroHeadingStdv;
		// 32 Reserved
		*(float*)(pb + 33 * 4) = Message.MDTSettlingTime;
		*(float*)(pb + 34 * 4) = Message.MDT1AngularRate;
		*(float*)(pb + 35 * 4) = Message.MD3NominalFN3dB;
		*(float*)(pb + 36 * 4) = Message.MD3InstantFN3dB;
		*(float*)(pb + 37 * 4) = Message.MDT3AngularRateInstant;
		*(float*)(pb + 38 * 4) = Message.MDT4LinearAcceleration;
		*(float*)(pb + 39 * 4) = Message.MDT2SpeedStdv;
		// 40 - 45 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x1004 Configuration Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1004Message(UINT8 *buffer, int startOffset, struct Hg0x1004ConfigurationInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X1004;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1004; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		 // Word 3 Reserved for Checksum
		 
		 // TBD - 2 bits ?!
		*(UINT32*)(pb + 16) = ((Message.UpdateConfig.VehicleFrameToCaseFrameTransformation & 0x03) << 4 |
			(Message.UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame & 0x03)<< 6 |
			(Message.UpdateConfig.MainAntennaLeverArm & 0x03)<< 8);
		// 5 - 8 Reserved
		*(INT32*)(pb + 36) = (INT32)(Message.VehicleEulerAngles[0] / (3.14159*4.6566e-10)); // LSB = pi*2-31
		*(INT32*)(pb + 40) = (INT32)(Message.VehicleEulerAngles[1] / (3.14159*4.6566e-10));
		*(INT32*)(pb + 44) = (INT32)(Message.VehicleEulerAngles[2] / (3.14159*4.6566e-10));
		// 12 - 17 Reserved
		*(float*)(pb + 72) = Message.VehicleLeverArms[0];
		*(float*)(pb + 76) = Message.VehicleLeverArms[1];
		*(float*)(pb + 80) = Message.VehicleLeverArms[2];
		*(float*)(pb + 84) = Message.MainAntennaLeverArms[0];
		*(float*)(pb + 88) = Message.MainAntennaLeverArms[1];
		*(float*)(pb + 92) = Message.MainAntennaLeverArms[2];
		// 22 - 44 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x1101 Barometric Altitude Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1101Message(UINT8 *buffer, int startOffset, struct Hg0x1101BarometricAltitudeInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X1101;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1101; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(double*)(pb + 16) = Message.BarometricAltitudeTov;
		*(UINT32*)(pb + 24) = ((INT32)Message.BarometricAltitudeValid | (INT32)Message.TovMode << 1);
		*(float*)(pb + 28) = Message.BarometricAltitudeMslGeoid;
		// 8 - 9 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x1105 Magnetic Heading Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1105Message(UINT8 *buffer, int startOffset, struct Hg0x1105MagneticHeading Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X1105;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1105; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(double*)(pb + 16) = Message.MagneticHeadingTov;
		*(UINT32*)(pb + 24) = ((INT32)Message.MagneticHeadingValid | (INT32)Message.TovMode << 1 | (INT32)Message.MagneticVariationValid << 2);
		*(float *)(pb + 28) = Message.MagneticHeading;
		*(float *)(pb + 32) = Message.MagneticVariation;
		// 9 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x1401 Navigation Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x1401Message(UINT8 *buffer, int startOffset, struct Hg0x1401NavigationInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;
		UINT32 MsgLength = MSG_LEN_X1401;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x1401; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(UINT32*)(pb + 16) = Message.AckRequested;
		// 5 Reserved
		*(double*)(pb + 24) = Message.PositionTov;
		*(UINT32*)(pb + 32) = ((INT32)Message.PositionTovSettings.Valid | (INT32)Message.PositionTovSettings.TovMode << 1 | (INT32)Message.PositionTovSettings.SetFrame << 2 | (INT32)Message.PositionTovSettings.StdvValid << 3);
		*(INT32*)(pb + 36) = (INT32) (Message.Latitude / 1.462918E-09f); // pi radians
		*(INT32*)(pb + 40) = (INT32) (Message.Longitude / 1.462918E-09f); // pi radians
		*(INT32*)(pb + 44) = (INT32) (Message.AltitudeAboveElipsoid / 6.103516E-05f); // meters

		*(double*)(pb + 48) = Message.VelocityTov;
		*(UINT32*)(pb + 56) = ((INT32)Message.VelocityTovSettings.Valid | (INT32)Message.VelocityTovSettings.TovMode << 1 | (INT32)Message.VelocityTovSettings.SetFrame << 2 | (INT32)Message.VelocityTovSettings.StdvValid << 3);
		*(INT32*)(pb + 60) = (INT32)(Message.NEDVelocity[0] / 7.629395E-06f);
		*(INT32*)(pb + 64) = (INT32)(Message.NEDVelocity[1] / 7.629395E-06f);
		*(INT32*)(pb + 68) = (INT32)(Message.NEDVelocity[2] / 7.629395E-06f);
		
		*(double*)(pb + 72) = Message.AttitudeTov;
		*(UINT32*)(pb + 80) = ((INT32)Message.AttitudeTovSettings.Valid | (INT32)Message.AttitudeTovSettings.TovMode << 1 | (INT32)Message.AttitudeTovSettings.SetFrame << 2 | (INT32)Message.AttitudeTovSettings.StdvValid << 3);
		*(INT32*)(pb + 84) = (INT32)(Message.VehicleEulerAngles[0] / 4.656613E-10f); // 1.462918E-09f; // TBD difference from manual - pi radians 
		*(INT32*)(pb + 88) = (INT32)(Message.VehicleEulerAngles[1] / 4.656613E-10f); //1.462918E-09f; // pi radians
		*(INT32*)(pb + 92) = (INT32)(Message.VehicleEulerAngles[2] / 4.656613E-10f); //1.462918E-09f; // pi radians
		// 24 - 29 Reserved
		*(INT32*)(pb + 120) = (INT32)(Message.NEDPositionStdv[0] / 0.0078125f);
		*(INT32*)(pb + 124) = (INT32)(Message.NEDPositionStdv[1] / 0.0078125f);
		*(INT32*)(pb + 128) = (INT32)(Message.NEDPositionStdv[2] / 0.0078125f);
		*(INT32*)(pb + 132) = (INT32)(Message.NEDVelocityStdv[0] / 7.629395E-06f);
		*(INT32*)(pb + 136) = (INT32)(Message.NEDVelocityStdv[1] / 7.629395E-06f);
		*(INT32*)(pb + 140) = (INT32)(Message.NEDVelocityStdv[2] / 7.629395E-06f);
		*(INT32*)(pb + 144) = (INT32)(Message.EulerAnglesStdv[0] / 4.656613E-10f); //1.462918E-09f; // TBD difference from manual - pi radians 
		*(INT32*)(pb + 148) = (INT32)(Message.EulerAnglesStdv[1] / 4.656613E-10f); //1.462918E-09f; // pi radians
		*(INT32*)(pb + 152) = (INT32)(Message.EulerAnglesStdv[2] / 4.656613E-10f); //1.462918E-09f; // pi radians

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x4204 Antenna Lever Arm Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x4204Message(UINT8 *buffer, int startOffset, struct Hg0x4204AntennaLeverArmInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X4204;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x4204; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(float*)(pb + 16) = Message.MainAntennaLeverArms[0];
		*(float*)(pb + 20) = Message.MainAntennaLeverArms[1];
		*(float*)(pb + 24) = Message.MainAntennaLeverArms[2];

		*(float*)(pb + 28) = Message.AuxAntennaLeverArms[0];
		*(float*)(pb + 32) = Message.AuxAntennaLeverArms[1];
		*(float*)(pb + 36) = Message.AuxAntennaLeverArms[2];

		*(float*)(pb + 40) = Message.MainAntennaLAUncertainty;
		*(float*)(pb + 44) = Message.AuxAntennaLAUncertainty;

		*(UINT32*)(pb + 48) = ((UINT32)Message.LAStoreToFlash.InputMainLA  | (UINT32)Message.LAStoreToFlash.InputAuxLA << 1 | (UINT32)Message.LAStoreToFlash.InputAuxLAStdv << 2 | (UINT32)Message.LAStoreToFlash.InputAuxLAStdv << 3);

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x4404 Vehicle Frame Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x4404Message(UINT8 *buffer, int startOffset, struct Hg0x4404VehicleFrameInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X4404;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x4404; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(float*)(pb + 16) = Message.VehicleLeverArms[0];
		*(float*)(pb + 20) = Message.VehicleLeverArms[1];
		*(float*)(pb + 24) = Message.VehicleLeverArms[2];

		*(float*)(pb + 28) = Message.VehicleEulerAngles[0];
		*(float*)(pb + 32) = Message.VehicleEulerAngles[1];
		*(float*)(pb + 36) = Message.VehicleEulerAngles[2];
		// 10 - 11 Reserved

		*(UINT32*)(pb + 48) = (Message.StoreToFlashEulerAngles<<4|Message.StoreToFlashVehicleLA<<5);

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Create 0x4110 Vehicle Speed Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x4110Message(UINT8 *buffer, int startOffset, struct Hg0x4110VelocityAidingInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X4110;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x4110; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
		// Word 3 Reserved for Checksum

		*(double*)(pb + 16) = Message.OdometerTimeDelay;
		*(double*)(pb + 24) = Message.OdometerTov;
		*(float*)(pb + 32) = Message.VehicleVelocity[0];
		*(float*)(pb + 36) = Message.VehicleVelocity[1];
		*(float*)(pb + 40) = Message.VehicleVelocity[2];

		*(INT32*)(pb + 44) = Message.OdometerCumulativePulses;
		*(float*)(pb + 48) = Message.DistancePerPulse;
		// 13 - 14 Reserved
		*(UINT32*)(pb + 60) = (int)Message.OdometerSettings.VehicleVelocityValid | (int)Message.OdometerSettings.OdometerValid << 1 | (int)Message.OdometerSettings.TovMode << 2 | (int)Message.OdometerSettings.AidingStatus << 3 | (int)Message.OdometerSettings.ZeroVelocityDetected << 4;
		*(UINT32*)(pb + 64) = Message.Counter;
		// 17 - 23 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/// <summary>Retrieve the 0x4110 Euler Attitudes Message from byte array</summary>
	/// <param name="buffer">Buffer containing binary/hex data</param>  
	/// <param name="startOffset">Byte containg the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">0x4110 meesage container structure</param>  
	/// <param name="endOffset">bytes processed + initial offset</param>  
	/// <returns>0 - OK
	///			-1 - incorrect start offset
	///			-2 - incorrect message type
	///			1 - Incorrect Checksum
	///			-3 - Incorrect start of Buffer
	///			-9 - Incorrect wordLength
	///			-10 - Incorrect Message Type</returns>  
	int Get0x4110Message(UINT8 *buffer, int startOffset, struct Hg0x4110VelocityAidingInput * Message, int * endOffset)
	{
		if (startOffset < 0)
			return -1;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		// 0 Header
		// 1 Message Type
		UINT32 WordLength = *(UINT32*)(pb + 2 * 4);
		UINT32 Checksum = *(UINT32*)(pb + 3 * 4);

		Message->OdometerTimeDelay = *(double*)(pb + 4 * 4);
		Message->OdometerTov = *(double*)(pb + 6 * 4);

		Message->VehicleVelocity[0] = *(float*)(pb + 8 * 4);
		Message->VehicleVelocity[1] = *(float*)(pb + 9 * 4);
		Message->VehicleVelocity[2] = *(float*)(pb + 10 * 4);
		
		Message->OdometerCumulativePulses = *(INT32*)(pb + 11 * 4);
		Message->DistancePerPulse = *(float*)(pb + 12 * 4);
		// 13-14 Reserved
		Message->OdometerSettings.VehicleVelocityValid = *(UINT8*)(pb + 15 * 4) & 0x01;
		Message->OdometerSettings.OdometerValid = (*(UINT8*)(pb + 15 * 4) & 0x02) >> 1;
		Message->OdometerSettings.TovMode = (*(UINT8*)(pb + 15 * 4) & 0x04) >> 2;
		Message->OdometerSettings.AidingStatus = (*(UINT8*)(pb + 15 * 4) & 0x08) >> 3;
		Message->OdometerSettings.ZeroVelocityDetected = (*(UINT8*)(pb + 15 * 4) & 0x10) >> 4;

		Message->Counter = *(UINT32*)(pb + 16 * 4);
		// 17 - 23 Reserved
		*endOffset = startOffset + MSG_LEN_X4110*4;
		//Calculate Checksum
		if (WordLength == MSG_LEN_X4110)
			return HgInsChecksum(buffer, startOffset, WordLength);
		else
			return -9;
	}

	/// <summary>Create 0x4438 Odometer Configuration Input Message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of Start Of Message Word (0 based)</param>  
	/// <param name="Message">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - TBD</returns>  
	int Create0x4438Message(UINT8 *buffer, int startOffset, Hg0x4438OdometerConfigurationInput Message, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		UINT32 MsgLength = MSG_LEN_X4438;
		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		//Zero the buffer
		for (UINT32 i = 0; i < MsgLength; i++)
			*(UINT32*)(pb + i * 4) = 0;

		*(UINT32*)pb = 0xA5C381FF; // Start of Message
		*(UINT32*)(pb + 4) = 0x4438; // Message ID
		*(UINT32*)(pb + 8) = MsgLength; // Message Length
										// Word 3 Reserved for Checksum

		// 4 - 7 Reserved
		*(float*)(pb + 8 * 4) = Message.LeverArms[0];
		*(float*)(pb + 9 * 4) = Message.LeverArms[1];
		*(float*)(pb + 10 * 4) = Message.LeverArms[2];
		*(float*)(pb + 11 * 4) = Message.MeasurementNoise;
		*(float*)(pb + 12 * 4) = Message.Threshold;
		*(float*)(pb + 13 * 4) = Message.ScaleFactorUncertainty;
		*(float*)(pb + 14 * 4) = Message.ScaleFactorProcessNoise;
		*(float*)(pb + 15 * 4) = Message.YawBoresightStdv;
		*(float*)(pb + 16 * 4) = Message.YawBoresightProcessNoise;
		*(float*)(pb + 17 * 4) = Message.PitchBoresightUncertainty;
		*(float*)(pb + 18 * 4) = Message.PitchBoresightProcessNoise;
		*(UINT8*)(pb + 19 * 4) = (int) Message.SaveToFlash;
		// 20 - 23 Reserved

		*byteLength = *(UINT32*)(pb + 8) * 4;
		//Message Checksum
		*(UINT32*)(pb + 12) = CalcHgInsChecksum(buffer, startOffset, MsgLength);
		return 0;
	}

	/*-----------------------------------------------------------------------------------------------------------------------------*/
	/*-CREATE-NMEA-OUTPUT-MESSAGES-------------------------------------------------------------------------------------------------*/
	/*-----------------------------------------------------------------------------------------------------------------------------*/

	/// <summary>Calculate NMEA Checksum</summary>
	/// <param name="buffer">Buffer with the stored message, including '$' and '*' symbos</param>  
	/// <param name="startOffset">Byte which contain the start of the message '$' - can be 0 as the function is searching for it</param>  
	/// <returns>0 - calculation success | -1 - haven't found '$' or '*' before 100th byte</returns>  
	char NmeaChecksum(UINT8* buffer, int startOffset)
	{
		char checksum = 0;
		if (startOffset < 0)
			return 0x00;

		int status = 0;
		UINT8 *pb = &buffer[startOffset];

		int i = 0;

		while (pb[i] != '$' && i <= 100)  // look for NMEA Start
		{
			i++;
		}
		i++;

		while (pb[i] != '*' && i <= 100)  // the message is never longer
		{
			checksum ^= (char)pb[i];
			i++;
		}

		if (i > 100)
			return 0x00;
		else
			return checksum;

	}

	/// <summary>Calculate UTC time from GPS time of week</summary>
	/// <param name="gpsTov">GPS time of week</param>   
	/// <returns>UTC time in HHMMSS format</returns>  
	double gpsTovToUtcTov(double gpsTov)
	{
		gpsTov += 18; //Leap Seconds

		double hour = std::fmod(gpsTov, 86400) / 3600;
		double minute = std::fmod(hour, 1) * 60;
		double second = std::fmod(minute, 1) * 60;

		return floor(hour) * 10000 + floor(minute) * 100 + second;
	}

	/// <summary>Fill NMEA latitude and longitude data</summary>
	/// <param name="lat_f">Latitude decimal</param>  
	/// <param name="lat_c">Latitude indicator</param>  
	/// <param name="lon_f">Longitude decimal</param>  
	/// <param name="lon_c">Longitude indicator</param> 
	/// <param name="lat">Latitude in Radians</param>  
	/// <param name="lon">Longitude in Radians</param> 
	/// <returns>0 - TBD</returns>  
	int getNmeaLatLon(double * lat_f, char * lat_c, double * lon_f, char * lon_c, double lat, double lon)
	{
		//Recalculate Lat/Lon to NMEA type
		int lat_d = (int)floor(lat*RAD_TO_DEG);
		double lat_m = std::fmod(std::abs(lat*RAD_TO_DEG), 1) * 60.0f;
		if (lat_d >= 0)
			*lat_c = 'N';
		else
		{
			*lat_c = 'S';
			lat_d *= -1;
		}
		*lat_f = lat_d * 100 + lat_m;

		int lon_d = (int)floor(lon*RAD_TO_DEG);
		double lon_m = std::fmod(std::abs(lon*RAD_TO_DEG), 1) * 60.0f;
		if (lon_d >= 0)
			*lon_c = 'E';
		else
		{
			*lon_c = 'W';
			lon_d *= -1;
		}
		*lon_f = lon_d * 100 + lon_m;
		return 0;
	}

	/// <summary>Create NMEA GPGGA message from 0x6403 Geodetic position and Altitude above Geoid from 0x2402 Navigation Output messages</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaGpGGA(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//Recalculate Lat/Lon to NMEA type
		char lat_c, lon_c;
		double lat_f, lon_f;
		getNmeaLatLon(&lat_f, &lat_c, &lon_f, &lon_c, MessageSet.GeoPosition.Latitude, MessageSet.GeoPosition.Longitude);

		//Determinate Fix quality
		int fixQuality = 0;
		switch (MessageSet.GeoPosition.InsGnssBIT.GpsMode)
		{
		case 0: fixQuality = 1; break;
		case 1:
		case 2: fixQuality = 2; break;
		case 3: fixQuality = 5; break;
		case 4: fixQuality = 4; break;
		default: fixQuality = 0;
		}

		//Calculate UTC Time
		double utcTime = gpsTovToUtcTov(MessageSet.GeoPosition.GpsTov);

		//Calculate 2D position precision
		float dilutionOfPosition = std::sqrt(std::pow(MessageSet.GeoPosition.StdvNED[0], 2) + std::pow(MessageSet.GeoPosition.StdvNED[1], 2));

		//Print the GPGGA string
		if (MessageSet.NavOutput.SystemTov != 0)
		{
			i = snprintf((char*)pb,512, "$GPGGA,%010.3f,"
				"%010.5f,%C,"
				"%011.5f,%C,"
				"%d,,%f," //dont fill Number of satellites
				"%f,M,%f,M,,*", utcTime, lat_f, lat_c, lon_f, lon_c, fixQuality, dilutionOfPosition, MessageSet.GeoPosition.AltitudeAboveEllipsoid, MessageSet.NavOutput.AltitudeGeoid);
		}
		else //no data in 0x2402 message
		{
			i = snprintf((char*)pb,512, "$GPGGA,%010.3f,"
				"%010.5f,%C,"
				"%011.5f,%C,"
				"%d,,%f," //dont fill Number of satellites
				"%f,M,,,,*", utcTime, lat_f, lat_c, lon_f, lon_c, fixQuality, dilutionOfPosition, MessageSet.GeoPosition.AltitudeAboveEllipsoid);
		}


		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i += snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

	/// <summary>Create NMEA GPRMC message from 0x6403 Geodetic position, 0x6405 Euler Attitudes and 0x6503 NED Velocity messages</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaGpRMC(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//Recalculate Lat/Lon to NMEA type
		char lat_c, lon_c;
		double lat_f, lon_f;
		getNmeaLatLon(&lat_f, &lat_c, &lon_f, &lon_c, MessageSet.GeoPosition.Latitude, MessageSet.GeoPosition.Longitude);

		//Calculate UTC Time
		double utcTime = gpsTovToUtcTov(MessageSet.GeoPosition.GpsTov);

		//Calculate ground speed in knots
		double speedOverGround = std::sqrt(std::pow(MessageSet.NEDVelocity.VelocityNED[0], 2) + std::pow(MessageSet.NEDVelocity.VelocityNED[1], 2))*1.94384f;

		//True Heading in degrees
		double trueHeading = MessageSet.EulerAttitudes.EulerAttitude[2] * 57.2957f;

		//Print the GPGGA string
		i = snprintf((char*)pb,512, "$GPRMC,%010.3f,A,"
			"%010.5f,%C,"
			"%011.5f,%C,"
			"%f,%f,"//don't fill date
			",,*", utcTime, lat_f, lat_c, lon_f, lon_c, speedOverGround, trueHeading);


		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i+=snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

	/// <summary>Create NMEA GPGLL message from 0x6403 Geodetic position message</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaGpGLL(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//Recalculate Lat/Lon to NMEA type
		char lat_c, lon_c;
		double lat_f, lon_f;
		getNmeaLatLon(&lat_f, &lat_c, &lon_f, &lon_c, MessageSet.GeoPosition.Latitude, MessageSet.GeoPosition.Longitude);

		//Calculate UTC Time
		double utcTime = gpsTovToUtcTov(MessageSet.GeoPosition.GpsTov);

		//Print the GPGGA string
		i = snprintf((char*)pb,256, "$GPGLL,"
			"%010.5f,%C,"
			"%011.5f,%C,"
			",%010.3f,A*", lat_f, lat_c, lon_f, lon_c, utcTime);


		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i += snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

	/// <summary>Create NMEA GPVTG message from 0x6405 Euler Attitudes and 0x6503 NED Velocity messages</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaGpVTG(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//Calculate ground speed in km/h
		double speedOverGround = std::sqrt(std::pow(MessageSet.NEDVelocity.VelocityNED[0], 2) + std::pow(MessageSet.NEDVelocity.VelocityNED[1], 2))*3.6f;
		double speedOverGroundKnots = speedOverGround * 0.539957f;

		//True Heading in degrees
		double trueHeading = MessageSet.EulerAttitudes.EulerAttitude[2] * RAD_TO_DEG;

		//Print the GPVTG string
		i = snprintf((char*)pb,256, "$GPVTG,%03f,T,,,%03f,N,%03f,K*", trueHeading, speedOverGroundKnots, speedOverGround);

		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i += snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

	/// <summary>Create NMEA PASHR message from 0x6405 Euler Attitudes</summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaPASHR(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//Calculate UTC Time
		double utcTime = gpsTovToUtcTov(MessageSet.GeoPosition.GpsTov);
		
		//Roll in degrees
		double roll = MessageSet.EulerAttitudes.EulerAttitude[0] * RAD_TO_DEG;
		double rollStdv = MessageSet.EulerAttitudes.EulerAttitudeStdv[0] * RAD_TO_DEG;
		//Pitch in degrees
		double pitch = MessageSet.EulerAttitudes.EulerAttitude[1] * RAD_TO_DEG;
		double pitchStdv = MessageSet.EulerAttitudes.EulerAttitudeStdv[1] * RAD_TO_DEG;
		//True Heading in degrees
		double trueHeading = MessageSet.EulerAttitudes.EulerAttitude[2] * RAD_TO_DEG;
		double trueHeadingStdv = MessageSet.EulerAttitudes.EulerAttitudeStdv[2] * RAD_TO_DEG;

		int qualityFlag = 0;
		if (MessageSet.EulerAttitudes.InsGnssBIT.GpsMode < 15) //ALL NON-RTK FIX
			qualityFlag = 1;
		if (MessageSet.EulerAttitudes.InsGnssBIT.GpsMode == 4) //RTK FIXED
			qualityFlag = 2;

		//Print the GPGGA string
		i = snprintf((char*)pb,256, "$PASHR,%010.3f,"
			"%06.2f,T,%06.2f,%06.2f,"
			"0," //placeholder for HEAVE
			"%06.3f,%06.3f,%06.3f,"//don't fill date
			"%d*", utcTime, trueHeading, roll, pitch, rollStdv, pitchStdv, trueHeadingStdv, qualityFlag);


		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i += snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

	/// <summary>Create NMEA GPHDT message from 0x6405 Euler Attitudes </summary>
	/// <param name="buffer">Buffer to store the message to</param>  
	/// <param name="startOffset">Byte which will contain the 1st byte of NMEA message - '$'</param>  
	/// <param name="MessageSet">Structure containing the data</param>  
	/// <param name="byteLength">Message Byte Length</param>  
	/// <returns>0 - OK | -1 - incorrect start offset | -2 - incorrect checksum start</returns>  
	int CreateNmeaGpHDT(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength)
	{
		if (startOffset < 0)
			return -1;

		int i = 0;
		UINT8 *pb = &buffer[startOffset];

		//True Heading in degrees
		double trueHeading = MessageSet.EulerAttitudes.EulerAttitude[2] * RAD_TO_DEG;

		//Print the GPVTG string
		i = snprintf((char*)pb,48, "$GPHDT,%f,T*", trueHeading);

		int checksum = NmeaChecksum(pb, 0);

		if (checksum == 0x00) // error calculating checksum
			return -2;

		//Print calculated Checksum
		i += snprintf((char*)(pb + i),4, "%.2X", checksum);

		*byteLength = i;
		return 0;
	}

}

