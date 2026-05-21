/*
HONEYWELL hereby grants to you a, perpetual, free of charge, worldwide, irrevocable, non-exclusive license to use, copy, modify, merge,
publish, distribute, sublicense the software and associated documentation (the �Software�), subject to the following conditions:

YOU AGREE THAT YOU ASSUME ALL THE RESPONSIBILITY AND RISK FOR YOUR USE OF THE SOFTWARE AND THE RESULTS AND PERFORMANCE THEREOF.
THE SOFTWARE IS PROVIDED TO YOU ON AN �AS IS� AND �AS AVAILABLE� BASIS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
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

#pragma once

#include "HgDataParser.h"
#include <cmath>
#include <stdio.h>

#ifdef _MSC_VER
#define snprintf  _snprintf 
#endif
/*--------------------------------------*/
/*-----------TYPE DEFINITIONS-----------*/
/*--------------------------------------*/
#ifndef HGTYPES
	typedef signed char         INT8, *PINT8;
	typedef signed short        INT16, *PINT16;
	typedef signed int          INT32, *PINT32;
	typedef unsigned char       UINT8, *PUINT8;
	typedef unsigned short      UINT16, *PUINT16;
	typedef unsigned int        UINT32, *PUINT32;

	typedef signed int LONG32, *PLONG32;

	typedef unsigned int ULONG32, *PULONG32;
	typedef unsigned int DWORD32, *PDWORD32;
	#define HGTYPES
#endif

#ifndef HG_UTILS
	#define PI 3.1415926535897932
	#define RAD_TO_DEG 180/PI
	#define DEG_TO_RAD PI/180
	#define HG_UTILS
#endif

#define MSG_LEN_X1001 16
#define MSG_LEN_X1002 46
#define MSG_LEN_X1004 45
#define MSG_LEN_X1101 10
#define	MSG_LEN_X1105 10
#define MSG_LEN_X1401 39
#define MSG_LEN_X4110 24
#define MSG_LEN_X4204 13
#define MSG_LEN_X4404 13
#define MSG_LEN_X4438 24
#define MSG_LEN_X2001 80
#define MSG_LEN_X2011 55
#define MSG_LEN_X20FF 11
#define MSG_LEN_X2201 24
#define MSG_LEN_X2311 17
#define MSG_LEN_X2402 57
#define MSG_LEN_X6403 18
#define MSG_LEN_X6405 15
#define MSG_LEN_X6504 15
#define MSG_LEN_X6108 32
#define MSG_LEN_X6109 19
#define MSG_LEN_X6110 24
#define MSG_LEN_X6111 40
#define MSG_LEN_X6424 42
#define MSG_LEN_X6438 48

#define MSG_LEN_MAX MSG_LEN_X2001 //maximum message length

namespace NavDataParser
{
	/*--------------------------------------*/
	/*-----------COMMON-MESSAGES------------*/
	/*--------------------------------------*/
	//List of all messages
	struct EnabledMessages
	{
		bool EnableX2001INSConfigurationOneShot; // 0 = Disable | 1 = Enable
		bool EnableX2011INSModeStatusBIT1Hz; // 0 = Disable | 1 = Enable
		bool EnableX2201TimeMark; // 0 = Disable | 1 = Enable
		bool EnableX2311InertialDataOutputMessage100Hz; // 0 = Disable | 1 = Enable
		bool SaveToFlash; // 0 = Do not Save to Falsh | 1 = Save to Flash
		bool EnableX2402NavigationOutputMessage50Hz; // 0 = Disable | 1 = Enable
		bool EnableX6108GNSSPositionFromReceiver; // 0 = Disable | 1 = Enable
		bool EnableX6403GeodeticPosition; // 0 = Disable | 1 = Enable
		bool EnableX6405EulerAttitude; // 0 = Disable | 1 = Enable
		bool EnableX6504NEDVelocity; // 0 = Disable | 1 = Enable
		bool EnableX6109GNSSAttitudeFromReceiver; // 0 = Disable | 1 = Enable
		bool EnableX6110DistanceTraveled; // 0 = Disable | 1 = Enable
		bool EnableDebugMessages; // 0 = Disable | 1 = Enable (send to Honeywell Apps engineer for debug)
		UINT32 MessageWord1;
		UINT32 MessageWord2;
		//Sets all values to zero / false
		void ZeroMessage()
		{
			EnableX2001INSConfigurationOneShot = false;
			EnableX2011INSModeStatusBIT1Hz = false;
			EnableX2201TimeMark = false;
			EnableX2311InertialDataOutputMessage100Hz = false;
			EnableX2402NavigationOutputMessage50Hz = false;
			EnableX6108GNSSPositionFromReceiver = false;
			EnableX6109GNSSAttitudeFromReceiver = false;
			EnableX6403GeodeticPosition = false;
			EnableX6405EulerAttitude = false;
			EnableX6504NEDVelocity = false;
			EnableX6110DistanceTraveled = false;
			EnableDebugMessages = false;
			SaveToFlash = false;
			MessageWord1=0;
			MessageWord2=0;
		}
		
		// == operator override
		bool operator==(const EnabledMessages& Message) const
		{
			if (EnableX2001INSConfigurationOneShot != Message.EnableX2001INSConfigurationOneShot) return false;
			if (EnableX2011INSModeStatusBIT1Hz != Message.EnableX2011INSModeStatusBIT1Hz) return false;
			if (EnableX2201TimeMark != Message.EnableX2201TimeMark)return false;
			if (EnableX2311InertialDataOutputMessage100Hz != Message.EnableX2311InertialDataOutputMessage100Hz)return false;
			if (SaveToFlash != Message.SaveToFlash)return false;
			if (EnableX2402NavigationOutputMessage50Hz != Message.EnableX2402NavigationOutputMessage50Hz)return false;
			if (EnableX6108GNSSPositionFromReceiver != Message.EnableX6108GNSSPositionFromReceiver)return false;
			if (EnableX6403GeodeticPosition != Message.EnableX6403GeodeticPosition)return false;
			if (EnableX6405EulerAttitude != Message.EnableX6405EulerAttitude)return false;
			if (EnableX6504NEDVelocity != Message.EnableX6504NEDVelocity)return false;
			if (EnableX6109GNSSAttitudeFromReceiver != Message.EnableX6109GNSSAttitudeFromReceiver)return false;
			if (EnableX6110DistanceTraveled != Message.EnableX6110DistanceTraveled)return false;
			if (EnableDebugMessages != Message.EnableDebugMessages)return false;
			if (MessageWord1 != Message.MessageWord1)return false;
			if (MessageWord2 != Message.MessageWord2)return false;
			else
				return true;
		}
	};

	/*--------------------------------------*/
	/*------------INPUT-MESSAGES------------*/
	/*--------------------------------------*/
	//Configure which messages should be sent by the INS
	struct Hg0x1001EnableInput
	{
		struct EnabledMessages Messages;
		//Sets all values to zero / false
		void ZeroMessage()
		{
			Messages.ZeroMessage();
		}

		// == operator override
		bool operator==(const Hg0x1001EnableInput& Message) const
		{
			if (Messages == Message.Messages)
				return true;
			else
				return false;
		}
	};

	//Select which Navigation Aiding Sources should be enabled
	struct EnableNavAidingSources
	{
		bool GnssPvtVelocity; // 0 = Disable | 1 = Enable
		bool GnssPvtPosition; // 0 = Disable | 1 = Enable
		bool ZeroVelocity; // 0 = Disable | 1 = Enable
		bool ZeroHeadingChange; // 0 = Disable | 1 = Enable
		bool BarometricAltitude; // 0 = Disable | 1 = Enable
		bool MagneticHeading; // 0 = Disable | 1 = Enable
		bool AidingSourcesEnable; //Must be 'true' in order to use values in structure
	};

	//Configure the INS Navigation mode
	struct Hg0x1002NavigationModeInput
	{
		UINT32 INSMode; // 0 = No Change | 1 = Standby | 2 = Coarse Level | 4 = Aided Navigation
		UINT32 TimeMarkPpsMode; // 0 = No Change | 1 = Disable | 2 = Pass Through | 3 = Time Mark via System
		float CoarseLevelDuration; // [s] Time in Coarse level
		struct EnableNavAidingSources NavAidingSourcesEn;
		float ZeroVelocityStdv; // [m/s2] Must be enabled in NavAidingSourcesEn
		float ZeroHeadingStdv; // [rad] Must be enabled in NavAidingSourcesEn
		float MDTSettlingTime; // [sec] Time before activation of Zero velocity (default = 5s)
		float MDT1AngularRate; // [rad/s] Total Angular rate considered zero heading change (default = 0.002 rad/s)
		float MDT2SpeedStdv; // [m/s] Total speed upper limit which triggers the ZUPT (default = TBD)
		float MDT3AngularRateInstant; // [rad/s] Angular rate which will trigger exit of ZUPT (default = 0.005 rad/s)
		float MDT4LinearAcceleration; // [m/s2] Linear Acceleration wich will trigger exit of ZUPT (default = 0.03 m/s2)

		float MD3NominalFN3dB; // [Hz] filter for 5s settling time (defauly = 0.2Hz)
		float MD3InstantFN3dB; // [Hz] filter for immediate exit of ZUPT (default = 15Hz)

		//Sets all values to zero / false
		void ZeroMessage()
		{
			INSMode = 0;
			TimeMarkPpsMode = 0;
			CoarseLevelDuration = 0;
			NavAidingSourcesEn.GnssPvtPosition= 0;
			NavAidingSourcesEn.GnssPvtVelocity= 0;
			NavAidingSourcesEn.AidingSourcesEnable = 0;
			NavAidingSourcesEn.BarometricAltitude = 0;
			NavAidingSourcesEn.MagneticHeading = 0;
			NavAidingSourcesEn.ZeroHeadingChange = 0;
			NavAidingSourcesEn.ZeroVelocity = 0;
			ZeroVelocityStdv = 0;
			ZeroHeadingStdv = 0;
			MDTSettlingTime= 0; 
			MDT1AngularRate = 0;
			MDT2SpeedStdv = 0; 
			MDT3AngularRateInstant = 0; 
			MDT4LinearAcceleration = 0; 
			MD3NominalFN3dB = 0; 
			MD3InstantFN3dB = 0;
		}

		// == operator override
		bool operator==(const Hg0x1002NavigationModeInput& Message) const
		{
			if (INSMode != Message.INSMode) return false;
			if (TimeMarkPpsMode != Message.TimeMarkPpsMode) return false;
			if (CoarseLevelDuration != Message.CoarseLevelDuration) return false;
			if (NavAidingSourcesEn.GnssPvtPosition != Message.NavAidingSourcesEn.GnssPvtPosition) return false;
			if (NavAidingSourcesEn.GnssPvtVelocity!= Message.NavAidingSourcesEn.GnssPvtVelocity) return false;
			if (NavAidingSourcesEn.AidingSourcesEnable != Message.NavAidingSourcesEn.AidingSourcesEnable) return false;
			if (NavAidingSourcesEn.BarometricAltitude != Message.NavAidingSourcesEn.BarometricAltitude) return false;
			if (NavAidingSourcesEn.MagneticHeading != Message.NavAidingSourcesEn.MagneticHeading) return false;
			if (NavAidingSourcesEn.ZeroHeadingChange != Message.NavAidingSourcesEn.ZeroHeadingChange) return false;
			if (NavAidingSourcesEn.ZeroVelocity != Message.NavAidingSourcesEn.ZeroVelocity) return false;
			if (ZeroVelocityStdv != Message.ZeroVelocityStdv) return false;
			if (MDTSettlingTime != Message.MDTSettlingTime) return false;
			if (MDT1AngularRate != Message.MDT1AngularRate) return false;
			if (MDT2SpeedStdv != Message.MDT2SpeedStdv) return false;
			if (MDT3AngularRateInstant != Message.MDT3AngularRateInstant) return false;
			if (MDT4LinearAcceleration != Message.MDT4LinearAcceleration) return false;
			if (MD3NominalFN3dB != Message.MD3NominalFN3dB) return false;
			if (MD3InstantFN3dB != Message.MD3InstantFN3dB) return false;
			else
				return true;
		}

		// limit test
		bool EqualWithMargin(const Hg0x1002NavigationModeInput& Message, double Margin) const
		{
			if (INSMode != Message.INSMode) return false;
			if (TimeMarkPpsMode != Message.TimeMarkPpsMode) return false;
			if (std::abs(CoarseLevelDuration - Message.CoarseLevelDuration)>std::abs(Message.CoarseLevelDuration*Margin)) return false;
			if (NavAidingSourcesEn.AidingSourcesEnable != Message.NavAidingSourcesEn.AidingSourcesEnable) return false;
			if (NavAidingSourcesEn.GnssPvtPosition != Message.NavAidingSourcesEn.GnssPvtPosition) return false;
			if (NavAidingSourcesEn.GnssPvtVelocity != Message.NavAidingSourcesEn.GnssPvtVelocity) return false;
			if (NavAidingSourcesEn.BarometricAltitude != Message.NavAidingSourcesEn.BarometricAltitude) return false;
			if (NavAidingSourcesEn.MagneticHeading != Message.NavAidingSourcesEn.MagneticHeading) return false;
			if (NavAidingSourcesEn.ZeroHeadingChange != Message.NavAidingSourcesEn.ZeroHeadingChange) return false;
			if (NavAidingSourcesEn.ZeroVelocity != Message.NavAidingSourcesEn.ZeroVelocity) return false;
			if (std::abs(ZeroVelocityStdv - Message.ZeroVelocityStdv)>std::abs(Message.ZeroVelocityStdv*Margin)) return false;
			if (std::abs(ZeroHeadingStdv - Message.ZeroHeadingStdv )>std::abs(Message.ZeroHeadingStdv*Margin)) return false;
			if (std::abs(MDTSettlingTime - Message.MDTSettlingTime)>std::abs(Message.MDTSettlingTime*Margin)) return false;
			if (std::abs(MDT1AngularRate - Message.MDT1AngularRate)>std::abs(Message.MDT1AngularRate*Margin)) return false;
			if (std::abs(MDT2SpeedStdv - Message.MDT2SpeedStdv)>std::abs(Message.MDT2SpeedStdv*Margin)) return false;
			if (std::abs(MDT3AngularRateInstant - Message.MDT3AngularRateInstant)>std::abs(Message.MDT3AngularRateInstant*Margin)) return false;
			if (std::abs(MDT4LinearAcceleration - Message.MDT4LinearAcceleration)>std::abs(Message.MDT4LinearAcceleration*Margin)) return false;
			if (std::abs(MD3NominalFN3dB - Message.MD3NominalFN3dB)>std::abs(Message.MD3NominalFN3dB*Margin)) return false;
			if (std::abs(MD3InstantFN3dB - Message.MD3InstantFN3dB)>std::abs(Message.MD3InstantFN3dB*Margin)) return false;
			else
				return true;
		}
	};

	//Select which values to update
	struct UpdateConfiguration
	{
		UINT8 VehicleFrameToCaseFrameTransformation; // 0 = No Change | 2 = Update with EulerAngles[3]
		UINT8 CaseToVehicleFrameLeverArmsInCaseFrame; // 0 = No Change | 2 = Update with VehicleLeverArms[3]
		UINT8 MainAntennaLeverArm;// 0 = No Change | 2 = Update with MainAntennaLeverArms[3]
	};

	//Configure vehicle/antenna mechanical structure and offsets - Legacy - use 0x4204 and 0x4404 respectively
	struct Hg0x1004ConfigurationInput
	{
		struct UpdateConfiguration UpdateConfig;
		float VehicleEulerAngles[3]; //[rad] Roll, Pitch, Yaw Vehicle frame offset - will be sent as pi*2-31
		float VehicleLeverArms[3]; // [m] X, Y, Z Vehicle frame offset
		float MainAntennaLeverArms[3]; //[m] X, Y, Z of Main Lever Arm
		//Sets all values to zero / false
		void ZeroMessage()
		{
			UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame = 0;
			UpdateConfig.MainAntennaLeverArm = 0;
			UpdateConfig.VehicleFrameToCaseFrameTransformation = 0;
			VehicleEulerAngles[0] = VehicleEulerAngles[1] = VehicleEulerAngles[2] = 0;
			VehicleLeverArms[0] = VehicleLeverArms[1] = VehicleLeverArms[2] = 0;
			MainAntennaLeverArms[0] = MainAntennaLeverArms[1] = MainAntennaLeverArms[2] = 0;
		}

		// == operator override
		bool operator==(const Hg0x1004ConfigurationInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (VehicleEulerAngles[i] != Message.VehicleEulerAngles[i])return false;
				if (VehicleLeverArms[i] != Message.VehicleLeverArms[i])return false;
				if (MainAntennaLeverArms[i] != Message.MainAntennaLeverArms[i])return false;
			}
			if (UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame != Message.UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame)return false;
			if (UpdateConfig.MainAntennaLeverArm != Message.UpdateConfig.MainAntennaLeverArm)return false;
			if (UpdateConfig.VehicleFrameToCaseFrameTransformation != Message.UpdateConfig.VehicleFrameToCaseFrameTransformation)return false;
			else
				return true;
		}

		// limit test
		bool EqualWithMargin(const Hg0x1004ConfigurationInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VehicleEulerAngles[i] - Message.VehicleEulerAngles[i])>std::abs(Message.VehicleEulerAngles[i]*Margin)) return false;
				if (std::abs(VehicleLeverArms[i] - Message.VehicleLeverArms[i])>std::abs(Message.VehicleLeverArms[i]*Margin)) return false;
				if (std::abs(MainAntennaLeverArms[i] - Message.MainAntennaLeverArms[i])>std::abs(Message.MainAntennaLeverArms[i]*Margin)) return false;
			}
			if (UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame != Message.UpdateConfig.CaseToVehicleFrameLeverArmsInCaseFrame)return false;
			if (UpdateConfig.MainAntennaLeverArm != Message.UpdateConfig.MainAntennaLeverArm)return false;
			if (UpdateConfig.VehicleFrameToCaseFrameTransformation != Message.UpdateConfig.VehicleFrameToCaseFrameTransformation)return false;	
			else
				return true;
		}
	};

	//Barometric Altitude Aiding input - Must be enabled in 0x1002
	struct Hg0x1101BarometricAltitudeInput
	{
		double BarometricAltitudeTov; // [s] data validity time (based on TovMode)
		bool BarometricAltitudeValid; // 0 = invalid | 1 = valid
		bool TovMode; // 0 = gps time | 1 = Message Receipt Timestamp
		float BarometricAltitudeMslGeoid; // [m] altitude above mean sea level geoid
		//Sets all values to zero / false
		void ZeroMessage()
		{
			BarometricAltitudeTov = 0;
			BarometricAltitudeValid = false;
			TovMode = false;
			BarometricAltitudeMslGeoid = 0;
		}

		// == operator override
		bool operator==(const Hg0x1101BarometricAltitudeInput& Message) const
		{
			if (BarometricAltitudeTov != Message.BarometricAltitudeTov) return false;
			if (BarometricAltitudeValid != Message.BarometricAltitudeValid) return false;
			if (TovMode != Message.TovMode) return false;
			if (BarometricAltitudeMslGeoid != Message.BarometricAltitudeMslGeoid) return false;
			else 
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x1101BarometricAltitudeInput& Message, double Margin) const
		{
			if (std::abs(BarometricAltitudeTov - Message.BarometricAltitudeTov)>std::abs(Message.BarometricAltitudeTov*Margin)) return false;
			if (BarometricAltitudeValid != Message.BarometricAltitudeValid) return false;
			if (TovMode != Message.TovMode) return false;
			if (std::abs(BarometricAltitudeMslGeoid - Message.BarometricAltitudeMslGeoid)>std::abs(Message.BarometricAltitudeMslGeoid*Margin)) return false;
			else
				return true;
		}
	};

	//Magnetic Heading Aiding input - Must be enabled in 0x1002
	struct Hg0x1105MagneticHeading
	{
		double MagneticHeadingTov; //[s] data validity time (based on TovMode)
		bool MagneticHeadingValid; // 0 = invalid | 1 = valid
		bool TovMode; // 0 = gps time | 1 = Message Receipt Timestamp
		bool MagneticVariationValid; // 0 = invalid | 1 = valid
		float MagneticHeading; //[rad] magnetic heading
		float MagneticVariation;
		//Sets all values to zero / false
		void ZeroMessage()
		{
			MagneticHeadingTov = 0;
			MagneticHeadingValid = false;
			TovMode = false;
			MagneticVariationValid = false;
			MagneticHeading = 0;
			MagneticVariation = 0;
		}

		// == operator override
		bool operator==(const Hg0x1105MagneticHeading& Message) const
		{
			if (MagneticHeadingTov != Message.MagneticHeadingTov) return false;
			if (MagneticHeadingValid != Message.MagneticHeadingValid) return false;
			if (TovMode != Message.TovMode) return false;
			if (MagneticVariationValid != Message.MagneticVariationValid) return false;
			if (MagneticHeading != Message.MagneticHeading) return false;
			if (MagneticVariation != Message.MagneticVariation) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x1105MagneticHeading& Message, double Margin) const
		{
			if (std::abs(MagneticHeadingTov - Message.MagneticHeadingTov)>std::abs(Message.MagneticHeadingTov*Margin)) return false;
			if (MagneticHeadingValid != Message.MagneticHeadingValid) return false;
			if (TovMode != Message.TovMode) return false;
			if (MagneticVariationValid != Message.MagneticVariationValid) return false;
			if (std::abs(MagneticHeading - Message.MagneticHeading)>std::abs(Message.MagneticHeading*Margin)) return false;
			if (std::abs(MagneticVariation - Message.MagneticVariation)>std::abs(Message.MagneticVariation*Margin)) return false;
			else
				return true;
		}
	};
	
	//TOV Settings of appropriate section
	struct TovSettings
	{
		bool Valid; // 0 = invalid | 1 = valid
		bool TovMode; // 0 = gps time | 1 = Message Receipt Timestamp
		bool SetFrame; // 0 = N-Frame | 1 = Euler
		bool StdvValid; // 0 = invalid | 1 = valid

		// == operator override
		bool operator==(const TovSettings& Settings) const
		{
			if (Valid != Settings.Valid) return false;
			if (TovMode != Settings.TovMode) return false;
			if (SetFrame != Settings.SetFrame) return false;
			if (StdvValid != Settings.StdvValid) return false;
			else
				return true;
		}

		// != operator override
		bool operator!=(const TovSettings& Settings) const
		{
			if (Valid == Settings.Valid) return false;
			if (TovMode == Settings.TovMode) return false;
			if (SetFrame == Settings.SetFrame) return false;
			if (StdvValid == Settings.StdvValid) return false;
			else
				return true;
		}
	};
	
	//Navigation input for the INS
	struct Hg0x1401NavigationInput
	{
		bool AckRequested; // 0 = false | 1 = true (default)

		double PositionTov; // [s] data validity time (based on TovMode)
		struct TovSettings PositionTovSettings;
		float Latitude; // [rad] latitude position
		float Longitude; // [rad] longitude position
		float AltitudeAboveElipsoid; //[m] altitude above ellipsoid

		double VelocityTov; //[s] data validity time (based on TovMode)
		struct TovSettings VelocityTovSettings;
		float NEDVelocity[3]; // [m/s] North, East, Down velocity

		double AttitudeTov; //[s] data validity time (based on TovMode)
		struct TovSettings AttitudeTovSettings;
		float VehicleEulerAngles[3]; // [rad] Roll, Pitch, True Heading

		float NEDPositionStdv[3]; // [m] STDV North, East, Down (passing 0 equals default = 100 m)
		float NEDVelocityStdv[3]; // [m/s] STDV North, East, Down (passing 0 means default = 1 m/s)
		float EulerAnglesStdv[3]; // [rad] STDV Roll, Pitch, True Heading (passing 0 means default = 0.0873 rad)

		//Sets all values to zero / false
		void ZeroMessage()
		{
			AckRequested = 1; // Default set to 1
			PositionTov = 0;
			PositionTovSettings.SetFrame = 0;
			PositionTovSettings.Valid = 0;
			PositionTovSettings.TovMode = 0;
			PositionTovSettings.StdvValid = 0;
			Latitude = 0;
			Longitude = 0;
			AltitudeAboveElipsoid = 0;
			VelocityTov = 0;
			VelocityTovSettings.SetFrame = 0;
			VelocityTovSettings.Valid = 0;
			VelocityTovSettings.TovMode = 0;
			VelocityTovSettings.StdvValid = 0;
			NEDVelocity[0] = NEDVelocity[1] = NEDVelocity[2] = 0;
			AttitudeTov = 0;
			AttitudeTovSettings.SetFrame = 1; // Always set to 1 = Euler
			AttitudeTovSettings.StdvValid = 0;
			AttitudeTovSettings.TovMode = 0;
			AttitudeTovSettings.Valid = 0;
			VehicleEulerAngles[0] = VehicleEulerAngles[1] = VehicleEulerAngles[2] = 0;
			NEDPositionStdv[0] = NEDPositionStdv[1] = NEDPositionStdv[2] = 0;
			NEDVelocityStdv[0] = NEDVelocityStdv[1] = NEDVelocityStdv[2] = 0;
			EulerAnglesStdv[0] = EulerAnglesStdv[1] = EulerAnglesStdv[2] = 0;
		}

		// == operator override
		bool operator==(const Hg0x1401NavigationInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (NEDVelocity[i] != Message.NEDVelocity[i]) return false;
				if (VehicleEulerAngles[i] != Message.VehicleEulerAngles[i]) return false;
				if (NEDPositionStdv[i] != Message.NEDPositionStdv[i]) return false;
				if (NEDVelocityStdv[i] != Message.NEDVelocityStdv[i]) return false;
				if (EulerAnglesStdv[i] != Message.EulerAnglesStdv[i]) return false;
			}
			if (AckRequested != Message.AckRequested) return false;
			if (PositionTov != Message.PositionTov) return false;
			if (PositionTovSettings != Message.PositionTovSettings) return false;
			if (Latitude != Message.Latitude) return false;
			if (Longitude != Message.Longitude) return false;
			if (AltitudeAboveElipsoid != Message.AltitudeAboveElipsoid) return false;
			if (VelocityTov != Message.VelocityTov) return false;
			if (VelocityTovSettings != Message.VelocityTovSettings) return false;
			if (AttitudeTov != Message.AttitudeTov) return false;
			if (AttitudeTovSettings != Message.AttitudeTovSettings) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x1401NavigationInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(NEDVelocity[i] - Message.NEDVelocity[i])>std::abs(Message.NEDVelocity[i]*Margin)) return false;
				if (std::abs(VehicleEulerAngles[i] - Message.VehicleEulerAngles[i])>std::abs(Message.VehicleEulerAngles[i]*Margin)) return false;
				if (std::abs(NEDPositionStdv[i] - Message.NEDPositionStdv[i])>std::abs(Message.NEDPositionStdv[i]*Margin)) return false;
				if (std::abs(NEDVelocityStdv[i] - Message.NEDVelocityStdv[i])>std::abs(Message.NEDVelocityStdv[i]*Margin)) return false;
				if (std::abs(EulerAnglesStdv[i] - Message.EulerAnglesStdv[i])>std::abs(Message.EulerAnglesStdv[i]*Margin)) return false;
			}

			if (AckRequested != Message.AckRequested) return false;
			if (std::abs(PositionTov - Message.PositionTov)>std::abs(Message.PositionTov*Margin)) return false;
			if (PositionTovSettings != Message.PositionTovSettings) return false;
			if (std::abs(Latitude - Message.Latitude)>std::abs(Message.Latitude*Margin)) return false;
			if (std::abs(Longitude - Message.Longitude)>std::abs(Message.Longitude*Margin)) return false;
			if (std::abs(AltitudeAboveElipsoid - Message.AltitudeAboveElipsoid)>std::abs(Message.AltitudeAboveElipsoid*Margin)) return false;
			if (std::abs(VelocityTov - Message.VelocityTov)>std::abs(Message.VelocityTov*Margin)) return false;
			if (VelocityTovSettings != Message.VelocityTovSettings) return false;
			if (std::abs(AttitudeTov - Message.AttitudeTov)>std::abs(Message.AttitudeTov*Margin)) return false;
			if (AttitudeTovSettings != Message.AttitudeTovSettings) return false;
			else
				return true;
		}
	};
	
	//Select which values to be saved to flash
	struct LAStoreToFlash
	{
		bool InputMainLA; // 0 = No Action | 1 = Store
		bool InputAuxLA; // 0 = No Action | 1 = Store
		bool InputMainLAStdv; // 0 = No Action | 1 = Store
		bool InputAuxLAStdv; // 0 = No Action | 1 = Store
	};

	//Antenna lever arms settings
	struct Hg0x4204AntennaLeverArmInput
	{
		float MainAntennaLeverArms[3]; // [m] X, Y, Z main antenna lever arms
		float AuxAntennaLeverArms[3]; // [m] X, Y, Z aux antenna lever arms
		float MainAntennaLAUncertainty; // [m] radius of uncertainty
		float AuxAntennaLAUncertainty; // [m] radius of uncertainty
		struct LAStoreToFlash LAStoreToFlash;
		//Sets all values to zero / false
		void ZeroMessage()
		{
			MainAntennaLeverArms[0] = MainAntennaLeverArms[1] = MainAntennaLeverArms[2] = 0;
			AuxAntennaLeverArms[0] = AuxAntennaLeverArms[1] = AuxAntennaLeverArms[2] = 0;
			MainAntennaLAUncertainty = 0;
			AuxAntennaLAUncertainty = 0;
			LAStoreToFlash.InputAuxLA = false;
			LAStoreToFlash.InputAuxLAStdv = false;
			LAStoreToFlash.InputMainLA = false;
			LAStoreToFlash.InputMainLAStdv = false;
		}

		// == operator override
		bool operator==(const Hg0x4204AntennaLeverArmInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (MainAntennaLeverArms[i] != Message.MainAntennaLeverArms[i]) return false;
				if (AuxAntennaLeverArms[i] != Message.AuxAntennaLeverArms[i]) return false;
			}
			if (MainAntennaLAUncertainty != Message.MainAntennaLAUncertainty) return false;
			if (AuxAntennaLAUncertainty != Message.AuxAntennaLAUncertainty) return false;
			if (LAStoreToFlash.InputAuxLA != Message.LAStoreToFlash.InputAuxLA) return false;
			if (LAStoreToFlash.InputAuxLAStdv != Message.LAStoreToFlash.InputAuxLAStdv) return false;
			if (LAStoreToFlash.InputMainLA != Message.LAStoreToFlash.InputMainLA) return false;
			if (LAStoreToFlash.InputMainLAStdv != Message.LAStoreToFlash.InputMainLAStdv) return false;
			else
				return true;
		}

		// limit test
		bool EqualWithMargin(const Hg0x4204AntennaLeverArmInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(MainAntennaLeverArms[i] - Message.MainAntennaLeverArms[i])>std::abs(Message.MainAntennaLeverArms[i]*Margin)) return false;
				if (std::abs(AuxAntennaLeverArms[i] - Message.AuxAntennaLeverArms[i])>std::abs(Message.AuxAntennaLeverArms[i]*Margin)) return false;
			}
			if (std::abs(MainAntennaLAUncertainty - Message.MainAntennaLAUncertainty)>std::abs(Message.MainAntennaLAUncertainty*Margin)) return false;
			if (std::abs(AuxAntennaLAUncertainty - Message.AuxAntennaLAUncertainty)>std::abs(Message.AuxAntennaLAUncertainty*Margin)) return false;
			if (LAStoreToFlash.InputAuxLA != Message.LAStoreToFlash.InputAuxLA) return false;
			if (LAStoreToFlash.InputAuxLAStdv != Message.LAStoreToFlash.InputAuxLAStdv) return false;
			if (LAStoreToFlash.InputMainLA != Message.LAStoreToFlash.InputMainLA) return false;
			if (LAStoreToFlash.InputMainLAStdv != Message.LAStoreToFlash.InputMainLAStdv) return false;
			else
				return true;
		}
	};

	//Configure vehicle frame
	struct Hg0x4404VehicleFrameInput
	{
		float VehicleLeverArms[3]; // [m] X, Y, Z Vehicle frame offset
		float VehicleEulerAngles[3]; // [rad] Roll, Pitch, Yaw Vehicle frame offset
		bool StoreToFlashEulerAngles; // 0 = No Action | 1 = Store
		bool StoreToFlashVehicleLA; // 0 = No Action | 1 = Store
		//Sets all values to zero / false
		void ZeroMessage()
		{
			VehicleLeverArms[0] = VehicleLeverArms[1] = VehicleLeverArms[2] = 0;
			VehicleEulerAngles[0] = VehicleEulerAngles[1] = VehicleEulerAngles[2] = 0;
			StoreToFlashEulerAngles = 0;
			StoreToFlashVehicleLA = 0;
		}

		// == operator override
		bool operator==(const Hg0x4404VehicleFrameInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (VehicleLeverArms[i] != Message.VehicleLeverArms[i]) return false;
				if (VehicleEulerAngles[i] != Message.VehicleEulerAngles[i]) return false;
			}
			if (StoreToFlashEulerAngles != Message.StoreToFlashEulerAngles) return false;
			if (StoreToFlashVehicleLA != Message.StoreToFlashVehicleLA) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x4404VehicleFrameInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VehicleLeverArms[i] - Message.VehicleLeverArms[i])>std::abs(Message.VehicleLeverArms[i]*Margin)) return false;
				if (std::abs(VehicleEulerAngles[i] - Message.VehicleEulerAngles[i])>std::abs(Message.VehicleEulerAngles[i]*Margin)) return false;
			}
			if (StoreToFlashEulerAngles != Message.StoreToFlashEulerAngles) return false;
			if (StoreToFlashVehicleLA != Message.StoreToFlashVehicleLA) return false;
			else
				return true;
		}
	};

	//Settings of Speed and Odometer Aiding
	struct OdometerSettings
	{
		bool VehicleVelocityValid; // 0 = Invalid | 1 = Valid
		bool TovMode; // 0 = GPS time | 1 = Message Receipt Timestamp
		bool OdometerValid; // 0 = Invalid | 1 = Valid
		bool AidingStatus; // 0 = Issues with aiding | 1 = OK (use to indicate issues with the aiding to the unit)
		bool ZeroVelocityDetected; // 0 = normal operation | 1 = Zero Velocity (default 0.01 m/s)

		// == operator override
		bool operator==(const OdometerSettings& Settings) const
		{
			if (VehicleVelocityValid != Settings.VehicleVelocityValid) return false;
			if (TovMode != Settings.TovMode) return false;
			if (OdometerValid != Settings.OdometerValid) return false;
			if (AidingStatus != Settings.AidingStatus) return false;
			if (ZeroVelocityDetected != Settings.ZeroVelocityDetected) return false;
			else
				return true;
		}

		// != operator override
		bool operator!=(const OdometerSettings& Settings) const
		{
			if (VehicleVelocityValid == Settings.VehicleVelocityValid) return false;
			if (TovMode == Settings.TovMode) return false;
			if (OdometerValid == Settings.OdometerValid) return false;
			if (AidingStatus == Settings.AidingStatus) return false;
			if (ZeroVelocityDetected == Settings.ZeroVelocityDetected) return false;
			else
				return true;
		}
	};


	//Send vehicle speed aiding
	struct Hg0x4110VelocityAidingInput
	{
		float VehicleVelocity[3]; // [m/s] X, Y, Z Vehicle Body Speed
		double OdometerTimeDelay; // [s] Odomter measurement delay
		double OdometerTov; // [s] Message time of validity - Mode set by OdometerSettings.TovMode
		INT32 OdometerCumulativePulses; // [-] Number of pulses from start
		float DistancePerPulse; // [m/pulse] Distance traveled per Odometer pulse (2*PI*WheelRadius / Vehicle Pulse rate / Second)
		struct OdometerSettings OdometerSettings; //Settings of Speed and Odometer Aiding
		UINT32 Counter; // [-] Incremental Counter
		
		//Sets all values to zero / false
		void ZeroMessage()
		{
			VehicleVelocity[0] = VehicleVelocity[1] = VehicleVelocity[0] = 0;
			OdometerTimeDelay = 0;
			OdometerTov= 0; 
			OdometerCumulativePulses= 0;
			DistancePerPulse = 0;
			OdometerSettings.AidingStatus = false;
			OdometerSettings.OdometerValid = false;
			OdometerSettings.TovMode = false;
			OdometerSettings.VehicleVelocityValid = false;
			Counter = 0;
		}

		// == operator override
		bool operator==(const Hg0x4110VelocityAidingInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (VehicleVelocity[i] != Message.VehicleVelocity[i]) return false;
			}
			if (OdometerTimeDelay != Message.OdometerTimeDelay) return false;
			if (OdometerTov != Message.OdometerTov) return false;
			if (OdometerCumulativePulses != Message.OdometerCumulativePulses) return false;
			if (DistancePerPulse != Message.DistancePerPulse) return false;
			if (OdometerSettings != Message.OdometerSettings) return false;
			if (Counter != Message.Counter) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x4110VelocityAidingInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VehicleVelocity[i] - Message.VehicleVelocity[i])>std::abs(Message.VehicleVelocity[i] * Margin)) return false;
			}
			if (std::abs(OdometerTimeDelay - Message.OdometerTimeDelay)>std::abs(Message.OdometerTimeDelay * Margin)) return false;
			if (std::abs(OdometerTov - Message.OdometerTov)>std::abs(Message.OdometerTov * Margin)) return false;
			if (std::abs(OdometerCumulativePulses - Message.OdometerCumulativePulses)>std::abs(Message.OdometerCumulativePulses * Margin)) return false;
			if (std::abs(DistancePerPulse - Message.DistancePerPulse)>std::abs(Message.DistancePerPulse * Margin)) return false;
			if (std::abs((long)Counter - (long)Message.Counter)>std::abs((long)Message.Counter * Margin)) return false;
			if (OdometerSettings != Message.OdometerSettings) return false;
			else
				return true;
		}
	};

	//Configure Kalman Filter for specific Odometer
	struct Hg0x4438OdometerConfigurationInput
	{
		float LeverArms[3]; // [m] X, Y, Z Lever Arms from Body Frame
		float MeasurementNoise; // [m/sample] Expected noise in measurement
		float Threshold; // [sigma] 
		float ScaleFactorUncertainty; //
		float ScaleFactorProcessNoise; // [1/sqrt(hz)]
		float YawBoresightStdv; // [rad]
		float YawBoresightProcessNoise; // [rad/sqrt(hz)]
		float PitchBoresightUncertainty; // [rad]
		float PitchBoresightProcessNoise; // [rad/sqrt(hz)]
		bool SaveToFlash; // 0 = No Action | 1 = Store

		//Sets all values to zero / false
		void ZeroMessage()
		{
			LeverArms[0] = LeverArms[1] = LeverArms[0] = 0;
			MeasurementNoise = 0;
			Threshold = 0;
			ScaleFactorUncertainty = 0;
			ScaleFactorProcessNoise = 0;
			YawBoresightStdv = 0;
			YawBoresightProcessNoise = 0;
			PitchBoresightUncertainty = 0;
			PitchBoresightProcessNoise = 0;
			SaveToFlash = 0;
		}

		// == operator override
		bool operator==(const Hg0x4438OdometerConfigurationInput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (LeverArms[i] != Message.LeverArms[i]) return false;
			}
			if (MeasurementNoise != Message.MeasurementNoise) return false;
			if (Threshold != Message.Threshold) return false;
			if (ScaleFactorUncertainty != Message.ScaleFactorUncertainty) return false;
			if (ScaleFactorProcessNoise != Message.ScaleFactorProcessNoise) return false;
			if (YawBoresightStdv != Message.YawBoresightStdv) return false;
			if (YawBoresightProcessNoise != Message.YawBoresightProcessNoise) return false;
			if (PitchBoresightUncertainty != Message.PitchBoresightUncertainty) return false;
			if (PitchBoresightProcessNoise != Message.PitchBoresightProcessNoise) return false;
			if (SaveToFlash != Message.SaveToFlash) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x4438OdometerConfigurationInput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(LeverArms[i] - Message.LeverArms[i])>std::abs(Message.LeverArms[i] * Margin)) return false;
			}
			if (std::abs(MeasurementNoise - Message.MeasurementNoise)>std::abs(Message.MeasurementNoise * Margin)) return false;
			if (std::abs(Threshold - Message.Threshold)>std::abs(Message.Threshold * Margin)) return false;
			if (std::abs(ScaleFactorUncertainty - Message.ScaleFactorUncertainty)>std::abs(Message.ScaleFactorUncertainty * Margin)) return false;
			if (std::abs(ScaleFactorProcessNoise - Message.ScaleFactorProcessNoise)>std::abs(Message.ScaleFactorProcessNoise * Margin)) return false;
			if (std::abs(YawBoresightStdv - Message.YawBoresightStdv)>std::abs(Message.YawBoresightStdv * Margin)) return false;
			if (std::abs(YawBoresightProcessNoise - Message.YawBoresightProcessNoise)>std::abs(Message.YawBoresightProcessNoise * Margin)) return false;
			if (std::abs(PitchBoresightUncertainty - Message.PitchBoresightUncertainty)>std::abs(Message.PitchBoresightUncertainty * Margin)) return false;
			if (std::abs(PitchBoresightProcessNoise - Message.PitchBoresightProcessNoise)>std::abs(Message.PitchBoresightProcessNoise * Margin)) return false;
			if (std::abs(SaveToFlash - Message.SaveToFlash)>std::abs(Message.SaveToFlash * Margin)) return false;
			else
				return true;
		}
	};

	// Create 0x1001 Enable outputs Input message
	HGDATAPARSER_API int Create0x1001Message(UINT8 *buffer, int startOffset, Hg0x1001EnableInput Message, int *byteLength);
	// Create 0x1002 Navigation Mode Input message
	HGDATAPARSER_API int Create0x1002Message(UINT8 *buffer, int startOffset, Hg0x1002NavigationModeInput Message, int *byteLength);
	// Create 0x1004 Configuration Input message
	HGDATAPARSER_API int Create0x1004Message(UINT8 *buffer, int startOffset, Hg0x1004ConfigurationInput Message, int *byteLength);
	// Create 0x1101 Barometric Altitude Input message
	HGDATAPARSER_API int Create0x1101Message(UINT8 *buffer, int startOffset, Hg0x1101BarometricAltitudeInput Message, int *byteLength);
	// Create 0x1105 Magnetic Heading Input message
	HGDATAPARSER_API int Create0x1105Message(UINT8 *buffer, int startOffset, Hg0x1105MagneticHeading Message, int *byteLength);
	// Create 0x1401 Navigation Input message
	HGDATAPARSER_API int Create0x1401Message(UINT8 *buffer, int startOffset, Hg0x1401NavigationInput Message, int *byteLength);
	// Create 0x4204 Antenna Lever Arms Input message
	HGDATAPARSER_API int Create0x4204Message(UINT8 *buffer, int startOffset, Hg0x4204AntennaLeverArmInput Message, int *byteLength);
	// Create 0x4404 Vehicle Frame Input message
	HGDATAPARSER_API int Create0x4404Message(UINT8 *buffer, int startOffset, Hg0x4404VehicleFrameInput Message, int *byteLength);
	// Create 0x4110 Vehicle Speed Input message
	HGDATAPARSER_API int Create0x4110Message(UINT8 *buffer, int startOffset, Hg0x4110VelocityAidingInput Message, int *byteLength);
	// Create 0x4438 Odometer Configuration Input message
	HGDATAPARSER_API int Create0x4438Message(UINT8 *buffer, int startOffset, Hg0x4438OdometerConfigurationInput Message, int *byteLength);

	//Return Calculated UINT32 Checksum from input buffer
	HGDATAPARSER_API UINT32 CalcHgInsChecksum(UINT8 *buffer, int startOffset, int wordLength);


	/*--------------------------------------*/
	/*-----------OUTPUT-MESSAGES------------*/
	/*--------------------------------------*/

	//INS / GNSS Built-in test results
	struct InsGnssBIT
	{
		bool GnssStatus; // 0 = OK | 1 = Failed
		bool ImuStatus; // 0 = OK | 1 = Failed
		bool InsStatus; // 0 = OK | 1 = Failed
		UINT8 GpsMode; // 0 = Standalone | 1 = SBAS | 2 = DGPS | 3 = RTK Float | 4 = RTK Fixed | 15 = Invalid Solution
		UINT8 InsMode; // 0 = No Change | 1 = Standby (default) | 2 = Coarse Level | 3 = Reserved | 4 = Aided Navigation
		
		bool MotionDetectActive; // 0 = Inactive | 1 = Active
		bool ZeroVelocity; // 0 = Not Detected | 1 = Zero Velocity Detected
		bool ZeroVelocityPending; // 0 = Conditions Not Met | 1 = Conditions Met, Waiting for Settling Time

		//Sets all values to zero / false
		void ZeroMessage()
		{
			GnssStatus = 0;
			ImuStatus = 0;
			InsStatus = 0;
			GpsMode = 0;
			InsMode = 0;
			MotionDetectActive = 0;
			ZeroVelocity = 0;
			ZeroVelocityPending = 0;
		}

		// == operator override
		bool operator==(const InsGnssBIT& Status) const
		{
			if (GnssStatus != Status.GnssStatus) return false;
			if (ImuStatus != Status.ImuStatus) return false;
			if (InsStatus != Status.InsStatus) return false;
			if (GpsMode != Status.GpsMode) return false;
			if (InsMode != Status.InsMode) return false;
			if (MotionDetectActive != Status.MotionDetectActive) return false;
			if (ZeroVelocity != Status.ZeroVelocity) return false;
			if (ZeroVelocityPending != Status.ZeroVelocityPending) return false;
			else 
				return true;
		}

		// != operator override
		bool operator!=(const InsGnssBIT& Status) const
		{
			if (GnssStatus == Status.GnssStatus) return false;
			if (ImuStatus == Status.ImuStatus) return false;
			if (InsStatus == Status.InsStatus) return false;
			if (GpsMode == Status.GpsMode) return false;
			if (InsMode == Status.InsMode) return false;
			if (MotionDetectActive == Status.MotionDetectActive) return false;
			if (ZeroVelocity == Status.ZeroVelocity) return false;
			if (ZeroVelocityPending == Status.ZeroVelocityPending) return false;
			else
				return true;
		}

	};

	//INS Navigation output - LEGACY
	struct Hg0x2402NavigationOutput
	{
		struct InsGnssBIT InsGnssBIT;
		UINT32 INSMode; // 0 = No Change | 1 = Standby | 2 = Coarse Level | 4 = Aided Navigation
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		double SystemTov; // [s] Time since INS power up
		INT16 GpsWeek; // [-] No. of current GPS week
		//INT16 UtcTimeFom;
		INT16 GnssTimeFom;
		INT16 InsBlendedFom;
		double PositionSystemTov; // [s] System time of validity of the position data
		float Latitude; // [rad] latitude position
		float Longitude; // [rad] longitude position
		float AltitudeElips; // [m] altitude above ellipsoid
		float AltitudeGeoid; // [m] altitude above geoid
		float ECEFPosition[3]; // [m] X, Y, Z Earth centered position 
		double VelocitySystemTov; // [s] System time of validity of the velocity data
		float NEDVelocity[3]; // [m/s] North, east, down velocity
		float ECEFVelocity[3]; // [m/s] X, Y, Z Earth centered velocity
		//double AttitudeTov;
		float VehicleEulerAngles[3]; // [rad] Roll, Pitch, True Heading
		float WanderAngle; // [rad] Wander Angle
		float DCM[3][3]; // Direction Cosine matrix
		float VehicleBodyAngularRate[3]; // [rad/s] Vehicle body X, Y, Z angular rate
		float VehicleBodyAcceleration[3]; // [m/s2] Vehicle body X, Y, Z acceleration
		INT32 AttitudeFom;
		float Quaternion[4]; // a, b, c, d

		//Sets all values to zero / false
		void ZeroMessage()
		{
			InsGnssBIT.ZeroMessage();
			INSMode = 0;
			GpsTov = 0;
			SystemTov = 0;
			GpsWeek = 0;
			//UtcTimeFom = 0;
			GnssTimeFom = 0;
			//InsBlendedFom = 0;
			PositionSystemTov = 0;
			Latitude=0;
			Longitude = 0;
			AltitudeElips = 0;
			AltitudeGeoid = 0;
			ECEFPosition[0] = ECEFPosition[1] = ECEFPosition[2] = 0; // X, Y, Z
			VelocitySystemTov = 0;
			NEDVelocity[0] = NEDVelocity[1] = NEDVelocity[2] = 0; // North, East, Down
			ECEFVelocity[0] = ECEFVelocity[1] = ECEFVelocity[2] = 0; // X, Y, Z
			//AttitudeTov = 0;
			VehicleEulerAngles[0] = VehicleEulerAngles[1] = VehicleEulerAngles[2] = 0; // Roll, Pitch, True Heading
			WanderAngle = 0;
			DCM[0][0] = DCM[0][1] = DCM[0][2] = DCM[1][0] = DCM[1][1] = DCM[1][2] = DCM[2][0] = DCM[2][1] = DCM[2][2] = 0;
			VehicleBodyAngularRate[0] = VehicleBodyAngularRate[1] = VehicleBodyAngularRate[2] = 0; // X, Y, Z
			VehicleBodyAcceleration[0] = VehicleBodyAcceleration[1] = VehicleBodyAcceleration[2] = 0; // X, Y, Z
			AttitudeFom = 0;
			Quaternion[0] = Quaternion[1] = Quaternion[2] = Quaternion[3] = 0;
		}

		// == operator override
		bool operator==(const  Hg0x2402NavigationOutput& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (ECEFPosition[i] != Message.ECEFPosition[i]) return false;
				if (NEDVelocity[i] != Message.NEDVelocity[i]) return false;
				if (ECEFVelocity[i] != Message.ECEFVelocity[i]) return false;
				if (Quaternion[i] != Message.Quaternion[i]) return false;
				if (VehicleEulerAngles[i] != Message.VehicleEulerAngles[i]) return false;
				if (VehicleBodyAngularRate[i] != Message.VehicleBodyAngularRate[i]) return false;
				if (VehicleBodyAcceleration[i] != Message.VehicleBodyAcceleration[i]) return false;

				for (int x = 0; x < 3; x++) {if (DCM[i][x] != Message.DCM[i][x]) return false;}
			}
			if (Quaternion[3] != Message.Quaternion[3]) return false;

			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (INSMode != Message.INSMode) return false;
			if (GpsTov != Message.GpsTov) return false;
			if (SystemTov != Message.SystemTov) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			//if (UtcTimeFom != Message.UtcTimeFom) return false;
			if (GnssTimeFom != Message.GnssTimeFom) return false;
			//if (InsBlendedFom != Message.InsBlendedFom) return false;
			if (PositionSystemTov != Message.PositionSystemTov) return false;;
			if (Latitude != Message.Latitude) return false;
			if (Longitude != Message.Longitude) return false;
			if (AltitudeElips != Message.AltitudeElips) return false;
			if (AltitudeGeoid != Message.AltitudeGeoid) return false;
			if (VelocitySystemTov != Message.VelocitySystemTov) return false;
			//if (AttitudeTov != Message.AttitudeTov) return false;
			if (WanderAngle != Message.WanderAngle) return false;
			if (AttitudeFom != Message.AttitudeFom) return false;
			else 
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x2402NavigationOutput& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(ECEFPosition[i] - Message.ECEFPosition[i])>std::abs(Message.ECEFPosition[i]*Margin)) return false;
				if (std::abs(NEDVelocity[i] - Message.NEDVelocity[i])>std::abs(Message.NEDVelocity[i]*Margin)) return false;
				if (std::abs(ECEFVelocity[i] - Message.ECEFVelocity[i])>std::abs(Message.ECEFVelocity[i]*Margin)) return false;
				if (std::abs(Quaternion[i] - Message.Quaternion[i])>std::abs(Message.Quaternion[i]*Margin)) return false;
				if (std::abs(VehicleEulerAngles[i] - Message.VehicleEulerAngles[i])>std::abs(Message.VehicleEulerAngles[i]*Margin)) return false;
				if (std::abs(VehicleBodyAngularRate[i] - Message.VehicleBodyAngularRate[i])>std::abs(Message.VehicleBodyAngularRate[i]*Margin)) return false;
				if (std::abs(VehicleBodyAcceleration[i] - Message.VehicleBodyAcceleration[i])>std::abs(Message.VehicleBodyAcceleration[i]*Margin)) return false;

				for (int x = 0; x < 3; x++) { if (std::abs(DCM[i][x] - Message.DCM[i][x])>std::abs(Message.DCM[i][x]*Margin)) return false; }
			}
			if (std::abs(Quaternion[3] - Message.Quaternion[3])>std::abs(Message.Quaternion[3]*Margin)) return false;

			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (INSMode != Message.INSMode) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			//if (UtcTimeFom != Message.UtcTimeFom) return false;
			if (GnssTimeFom != Message.GnssTimeFom) return false;
			//if (InsBlendedFom != Message.InsBlendedFom) return false;
			if (std::abs(PositionSystemTov - Message.PositionSystemTov)>std::abs(Message.PositionSystemTov*Margin)) return false;;
			if (std::abs(Latitude - Message.Latitude)>std::abs(Message.Latitude*Margin)) return false;
			if (std::abs(Longitude - Message.Longitude)>std::abs(Message.Longitude*Margin)) return false;
			if (std::abs(AltitudeElips - Message.AltitudeElips)>std::abs(Message.AltitudeElips*Margin)) return false;
			if (std::abs(AltitudeGeoid - Message.AltitudeGeoid)>std::abs(Message.AltitudeGeoid*Margin)) return false;
			if (std::abs(VelocitySystemTov - Message.VelocitySystemTov)>std::abs(Message.VelocitySystemTov*Margin)) return false;
			//if (std::abs(AttitudeTov - Message.AttitudeTov)>std::abs(Message.AttitudeTov*Margin)) return false;
			if (std::abs(WanderAngle - Message.WanderAngle)>std::abs(Message.WanderAngle*Margin)) return false;
			if (AttitudeFom != Message.AttitudeFom) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,GpsWeek,INSMode,GnssTimeFomPositionSystemTov,Latitude,Longitude,AltitudeElips,AltitudeGeoid,ECEFPosition[0],ECEFPosition[1],ECEFPosition[2],VelocitySystemTov,NVelocity,EVelocity,DVelocity,ECEFVelocity[0],ECEFVelocity[1],ECEFVelocity[2],AttitudeTov,VehicleEulerAngles[0],VehicleEulerAngles[1],VehicleEulerAngles[2],WanderAngle,DCM[0][0],DCM[0][1],DCM[0][2],DCM[1][0],DCM[1][1],DCM[1][2],DCM[2][0],DCM[2][1],DCM[2][2],VehicleBodyAngularRate[0],VehicleBodyAngularRate[1],VehicleBodyAngularRate[2],VehicleBodyAcceleration[0],VehicleBodyAcceleration[1],VehicleBodyAcceleration[2],Quaternion[0],Quaternion[1],Quaternion[2],Quaternion[3]";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%d,%d,%d,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%d,%.12f,%.12f,%.12f,%.12f",
				SystemTov ,GpsTov , GpsWeek, INSMode, GnssTimeFom , PositionSystemTov ,
				Latitude , Longitude , AltitudeElips , AltitudeGeoid , ECEFPosition[0] , ECEFPosition[1] , ECEFPosition[2] ,
				VelocitySystemTov , NEDVelocity[0] , NEDVelocity[1] , NEDVelocity[2] , 
				ECEFVelocity[0] , ECEFVelocity[1] , ECEFVelocity[2] , 
				VehicleEulerAngles[0] , VehicleEulerAngles[1] , VehicleEulerAngles[2] ,
				WanderAngle , DCM[0][0] , DCM[0][1] , DCM[0][2] , DCM[1][0] , DCM[1][1] , DCM[1][2] , DCM[2][0] , DCM[2][1] , DCM[2][2] ,
				VehicleBodyAngularRate[0] , VehicleBodyAngularRate[1] , VehicleBodyAngularRate[2] , VehicleBodyAcceleration[0] , VehicleBodyAcceleration[1] , VehicleBodyAcceleration[2] ,
				AttitudeFom , Quaternion[0] , Quaternion[1] , Quaternion[2] , Quaternion[3]);
			return OutStr;
		}
	};
	
	//INS Configuration data
	struct Hg0x2001Configuration
	{
		char ImuSerialNumber[8]; //8x ASCII
		char ImuSwVersion[16]; //16x ASCII
		char HGuideSerialNumber[8]; //8x ASCII
		char HGuidePartNumber[8]; //8x ASCII
		char HGuideSwVersion[16]; //16x ASCII
		char HGuideSwBuildDate[16]; //16x ASCII

		float VehicleLeverArms[3]; // [m] X, Y, Z displacement from vehicle center
		float MainAntennaLeverArms[3]; // [m] X, Y, Z lever arms of main antenna to vehicle frame
		float AuxAntennaLeverArms[3]; // [m] X, Y, Z lever arms of aux antenna
		//Sets all values to zero / false
		void ZeroMessage()
		{
			for (int i = 0; i < 8; i++)
			{
				ImuSerialNumber[i] = 0;
				HGuideSerialNumber[i] = 0;
				HGuidePartNumber[i] = 0;
			}
			for (int i = 0; i < 16; i++)
			{
				ImuSwVersion[i] = 0;
				HGuideSwVersion[i] = 0;
				HGuideSwBuildDate[i] = 0;
			}
			VehicleLeverArms[0] = VehicleLeverArms[1] = VehicleLeverArms[2] = 0;
			MainAntennaLeverArms[0] = MainAntennaLeverArms[1] = MainAntennaLeverArms[2] = 0;
			AuxAntennaLeverArms[0] = AuxAntennaLeverArms[1] = AuxAntennaLeverArms[2] = 0;
		}

		// == operator override
		bool operator==(const Hg0x2001Configuration& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (VehicleLeverArms[i] != Message.VehicleLeverArms[i]) return false;
				if (MainAntennaLeverArms[i] != Message.MainAntennaLeverArms[i]) return false;
				if (AuxAntennaLeverArms[i] != Message.AuxAntennaLeverArms[i]) return false;
			}
			for (int i = 0; i < 8; i++)
			{
				if (ImuSerialNumber[i] != Message.ImuSerialNumber[i]) return false;
				if (HGuideSerialNumber[i] != Message.HGuideSerialNumber[i]) return false;
				if (HGuidePartNumber[i] != Message.HGuidePartNumber[i]) return false;
			}
			for (int i = 0; i < 16; i++)
			{
				if (ImuSwVersion[i] != Message.ImuSwVersion[i]) return false;
				if (HGuideSwVersion[i] != Message.HGuideSwVersion[i]) return false;
				if (HGuideSwBuildDate[i] != Message.HGuideSwBuildDate[i]) return false;
			}
			return true;
		}

		// limit test
		bool EqualWithMargin(const Hg0x2001Configuration& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VehicleLeverArms[i] - Message.VehicleLeverArms[i])>std::abs(Message.VehicleLeverArms[i]*Margin)) return false;
				if (std::abs(MainAntennaLeverArms[i] - Message.MainAntennaLeverArms[i])>std::abs(Message.MainAntennaLeverArms[i]*Margin)) return false;
				if (std::abs(AuxAntennaLeverArms[i] - Message.AuxAntennaLeverArms[i])>std::abs(Message.AuxAntennaLeverArms[i]*Margin)) return false;
			}
			for (int i = 0; i < 8; i++)
			{
				if (ImuSerialNumber[i] != Message.ImuSerialNumber[i]) return false;
				if (HGuideSerialNumber[i] != Message.HGuideSerialNumber[i]) return false;
				if (HGuidePartNumber[i] != Message.HGuidePartNumber[i]) return false;
			}
			for (int i = 0; i < 16; i++)
			{
				if (ImuSwVersion[i] != Message.ImuSwVersion[i]) return false;
				if (HGuideSwVersion[i] != Message.HGuideSwVersion[i]) return false;
				if (HGuideSwBuildDate[i] != Message.HGuideSwBuildDate[i]) return false;
			}
			return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "HGuidePartNumber,HGuideSerialNumber,HGuideSwVersion,HGuideSwBuildDate,ImuSerialNumber,ImuSwVersion,MainAntennaLeverArms[0],MainAntennaLeverArms[1],MainAntennaLeverArms[2],AuxAntennaLeverArms[0],AuxAntennaLeverArms[1],AuxAntennaLeverArms[2],VehicleLeverArms[0],VehicleLeverArms[1],VehicleLeverArms[2]";
		}
		
		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.8s,%.8s,%.16s,%.16s,%.8s,%.16s,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f", HGuideSerialNumber, HGuidePartNumber, HGuideSwVersion, HGuideSwBuildDate,ImuSerialNumber, ImuSwVersion, MainAntennaLeverArms[0], MainAntennaLeverArms[1], MainAntennaLeverArms[2], AuxAntennaLeverArms[0], AuxAntennaLeverArms[1], AuxAntennaLeverArms[2], VehicleLeverArms[0], VehicleLeverArms[1], VehicleLeverArms[2]);
			return OutStr;
		}
	};

	//Aiding data validity bits
	struct GnssAidingStatus
	{
		bool BaroAidingValid; // 0 = Invalid | 1 = Valid
		bool BaroAidingUse; // 0 = Not Used | 1 = In Use
		bool MagAidingValid; // 0 = Invalid | 1 = Valid
		bool MagAidingUse; // 0 = Not Used | 1 = In Use
	};

	//INS status bits
	struct InsWord2BitStatus
	{
		bool FirstStageBootLoader;
		bool FlashLoaderTable;
		bool RegisterInitializationTable;
	};

	//GNSS status bits
	struct GnssWord1BitStatus
	{
		bool GNSSFunction;
		bool GNSSCommunication;
		bool GNSSTimeMark;
		bool GNSST20Synchronization;
	};
	
	//INS General Status Message
	struct Hg0x2011Status
	{
		UINT32 NavigationFlag; // Factory Use
		double SystemTov; //[s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		INT32 GpsWeek; // [-] No. of current GPS week
		INT32 PowerCycleCount; //[-] Counter of power cycles
		double DeviceElapsedTime; // [s] Total online time of INS
		float InsDeviceTemperature; // [°C] Temperature of INS
		struct InsGnssBIT InsGnssBIT;
		INT32 InsFom; // INS Figure of Merit
		INT32 GnssFom; // GNSS Figure of Merit
		INT32 UtcFom; // UTC Time Figure of Merit
		struct EnabledMessages EnabledMessages;
		bool ImuBitStatus; // 0 = OK | 1 = Failed
		INT32 NumberOfSatellitesUsed;
		UINT32 PseudoRangeValidity;
		UINT32 DeltaRangeValidity;
		UINT32 SolutionConvergence; // 0 = Not Converged | 1 = Converged
		INT32 AttitudeFom; // Attitude Figure of Merit
		struct GnssAidingStatus GnssAidingStatus;
		struct InsWord2BitStatus InsWord2BitStatus;
		struct GnssWord1BitStatus GnssWord1BitStatus;
		float AccelTemperature; //[°C] Temperature of the Accels in IMU
		float GyroTemperature; //[°C] Temperature of the Gyros in IMU

		//Sets all values to zero / false
		void ZeroMessage()
		{
			NavigationFlag = 0;
			SystemTov = 0;
			GpsTov = 0;
			GpsWeek = 0;
			PowerCycleCount = 0;
			DeviceElapsedTime = 0;
			InsDeviceTemperature = 0;
			InsGnssBIT.ZeroMessage();
			InsFom = 0;
			GnssFom = 0;
			UtcFom = 0;
			EnabledMessages.ZeroMessage();
			ImuBitStatus = 0;
			NumberOfSatellitesUsed = 0;
			PseudoRangeValidity = 0;
			DeltaRangeValidity = 0;
			SolutionConvergence = 0;
			AttitudeFom = 0;
			GnssAidingStatus.BaroAidingUse = 0;
			GnssAidingStatus.BaroAidingValid = 0;
			GnssAidingStatus.MagAidingUse = 0;
			GnssAidingStatus.MagAidingValid = 0;
			InsWord2BitStatus.FirstStageBootLoader = 0;
			InsWord2BitStatus.FlashLoaderTable = 0;
			InsWord2BitStatus.RegisterInitializationTable = 0;
			GnssWord1BitStatus.GNSSCommunication = 0;
			GnssWord1BitStatus.GNSSFunction = 0;
			GnssWord1BitStatus.GNSST20Synchronization = 0;
			GnssWord1BitStatus.GNSSTimeMark = 0;
			AccelTemperature = 0;
			GyroTemperature = 0;
		}

		// == operator override
		bool operator==(const  Hg0x2011Status& Message) const
		{
			if (NavigationFlag != Message.NavigationFlag) return false;
			if (SystemTov != Message.SystemTov) return false;
			if (GpsTov != Message.GpsTov) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			if (PowerCycleCount != Message.PowerCycleCount) return false;
			if (DeviceElapsedTime != Message.DeviceElapsedTime) return false;
			if (InsDeviceTemperature != Message.InsDeviceTemperature) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (InsFom != Message.InsFom) return false;
			if (GnssFom != Message.GnssFom) return false;
			if (UtcFom != Message.UtcFom) return false;
			if (!(EnabledMessages == Message.EnabledMessages)) return false;
			if (ImuBitStatus != Message.ImuBitStatus) return false;
			if (NumberOfSatellitesUsed != Message.NumberOfSatellitesUsed) return false;
			if (PseudoRangeValidity != Message.PseudoRangeValidity) return false;
			if (DeltaRangeValidity != Message.DeltaRangeValidity) return false;
			if (SolutionConvergence != Message.SolutionConvergence) return false;
			if (AttitudeFom != Message.AttitudeFom) return false;
			if (GnssAidingStatus.BaroAidingUse != Message.GnssAidingStatus.BaroAidingUse) return false;
			if (GnssAidingStatus.BaroAidingValid != Message.GnssAidingStatus.BaroAidingValid) return false;
			if (GnssAidingStatus.MagAidingUse != Message.GnssAidingStatus.MagAidingUse) return false;
			if (GnssAidingStatus.MagAidingValid != Message.GnssAidingStatus.MagAidingValid) return false;
			if (InsWord2BitStatus.FirstStageBootLoader != Message.InsWord2BitStatus.FirstStageBootLoader) return false;
			if (InsWord2BitStatus.FlashLoaderTable != Message.InsWord2BitStatus.FlashLoaderTable) return false;
			if (InsWord2BitStatus.RegisterInitializationTable != Message.InsWord2BitStatus.RegisterInitializationTable) return false;
			if (GnssWord1BitStatus.GNSSCommunication != Message.GnssWord1BitStatus.GNSSCommunication) return false;
			if (GnssWord1BitStatus.GNSSFunction != Message.GnssWord1BitStatus.GNSSFunction) return false;
			if (GnssWord1BitStatus.GNSST20Synchronization != Message.GnssWord1BitStatus.GNSST20Synchronization) return false;
			if (GnssWord1BitStatus.GNSSTimeMark != Message.GnssWord1BitStatus.GNSSTimeMark) return false;
			if (AccelTemperature != Message.AccelTemperature) return false;
			if (GyroTemperature != Message.GyroTemperature) return false;
			else
				return true;
		}

		// limit test
		bool EqualWithMargin(const Hg0x2011Status& Message, double Margin) const
		{
			if (NavigationFlag != Message.NavigationFlag) return false;
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			if (PowerCycleCount != Message.PowerCycleCount) return false;
			if (std::abs(DeviceElapsedTime - Message.DeviceElapsedTime)>std::abs(Message.DeviceElapsedTime*Margin)) return false;
			if (std::abs(InsDeviceTemperature - Message.InsDeviceTemperature)>std::abs(Message.InsDeviceTemperature*Margin)) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (InsFom != Message.InsFom) return false;
			if (GnssFom != Message.GnssFom) return false;
			if (UtcFom != Message.UtcFom) return false;
			if (!(EnabledMessages == Message.EnabledMessages)) return false;
			if (ImuBitStatus != Message.ImuBitStatus) return false;
			if (NumberOfSatellitesUsed != Message.NumberOfSatellitesUsed) return false;
			if (PseudoRangeValidity != Message.PseudoRangeValidity) return false;
			if (DeltaRangeValidity != Message.DeltaRangeValidity) return false;
			if (SolutionConvergence != Message.SolutionConvergence) return false;
			if (AttitudeFom != Message.AttitudeFom) return false;
			if (GnssAidingStatus.BaroAidingUse != Message.GnssAidingStatus.BaroAidingUse) return false;
			if (GnssAidingStatus.BaroAidingValid != Message.GnssAidingStatus.BaroAidingValid) return false;
			if (GnssAidingStatus.MagAidingUse != Message.GnssAidingStatus.MagAidingUse) return false;
			if (GnssAidingStatus.MagAidingValid != Message.GnssAidingStatus.MagAidingValid) return false;
			if (InsWord2BitStatus.FirstStageBootLoader != Message.InsWord2BitStatus.FirstStageBootLoader) return false;
			if (InsWord2BitStatus.FlashLoaderTable != Message.InsWord2BitStatus.FlashLoaderTable) return false;
			if (InsWord2BitStatus.RegisterInitializationTable != Message.InsWord2BitStatus.RegisterInitializationTable) return false;
			if (GnssWord1BitStatus.GNSSCommunication != Message.GnssWord1BitStatus.GNSSCommunication) return false;
			if (GnssWord1BitStatus.GNSSFunction != Message.GnssWord1BitStatus.GNSSFunction) return false;
			if (GnssWord1BitStatus.GNSST20Synchronization != Message.GnssWord1BitStatus.GNSST20Synchronization) return false;
			if (GnssWord1BitStatus.GNSSTimeMark != Message.GnssWord1BitStatus.GNSSTimeMark) return false;
			if (std::abs(AccelTemperature - Message.AccelTemperature)>std::abs(Message.AccelTemperature*Margin)) return false;
			if (std::abs(GyroTemperature - Message.GyroTemperature)>std::abs(Message.GyroTemperature*Margin)) return false;
			else
				return true;
		}

		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,NavigationFlag,GpsWeek,PowerCycleCount,DeviceElapsedTime,InsDeviceTemperature,\
					InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity,\
					InsFom,GnssFom,UtcFom,ImuBitStatus,SolutionConvergence,AttitudeFom,GnssAidingStatus.BaroAidingUse,GnssAidingStatus.BaroAidingValid,GnssAidingStatus.MagAidingUse,GnssAidingStatus.MagAidingValid,AccelTemperature,GyroTemperature,EnabledMessages.MessageWord1,EnabledMessages.MessageWord2";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%d,%d,%d,%.12f,%.12f,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%d,%.12f,%.12f,%08X,%08X", SystemTov, GpsTov, NavigationFlag, GpsWeek, PowerCycleCount, DeviceElapsedTime,
				InsDeviceTemperature, (int)InsGnssBIT.InsMode, (int)InsGnssBIT.GpsMode, InsGnssBIT.InsStatus, InsGnssBIT.GnssStatus, InsGnssBIT.ImuStatus,(int)InsGnssBIT.MotionDetectActive ,(int)InsGnssBIT.ZeroVelocityPending,(int)InsGnssBIT.ZeroVelocity,
				InsFom, GnssFom, UtcFom, ImuBitStatus, SolutionConvergence, AttitudeFom, GnssAidingStatus.BaroAidingUse,	
				GnssAidingStatus.BaroAidingValid,GnssAidingStatus.MagAidingUse,GnssAidingStatus.MagAidingValid,AccelTemperature,
				GyroTemperature,EnabledMessages.MessageWord1,EnabledMessages.MessageWord2);
			return OutStr;
		}
	};

	//Message Ack / Nack notification
	struct Hg0x20ffAck
	{
		UINT32 Ack; // 0 = NAK | 1 = ACK
		UINT32 InputMessageID; // ID of the incoming message
		UINT32 NoOfValidMessagesSinceLast; // [-] Number of valid input messages since last 20FF
		UINT32 NoOfValidMessagesSincePowerUp; // [-] Number of valid input messages since powerup
		double MessageTimeOfReception; // [s] in reference to the system time
		//Sets all values to zero / false
		void ZeroMessage()
		{
			Ack = 0; 
			InputMessageID = 0;
			NoOfValidMessagesSinceLast = 0;
			NoOfValidMessagesSincePowerUp = 0;
			MessageTimeOfReception = 0;
		}

		// == operator override
		bool operator==(const  Hg0x20ffAck& Message) const
		{
			if (Ack != Message.Ack) return false;
			if (InputMessageID != Message.InputMessageID) return false;
			if (NoOfValidMessagesSinceLast != Message.NoOfValidMessagesSinceLast) return false;
			if (NoOfValidMessagesSincePowerUp != Message.NoOfValidMessagesSincePowerUp) return false;
			if (MessageTimeOfReception != Message.MessageTimeOfReception) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x20ffAck& Message, double Margin) const
		{
			if (Ack != Message.Ack) return false;
			if (InputMessageID != Message.InputMessageID) return false;
			if (NoOfValidMessagesSinceLast != Message.NoOfValidMessagesSinceLast) return false;
			if (NoOfValidMessagesSincePowerUp != Message.NoOfValidMessagesSincePowerUp) return false;
			if (std::abs(MessageTimeOfReception - Message.MessageTimeOfReception)>std::abs(Message.MessageTimeOfReception*Margin)) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "MessageTimeOfReception,InputMessageID [Hex],Ack,NoOfValidMessagesSinceLast,NoOfValidMessagesSincePowerUp";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.4x,%d,%d,%d", MessageTimeOfReception,InputMessageID ,Ack , NoOfValidMessagesSinceLast, NoOfValidMessagesSincePowerUp);
			return OutStr;
		}
	};

	//Status if the provided time is correct
	struct TimeValidityBits
	{
		bool GnssTime; // 0 = Invalid | 1 = Valid
		bool UtcTime; // 0 = Invalid | 1 = Valid
		UINT32 HardwarePulse;  // 0x000 = Disabled | 0x100 = INS System Time
	};

	//INS Time mark data
	struct Hg0x2201TimeMark
	{
		double EventInSystemTov; //[s] Time since INS power up
		double EventInGpsTov; //[s] GPS time in current week (Sunday 00:00 UTC)
		INT32 EventInCount; // [-] number of incoming pulses counter

		struct TimeValidityBits TimeValidityBits;
		INT32 UtcTimeFom; // UTC Time Figure of Merit
		INT32 GpsWeek; // [-] No. of current GPS week
		double PpsSystemTov; // [s] PPS output time in System Time
		double PpsGpsTov; // [s] PPS output time in GPS Time
		
		//Sets all values to zero / false
		void ZeroMessage()
		{
			TimeValidityBits.GnssTime=0;
			TimeValidityBits.HardwarePulse = 0;
			TimeValidityBits.UtcTime = 0;
			UtcTimeFom = 0;
			EventInSystemTov = 0;
			EventInGpsTov = 0;
			GpsWeek = 0;
			PpsSystemTov=0;
			PpsGpsTov=0;
			EventInCount=0;
		}

		// == operator override
		bool operator==(const  Hg0x2201TimeMark& Message) const
		{
			if (TimeValidityBits.GnssTime != Message.TimeValidityBits.GnssTime) return false;
			if (TimeValidityBits.HardwarePulse != Message.TimeValidityBits.HardwarePulse) return false;
			if (TimeValidityBits.UtcTime != Message.TimeValidityBits.UtcTime) return false;
			if (UtcTimeFom != Message.UtcTimeFom) return false;
			if (EventInSystemTov != Message.EventInSystemTov) return false;
			if (EventInGpsTov != Message.EventInGpsTov) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			if (PpsSystemTov != Message.PpsSystemTov) return false;
			if (PpsGpsTov != Message.PpsGpsTov) return false;
			if (EventInCount != Message.EventInCount) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x2201TimeMark& Message, double Margin) const
		{
			if (TimeValidityBits.GnssTime != Message.TimeValidityBits.GnssTime) return false;
			if (TimeValidityBits.HardwarePulse != Message.TimeValidityBits.HardwarePulse) return false;
			if (TimeValidityBits.UtcTime != Message.TimeValidityBits.UtcTime) return false;
			if (UtcTimeFom != Message.UtcTimeFom) return false;
			if (std::abs(EventInSystemTov - Message.EventInSystemTov)>std::abs(Message.EventInSystemTov*Margin)) return false;
			if (std::abs(EventInGpsTov - Message.EventInGpsTov)>std::abs(Message.EventInGpsTov*Margin)) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			if (PpsSystemTov != Message.PpsSystemTov) return false;
			if (PpsGpsTov != Message.PpsGpsTov) return false;
			if (EventInCount != Message.EventInCount) return false;
			else
				return true;
		}
		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "EventInSystemTov,EventIntGpsTov,EventInCount,GpsWeek,UtcTimeFom,PpsSystemTov,PpsGpsTov,GnssTimeValid,UtcTimeValid,HwPulseMode";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%d,%d,%d,%.12f,%.12f,%d,%d,%d",
				EventInSystemTov , EventInGpsTov, EventInCount,GpsWeek, UtcTimeFom ,PpsSystemTov ,
				PpsGpsTov  ,TimeValidityBits.GnssTime ,TimeValidityBits.UtcTime ,TimeValidityBits.HardwarePulse);
			return OutStr;
		}
	};

	//Raw Attitude from the Gnss receiver
	struct Hg0x6109GnssAttitude
	{
		double SystemTov; //[s] Time since INS power up
		double GpsTov; //[s] GPS time in current week (Sunday 00:00 UTC)
		float GnssAttitude[2]; // [rad] Pitch, True Heading
		float GnssAttitudeStdv[2]; // [rad] STDV Pitch, True Heading
		UINT32 HeadingValid; // 0 - Invalid | 1 - Valid
		//struct InsGnssBIT InsGnssBIT;
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			GnssAttitude[0] = GnssAttitude[1]  = 0;
			GnssAttitudeStdv[0] = GnssAttitudeStdv[1] = 0;
			HeadingValid = 0;
			//InsGnssBIT.ZeroMessage();
		}

		// == operator override
		bool operator==(const  Hg0x6109GnssAttitude& Message) const
		{
			if (SystemTov != Message.SystemTov) return false;
			if (GpsTov != Message.GpsTov) return false;
			for (int i = 0; i < 2; i++)
			{
				if (GnssAttitude[i] != Message.GnssAttitude[i]) return false;
				if (GnssAttitudeStdv[i] != Message.GnssAttitudeStdv[i]) return false;
			}
			if (HeadingValid != Message.HeadingValid) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6109GnssAttitude& Message, double Margin) const
		{
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			for (int i = 0; i < 2; i++)
			{
				if (std::abs(GnssAttitude[i] - Message.GnssAttitude[i])>std::abs(Message.GnssAttitude[i]*Margin)) return false;
				if (std::abs(GnssAttitudeStdv[i] - Message.GnssAttitudeStdv[i])>std::abs(Message.GnssAttitudeStdv[i]*Margin)) return false;
			}
			if (HeadingValid != Message.HeadingValid) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,GnssAttitude[0],GnssAttitude[1],GnssAttitudeStdv[0],GnssAttitudeStdv[1],HeadingValid"/*,\
			InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive ,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity"*/;
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%d"/*,%d,%d,%d,%d,%d,%d,%d,%d"*/, SystemTov,GpsTov,GnssAttitude[0],GnssAttitude[1],GnssAttitudeStdv[0],GnssAttitudeStdv[1],HeadingValid
				/*,InsGnssBIT.InsMode,InsGnssBIT.GpsMode,(int)InsGnssBIT.InsStatus,(int)InsGnssBIT.GnssStatus,(int)InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity*/);
			return OutStr;
		}
	};

	//Raw position from the Gnss receiver
	struct Hg0x6108GnssPosition
	{
		double SystemTov;// [s] Time since INS power up	
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		UINT32 GpsWeek;	// [-] No. of current GPS week
		//struct InsGnssBIT InsGnssBIT;
		double Latitude; // [rad] latitude position
		double Longitude; // [rad] longitude position
		double AltitudeAboveEllipsoid; // [m] altitude above ellipsoid
		float VelocityNED[3]; // [m/s] North, East, Down velocity
		float RxClkBias;
		UINT32 PVT_comp;
		UINT32 corr_info;
		UINT32 signal_information;
		UINT32 PPP_info;
		float GnssStdvLat; // [rad] Latitude position STDV
		float GnssStdvLon; // [rad] Longitude position STDV
		float GnssStdvAlt; // [m] Altitude position STDV
		float GnssStdvNEDVelocity[3]; // [m/s] North, East, Down Velocity STDV
	
		//Sets all values to zero / false
		void ZeroMessage()
		{
			GpsTov =0;
			GpsWeek = 0;
			//InsGnssBIT.ZeroMessage();
			Latitude = 0;
			Longitude = 0;
			AltitudeAboveEllipsoid = 0;
			VelocityNED[0] = VelocityNED[1] = VelocityNED[2] = 0;
			RxClkBias = 0;
			PVT_comp = 0;
			corr_info = 0;
			signal_information = 0;
			PPP_info = 0;
			GnssStdvLat = 0;
			GnssStdvLon = 0;
			GnssStdvAlt = 0;
			GnssStdvNEDVelocity[0] = GnssStdvNEDVelocity[1] = GnssStdvNEDVelocity[2] = 0;
			SystemTov = 0;
		}

		// == operator override
		bool operator==(const  Hg0x6108GnssPosition& Message) const
		{
			if (GpsTov != Message.GpsTov ) return false;
			if (GpsWeek != Message.GpsWeek ) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (Latitude != Message.Latitude ) return false;
			if (Longitude != Message.Longitude ) return false;
			if (AltitudeAboveEllipsoid != Message.AltitudeAboveEllipsoid ) return false;
			for (int i=0; i < 3; i++)
			{
				if (GnssStdvNEDVelocity[i] != Message.GnssStdvNEDVelocity[i])return false;
				if (VelocityNED[i] != Message.VelocityNED[i]) return false;
			}
			if (RxClkBias != Message.RxClkBias ) return false;
			if (PVT_comp != Message.PVT_comp ) return false;
			if (corr_info != Message.corr_info ) return false;
			if (signal_information != Message.signal_information ) return false;
			if (PPP_info != Message.PPP_info ) return false;
			if (GnssStdvLat != Message.GnssStdvLat ) return false;
			if (GnssStdvLon != Message.GnssStdvLon ) return false;
			if (GnssStdvAlt != Message.GnssStdvAlt ) return false;
			if (SystemTov != Message.SystemTov) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6108GnssPosition& Message, double Margin) const
		{
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (std::abs(Latitude - Message.Latitude)>std::abs(Message.Latitude*Margin)) return false;
			if (std::abs(Longitude - Message.Longitude)>std::abs(Message.Longitude*Margin)) return false;
			if (std::abs(AltitudeAboveEllipsoid - Message.AltitudeAboveEllipsoid)>std::abs(Message.AltitudeAboveEllipsoid*Margin)) return false;
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VelocityNED[i] - Message.VelocityNED[i])>std::abs(Message.VelocityNED[i]*Margin)) return false;
				if (std::abs(GnssStdvNEDVelocity[i] - Message.GnssStdvNEDVelocity[i])>std::abs(Message.GnssStdvNEDVelocity[i]*Margin)) return false;
			}
			if (std::abs(RxClkBias - Message.RxClkBias)>std::abs(Message.RxClkBias*Margin)) return false;
			if (PVT_comp != Message.PVT_comp) return false;
			if (corr_info != Message.corr_info) return false;
			if (signal_information != Message.signal_information) return false;
			if (PPP_info != Message.PPP_info) return false;
			if (std::abs(GnssStdvLat - Message.GnssStdvLat)>std::abs(Message.GnssStdvLat*Margin)) return false;
			if (std::abs(GnssStdvLon - Message.GnssStdvLon)>std::abs(Message.GnssStdvLon*Margin)) return false;
			if (std::abs(GnssStdvAlt - Message.GnssStdvAlt)>std::abs(Message.GnssStdvAlt*Margin)) return false;	
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			else
				return true;
		}
		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,GpsWeek,Latitude,Longitude,AltitudeAboveEllipsoid,VelocityN,VelocityE,VelocityD,RxClkBias,PVT_comp,corr_info,signal_information,PPP_info,\
					GnssStdvLat,GnssStdvLon,GnssStdvAlt,GnssStdvN,GnssStdvE,GnssStdvD"/*,\
					InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity"*/;
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%d,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%f,%d,%d,%d,%d,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f"/*,%d,%d,%d,%d,%d,%d,%d,%d"*/, SystemTov,GpsTov,GpsWeek,Latitude,Longitude,AltitudeAboveEllipsoid,
				VelocityNED[0],VelocityNED[1],VelocityNED[2],
				RxClkBias,PVT_comp,corr_info,signal_information,PPP_info,
				GnssStdvLat,GnssStdvLon,GnssStdvAlt, GnssStdvNEDVelocity[0], GnssStdvNEDVelocity[1], GnssStdvNEDVelocity[2]
				/*,InsGnssBIT.InsMode,InsGnssBIT.GpsMode,(int)InsGnssBIT.InsStatus,(int)InsGnssBIT.GnssStatus,(int)InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity*/);
			return OutStr;
		}
	};

	//INS Inertial Data
	struct Hg0x2311InertialData
	{
		double SystemTov; // [s] Time since INS power up
		//struct InsGnssBIT InsGnssBIT;
		float DeltaTheta[3]; // [rad] change in angle along X, Y, Z axes (FS = 0.25 Rads)
		float DeltaVelocity[3]; // [m/s] change in speed in X, Y, Z axes (FS = 4 m/s)
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			//InsGnssBIT.ZeroMessage();
			DeltaTheta[0] = DeltaTheta[1] = DeltaTheta[2] = 0;
			DeltaVelocity[0] = DeltaVelocity[1] = DeltaVelocity[2] = 0;
		}
		
		// == operator override
		bool operator==(const  Hg0x2311InertialData& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (DeltaTheta[i] != Message.DeltaTheta[i]) return false;
				if (DeltaVelocity[i] != Message.DeltaVelocity[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x2311InertialData& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(DeltaTheta[i] - Message.DeltaTheta[i])>std::abs(Message.DeltaTheta[i]*Margin)) return false;
				if (std::abs(DeltaVelocity[i] - Message.DeltaVelocity[i])>std::abs(Message.DeltaVelocity[i]*Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			//if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,DeltaTheta[0],DeltaTheta[1],DeltaTheta[2],DeltaVelocity[0],DeltaVelocity[1],DeltaVelocity[2]"/*,\
			InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity"*/;
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f"/*,%d,%d,%d,%d,%d,%d,%d,%d"*/, SystemTov,DeltaTheta[0],DeltaTheta[1],DeltaTheta[2],
				 DeltaVelocity[0],DeltaVelocity[1],DeltaVelocity[2]/*,
				 InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity*/);
			return OutStr;
		}
	};

	//INS NED Velocity data
	struct Hg0x6504NEDVelocity
	{
		double SystemTov; //[s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		float VelocityNED[3]; //[m/s] North, East, Down velocity
		struct InsGnssBIT InsGnssBIT;
		float VelocityStdvNED[3]; // [m/s] North, East, Down velocity STDV
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			VelocityNED[0] = VelocityNED[1] = VelocityNED[2] = 0;
			InsGnssBIT.ZeroMessage();
			VelocityStdvNED[0] = VelocityStdvNED[1] = VelocityStdvNED[2] = 0;
		}

		// == operator override
		bool operator==(const  Hg0x6504NEDVelocity& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (VelocityNED[i] != Message.VelocityNED[i]) return false;
				if (VelocityStdvNED[i] != Message.VelocityStdvNED[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (GpsTov != Message.GpsTov) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6504NEDVelocity& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(VelocityNED[i] - Message.VelocityNED[i])>std::abs(Message.VelocityNED[i]*Margin)) return false;
				if (std::abs(VelocityStdvNED[i] - Message.VelocityStdvNED[i])>std::abs(Message.VelocityStdvNED[i]*Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,VelocityNED[0],VelocityNED[1],VelocityNED[2],VelocityStdvNED[0],VelocityStdvNED[1],VelocityStdvNED[2],\
					InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%d,%d,%d,%d,%d,%d,%d,%d", SystemTov,GpsTov,VelocityNED[0],VelocityNED[1],VelocityNED[2],
				VelocityStdvNED[0],VelocityStdvNED[1],VelocityStdvNED[2],
				InsGnssBIT.InsMode, InsGnssBIT.GpsMode, (int) InsGnssBIT.InsStatus, (int) InsGnssBIT.GnssStatus, (int) InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity);
			return OutStr;
		}
	};

	//INS attitude data
	struct Hg0x6405EulerAttitudes
	{
		double SystemTov; // [s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		float EulerAttitude[3]; // [rad] Roll, Pitch, True Heading
		struct InsGnssBIT InsGnssBIT;
		float EulerAttitudeStdv[3]; // [rad] STDV Roll, Pitch, True Heading
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			EulerAttitude[0] = EulerAttitude[1] = EulerAttitude[2] = 0;
			InsGnssBIT.ZeroMessage();
			EulerAttitudeStdv[0] = EulerAttitudeStdv[2] = EulerAttitudeStdv[1] = 0;
		}

		// == operator override
		bool operator==(const  Hg0x6405EulerAttitudes& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (EulerAttitude[i] != Message.EulerAttitude[i]) return false;
				if (EulerAttitudeStdv[i] != Message.EulerAttitudeStdv[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (GpsTov != Message.GpsTov) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6405EulerAttitudes& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(EulerAttitude[i] - Message.EulerAttitude[i])>std::abs(Message.EulerAttitude[i]*Margin)) return false;
				if (std::abs(EulerAttitudeStdv[i] - Message.EulerAttitudeStdv[i])>std::abs(Message.EulerAttitudeStdv[i]*Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,EulerAttitude[0],EulerAttitude[1],EulerAttitude[2],EulerAttitudeStdv[0],EulerAttitudeStdv[1],EulerAttitudeStdv[2],\
					InsGnssBIT.InsMode,InsGnssBIT.GpsMode,InsGnssBIT.InsStatus,InsGnssBIT.GnssStatus,InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%.12f,%d,%d,%d,%d,%d,%d,%d,%d", 
				SystemTov, GpsTov, EulerAttitude[0], EulerAttitude[1], EulerAttitude[2],
				EulerAttitudeStdv[0], EulerAttitudeStdv[1], EulerAttitudeStdv[2],
				 InsGnssBIT.InsMode, InsGnssBIT.GpsMode, (int)InsGnssBIT.InsStatus, (int)InsGnssBIT.GnssStatus, (int)InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity);
			return OutStr;
		}
	};

	// INS Geodetic position
	struct Hg0x6403GeodeticPosition
	{
		double SystemTov; // [s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		double Latitude; // [rad] latitude position
		double Longitude; // [rad] longitude position
		double AltitudeAboveEllipsoid; // [m] altitude above ellipsoid
		struct InsGnssBIT InsGnssBIT;
		float StdvNED[3]; // [m] North, East, Down
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			Latitude = 0;
			Longitude = 0;
			AltitudeAboveEllipsoid = 0;
			InsGnssBIT.ZeroMessage();
			StdvNED[0] = StdvNED[1] = StdvNED[2] = 0;
		}

		// == operator override
		bool operator==(const Hg0x6403GeodeticPosition& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (StdvNED[i] != Message.StdvNED[i]) return false;
			}
			if (SystemTov != Message.SystemTov ) return false;
			if (GpsTov != Message.GpsTov ) return false;
			if (Latitude != Message.Latitude ) return false;
			if (Longitude != Message.Longitude ) return false;
			if (AltitudeAboveEllipsoid != Message.AltitudeAboveEllipsoid ) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6403GeodeticPosition& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(StdvNED[i] - Message.StdvNED[i])>std::abs(Message.StdvNED[i]*Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (std::abs(Latitude - Message.Latitude)>std::abs(Message.Latitude*Margin)) return false;
			if (std::abs(Longitude - Message.Longitude)>std::abs(Message.Longitude*Margin)) return false;
			if (std::abs(AltitudeAboveEllipsoid - Message.AltitudeAboveEllipsoid)>std::abs(Message.AltitudeAboveEllipsoid*Margin)) return false;
			if (InsGnssBIT != Message.InsGnssBIT) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,Latitude,Longitude,AltitudeAboveEllipsoid,InsGnssBIT.InsMode, InsGnssBIT.GpsMode, InsGnssBIT.InsStatus, InsGnssBIT.GnssStatus, InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity,StdvNED[0],StdvNED[1],StdvNED[2]";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr,length, "%.12f,%.12f,%.12f,%.12f,%.12f,%d,%d,%d,%d,%d,%d,%d,%d,%.12f,%.12f,%.12f", 
				SystemTov, GpsTov, Latitude, Longitude, AltitudeAboveEllipsoid,
				InsGnssBIT.InsMode, InsGnssBIT.GpsMode, (int)InsGnssBIT.InsStatus, (int)InsGnssBIT.GnssStatus, (int)InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity,
				StdvNED[0], StdvNED[1], StdvNED[2]);
			return OutStr;

		}
	};

	// Distance traveled based on 0x4110 Velocity Input message
	struct Hg0x6110DistanceTraveled
	{
		double SystemTov; // [s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		float DistanceTraveled[3]; // [m] X, Y, Z Distance traveled
		UINT32 OdometerCumulativePulses; // [-] Number of pulses from start
		struct OdometerSettings OdometerSettings; // Settings of Speed and Odometer Aiding
		UINT32 Counter; // [-] Incremental Counter

		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			DistanceTraveled[0] = DistanceTraveled[1] = DistanceTraveled[2] = 0;
			OdometerCumulativePulses = 0;
			OdometerSettings.VehicleVelocityValid = false;
			OdometerSettings.TovMode = false;
			OdometerSettings.OdometerValid = false;
			OdometerSettings.AidingStatus = false;
			OdometerSettings.ZeroVelocityDetected = false;
			Counter = 0;
		}

		// == operator override
		bool operator==(const Hg0x6110DistanceTraveled& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (DistanceTraveled[i] != Message.DistanceTraveled[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			if (GpsTov != Message.GpsTov) return false;
			if (OdometerCumulativePulses != Message.OdometerCumulativePulses) return false;
			if (OdometerSettings.VehicleVelocityValid!= Message.OdometerSettings.VehicleVelocityValid) return false;
			if (OdometerSettings.TovMode != Message.OdometerSettings.TovMode) return false;
			if (OdometerSettings.OdometerValid != Message.OdometerSettings.OdometerValid) return false;
			if (OdometerSettings.AidingStatus!= Message.OdometerSettings.AidingStatus) return false;
			if (OdometerSettings.ZeroVelocityDetected!= Message.OdometerSettings.ZeroVelocityDetected) return false;
			if (Counter != Message.Counter) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6110DistanceTraveled& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(DistanceTraveled[i] - Message.DistanceTraveled[i])>std::abs(Message.DistanceTraveled[i] * Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (std::abs((long)OdometerCumulativePulses - (long)Message.OdometerCumulativePulses)>std::abs((long)Message.OdometerCumulativePulses*Margin)) return false;
			if (std::abs((long)Counter - (long)Message.Counter)>std::abs((long)Message.Counter*Margin)) return false;
			if (OdometerSettings.VehicleVelocityValid != Message.OdometerSettings.VehicleVelocityValid) return false;
			if (OdometerSettings.TovMode != Message.OdometerSettings.TovMode) return false;
			if (OdometerSettings.OdometerValid != Message.OdometerSettings.OdometerValid) return false;
			if (OdometerSettings.AidingStatus != Message.OdometerSettings.AidingStatus) return false;
			if (OdometerSettings.ZeroVelocityDetected != Message.OdometerSettings.ZeroVelocityDetected) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,DistanceTraveled[0],DistanceTraveled[1],DistanceTraveled[2],OdometerCumulativePulses,Counter,AidingStatus,VehicleVelocityValid,OdometerValid,ZeroVelocityDetected,TovMode";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr, length, "%.12f,%.12f,%.12f,%.12f,%.12f,%d,%d,%d,%d,%d,%d,%d",SystemTov, GpsTov, DistanceTraveled[0], DistanceTraveled[1], DistanceTraveled[2], OdometerCumulativePulses, Counter,
				OdometerSettings.AidingStatus, OdometerSettings.VehicleVelocityValid, OdometerSettings.OdometerValid, OdometerSettings.ZeroVelocityDetected, OdometerSettings.TovMode);
			return OutStr;
		}
	};

	// Results of Motion detection algorithm in kalman filter
	struct Hg0x6111MotionDetection
	{
		double SystemTov; // [s] Time since INS power up
		double TriggerTime; // [s] Time of zero motion detection in System Time
		//struct InsGnssBIT InsGnssBIT;
		double Latitude; // [rad] Latitude of ZUPT detection
		double Longitude; // [rad] Longitude of ZUPT detection

		//Valid for All Tests
		float MDTimeStationary;
		float MDTSettlingTime;
		float MD2NavigationValid;
		//Test 1
		float MD1Rotation[3]; // [rad] Roll, Pitch, Yaw Values of rotation
		float MDD1AngularRateTotal; // [rad/s] Actual value of sum of Angular Rates in all axes
		float MDT1Rotation; // [rad] Threshold of sum of rotation in all axes
		//Test 2
		float MDD2SpeedStdv; // [m/s] Actual sum of speed in all axes
		float MDT2SpeedStdv; // [m/s] Threshold of sum of speed in all axes
		//Test 3
		float MD3AngularRateInstant[3]; // [rad/s] X, Y, Z Angular Rate Filtered by Instant FN 3dB filter
		float MD3InstantFN3dB; // [Hz] Instant Filter for exit of ZUPT
		float MD3AngularRateNominal[3]; // [rad/s] X, Y, Z Angular Rate Filtered by Nominal FN 3dB filter
		float MD3NominalFN3dB; // [Hz] Nominal Filter for entering of ZUPT
		float MDD3AngularRate[3]; // [rad/s] X, Y, Z Actual Angular rate
		float MDT3AngularRateInstant; // [rad/s] Threshold of sum of angular rate in all axes
		//Test 4
		float MDD4LinearAcceleration; // [m/s2] Actual value of sum of linar accelerations in all axes
		float MDT4LinearAcceleration; // [m/s2] Threshold of sum of linear accelerations in all axes
		//Test 5
		float MDD5OdometerDeltaDistance; // [m] Change in distance based on Odometer Input
		float MDT5Odometer; // [m] Threshold of change in distance based on Odometer Input
		float MDOdometerTimeAtRest; // [s] Number of seconds when Odometer sent no pulses
		
		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			TriggerTime = 0;
			//InsGnssBIT.ZeroMessage();
			Latitude = 0;
			Longitude = 0;
			MDTimeStationary = 0;
			MDTSettlingTime = 0;
			MD2NavigationValid = 0;
			MD1Rotation[0]= MD1Rotation[1]= MD1Rotation[2] = 0;
			MDD1AngularRateTotal = 0;
			MDT1Rotation = 0;
			MDD2SpeedStdv = 0;
			MDT2SpeedStdv = 0;
			MD3AngularRateInstant[0] = MD3AngularRateInstant[1] = MD3AngularRateInstant[2] = 0;
			MD3InstantFN3dB = 0;
			MD3AngularRateNominal[0] = MD3AngularRateNominal[1] = MD3AngularRateNominal[2] = 0;
			MD3NominalFN3dB = 0;
			MDD3AngularRate[0] = MDD3AngularRate[1] = MDD3AngularRate[2] = 0;
			MDT3AngularRateInstant = 0;
			MDD4LinearAcceleration = 0;
			MDT4LinearAcceleration = 0;
			MDD5OdometerDeltaDistance = 0;
			MDT5Odometer = 0;
			MDOdometerTimeAtRest = 0;
		}

		// == operator override
		bool operator==(const Hg0x6111MotionDetection& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (MD1Rotation[i] != Message.MD1Rotation[i]) return false;
				if (MD3AngularRateInstant[i] != Message.MD3AngularRateInstant[i]) return false;
				if (MD3AngularRateNominal[i] != Message.MD3AngularRateNominal[i]) return false;
				if (MDD3AngularRate[i] != Message.MDD3AngularRate[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			if (TriggerTime != Message.TriggerTime) return false;
			if (Latitude != Message.Latitude) return false;
			if (Longitude != Message.Longitude) return false;
			if (MDTimeStationary != Message.MDTimeStationary) return false;
			if (MDTSettlingTime != Message.MDTSettlingTime) return false;
			if (MD2NavigationValid != Message.MD2NavigationValid) return false;
			if (MDD1AngularRateTotal != Message.MDD1AngularRateTotal) return false;
			if (MDT1Rotation != Message.MDT1Rotation) return false;
			if (MDD2SpeedStdv != Message.MDD2SpeedStdv) return false;
			if (MDT2SpeedStdv != Message.MDT2SpeedStdv) return false;
			if (MD3InstantFN3dB != Message.MD3InstantFN3dB) return false;
			if (MD3NominalFN3dB != Message.MD3NominalFN3dB) return false;
			if (MDT3AngularRateInstant != Message.MDT3AngularRateInstant) return false;
			if (MDD4LinearAcceleration != Message.MDD4LinearAcceleration) return false;
			if (MDT4LinearAcceleration != Message.MDT4LinearAcceleration) return false;
			if (MDD5OdometerDeltaDistance != Message.MDD5OdometerDeltaDistance) return false;
			if (MDT5Odometer != Message.MDT5Odometer) return false;
			if (MDOdometerTimeAtRest != Message.MDOdometerTimeAtRest) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6111MotionDetection& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(MD1Rotation[i] - Message.MD1Rotation[i])>std::abs(Message.MD1Rotation[i] * Margin)) return false;
				if (std::abs(MD3AngularRateInstant[i] - Message.MD3AngularRateInstant[i])>std::abs(Message.MD3AngularRateInstant[i] * Margin)) return false;
				if (std::abs(MD3AngularRateNominal[i] - Message.MD3AngularRateNominal[i])>std::abs(Message.MD3AngularRateNominal[i] * Margin)) return false;
				if (std::abs(MDD3AngularRate[i] - Message.MDD3AngularRate[i])>std::abs(Message.MDD3AngularRate[i] * Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(TriggerTime - Message.TriggerTime)>std::abs(Message.TriggerTime*Margin)) return false;
			if (std::abs(Latitude - Message.Latitude)>std::abs(Message.Latitude*Margin)) return false;
			if (std::abs(Longitude - Message.Longitude)>std::abs(Message.Longitude*Margin)) return false;
			if (std::abs(MDTimeStationary - Message.MDTimeStationary)>std::abs(Message.MDTimeStationary*Margin)) return false;
			if (std::abs(MDTSettlingTime - Message.MDTSettlingTime)>std::abs(Message.MDTSettlingTime*Margin)) return false;
			if (std::abs(MD2NavigationValid - Message.MD2NavigationValid)>std::abs(Message.MD2NavigationValid*Margin)) return false;
			if (std::abs(MDD1AngularRateTotal - Message.MDD1AngularRateTotal)>std::abs(Message.MDD1AngularRateTotal*Margin)) return false;
			if (std::abs(MDT1Rotation - Message.MDT1Rotation)>std::abs(Message.MDT1Rotation*Margin)) return false;
			if (std::abs(MDD2SpeedStdv - Message.MDD2SpeedStdv)>std::abs(Message.MDD2SpeedStdv*Margin)) return false;
			if (std::abs(MDT2SpeedStdv - Message.MDT2SpeedStdv)>std::abs(Message.MDT2SpeedStdv*Margin)) return false;
			if (std::abs(MD3InstantFN3dB - Message.MD3InstantFN3dB)>std::abs(Message.MD3InstantFN3dB*Margin)) return false;
			if (std::abs(MD3NominalFN3dB - Message.MD3NominalFN3dB)>std::abs(Message.MD3NominalFN3dB*Margin)) return false;
			if (std::abs(MDT3AngularRateInstant - Message.MDT3AngularRateInstant)>std::abs(Message.MDT3AngularRateInstant*Margin)) return false;
			if (std::abs(MDD4LinearAcceleration - Message.MDD4LinearAcceleration)>std::abs(Message.MDD4LinearAcceleration*Margin)) return false;
			if (std::abs(MDT4LinearAcceleration - Message.MDT4LinearAcceleration)>std::abs(Message.MDT4LinearAcceleration*Margin)) return false;
			if (std::abs(MDD5OdometerDeltaDistance - Message.MDD5OdometerDeltaDistance)>std::abs(Message.MDD5OdometerDeltaDistance*Margin)) return false;
			if (std::abs(MDT5Odometer - Message.MDT5Odometer)>std::abs(Message.MDT5Odometer*Margin)) return false;
			if (std::abs(MDOdometerTimeAtRest - Message.MDOdometerTimeAtRest)>std::abs(Message.MDOdometerTimeAtRest*Margin)) return false;

			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,TriggerTime,Latitude,Longitude,\
			InsGnssBIT.InsMode, InsGnssBIT.GpsMode, InsGnssBIT.InsStatus, InsGnssBIT.GnssStatus, InsGnssBIT.ImuStatus,InsGnssBIT.MotionDetectActive,InsGnssBIT.ZeroVelocityPending,InsGnssBIT.ZeroVelocity,\
			MDTimeStationary,MDTSettlingTime,MD2NavigationValid,\
			MD1Rotation[0],MD1Rotation[1],MD1Rotation[2],MDD1AngularRateTotal,MDT1Rotation,\
			MDD2SpeedStdv,MDT2SpeedStdv,\
			MD3AngularRateInstant[0],MD3AngularRateInstant[1],MD3AngularRateInstant[2],MD3InstantFN3dB,\
			MD3AngularRateNominal[0],MD3AngularRateNominal[1],MD3AngularRateNominal[2],MD3NominalFN3dB,\
			MDD3AngularRate[0],MDD3AngularRate[1],MDD3AngularRate[2],MDT3AngularRateInstant,\
			MDD4LinearAcceleration,MDT4LinearAcceleration,\
			MDD5OdometerDeltaDistance,MDT5Odometer,MDOdometerTimeAtRest";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr, length, "%.12f,%.12f,%.12f,%.12f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f", SystemTov, TriggerTime, Latitude, Longitude, \
				/*InsGnssBIT.InsMode, InsGnssBIT.GpsMode, (int)InsGnssBIT.InsStatus, (int)InsGnssBIT.GnssStatus, (int)InsGnssBIT.ImuStatus, (int)InsGnssBIT.MotionDetectActive, (int)InsGnssBIT.ZeroVelocityPending, (int)InsGnssBIT.ZeroVelocity,*/ \
				MDTimeStationary, MDTSettlingTime, MD2NavigationValid, \
				MD1Rotation[0], MD1Rotation[1], MD1Rotation[2], MDD1AngularRateTotal, MDT1Rotation, \
				MDD2SpeedStdv, MDT2SpeedStdv, \
				MD3AngularRateInstant[0],MD3AngularRateInstant[1],MD3AngularRateInstant[2], MD3InstantFN3dB, \
				MD3AngularRateNominal[0],MD3AngularRateNominal[1],MD3AngularRateNominal[2], MD3NominalFN3dB, \
				MDD3AngularRate[0],MDD3AngularRate[1],MDD3AngularRate[2], MDT3AngularRateInstant, \
				MDD4LinearAcceleration, MDT4LinearAcceleration, \
				MDD5OdometerDeltaDistance, MDT5Odometer, MDOdometerTimeAtRest);
			return OutStr;
		}
	};

	// Estimates of Antenna Lever Arms reported from Kalman Filter
	struct Hg0x6424AntennaLeverArmEstimates
	{
		double SystemTov; // [s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		UINT32 GpsWeek; // [-] No. of current GPS week

		float MainAntennaLeverArm[3]; // [m] X, Y, Z position of Main Antenna with respect to Navigation Frame
		float MainAntennaLeverArmStdv[3]; // [m] X, Y, Z standard deviations of Main Antenna position
		float AntennaBoresight[3]; // [rad] Roll, Pitch, Yaw Antenna Boresight - declination angle of antenna heading with respect to Navigation frame's Euler Angles
		float AntennaBoresightStdv[3]; // [rad] Roll, Pitch, Yaw standard deviations of Antenna Boresight angle

		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			GpsWeek = 0;

			MainAntennaLeverArm[0] = MainAntennaLeverArm[1] = MainAntennaLeverArm[2] = 0;
			MainAntennaLeverArmStdv[0] = MainAntennaLeverArmStdv[1] = MainAntennaLeverArmStdv[2] = 0;
			AntennaBoresight[0] = AntennaBoresight[1] = AntennaBoresight[2] = 0;
			AntennaBoresightStdv[0] = AntennaBoresightStdv[1] = AntennaBoresightStdv[2] = 0;
		}

		// == operator override
		bool operator==(const Hg0x6424AntennaLeverArmEstimates& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (MainAntennaLeverArm[i] != Message.MainAntennaLeverArm[i]) return false;
				if (MainAntennaLeverArmStdv[i] != Message.MainAntennaLeverArmStdv[i]) return false;
				if (AntennaBoresight[i] != Message.AntennaBoresight[i]) return false;
				if (AntennaBoresightStdv[i] != Message.AntennaBoresightStdv[i]) return false;
			}
			if (SystemTov != Message.SystemTov) return false;
			if (GpsTov != Message.GpsTov) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6424AntennaLeverArmEstimates& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(MainAntennaLeverArm[i] - Message.MainAntennaLeverArm[i])>std::abs(Message.MainAntennaLeverArm[i] * Margin)) return false;
				if (std::abs(MainAntennaLeverArmStdv[i] - Message.MainAntennaLeverArmStdv[i])>std::abs(Message.MainAntennaLeverArmStdv[i] * Margin)) return false;
				if (std::abs(AntennaBoresight[i] - Message.AntennaBoresight[i])>std::abs(Message.AntennaBoresight[i] * Margin)) return false;
				if (std::abs(AntennaBoresightStdv[i] - Message.AntennaBoresightStdv[i])>std::abs(Message.AntennaBoresightStdv[i] * Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (std::abs((long)GpsWeek - (long)Message.GpsWeek)>std::abs((long)Message.GpsWeek*Margin)) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov,GpsTov,GpsWeek,\
				MainAntennaLeverArm[0],MainAntennaLeverArm[1],MainAntennaLeverArm[2],\
				MainAntennaLeverArmStdv[0],MainAntennaLeverArmStdv[1],MainAntennaLeverArmStdv[2],\
				AntennaBoresight[0],AntennaBoresight[1],AntennaBoresight[2],\
				AntennaBoresightStdv[0],AntennaBoresightStdv[1],AntennaBoresightStdv[2]";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr, length, "%.12f,%.12f,%d,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f", SystemTov,GpsTov,GpsWeek,
				MainAntennaLeverArm[0],MainAntennaLeverArm[1],MainAntennaLeverArm[2],
				MainAntennaLeverArmStdv[0],MainAntennaLeverArmStdv[1],MainAntennaLeverArmStdv[2],
				AntennaBoresight[0],AntennaBoresight[1],AntennaBoresight[2],
				AntennaBoresightStdv[0],AntennaBoresightStdv[1],AntennaBoresightStdv[2]);
			return OutStr;
		}
	};

	// Odometer Calibration estimates from Kalman Filter
	struct Hg0x6438OdometerCalibration
	{
		double SystemTov; // [s] Time since INS power up
		double GpsTov; // [s] GPS time in current week (Sunday 00:00 UTC)
		UINT32 GpsWeek; // [-] No. of current GPS week

		float DistanceTraveled; // [m] Distance travelled based on Odometer input

		float ScaleFactorCorrection; // [<0;1>] Correction of Scalefactor (1 = 100% correction)
		float ScaleFactorStdv; // [] Standard deviation of Scalefactor correction

		float LeverArms[3]; // [m] X, Y, Z Estimated position of Odometer with respect to Navigation Frame
		float LeverArmsStdv[3]; // [m] X, Y, Z standard deviations of Odometer position
		float StoredLeverArms[3]; // [m] X, Y, Z Initially Stored position of Odometer with respect to Navigation Frame

		float Boresight[2]; // [rad] Yaw, Pitch Odometer Boresight with respect to Navigation Frame
		float BoresightStdv[2]; // [rad] Yaw, Pitch standard deviation of Odometer Boresight

		//Sets all values to zero / false
		void ZeroMessage()
		{
			SystemTov = 0;
			GpsTov = 0;
			GpsWeek = 0;

			LeverArms[0] = LeverArms[1] = LeverArms[2] = 0;
			LeverArmsStdv[0] = LeverArmsStdv[1] = LeverArmsStdv[2] = 0;
			StoredLeverArms[0] = StoredLeverArms[1] = StoredLeverArms[2] = 0;
			Boresight[0] = Boresight[1]= 0;
			BoresightStdv[0] = BoresightStdv[1] = 0;
		}

		// == operator override
		bool operator==(const Hg0x6438OdometerCalibration& Message) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (LeverArms[i] != Message.LeverArms[i]) return false;
				if (LeverArmsStdv[i] != Message.LeverArmsStdv[i]) return false;
				if (StoredLeverArms[i] != Message.StoredLeverArms[i]) return false;
			}
			for (int i = 0; i < 2; i++)
			{
				if (Boresight[i] != Message.Boresight[i]) return false;
				if (BoresightStdv[i] != Message.BoresightStdv[i]) return false;
			}

			if (SystemTov != Message.SystemTov) return false;
			if (GpsTov != Message.GpsTov) return false;
			if (GpsWeek != Message.GpsWeek) return false;
			else
				return true;
		}
		// limit test
		bool EqualWithMargin(const Hg0x6438OdometerCalibration& Message, double Margin) const
		{
			for (int i = 0; i < 3; i++)
			{
				if (std::abs(LeverArms[i] - Message.LeverArms[i])>std::abs(Message.LeverArms[i] * Margin)) return false;
				if (std::abs(LeverArmsStdv[i] - Message.LeverArmsStdv[i])>std::abs(Message.LeverArmsStdv[i] * Margin)) return false;
				if (std::abs(StoredLeverArms[i] - Message.StoredLeverArms[i])>std::abs(Message.StoredLeverArms[i] * Margin)) return false;
			}
			for (int i = 0; i < 2; i++)
			{
				if (std::abs(Boresight[i] - Message.Boresight[i])>std::abs(Message.Boresight[i] * Margin)) return false;
				if (std::abs(BoresightStdv[i] - Message.BoresightStdv[i])>std::abs(Message.BoresightStdv[i] * Margin)) return false;
			}
			if (std::abs(SystemTov - Message.SystemTov)>std::abs(Message.SystemTov*Margin)) return false;
			if (std::abs(GpsTov - Message.GpsTov)>std::abs(Message.GpsTov*Margin)) return false;
			if (std::abs((long)GpsWeek - (long)Message.GpsWeek)>std::abs((long)Message.GpsWeek*Margin)) return false;
			else
				return true;
		}

		//Print a Header to the CSV file
		const char* printHeaderToCsv()
		{
			return "SystemTov, GpsTov, GpsWeek,\
				DistanceTraveled,ScaleFactorCorrection,ScaleFactorStdv,\
				LeverArms[0],LeverArms[1],LeverArms[2],\
				LeverArmsStdv[0],LeverArmsStdv[1],LeverArmsStdv[2],\
				StoredLeverArms[0],StoredLeverArms[1],StoredLeverArms[2],\
				Boresight[0],Boresight[1],\
				BoresightStdv[0],BoresightStdv[1]";
		}

		//Print all structure data according to Header to the CSV file
		const char* printDataToCsv(char * OutStr, int length)
		{
			snprintf(OutStr, length, "%.12f,%.12f,%d,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f,%.6f", SystemTov, GpsTov, GpsWeek,
				DistanceTraveled,ScaleFactorCorrection,ScaleFactorStdv,
				LeverArms[0], LeverArms[1], LeverArms[2],
				LeverArmsStdv[0], LeverArmsStdv[1], LeverArmsStdv[2],
				StoredLeverArms[0], StoredLeverArms[1], StoredLeverArms[2],
				Boresight[0], Boresight[1],
				BoresightStdv[0], BoresightStdv[1]);
			return OutStr;
		}
	};


	//Collection of all messages sent by the INS - create and intialize for your application to use
	struct HgNavMessageSet
	{
		struct Hg0x2001Configuration Configuration; // INS Configuration Data
		struct Hg0x2011Status Status; // INS Status Data
		struct Hg0x20ffAck Ack; // Acknowledge of input message
		struct Hg0x2201TimeMark TimeMark; // Timemark and PPS message
		struct Hg0x6109GnssAttitude GnssAttitude; // Attitude from GNSS Receiver
		struct Hg0x6108GnssPosition GnssPosition; // Position from GNSS Receiver
		struct Hg0x2311InertialData InertialData; // IMU Inertial Data
		struct Hg0x6504NEDVelocity NEDVelocity; // Velocity in NED coordinates
		struct Hg0x6405EulerAttitudes EulerAttitudes; // INS Attitude Data
		struct Hg0x6403GeodeticPosition GeoPosition; // Geodetic position Data
		struct Hg0x6110DistanceTraveled DistanceTraveled; // Distance traveled calculated from inputted Speed aiding
		struct Hg0x6111MotionDetection MotionDetection; // Results of Motion detection algorithm in kalman filter
		struct Hg0x6424AntennaLeverArmEstimates AntennaLeverArmEstimates; //Estimates of Antenna Lever Arms from Kalman Filter
		struct Hg0x6438OdometerCalibration OdometerCalibration; // Odometer Calibration estimates from Kalman Filter
		struct Hg0x2402NavigationOutput NavOutput; // Legacy Navigation output
		
		/* Add additional Messages here*/

		//Initialization Method - call all Zero Methods
		void Init()
		{
			Configuration.ZeroMessage();
			Status.ZeroMessage();
			Ack.ZeroMessage();
			TimeMark.ZeroMessage();
			GnssAttitude.ZeroMessage();
			GnssPosition.ZeroMessage();
			InertialData.ZeroMessage();
			NEDVelocity.ZeroMessage();
			EulerAttitudes.ZeroMessage();
			GeoPosition.ZeroMessage();
			NavOutput.ZeroMessage();
			DistanceTraveled.ZeroMessage();
			MotionDetection.ZeroMessage();
			AntennaLeverArmEstimates.ZeroMessage();
			OdometerCalibration.ZeroMessage();
		}

		// == operator override
		bool operator==(const HgNavMessageSet& Message) const
		{
			if (!(Configuration == Message.Configuration)) return false;
			if (!(Status == Message.Status)) return false;
			if (!(Ack == Message.Ack)) return false;
			if (!(TimeMark== Message.TimeMark)) return false;
			if (!(GnssAttitude== Message.GnssAttitude)) return false;
			if (!(GnssPosition== Message.GnssPosition)) return false;
			if (!(InertialData== Message.InertialData)) return false;
			if (!(NEDVelocity== Message.NEDVelocity)) return false;
			if (!(EulerAttitudes== Message.EulerAttitudes)) return false;
			if (!(GeoPosition== Message.GeoPosition)) return false;
			if (!(NavOutput== Message.NavOutput)) return false;
			if (!(DistanceTraveled == Message.DistanceTraveled)) return false;
			if (!(MotionDetection == Message.MotionDetection)) return false;
			if (!(AntennaLeverArmEstimates == Message.AntennaLeverArmEstimates)) return false;
			if (!(OdometerCalibration == Message.OdometerCalibration)) return false;
			else
				return true;
		}

		// limit test
		bool EqualWithMargin(const HgNavMessageSet& Message, double Margin) const
		{
			if (!Configuration.EqualWithMargin(Message.Configuration, Margin)) return false;
			if (!Status.EqualWithMargin(Message.Status, Margin)) return false;
			if (!Ack.EqualWithMargin(Message.Ack, Margin)) return false;
			if (!TimeMark.EqualWithMargin(Message.TimeMark, Margin)) return false;
			if (!GnssAttitude.EqualWithMargin(Message.GnssAttitude, Margin)) return false;
			if (!GnssPosition.EqualWithMargin(Message.GnssPosition, Margin)) return false;
			if (!InertialData.EqualWithMargin(Message.InertialData, Margin)) return false;
			if (!NEDVelocity.EqualWithMargin(Message.NEDVelocity, Margin)) return false;
			if (!EulerAttitudes.EqualWithMargin(Message.EulerAttitudes, Margin)) return false;
			if (!GeoPosition.EqualWithMargin(Message.GeoPosition, Margin)) return false;
			if (!NavOutput.EqualWithMargin(Message.NavOutput, Margin)) return false;
			if (!DistanceTraveled.EqualWithMargin(Message.DistanceTraveled, Margin)) return false;
			if (!MotionDetection.EqualWithMargin(Message.MotionDetection, Margin)) return false;
			if (!AntennaLeverArmEstimates.EqualWithMargin(Message.AntennaLeverArmEstimates, Margin)) return false;
			if (!OdometerCalibration.EqualWithMargin(Message.OdometerCalibration, Margin)) return false;
			else
				return true;
		}
	};

	// GetNavMessage function - selects appropriate message and fills the structure
	HGDATAPARSER_API int GetNavMessage(UINT8 *buffer, int startOffset, struct HgNavMessageSet *Message, int *endOffset);

	HGDATAPARSER_API int Get0x2402Message(UINT8 *buffer, int startOffset, struct Hg0x2402NavigationOutput *Message, int *endOffset);
	HGDATAPARSER_API int Get0x2001Message(UINT8 *buffer, int startOffset, struct Hg0x2001Configuration *Message, int *endOffset);
	HGDATAPARSER_API int Get0x2011Message(UINT8 *buffer, int startOffset, struct Hg0x2011Status *Message, int *endOffset);
	HGDATAPARSER_API int Get0x20ffMessage(UINT8 *buffer, int startOffset, struct Hg0x20ffAck *Message, int *endOffset);
	HGDATAPARSER_API int Get0x2201Message(UINT8 *buffer, int startOffset, struct Hg0x2201TimeMark *Message, int *endOffset);
	HGDATAPARSER_API int Get0x2311Message(UINT8 *buffer, int startOffset, struct Hg0x2311InertialData *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6403Message(UINT8 *buffer, int startOffset, struct Hg0x6403GeodeticPosition *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6504Message(UINT8 *buffer, int startOffset, struct Hg0x6504NEDVelocity *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6405Message(UINT8 *buffer, int startOffset, struct Hg0x6405EulerAttitudes *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6108Message(UINT8 *buffer, int startOffset, struct Hg0x6108GnssPosition *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6109Message(UINT8 *buffer, int startOffset, struct Hg0x6109GnssAttitude *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6110Message(UINT8 *buffer, int startOffset, struct Hg0x6110DistanceTraveled *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6111Message(UINT8 *buffer, int startOffset, struct Hg0x6111MotionDetection *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6424Message(UINT8 *buffer, int startOffset, struct Hg0x6424AntennaLeverArmEstimates *Message, int *endOffset);
	HGDATAPARSER_API int Get0x6438Message(UINT8 *buffer, int startOffset, struct Hg0x6438OdometerCalibration *Message, int *endOffset);

	//Calculate the HGuide INS Checksum
	HGDATAPARSER_API int HgInsChecksum(UINT8 *buffer, int startOffset, int wordLength);

	//0x4110 Vehicle Speed Input Message
	HGDATAPARSER_API int Get0x4110Message(UINT8 *buffer, int startOffset, struct Hg0x4110VelocityAidingInput * Message, int * endOffset);

	//Utilities to create NMEA messages from appropriate Output Messages

	HGDATAPARSER_API int CreateNmeaGpGGA(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
	HGDATAPARSER_API int CreateNmeaGpRMC(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
	HGDATAPARSER_API int CreateNmeaGpGLL(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
	HGDATAPARSER_API int CreateNmeaGpVTG(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
	HGDATAPARSER_API int CreateNmeaPASHR(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
	HGDATAPARSER_API int CreateNmeaGpHDT(UINT8 *buffer, int startOffset, struct HgNavMessageSet MessageSet, int *byteLength);
}
