using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OperateIniFile;
using Spire.Xls; 
using System.ComponentModel.DataAnnotations;
using Farmework.Other.enumTools;
using Framework.Other;

namespace XRDataProcess
{
    /// <summary>
    /// 养护标准类型
    /// </summary>
  public  enum StandardParmType
    {
        /// <summary>
        /// 0--等级公路2007
        /// </summary>
        DegreeRoad2007,

        /// <summary>
        /// 1--城镇道路
        /// </summary>
        CityRoad,

        /// <summary>
        /// 2--北京农村公路
        /// </summary>
        RuralRoadBeijing,

        /// <summary>
        /// 3--等级公路2018
        /// </summary>
        DegreeRoad2018,

        /// <summary>
        /// 4--等级公路2001
        /// </summary>
        DegreeRoad2001,

        /// <summary>
        /// 5--上海城市道路
        /// </summary>
        CityRoadShanghai,

        /// <summary>
        /// 6--辽宁农村路
        /// </summary>
        RuralRoadLiaoning,

        /// <summary>
        /// 7-广西农村路
        /// </summary>
        RuralRoadGuangxi,

        /// <summary>
        /// 8-重庆农村路
        /// </summary>
        RuralRoadChongqing,

        /// <summary>
        /// 9-湖南农村路
        /// </summary>
        RuralRoadHunan,

      
        /// <summary>
        /// 10-低等级农村公路
        /// </summary>
        RuralRoadlowLevel,

        ///// <summary>
        ///// 11-湖南农村路2024 
        ///// </summary>
        //RuralRoadHunan2024
    }

    /// <summary>
    /// 软件安装目录下面的XRSetting.ini文件配置内容
    /// </summary>
    class XRSetting
    {
        private XRSetting() { }

        private static XRSetting singleInstance = null;

        public static XRSetting GetInstance()
        {
            return singleInstance;
        }

        static XRSetting()
        {
            singleInstance = new XRSetting();
        }
        private const string IniFileName = "XRSetting.ini";
        private const int MaxExcelConfigCount = 100;
        //private const string IniDefaultFileName = "XRSetting_Default.ini"; // 可选：安装目录放默认模板
        /// <summary>
        /// 获取用户专属的 XRSetting.ini 路径（%LocalAppData%）
        /// </summary>
        private static string GetUserIniPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);
            return Path.Combine(appFolder, IniFileName);
        }

        /// <summary>
        /// 获取安装目录下的默认 ini 路径（只读）
        /// </summary>
        private static string GetDefaultIniPath()
        {
            return Path.Combine(System.Windows.Forms.Application.StartupPath, IniFileName);
        }

        /// <summary>
        /// 首次运行时复制默认配置到用户目录
        /// </summary>
        private static void EnsureDefaultIniCopied()
        {
            string userPath = GetUserIniPath();
            string defaultPath = GetDefaultIniPath();

            if (File.Exists(userPath))
            {
                return;
            }

            try
            {
                if (File.Exists(defaultPath))
                {
                    File.Copy(defaultPath, userPath);
                }
                else
                {
                    File.WriteAllText(userPath, "#SettingPara", Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("XRSetting 默认配置复制失败: " + ex.Message);
            }
        }

        private static int ClampCount(int value)
        {
            if (value < 0)
            {
                return 0;
            }

            if (value > MaxExcelConfigCount)
            {
                return MaxExcelConfigCount;
            }

            return value;
        }

        private static StandardParmType ReadStandardParmType(IniFiles inisetting)
        {
            int parmStyle = inisetting.ReadInteger("UI", "ParmStyle", 0);
            if (!Enum.IsDefined(typeof(StandardParmType), parmStyle))
            {
                return StandardParmType.DegreeRoad2007;
            }

            return (StandardParmType)parmStyle;
        }

        private static int GetWritableCount(int count, params Array[] arrays)
        {
            count = ClampCount(count);
            foreach (Array array in arrays)
            {
                if (array == null)
                {
                    return 0;
                }

                count = Math.Min(count, array.Length);
            }

            return count;
        }

        private static void PrepareTempIniFile(string userIniPath, string tempIniPath)
        {
            TryDeleteFile(tempIniPath);

            if (File.Exists(userIniPath))
            {
                File.Copy(userIniPath, tempIniPath);
                return;
            }

            string defaultPath = GetDefaultIniPath();
            if (File.Exists(defaultPath))
            {
                File.Copy(defaultPath, tempIniPath);
            }
            else
            {
                File.WriteAllText(tempIniPath, "#SettingPara", Encoding.UTF8);
            }
        }

        private static void ReplaceIniFile(string tempIniPath, string userIniPath)
        {
            if (File.Exists(userIniPath))
            {
                File.Replace(tempIniPath, userIniPath, null, true);
            }
            else
            {
                File.Move(tempIniPath, userIniPath);
            }
        }

        private static void ValidateIniFile(string iniPath)
        {
            FileInfo fileInfo = new FileInfo(iniPath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                throw new InvalidDataException("XRSetting.ini 为空。");
            }

            IniFiles inisetting = new IniFiles(iniPath);
            int parmStyle = inisetting.ReadInteger("UI", "ParmStyle", 0);
            if (!Enum.IsDefined(typeof(StandardParmType), parmStyle))
            {
                throw new InvalidDataException("XRSetting.ini 中 ParmStyle 无效。");
            }

            int lenExcelNum = inisetting.ReadInteger("UI", "LenExcelNum", 0);
            int streetLenExcelNum = inisetting.ReadInteger("UI", "StreetLenExcelNum", 0);
            if (lenExcelNum < 0 || lenExcelNum > MaxExcelConfigCount ||
                streetLenExcelNum < 0 || streetLenExcelNum > MaxExcelConfigCount)
            {
                throw new InvalidDataException("XRSetting.ini 中报表配置数量无效。");
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("删除临时配置文件失败: " + ex.Message);
            }
        }
        public void ReadData()
        {
            // 1. 确保默认配置已复制
            EnsureDefaultIniCopied();

            // 2. 读取用户配置（一定可读）
            string userIniPath = GetUserIniPath();
            IniFiles inisetting = new IniFiles(userIniPath);
            //IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\XRSetting.ini");
            SkinName = inisetting.ReadString("UI", "SkinName", "");
            ICO = inisetting.ReadString("UI", "ICO", "");
            ICODX = inisetting.ReadString("UI", "ICODX", "");
            CompanyInfo = inisetting.ReadString("UI", "CompanyInfo", "");
           
            ErrorVal = inisetting.ReadInteger("UI", "ErrorVal", 0);
            ErrorIRI = inisetting.ReadDouble("UI", "ErrorIRI", 0);
            ErrorMTD = inisetting.ReadDouble("UI", "ErrorMTD", 0);
            ErrorRut = inisetting.ReadDouble("UI", "ErrorRut", 0);
            ErrorRutTh1 = inisetting.ReadDouble("UI", "ErrorRutTh1", 0);
            ErrorRutTh2 = inisetting.ReadDouble("UI", "ErrorRutTh2", 0);
            IsThresholdRut = inisetting.ReadBool("UI", "IsThresholdRut", false);
            OutRutIndex = inisetting.ReadInteger("UI", "OutRutIndex", -1);
            ParmStyle = ReadStandardParmType(inisetting);
            ExcelType = inisetting.ReadInteger("UI", "ExcelType", 0);
            IsExcelSort = inisetting.ReadBool("UI", "IsExcelSort", false);
            IsStatistics = inisetting.ReadBool("UI", "IsStatistics", false);
            IsRename = inisetting.ReadBool("UI", "IsRename", false);
            YHType = inisetting.ReadInteger("UI", "YHType", 0);
            DefaultPath = inisetting.ReadString("UI", "DefaultPath", "");
            ImgType = inisetting.ReadString("UI", "ImgType", "");

            LenExcelNum = ClampCount(inisetting.ReadInteger("UI", "LenExcelNum", 0));
            IsExcel = new bool[LenExcelNum];
            LenExcel = new string[LenExcelNum];
            for (int i = 0; i < LenExcelNum; ++i)
            {
                IsExcel[i] = inisetting.ReadBool("UI", "IsExcel" + i.ToString(), false);
                LenExcel[i] = inisetting.ReadString("UI", "LenExcel" + i.ToString(), "");
            }

            DetectYear = inisetting.ReadString("UI", "DetectYear", "");
            DetectNum = inisetting.ReadString("UI", "DetectNum", "");
            DistrictCode = inisetting.ReadString("UI", "DistrictCode", "");
            DutyUnit = inisetting.ReadString("UI", "DutyUnit", "");
            RoadSideType = inisetting.ReadString("UI", "RoadSideType", "");
            CADLength = inisetting.ReadInteger("UI", "CADLength", 0);
            IsRepair = inisetting.ReadBool("UI", "IsRepair", false);
            OutRut = inisetting.ReadInteger("UI", "OutRut", 0);
            Qufen_dis_degree = inisetting.ReadInteger("UI", "Qufen_dis_degree", 0);
            Acc_IRI = inisetting.ReadInteger("UI", "Acc_IRI", 0);
            Acc_IRI_K_1 = inisetting.ReadDouble("UI", "Acc_IRI_K_1", 0);
            Acc_IRI_B_1 = inisetting.ReadDouble("UI", "Acc_IRI_B_1", 0);
            IRIk = inisetting.ReadDouble("UI", "IRIk", 0);
            IRIb = inisetting.ReadDouble("UI", "IRIb", 0);
             
            Las_Filter = inisetting.ReadBool("UI", "Las_Filter", false);
            Las_Filter_Thresh0 = inisetting.ReadDouble("UI", "Las_Filter_Thresh0", 0);
            Las_Filter_Thresh1 = inisetting.ReadDouble("UI", "Las_Filter_Thresh1", 0);
            Out_roadimg = inisetting.ReadInteger("UI", "Out_roadimg", 0);
            Is_Multfolder = inisetting.ReadInteger("UI", "Is_Multfolder", 0);
            IRI_threshval = inisetting.ReadDouble("UI", "IRI_threshval", 0);
            cmop_rows = inisetting.ReadInteger("UI", "cmop_rows", 0);
            SelectDrawDis = inisetting.ReadInteger("UI", "SelectDrawDis", 0);
            ZJGT_dismodel = inisetting.ReadInteger("UI", "ZJGT_dismodel", 0);
            BrokenPlatetype = inisetting.ReadInteger("UI", "BrokenPlatetype", 0);
            PlateWidth = inisetting.ReadDouble("UI", "PlateWidth", 0);
            PlateLength = inisetting.ReadDouble("UI", "PlateLength", 0);
            RutDisWidth = inisetting.ReadDouble("UI", "RutDisWidth", 0);
            Is_SnCarve = inisetting.ReadInteger("UI", "Is_SnCarve", 0);
            IsShowAnalysis = inisetting.ReadBool("UI", "IsShowAnalysis", false);

            MPD_K = inisetting.ReadDouble("UI", "MPD_K", 0);
            MPD_B = inisetting.ReadDouble("UI", "MPD_B", 0);

            IsWarning = inisetting.ReadBool("UI", "IsWarning", false);
            PartType = inisetting.ReadInteger("UI", "PartType", 0);
            PartType_Dmi_Len = inisetting.ReadInteger("UI", "PartType_Dmi_Len", 0);
            Out_roadinfo = inisetting.ReadInteger("UI", "Out_roadinfo", 0);
            sheetRoundingOffType = inisetting.ReadInteger("UI", "sheetRoundingOffType", 0);
            sheetRoundingOffNum = inisetting.ReadInteger("UI", "sheetRoundingOffNum", 0);

            StreetLenExcelNum = ClampCount(inisetting.ReadInteger("UI", "StreetLenExcelNum", 0));
            StreetIsExcel = new bool[StreetLenExcelNum];
            StreetLenExcel = new string[StreetLenExcelNum];
            for (int i = 0; i < StreetLenExcelNum; ++i)
            {
                StreetLenExcel[i] = inisetting.ReadString("UI", "StreetLenExcel" + i.ToString(), "");
                StreetIsExcel[i] = inisetting.ReadBool("UI", "StreetIsExcel" + i.ToString(), false);
            }
          
            IsForbidOverLapping = inisetting.ReadBool("UI", "IsForbidOverLapping", false);
            IsCrackRemark = inisetting.ReadBool("UI", "IsCrackRemark", false);
            GPSJumpTime = inisetting.ReadInteger("UI", "GPSJumpTime", 0);

            IRIExcelSide = inisetting.ReadInteger("UI", "IRIExcelSide", 2);

            IsOutputDisAreaSubtotal = inisetting.ReadBool("UI", "IsOutputDisAreaSubtotal", true);

            IsCheckIRIGPSTime = inisetting.ReadBool("UI", "IsCheckIRIGPSTime", true);
            hasCamsetting = inisetting.ReadBool("UI", "HasCamsetting", false);
            needSub = inisetting.ReadBool("UI", "NeedSub", false);
            subData = inisetting.ReadString("UI", "SubData", "");
            recordHumanDis = inisetting.ReadBool("UI", "RecordHumanDis", false);
           
            rutKCorrect = inisetting.ReadDouble("UI", "RutKCorrect", 0.0);
             rutBCorrect = inisetting.ReadDouble("UI", "RutBCorrect", 0.0);
            iriKCorrect = inisetting.ReadDouble("UI", "IriKCorrect", 0.0);
            iriBCorrect = inisetting.ReadDouble("UI", "IriBCorrect", 0.0);

            hefei2MinSplit = inisetting.ReadInteger("UI", "Hefei2MinSplit", 50);
            roadCrossingShow = inisetting.ReadBool("UI", "RoadCrossingShow", true);
            lasthreshvalFactor = inisetting.ReadDouble("UI", "LasthreshvalFactor", 1.5);
            multiExcelMergeType = inisetting.ReadInteger("UI", "MultiExcelMergeType", 0);
            
            rutLeftCorrect = inisetting.ReadDouble("UI", "rutLeftCorrect", 0);
            rutRightCorrect = inisetting.ReadDouble("UI", "rutRightCorrect", 0);
            mptLeftCorrect = inisetting.ReadDouble("UI", "mptLeftCorrect", 0);
            mptRightCorrect = inisetting.ReadDouble("UI", "mptRightCorrect", 0);
            mptMidCorrect = inisetting.ReadDouble("UI", "mptMidCorrect", 0);
            splitExcelDh = inisetting.ReadBool("UI", "splitExcelDh", true);
            gjLbiOutHight = inisetting.ReadBool("UI", "gjLbiOutHight",false);
            outDaqAccelerate = inisetting.ReadBool("UI", "outDaqAccelerate", false);
            JSAverageType = inisetting.ReadBool("UI", "JSAverageType", false); 

            czJudgeType = inisetting.ReadInteger("UI", "czJudgeType", 0);
            OutWordPasteDelay = inisetting.ReadInteger("UI", "OutWordPasteDelay",  500);
            outSmallDisCalculateArea=inisetting.ReadBool("UI", "outSmallDisCalculateArea", false);
            mpdInterveneFAactor = inisetting.ReadString("UI", "MpdInterveneFAactor", "");
            zcSplit = inisetting.ReadBool("UI","zcSplit",false);
            heFeiContineIndex = inisetting.ReadInteger("UI", "heFeiContineIndex",0);
            outHumanDeleteDisease = inisetting.ReadBool("UI", "outHumanDeleteDisease",true);
            outHumanDeleteDiseasePath = inisetting.ReadString("UI", " outHumanDeleteDiseasePath ", "");

            equipType = inisetting.ReadInteger("UI", "equipType", 0);
            gpsformat = inisetting.ReadBool("UI", "gpsformat", true);
            SplitPartDistance = inisetting.ReadInteger("UI", nameof(SplitPartDistance), 1000);
            RQIJudgeType = inisetting.ReadInteger("UI",nameof(RQIJudgeType),0);
            IsOutputLasval = inisetting.ReadBool("UI", "IsOutputLasval", false);
            SmallDiseaseDrawType = inisetting.ReadInteger("UI", "SmallDiseaseDrawType", 0);

            IRIAlgorithmInterval = inisetting.ReadDouble("UI", "IRIAlgorithmInterval", 0.25);

            outMoHaoData = inisetting.ReadBool("UI", "outMoHaoData",false);
            is5211MergeArea500 = inisetting.ReadBool("UI", "is5211MergeArea500", true);
        }

        public void WriteData()
        {

            try
            {
                string userIniPath = GetUserIniPath();
                string tempIniPath = userIniPath + ".tmp";

                try
                {
                    PrepareTempIniFile(userIniPath, tempIniPath);
                    IniFiles inisetting = new IniFiles(tempIniPath);

                    // === 下面代码保持不变 ===
                    inisetting.WriteString("UI", "SkinName", SkinName);
                    inisetting.WriteString("UI", "ICO", ICO);
                //IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\XRSetting.ini");
                inisetting.WriteString("UI", "SkinName", SkinName);
                inisetting.WriteString("UI", "ICO", ICO);
                inisetting.WriteString("UI", "ICODX", ICODX);
                inisetting.WriteString("UI", "CompanyInfo", CompanyInfo);

                inisetting.WriteInteger("UI", "ErrorVal", ErrorVal);
                inisetting.WriteDouble("UI", "ErrorIRI", ErrorIRI);
                inisetting.WriteDouble("UI", "ErrorMTD", ErrorMTD);
                inisetting.WriteDouble("UI", "ErrorRut", ErrorRut);
                inisetting.WriteDouble("UI", "ErrorRutTh1", ErrorRutTh1);
                inisetting.WriteDouble("UI", "ErrorRutTh2", ErrorRutTh2);
                inisetting.WriteBool("UI", "IsThresholdRut", IsThresholdRut);
                inisetting.WriteInteger("UI", "ParmStyle", (int)ParmStyle);
                inisetting.WriteInteger("UI", "ExcelType", ExcelType);
                inisetting.WriteBool("UI", "IsExcelSort", IsExcelSort);
                inisetting.WriteBool("UI", "IsStatistics", IsStatistics);
                inisetting.WriteBool("UI", "IsRename", IsRename);
                inisetting.WriteInteger("UI", "YHType", YHType);
                inisetting.WriteString("UI", "DefaultPath", DefaultPath);
                inisetting.WriteString("UI", "ImgType", ImgType);

                int lenExcelCount = GetWritableCount(LenExcelNum, IsExcel, LenExcel);
                inisetting.WriteInteger("UI", "LenExcelNum", lenExcelCount);
                for (int i = 0; i < lenExcelCount; ++i)
                {
                    inisetting.WriteBool("UI", "IsExcel" + i.ToString(), IsExcel[i]);
                    inisetting.WriteString("UI", "LenExcel" + i.ToString(), LenExcel[i]);
                }

                inisetting.WriteString("UI", "DetectYear", DetectYear);
                inisetting.WriteString("UI", "DetectNum", DetectNum);
                inisetting.WriteString("UI", "DistrictCode", DistrictCode);
                inisetting.WriteString("UI", "DutyUnit", DutyUnit);
                inisetting.WriteString("UI", "RoadSideType", RoadSideType);
                inisetting.WriteInteger("UI", "CADLength", CADLength);
                inisetting.WriteBool("UI", "IsRepair", IsRepair);
                inisetting.WriteInteger("UI", "OutRut", OutRut);
                inisetting.WriteInteger("UI", "Qufen_dis_degree", Qufen_dis_degree);
                inisetting.WriteInteger("UI", "Acc_IRI", Acc_IRI);
                inisetting.WriteDouble("UI", "Acc_IRI_K_1", Acc_IRI_K_1);
                inisetting.WriteDouble("UI", "Acc_IRI_B_1", Acc_IRI_B_1);
                inisetting.WriteDouble("UI", "IRIk", IRIk);
                inisetting.WriteDouble("UI", "IRIb", IRIb);
                inisetting.WriteBool("UI", "outSmallDisCalculateArea", outSmallDisCalculateArea);
                inisetting.WriteBool("UI", "IsOutputLasval", IsOutputLasval);
                inisetting.WriteBool("UI", "Las_Filter", Las_Filter);
                inisetting.WriteDouble("UI", "Las_Filter_Thresh0", Las_Filter_Thresh0);
                inisetting.WriteDouble("UI", "Las_Filter_Thresh1", Las_Filter_Thresh1);
                inisetting.WriteInteger("UI", "Out_roadimg", Out_roadimg);
                inisetting.WriteInteger("UI", "Is_Multfolder", Is_Multfolder);
                inisetting.WriteDouble("UI", "IRI_threshval", IRI_threshval);
                inisetting.WriteInteger("UI", "cmop_rows", cmop_rows);
                inisetting.WriteInteger("UI", "SelectDrawDis", SelectDrawDis);
                inisetting.WriteInteger("UI", "ZJGT_dismodel", ZJGT_dismodel);
                inisetting.WriteInteger("UI", "BrokenPlatetype", BrokenPlatetype);
                inisetting.WriteDouble("UI", "PlateWidth", PlateWidth);
                inisetting.WriteDouble("UI", "PlateLength", PlateLength);
                inisetting.WriteDouble("UI", "RutDisWidth", RutDisWidth);
                inisetting.WriteInteger("UI", "Is_SnCarve", Is_SnCarve);
                inisetting.WriteBool("UI", "IsShowAnalysis", IsShowAnalysis);

                inisetting.WriteDouble("UI", "MPD_K", MPD_K);
                inisetting.WriteDouble("UI", "MPD_B", MPD_B);
                inisetting.WriteBool("UI", "IsWarning", IsWarning);
                inisetting.WriteInteger("UI", "PartType", PartType);
                inisetting.WriteInteger("UI", "PartType_Dmi_Len", PartType_Dmi_Len);
                inisetting.WriteInteger("UI", "Out_roadinfo", Out_roadinfo);
                inisetting.WriteInteger("UI", "sheetRoundingOffType", sheetRoundingOffType);
                inisetting.WriteInteger("UI", "sheetRoundingOffNum", sheetRoundingOffNum);
                int streetLenExcelCount = GetWritableCount(StreetLenExcelNum, StreetIsExcel, StreetLenExcel);
                inisetting.WriteInteger("UI", "StreetLenExcelNum", streetLenExcelCount);
                for (int i = 0; i < streetLenExcelCount; ++i)
                {
                    inisetting.WriteString("UI", "StreetLenExcel" + i.ToString(), StreetLenExcel[i]);
                    inisetting.WriteBool("UI", "StreetIsExcel" + i.ToString(), StreetIsExcel[i]);
                }

                inisetting.WriteBool("UI", "IsForbidOverLapping", IsForbidOverLapping);
                inisetting.WriteBool("UI", "IsCrackRemark", IsCrackRemark);
                inisetting.WriteInteger("UI", "GPSJumpTime", GPSJumpTime);

                inisetting.WriteInteger("UI", "IRIExcelSide", IRIExcelSide);

                inisetting.WriteBool("UI", "IsOutputDisAreaSubtotal", IsOutputDisAreaSubtotal);

                inisetting.WriteBool("UI", "IsCheckIRIGPSTime", IsCheckIRIGPSTime);
                inisetting.WriteBool("UI", "HasCamsetting", hasCamsetting);
                inisetting.WriteBool("UI", "NeedSub", needSub);
                inisetting.WriteString("UI", "SubData", subData);
                inisetting.WriteBool("UI", "RecordHumanDis", recordHumanDis);
                inisetting.WriteDouble("UI", "RutKCorrect", rutKCorrect);
                inisetting.WriteDouble("UI", "RutBCorrect", rutBCorrect);
                inisetting.WriteDouble("UI", "IriKCorrect", iriKCorrect);
                inisetting.WriteDouble("UI", "IriBCorrect", iriBCorrect);

                inisetting.WriteInteger("UI", "Hefei2MinSplit", hefei2MinSplit);
                inisetting.WriteBool("UI", "RoadCrossingShow", roadCrossingShow);
                inisetting.WriteDouble("UI", "LasthreshvalFactor", lasthreshvalFactor);
                inisetting.WriteInteger("UI", "MultiExcelMergeType", multiExcelMergeType);

                inisetting.WriteDouble("UI", "rutLeftCorrect", rutLeftCorrect);
                inisetting.WriteDouble("UI", "rutRightCorrect", rutRightCorrect);
                inisetting.WriteDouble("UI", "mptLeftCorrect", mptLeftCorrect);
                inisetting.WriteDouble("UI", "mptRightCorrect", mptRightCorrect);
                inisetting.WriteDouble("UI", "mptMidCorrect", mptMidCorrect);
                inisetting.WriteBool("UI", "splitExcelDh", splitExcelDh);
                inisetting.WriteBool("UI", "gjLbiOutHight", gjLbiOutHight);
                inisetting.WriteBool("UI", "outDaqAccelerate", outDaqAccelerate);
                inisetting.WriteBool("UI", "JSAverageType", JSAverageType);

                inisetting.WriteInteger("UI", "czJudgeType", czJudgeType);
                inisetting.WriteInteger("UI", "OutWordPasteDelay", OutWordPasteDelay);
                inisetting.WriteString("UI", "MpdInterveneFAactor", mpdInterveneFAactor);
                inisetting.WriteBool("UI", "zcSplit", zcSplit);
                inisetting.WriteInteger("UI", "heFeiContineIndex", heFeiContineIndex);

                inisetting.WriteInteger("UI", "equipType", equipType);

                inisetting.WriteBool("UI", "outHumanDeleteDisease", outHumanDeleteDisease);
                inisetting.WriteString("UI", "outHumanDeleteDiseasePath", outHumanDeleteDiseasePath);
                inisetting.WriteBool("UI", "gpsformat", gpsformat);

                inisetting.WriteInteger("UI", nameof(SplitPartDistance), SplitPartDistance);

                inisetting.WriteInteger("UI", nameof(RQIJudgeType), RQIJudgeType);
                inisetting.WriteInteger("UI", nameof(OutRutIndex), OutRutIndex);

                inisetting.WriteInteger("UI", nameof(SmallDiseaseDrawType), SmallDiseaseDrawType);

                inisetting.WriteDouble("UI", nameof(IRIAlgorithmInterval), IRIAlgorithmInterval);

                inisetting.WriteBool("UI", "outMoHaoData", outMoHaoData);
                inisetting.WriteBool("UI", "is5211MergeArea500", is5211MergeArea500);

                    ValidateIniFile(tempIniPath);
                    ReplaceIniFile(tempIniPath, userIniPath);
                }
                finally
                {
                    TryDeleteFile(tempIniPath);
                }
            }
            catch (Exception ex)
            {
                // 防止写入失败导致崩溃
                System.Diagnostics.Debug.WriteLine("XRSetting WriteData 失败: " + ex.Message);
                // 如使用 log4net：log.Error("XRSetting 保存失败", ex);
            }

            
        }

        /// <summary>
        /// 界面风格
        /// </summary>
        public string SkinName;

        /// <summary>
        /// 软件大图标
        /// </summary>
        public string ICO;

        /// <summary>
        /// 左上角软件小图标
        /// </summary>
        public string ICODX;

        /// <summary>
        /// 软件的公司信息
        /// </summary>
        public string CompanyInfo;

        /// <summary>
        /// IRM的异常值处理方式，0--异常值不处理，1--异常值根据设置方式调整
        /// </summary>
        public int ErrorVal;

        /// <summary>
        /// IRI要调整处理的异常值上限
        /// </summary>
        public double ErrorIRI;

        /// <summary>
        /// 构造要处理的异常值上限
        /// </summary>
        public double ErrorMTD;
        
      

        /// <summary>
        /// 车辙要处理的异常值上限
        /// </summary>
        public double ErrorRut;

        /// <summary>
        /// 左右车辙调整的异常差值上限
        /// </summary>
        public double ErrorRutTh1;

        /// <summary>
        /// 前后相邻断面需要处理的车辙异常差值上限
        /// </summary>
        public double ErrorRutTh2;

        /// <summary>
        /// 是否要将，车辙值控制在 ErrorRut 异常值上限的范围内
        /// </summary>
        public bool IsThresholdRut;

        /// <summary>
        /// 输出调试车辙索引 默认-1
        /// </summary>
        public int OutRutIndex;

        /// <summary>
        /// 0--等级公路2007，1--城镇道路，2--北京农村公路，3--等级公路2018, 4--等级公路2001, 5--上海城市道路，6--辽宁农村路，7-广西农村路，8-重庆农村路，9-湖南农村路  10-低等级农村公路
        /// </summary>
        public StandardParmType ParmStyle;

        /// <summary>
        /// 等级路 0--各项指标单独出表，1--所有指标综合出表，2--中南安环，3--中交国通，4--带GPS模板，5--奥路通，7--上海浦公，8-厦门捷航，9-河南焦作, 10-广东华路 ,12-csv报表
        ///  农村路7--合肥报表
        /// </summary>
        public int ExcelType;

        //多车道统计选择
        public int multiExcelMergeType;

        /// <summary>
        /// false--导出的报表内容不排序，true--导出的报表内容根据桩号从小到大排序
        /// </summary>
        public bool IsExcelSort;

        /// <summary>
        /// false--报表不输出指标统计信息，true--报表输出指标统计信息
        /// </summary>
        public bool IsStatistics;

        /// <summary>
        /// 是否给图像名中添加桩号信息
        /// </summary>
        public bool IsRename;

        /// <summary>
        /// 养护类型，0--广西标准，1--辽宁标准，2--广西PCI标准，生成广西桂兴达报表时会用到
        /// </summary>
        public int YHType;

        /// <summary>
        /// 导入的工程数据默认路径
        /// </summary>
        public string DefaultPath;

        /// <summary>
        /// 图像文件的后缀
        /// </summary>
        public string ImgType;

        /// <summary>
        /// 不同指标报表数量
        /// </summary>
        public int LenExcelNum;

        /// <summary>
        /// 是否要导出不同指标的报表
        /// </summary>
        public bool[] IsExcel;

        /// <summary>
        /// 不同指标报表的单元区间长度
        /// </summary>
        public string[] LenExcel;

        /// <summary>
        /// 检测年，报表转换为奥路通平台输入模板时会用到
        /// </summary>
        public string DetectYear;
        /// <summary>
        /// 检测次数，报表转换为奥路通平台输入模板时会用到
        /// </summary>
        public string DetectNum;
        /// <summary>
        /// 县区代码，报表转换为奥路通平台输入模板时会用到
        /// </summary>
        public string DistrictCode;

        /// <summary>
        /// 管养单位
        /// </summary>
        public string DutyUnit;

        /// <summary>
        /// 道路的车道描述，比如双向四车道，生成城镇路的报告报表时会用到
        /// </summary>
        public string RoadSideType;

        /// <summary>
        /// 将城镇路的病害合并，让客户能自己再画到CAD上的，病害合并长度
        /// </summary>
        public int CADLength;
        
        public bool IsRepair;

        /// <summary>
        /// 导出路面车辙病害  0--不导出车辙病害 1--导出所有等级路车辙病害 2--只导出二三四级公路车辙病害
        /// </summary>
        public int OutRut;
        
        /// <summary>
        /// 区分病害程度  0--区分 1--所有病害程度按重度计算
        /// </summary>
        public int Qufen_dis_degree;
        
        /// <summary>
        /// 不同的平整度算法，正常应该设置为0
        /// </summary>
        public int Acc_IRI;
        public double Acc_IRI_K_1;
        public double Acc_IRI_B_1;
        public double IRIk;
        public double IRIb;

        /// <summary>
        /// 平整度计算时，激光测距机的数据是否需要剔除异常值，正常激光测距机数据质量较好不需要，当路面有水或者沥青雾封罩面工艺的新路，测距机数据里面异常值比较多需要特殊处理
        /// </summary>
        public bool  Las_Filter;

        /// <summary>
        /// 导出磨耗原始数据
        /// </summary>
        public bool outMoHaoData;
        public double Las_Filter_Thresh0;
        public double Las_Filter_Thresh1;

        /// <summary>
        /// 路面病害列表的报表里面是否要输出病害所在的路面图像名称和路径
        /// </summary>
        public int Out_roadimg;

        /// <summary>
        /// 导入多个工程的报表是否输出到同一个文件夹-0，或新建多个各自的文件夹-1
        /// </summary>
        public int Is_Multfolder;

        /// <summary>
        /// 当路面有水的时候，构造深度值特别小，用IRI_threshval来和路段的构造深度值比较，小于这个值就认为路面有水，调整IRI计算策略，只用加速度的位移来计算IRI
        /// </summary>
        public double IRI_threshval;
        
        /// <summary>
        /// CMOP调查表每页的行数，打印的时候用
        /// </summary>
        public int cmop_rows;

        /// <summary>
        /// 病害勾画选择   0--拉框  1--小方格，对于2018年的公路标准有区别，其他标准都用大框
        /// </summary>
        public int SelectDrawDis;

        /// <summary>
        /// 0--中交国通CPMS模板 1--按路面类型出整10* 米病害
        /// </summary>
        public int ZJGT_dismodel;

        /// <summary>
        /// 水泥路面，破碎板的面积计算方式，0--病害框面积，1--板块面积
        /// </summary>
        public int BrokenPlatetype;

        /// <summary>
        /// 水泥路面，水泥板块的宽度，单位m
        /// </summary>
        public double PlateWidth;

        /// <summary>
        /// 水泥路面，水泥板块的长度，单位m
        /// </summary>
        public double PlateLength;

        /// <summary>
        /// 路面车辙病害的影响宽度0.4m，只对于小方格的2018等级公路会用到
        /// </summary>
        public double RutDisWidth;

        /// <summary>
        /// 水泥路面是否有刻槽，0-没有刻槽，1-有刻槽，有刻槽的高等级水泥路面PWI不参与PQI的计算
        /// </summary>
        public int Is_SnCarve;

        /// <summary>
        /// IRM窗口数据分析的时候是否显示计算车辙的包络线，默认是不给客户显示的
        /// </summary>
        public bool IsShowAnalysis;

        /// <summary>
        /// MPD的计算系数k
        /// </summary>
        public double MPD_K;

        /// <summary>
        /// MPD的计算系数b 
        /// </summary>
        public double MPD_B;

        /// <summary>
        /// 出报表的时候是否弹窗显示IRM计算失败
        /// </summary>
        public bool IsWarning;

        /// <summary>
        /// 0-整桩号分段，1-整里程分段，仅对城镇有区别，等级公路始终使用整桩号分段
        /// </summary>
        public int PartType;

        /// <summary>
        /// 整里程分段，分段区间的长度
        /// </summary>
        public int PartType_Dmi_Len;

        /// <summary>
        /// 生成报表的时候是否输出车速和打标备注，有的客户要，有的客户不要
        /// </summary>
        public int Out_roadinfo;

        /// <summary>
        /// 报表数值修约方式，0-四舍五入修约，1-奇进偶舍修约
        /// </summary>
        public int sheetRoundingOffType;
        
        /// <summary>
        /// 导出的报表中数值小数修约位数
        /// </summary>
        public int sheetRoundingOffNum;

        /// <summary>
        /// 和景观相关的，SCI和TCI报表数量
        /// </summary>
        public int StreetLenExcelNum;

        /// <summary>
        /// SCI和TCI的单元区间长度
        /// </summary>
        public string[] StreetLenExcel;

        /// <summary>
        /// 是否输出SCI和TCI的报表
        /// </summary>
        public bool[] StreetIsExcel;

        /// <summary>
        /// 是否将构造的激光测距值转换成明码输出，调试用
        /// </summary>
        public bool IsOutputLasval =false;

        /// <summary>
        /// 是否禁止病害框有重叠区域
        /// </summary>
        public bool IsForbidOverLapping;

        /// <summary>
        /// 是否给路面病害添加备注
        /// </summary>
        public bool IsCrackRemark;

        /// <summary>
        /// 跳秒，UTC时=GPS时-18（秒）
        /// </summary>
        public int GPSJumpTime;

        /// <summary>
        /// 平整度评定方式，0-取双轮迹平整度的平均值（默认），1-取双轮迹平整度的最大值 ,2-低等级农村路湖南规范平整度计算标准
        /// </summary>
        public int RQIJudgeType = 0;

        /// <summary>
        /// 高速路采集了双轮迹的平整度报表导出设置，0-只导出左侧DAQ0，1-只导出右侧DAQ1，2-默认选项所有数据都导出
        /// </summary>
        public int IRIExcelSide = 2;

        /// <summary>
        /// 病害汇总表是否导出整公里小计，true-导出（默认），false-不导出
        /// </summary>
        public bool IsOutputDisAreaSubtotal = true;

        /// <summary>
        /// 是否检查平整度的同步时间，减少额外计算
        /// </summary>
        public bool IsCheckIRIGPSTime = false;
        //农村路是否需要进行惯导计算（惯导计算的工程项目 不支持 车辙，跳车等的出表 此处提供依据是否展示相应出表按钮）
        public bool isGDIriCalculate = false;

        ///// <summary>
        ///// 低等级农村路  是否进行图像校准   0 模块化 1高配 2低配 3畸变矫正 4 lm300 5 MM800 6 HM800 7 自动获取
        ///// </summary>
        //public int isImageCorrect = 0;
        ////低等级农村路设备参数 格式 真实宽度|真实高度（例如：3.5|2）
        //   public string real_HM800 = "";
        //public string real_lm300 = "";
        //public string real_MM800 = "";
        //农村路是否存在配置文件
        public bool hasCamsetting = false;
        //需要分段出表
        public bool needSub = false;
        public string subData = null;
        /// <summary>
        /// 用于记录分段出表当前区间
        /// 供生成桩号时读取  以确定该用哪个区间进行过滤
        /// </summary>
        public string nowSubIndexStr = "";
        //是否记录人工病害
        public bool recordHumanDis = false;
        //合肥报表标志
     
        public bool hefeiOutExcel2 = false;
        //车辙原始值:x  使用公式 x* rutKCorrect *rutBCorrect
        public double rutKCorrect;
        public double rutBCorrect;
        public double iriKCorrect;
        public double iriBCorrect;

        //左车辙调整值  直接在出表前调整 不修改原始数据
        public double rutLeftCorrect;
        //右车辙调整值  直接在出表前调整 不修改原始数据
        public double rutRightCorrect;
        //构造深度调整值  直接在出表前调整 不修改原始数据
        public double mptLeftCorrect  ;
        public double mptRightCorrect;
        public double mptMidCorrect   ;
        /// <summary>
        ///  合肥2最小分段区间控制  单位m
        /// </summary>
        public int hefei2MinSplit = 50;
        /// <summary>
        /// 路口标志是否显示(城镇道路)
        /// </summary>
        public bool roadCrossingShow = true;

        public double lasthreshvalFactor = 1.5;
       

        //调绘表是否分表  按照1000行
        public bool splitExcelDh = true;
        //
        /// <summary>
        /// 国检转换规范
        /// 0 交通部2023规范 1 交通部2024规范 2 河南(中交国通)2024定制
        /// </summary>
        public hnEnumTools.CityModelItem gjStandardNew = 0;
        //国检转换是否导出高程
        public bool gjLbiOutHight = false;

        //惯导daq文件直接导出文本文件
        public bool outDaqAccelerate = false;

        //车辙评定方式  0平均值 1最大值 2最大值平均
        public int czJudgeType = 0;

        public int OutWordPasteDelay = 500;


        public bool outSmallDisCalculateArea;

        //技术状况评定明细表合并 使用平均规则
        /// <summary>
        /// true  算数平均
        /// false  加权平均
        /// </summary>
        public bool JSAverageType;

        public string mpdInterveneFAactor;

        public bool zcSplit = false;
         public int heFeiContineIndex = 0;

        /// <summary>
        /// 屏蔽打标
        /// </summary>
        public bool shieldMark = false;

        /// <summary>
        /// 显示gps信息到图片
        /// </summary>
        public bool showGpsInfoToPicture = false;
        /// <summary>
        /// 0 模块化设备
        /// 1 二三维设备
        /// 在高精度定位 gps桩号匹配时 供用户选择确定 
        /// </summary>
        public int equipType = 0;

        /// <summary>
        /// 高精度模块 统一减去-s
        /// </summary>
        public bool allSub1s = false;

        //出表禁用打标 分段
        public bool banMarkSign = false;

        /// <summary>
        /// 使用道路情况进行分段
        /// </summary>
        public bool userRoadCondition = false;
        /// <summary>
        /// 打标分段距离（需在500~1500）
        /// </summary>
        public int SplitPartDistance = 500;

        public bool outHumanDeleteDisease = true;
        public string outHumanDeleteDiseasePath = "";

        /// <summary>
        /// true 显示gps
        /// false 显示大地坐标
        /// </summary>
        public bool gpsformat = true;


        /// <summary>
        /// 小框绘制模式
        /// 0 常规绘制
        /// 1 片状绘制
        /// 2 线状绘制
        /// </summary>
        public int SmallDiseaseDrawType = 0;

        public int OneTimeTip = 1; //一次性提示次数

      public  double IRIAlgorithmInterval = 0.25; //平整度算法计算间隔

        /// <summary>
        /// 低等级 5211规范要求,是否合并小于500的区间
        /// </summary>
        public bool is5211MergeArea500 = true;
    }
}
