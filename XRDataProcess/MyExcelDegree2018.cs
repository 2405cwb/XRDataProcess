using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms;
using Framework.Office.Excel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using Spire.Xls.Core;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.PTG;
using Framework.Other;

namespace XRDataProcess
{
    /// <summary>
    /// 等级公路规范2018，JTG 5210-2018 公路技术状况评定标准，大框
    /// </summary>
    class MyExcelDegree2018
    {
        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

        private static double[][][] _RQIGrade;//道路等级 路面材质 等级区间
        private static double[][] _RDIGrade;
        private static double[][] _PCIGrade;
        private static double[][] _PQIGrade;
        private static double[][] _PBIGrade;
        private static double[][] _PWIGrade;
        private static double[][] _RDIRD;
        private static double[] _RDIa;
        private static double[] _PWIa;
        private static double[] _PBIThresh;
        private static double[] _PBIScore;

        private static double[][][] _RQIa;//公路等级 路面材质 参数序号
        private static double[][][] _PCIa;//公路等级 路面材质 参数序号

        /// <summary>
        /// WPCI WRQI WRDI WPBI WPWI
        /// </summary>
        private static double[][][] _PQIW;//公路等级 路面材质 参数序号

        /// <summary>
        /// WSCI WPQI WBCI WTCI
        /// </summary>
        private static double[] _MQIW;
        /// <summary>
        /// MQI、SCI、SRI、PSSI、BCI、TCI指标的优良中次差等级区间
        /// </summary>
        private static double[] _MQIGrade;

        private static double[][] _WeightParm;//0-沥青，1-水泥
        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        public static Dictionary<string, int> _RoadGradeDict;

        public static List<MilePart> _RoadPart = null;

        public static List<MilePart> _RoadPart10 = null;//整10米桩号分段
        private static double[] _SpeedVal10 = null;
        private static string[] _MarkVal10 = null;

        public static List<MilePart> _RoadPart1M = null;//1米桩号分段
        private static Disease[] _RoadDisList = null;
        private static Disease[] _RoadRepairList = null;

        private static double[] _LIRIMeanVal = null;
        private static double[] _RIRIMeanVal = null;

        private static double[] _LMTDMeanVal = null;
        private static double[] _RMTDMeanVal = null;
        private static double[] _CMTDMeanVal = null;

        private static double[] _LRutMeanVal = null;
        private static double[] _RRutMeanVal = null;
        private static double[] _SRutMeanVal = null;

        private static double[] _LRutMaxVal = null;
        private static double[] _RRutMaxVal = null;
        private static double[] _SRutMaxVal = null;

        private static double[] _SRutDisVal = null;
        private static int[] _SRutDisMile = null;
        private static double[] _rutThresh = new double[2];
        private static int[][] _PBIVal = null;
        private static double[] _LDeltaHVal = null;
        private static double[] _RDeltaHVal = null;
        private static double[] _LDeltaHVal_1M = null;
        private static double[] _RDeltaHVal_1M = null;
        private static double[] _SpeedVal = null;
        private static string[] _MarkVal = null;

        private static double[] _DeltaHVal = null;

        private static double[] _LMPDMeanVal = null;
        private static double[] _RMPDMeanVal = null;
        private static double[] _CMPDMeanVal = null;

       

        private static double[] _Curvature = null;
        private static double[] _CrossSlope = null;
        private static double[] _HeightSlope = null;


        private static ExcelGPS[] _GPSInfo = null;
        private static void InitXlsParm()
        {
            int len = _RoadGradeStr.Length;

            _RQIGrade = new double[len][][];
            _RDIGrade = new double[len][];
            _PCIGrade = new double[len][];
            _PQIGrade = new double[len][];
            _PBIGrade = new double[len][];
            _PWIGrade = new double[len][];

            _RQIa = new double[len][][];
            _PCIa = new double[len][][];
            _PQIW = new double[len][][];

            _MQIW = new double[4];
            _MQIGrade = new double[5];

            for (int i = 0; i < len; i++)
            {
                _RQIGrade[i] = new double[2][];
                _RDIGrade[i] = new double[5];
                _PCIGrade[i] = new double[5];
                _PQIGrade[i] = new double[5];
                _PBIGrade[i] = new double[5];
                _PWIGrade[i] = new double[5];

                _RQIa[i] = new double[2][];
                _PCIa[i] = new double[2][];
                _PQIW[i] = new double[2][];
                for (int j = 0; j < 2; j++)
                {
                    _RQIGrade[i][j] = new double[5];
                    _PCIa[i][j] = new double[2];
                    _PQIW[i][j] = new double[5];
                    _RQIa[i][j] = new double[2];
                }
            }
            _PBIThresh = new double[4];
            _PBIScore = new double[4];
            _RDIa = new double[2];
            _PWIa = new double[2];
            _RDIRD = new double[2][];
            for (int i = 0; i < 2; i++)
            {
                _RDIRD[i] = new double[2];
            }

            _RoadSocre = new Dictionary<string, CityRoadDis>[2];
            for (int i = 0; i < 2; i++)
            {
                _RoadSocre[i] = new Dictionary<string, CityRoadDis>();
            }

            _WeightParm = new double[2][];

            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }
        }

        public static void LoadXlsParm()
        {
            InitXlsParm();

            XmlDocument Doc = new XmlDocument();
            Doc = new XmlDocument();
            XmlElement Elem;
            XmlNodeList xmlNodes;


            //读取病害类型
            Doc.Load(System.Windows.Forms.Application.StartupPath + "\\ParaVal.xml");    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点  
            xmlNodes = Elem.ChildNodes;

            for (int i = 0; i < 2; i++)
            {
                foreach (XmlNode rootchild in Elem.ChildNodes)
                {
                    if (rootchild.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
                    {
                        foreach (XmlNode subnode in rootchild.ChildNodes)
                        {
                            if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面病害类型")
                            {
                                foreach (XmlNode node in subnode.ChildNodes)
                                {
                                    CityRoadDis roaddis = new CityRoadDis();
                                    roaddis._DisName = node.Name;
                                    roaddis._UseWidth = Convert.ToDouble(((XmlElement)node).GetAttribute("影响宽度"));
                                    roaddis._Weight = Convert.ToDouble(((XmlElement)node).GetAttribute("权重"));
                                    _RoadSocre[i].Add(roaddis._DisName, roaddis);
                                }
                            }
                        }
                    }
                }
            }

            //读取等级区间
            int dlen = _RoadGradeStr.Length;
            string strval;
            string[] s;
            double[] val;
            for (int i = 0; i < dlen; i++)
            {
                foreach (XmlNode rootchild in Elem.ChildNodes)
                {
                    if (rootchild.Name == Framework.Other.MyGlobal. Global.g_ParmStyles[(int)_Setting.ParmStyle])
                    {
                        foreach (XmlNode subnode in rootchild.ChildNodes)
                        {
                            if (subnode.Name == _RoadGradeStr[i])
                            {
                                foreach (XmlNode node in subnode.ChildNodes)
                                {
                                    if (node.Name == "RQI")
                                    {
                                        foreach (XmlNode nnode in node.ChildNodes)
                                        {
                                            strval = ((XmlElement)nnode).GetAttribute("等级区间");
                                            s = strval.Split(' ');
                                            val = new double[s.Length];
                                            for (int j = 0; j < s.Length; j++)
                                            {
                                                val[j] = Convert.ToDouble(s[j]);
                                            }
                                            val.CopyTo(_RQIGrade[i][RoadDiseaseTypes.roadtypedict[nnode.Name]], 0);
                                            _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("w1"));
                                            _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("w2"));
                                        }
                                    }
                                    else
                                    {
                                        strval = ((XmlElement)node).GetAttribute("等级区间");
                                        s = strval.Split(' ');
                                        val = new double[s.Length];
                                        for (int j = 0; j < s.Length; j++)
                                        {
                                            val[j] = Convert.ToDouble(s[j]);
                                        }
                                        if (node.Name == "RDI")
                                        {
                                            val.CopyTo(_RDIGrade[i], 0);
                                        }
                                        else if (node.Name == "PWI")
                                        {
                                            val.CopyTo(_PWIGrade[i], 0);
                                        }
                                        else if (node.Name == "PBI")
                                        {
                                            val.CopyTo(_PBIGrade[i], 0);
                                        }
                                        else if (node.Name == "PCI")
                                        {
                                            val.CopyTo(_PCIGrade[i], 0);
                                            foreach (XmlNode nnode in node.ChildNodes)
                                            {
                                                _PCIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a0"));
                                                _PCIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a1"));
                                            }
                                        }
                                        else if (node.Name == "PQI")
                                        {
                                            val.CopyTo(_PQIGrade[i], 0);
                                            foreach (XmlNode nnode in node.ChildNodes)
                                            {
                                                _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPCI"));
                                                _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRQI"));
                                                _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][2] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRDI"));
                                                _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][3] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPBI"));
                                                _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][4] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPWI"));
                                            }
                                        }
                                    }
                                }
                            }
                            //读取计算RDI的系数
                            else if (i == 0 && subnode.Name == "RDI系数")
                            {
                                _RDIRD[0][0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙常数a"));
                                _RDIRD[1][0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙常数b"));
                                _RDIRD[0][1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙RDa"));
                                _RDIRD[1][1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙RDb"));
                                _RDIa[0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙a0"));
                                _RDIa[1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙a1"));
                            }
                            //读取计算PBI的系数
                            else if (i == 0 && subnode.Name == "PBI系数")
                            {
                                strval = ((XmlElement)subnode).GetAttribute("划分标准");
                                s = strval.Split(' ');
                                val = new double[s.Length];
                                for (int j = 0; j < s.Length; j++)
                                {
                                    val[j] = Convert.ToDouble(s[j]);
                                }
                                val.CopyTo(_PBIThresh, 0);

                                strval = ((XmlElement)subnode).GetAttribute("扣分");
                                s = strval.Split(' ');
                                val = new double[s.Length];
                                for (int j = 0; j < s.Length; j++)
                                {
                                    val[j] = Convert.ToDouble(s[j]);
                                }
                                val.CopyTo(_PBIScore, 0);
                            }
                            //读取计算PWI的系数
                            else if (i == 0 && subnode.Name == "PWI系数")
                            {
                                _PWIa[0] = double.Parse(((XmlElement)subnode).GetAttribute("a0"));
                                _PWIa[1] = double.Parse(((XmlElement)subnode).GetAttribute("a1"));
                            }
                            else if (i == 0 && subnode.Name == "MQI系数")
                            {
                                _MQIW[0] = double.Parse(((XmlElement)subnode).GetAttribute("WSCI"));
                                _MQIW[1] = double.Parse(((XmlElement)subnode).GetAttribute("WPQI"));
                                _MQIW[2] = double.Parse(((XmlElement)subnode).GetAttribute("WBCI"));
                                _MQIW[3] = double.Parse(((XmlElement)subnode).GetAttribute("WTCI"));

                                strval = ((XmlElement)subnode).GetAttribute("等级区间");
                                s = strval.Split(' ');
                                for (int j = 0; j < s.Length; j++)
                                {
                                    _MQIGrade[j] = Convert.ToDouble(s[j]);
                                }
                            }
                        }
                    }
                }
            }
        }

        private class CityRoadDis
        {
            public string _DisType = null;
            public string _DisName = null;
            public double _UseWidth = 0.0;
            public double _Weight = 0.0;
        }
        public static List<MilePartD> _RoadPartF = null;//0.1米桩号分段
        private static double[] _LiriHVal = null;
        private static double[] _RiriHVal = null;

        public static bool InitProDataD(DirectoryInfo prjdir, ProjectInfo prjinfo, double disval,
       bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
        {
            _SpeedVal = null;

            bool IRIRes = true, RutRes = true, MTDRes = true, PBIRes = true, GPSRes = true, MaxRutRes = true, MPDRes = true;


            #region 国检转换TP
            if (_RoadPartF != null)
            {
                _RoadPartF.Clear();
                _RoadPartF = null;
            }
            _RoadPartF = new List<MilePartD>();
            MilePartD spartF = null;
            try
            {
                spartF = new MilePartD() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
            }
            catch
            {
                MessageBox.Show(string.Format("【低等级农村公路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
                System.Environment.Exit(0);
            }
            _RoadPartF.Add(spartF);
            if (prjinfo._IsIRIMTD)
            {

                GlobalExcel.GetAllMilePartD(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPartF, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                if (IsSpeed)
                {
                    GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPartF, ref _SpeedVal);
                }

                if (IsPBI)
                {
                    //平整度原始数据 纵断面高程TP   

                    GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF,disval, 0, ref _LiriHVal);
                    if (prjinfo._IsDIRIMTD)
                    {
                        GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF,disval ,1, ref _RiriHVal);
                    }
                }


            }
            else
            {
                IRIRes = true;
            }

            if (prjinfo._IsRut)
            {
                if (IsMeanRut)
                {
                    RutRes = GlobalExcel.GetRutMeanVal(prjinfo, prjdir, _RoadPartF, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, _Setting.IsWarning);
                    MaxRutRes = GlobalExcel.GetRutMaxVal(prjinfo, prjdir, _RoadPartF, ref _LRutMaxVal, ref _RRutMaxVal, ref _SRutMaxVal);
                }
            }
            else
            {
                RutRes = true;
            }
            #endregion
            if (_RoadPartF[0].roaddegree <= 1)
            {
                return IRIRes && RutRes && MTDRes && GPSRes && MPDRes;
            }
            else
            {
                return IRIRes && MPDRes;
            }
        }

        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
            bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false, bool IsGeoAlig = false)
        {
            _SpeedVal = null;

            bool IRIRes = true, RutRes = true, MTDRes = true, PBIRes = true, GPSRes = true, MaxRutRes = true, MPDRes = true, GeoAligRes = true;
            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }
            _RoadPart = new List<MilePart>();

            if (_RoadPart10 != null)
            {
                _RoadPart10.Clear();
                _RoadPart10 = null;
            }
            _RoadPart10 = new List<MilePart>();

            if (_RoadPart1M != null)
            {
                _RoadPart1M.Clear();
                _RoadPart1M = null;
            }
        
            _RoadPart1M = new List<MilePart>();
            MilePart spart = null;
            try
            {
                spart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
            }
            catch
            {
                MessageBox.Show(string.Format("【等级公路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
                System.Environment.Exit(0);
            }
            _RoadPart.Add(spart);
            _RoadPart10.Add(spart);
            _RoadPart1M.Add(spart);
            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
            GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart, ref _MarkVal);
            if (IsDis)
            {
                GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 1, prjinfo._Direction, _RoadGradeStr, ref _RoadPart1M, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);

                if (_Setting.OutRut == 1 || (_Setting.OutRut == 2 && (_RoadGradeDict[prjinfo._RoadGrade] > 1)))
                {
                    GlobalExcel.GetRutDisVal(prjinfo, prjdir, _RoadPart1M, ref _SRutDisVal, ref _SRutDisMile);
                }
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, _RoadGradeDict, _SRutDisVal, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh, _RoadPart);
            }
            if (prjinfo._IsIRIMTD)
            {
                if (IsSpeed)
                {
                    GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPart, ref _SpeedVal);
                }
                if (IsMeanIRI)
                {
                    IRIRes = GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, _Setting.IsWarning);
                }
                if (IsPBI)
                {
                    GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 10, prjinfo._Direction, _RoadGradeStr, ref _RoadPart10, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                    GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart10, 0, ref _LDeltaHVal);
                    GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart, 0, ref _LDeltaHVal_1M);
                    if (prjinfo._IsDIRIMTD)
                    {
                        GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart10, 1, ref _RDeltaHVal);
                        GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart, 1, ref _RDeltaHVal_1M);
                    }
                    PBIRes = GlobalExcel.GetPBVal(prjinfo, prjdir, _RoadPart, _RoadPart10, ref _PBIVal, _PBIThresh, _LDeltaHVal, _RDeltaHVal, 0, ref _DeltaHVal);
                    GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart10, ref _MarkVal10);
                    GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPart10, ref _SpeedVal10);
                }
                if (IsMeanMTD&& !_Setting.isGDIriCalculate)
                {
                    MTDRes = GlobalExcel.GetMTDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, _Setting.IsWarning);
                   
                }
                if (IsMeanMPD&& !_Setting.isGDIriCalculate)
                {
                    MPDRes = GlobalExcel.GetMPDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMPDMeanVal, ref _RMPDMeanVal, ref _CMPDMeanVal, _Setting.IsWarning);
                 
                }
            }
            else
            {
                IRIRes = true;
            }

            if (prjinfo._IsRut)
            {
                if (IsMeanRut)
                {
                    RutRes = GlobalExcel.GetRutMeanVal(prjinfo, prjdir, _RoadPart, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, _Setting.IsWarning);
                    MaxRutRes = GlobalExcel.GetRutMaxVal(prjinfo, prjdir, _RoadPart, ref _LRutMaxVal, ref _RRutMaxVal, ref _SRutMaxVal);
                }
                if (IsGeoAlig)
                {
                    GeoAligRes = GlobalExcel.GetGeoAligVal(prjinfo, prjdir, _RoadPart, ref _Curvature, ref _CrossSlope, ref _HeightSlope, _Setting.IsWarning);
                }
            }
            else
            {
                RutRes = true;
            }

            if (_Setting.ExcelType == 4 || _Setting.ExcelType == 18 || _Setting.ExcelType == 15) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);

            if (_RoadPart[0].roaddegree <= 1)
            {
                return IRIRes && RutRes && MTDRes && GPSRes && MPDRes && GeoAligRes;
            }
            else
            {
                return IRIRes && MPDRes && GeoAligRes;
            }
        }

        private static void WriteProjectInfo2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string roadtypeinfo = File.ReadAllText(prjdir.FullName + "\\RoadTypeInfo.txt").Replace(" ", Environment.NewLine);

            object[,] val = new object[15, 1];
            val[0, 0] = prjinfo._Province;
            val[1, 0] = prjinfo._City;
            val[2, 0] = prjinfo._District;
            val[3, 0] = prjinfo._RoadCode;
            val[4, 0] = prjinfo._RoadName;
            val[5, 0] = prjinfo._StartMile;
            val[6, 0] = prjinfo._Direction > 0 ? "上行" : "下行";
            val[7, 0] = prjinfo._RoadGrade;
            val[8, 0] = prjinfo._RoadNum;
            val[9, 0] = prjinfo._DataDate;
            val[10, 0] = prjinfo._DataTime;
            val[11, 0] = prjinfo._DataPerson;
            val[12, 0] = prjinfo._DataWeather;
            //val[13, 0] = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            val[13, 0] = roadtypeinfo;
            val[14, 0] = prjinfo._EndMile; 

            MSExcel.Range destrange = _Worksheet.get_Range("B2:B16");
            destrange.Value2 = val;
        }
        public static void 江西磨耗(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\路面磨耗.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}-MPD-路面磨耗-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                ((double)prjinfo._StartMile / 1000).ToString("f3"),
                                                                                                ((double)prjinfo._EndMile / 1000).ToString("f3"));

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

            var roadpart = _RoadPart;
            int rowCount = 0;
            int len = roadpart.Count - 1;
            object[,] datas = new object[len, 8];

            for (int i = 0; i < len; i++)//i区间索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                datas[rowCount, 0] = prjinfo._RoadCode;
                datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                datas[rowCount, 2] = prjinfo._RoadNum;
                if (prjinfo._Direction > 0)
                {
                    datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                }
                else
                {
                    datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                }

                // 获取并写入磨耗数据
                datas[rowCount, 5] = _LMPDMeanVal[i].ToString("f2");
                datas[rowCount, 6] = _RMPDMeanVal[i].ToString("f2");
                datas[rowCount, 7] = _CMPDMeanVal[i].ToString("f2");

                rowCount++;

            }

            int StartRow = 3;
            MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:H{StartRow + rowCount - 1}");
            destrange.Value2 = datas;
            GlobalExcel.SetBorderLine(destrange, 53);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

        public static void 江西跳车(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\路面跳车.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}-PB-路面跳车-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                ((double)prjinfo._StartMile / 1000).ToString("f3"),
                                                                                                ((double)prjinfo._EndMile / 1000).ToString("f3"));

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

            var roadpart = _RoadPart;
            int rowCount = 0;
            int len = roadpart.Count - 1;
            object[,] datas = new object[len, 8];

            for (int i = 0; i < len; i++)//i区间索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                datas[rowCount, 0] = prjinfo._RoadCode;
                datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                datas[rowCount, 2] = prjinfo._RoadNum;
                if (prjinfo._Direction > 0)
                {
                    datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                }
                else
                {
                    datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                }

                // 获取并写入跳车数据
                datas[rowCount, 5] = _PBIVal[i][1].ToString("f2");
                datas[rowCount, 6] = _PBIVal[i][2].ToString("f2");
                datas[rowCount, 7] = _PBIVal[i][3].ToString("f2");

                rowCount++;

            }

            int StartRow = 3;
            MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:H{StartRow + rowCount - 1}");
            destrange.Value2 = datas;
            GlobalExcel.SetBorderLine(destrange, 53);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

        public static void 江西车辙(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\路面车辙.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}-RD-路面车辙-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                ((double)prjinfo._StartMile / 1000).ToString("f3"),
                                                                                                ((double)prjinfo._EndMile / 1000).ToString("f3"));

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

            var roadpart = _RoadPart;
            int rowCount = 0;
            int len = roadpart.Count - 1;
            object[,] datas = new object[len, 8];

            for (int i = 0; i < len; i++)//i区间索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                datas[rowCount, 0] = prjinfo._RoadCode;
                datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                datas[rowCount, 2] = prjinfo._RoadNum;
                if (prjinfo._Direction > 0)
                {
                    datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                }
                else
                {
                    datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                }

                // 获取并写入车辙
                datas[rowCount, 5] = _LRutMeanVal[i].ToString("f2");
                datas[rowCount, 6] = _RRutMeanVal[i].ToString("f2");
                datas[rowCount, 7] = Math.Max(_LRutMeanVal[i], _RRutMeanVal[i]).ToString("f2");

                rowCount++;

            }

            int StartRow = 3;
            MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:H{StartRow + rowCount - 1}");
            destrange.Value2 = datas;
            GlobalExcel.SetBorderLine(destrange, 53);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

        public static void 江西平整度(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\路面平整度.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}-IRI-路面平整度-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                ((double)prjinfo._StartMile / 1000).ToString("f3"),
                                                                                                ((double)prjinfo._EndMile / 1000).ToString("f3"));

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

            var roadpart = _RoadPart;
            int rowCount = 0;
            int len = roadpart.Count - 1;
            object[,] datas = new object[len, 8];

            for (int i = 0; i < len; i++)//i区间索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                datas[rowCount, 0] = prjinfo._RoadCode;
                datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                datas[rowCount, 2] = prjinfo._RoadNum;
                if (prjinfo._Direction > 0)
                {
                    datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                }
                else
                {
                    datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                    datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                }

                // 获取并写入平整度
                datas[rowCount, 5] = _LIRIMeanVal[i].ToString("f2");
                datas[rowCount, 6] = _RIRIMeanVal[i].ToString("f2");
                datas[rowCount, 7] = Math.Max(_LIRIMeanVal[i], _RIRIMeanVal[i]).ToString("f2");

                rowCount++;

            }

            int StartRow = 3;
            MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:H{StartRow + rowCount - 1}");
            destrange.Value2 = datas;
            GlobalExcel.SetBorderLine(destrange, 53);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }


        public static void 江西公路沥青病害(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, RoadConfig roadConfig)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\高等级沥青路面破损.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

                

            foreach (var line in roadTypeSplit)
            {
                if (!line.Contains("沥青"))
                {
                    continue;
                }
                double startMileD = double.Parse(line.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(line.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                if(startMileD > endMileStrD)
                {
                    (startMileD, endMileStrD) = (endMileStrD, startMileD);
                }
                // 文件命名==============
                string Destxls = string.Format(@"{0}\{1}-DR-高等级沥青路面破损-{2}-{3}.xlsx", path, prjinfo._RoadCode +   ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                (startMileD).ToString("f3"),
                                                                                                (endMileStrD).ToString("f3"));

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
                _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
                MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

                var roadpart = _RoadPart;
                var arrdis = _RoadDisList;
                int roadType = 0;
                bool has = false;

                string errlog = prjdir.FullName + "\\errlog.txt";
                int rowCount = 0;
                int len = roadpart.Count - 1, dlen = arrdis.Length;
                object[,] datas = new object[len, 19];
                int typeidx = 0;
                bool res = false;

                int colcnt = 1;
                for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
                {
                    int smile = roadpart[i].mile;
                    int emile = roadpart[i + 1].mile;
                    //统计位于这个区域的病害
                    RoadDiseaseTypes.Clear();
                    while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                          || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                                arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                        ++j;
                    }
                    if (roadpart[i].roadtype == roadType &&(startMileD*1000 <= roadpart[i].mile && roadpart[i].mile <= endMileStrD * 1000))
                    {
                        //有病害
                        has = true;
                        smile = roadpart[i].mile;
                        emile = roadpart[i + 1].mile;
                        int milelength = Math.Abs(smile - emile);

                        //病害汇总表
                        datas[rowCount, 0] = prjinfo._RoadCode;
                        datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                        datas[rowCount, 2] = prjinfo._RoadNum;
                        if (prjinfo._Direction > 0)
                        {
                            datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                            datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                        }
                        else
                        {
                            datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                            datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                        }

                        var 龟裂 = RoadDiseaseTypes.roaddis[roadType][0].totalarea + RoadDiseaseTypes.roaddis[roadType][1].totalarea + RoadDiseaseTypes.roaddis[roadType][2].totalarea;

                        datas[rowCount, 5] = 龟裂.ToString("f2");
                        colcnt = 6;

                        for (int dis = 3; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                        {
                            double val = 0;
                            if (dis < 19 && dis % 2 == 1)
                            {// 合并同一病害类型下的轻重等级
                                val = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                dis += 1;
                            }
                            val += RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                            datas[rowCount, colcnt++] = val.ToString("f3");
                        }

                        datas[rowCount, 17] = roadConfig.DetectWidth;
                        datas[rowCount, 18] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength); // 结果示例表中 未格式化输出
                        rowCount++;
                    }
                }
                if (has)
                {
                    int StartRow = 5;
                    MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:S{StartRow + rowCount - 1}");
                    destrange.Value2 = datas;
                    GlobalExcel.SetBorderLine(destrange, 53);
                }

                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();

            }

        }

        public static void 江西公路水泥病害(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, RoadConfig roadConfig)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\江西车检\高等级水泥路面破损.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (var line in roadTypeSplit)
            {
                if (!line.Contains("水泥"))
                {
                    continue;
                }
                double startMileD = double.Parse(line.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(line.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                if (startMileD > endMileStrD)
                {
                    (startMileD, endMileStrD) = (endMileStrD, startMileD);
                }

                // 文件命名==============
                string Destxls = string.Format(@"{0}\{1}-DR-高等级水泥路面破损-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                (startMileD).ToString("f3"),
                                                                                                (endMileStrD).ToString("f3"));

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
                _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
                MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

                var roadpart = _RoadPart;
                var arrdis = _RoadDisList;
                int roadType = 1;
                bool has = false;

                string errlog = prjdir.FullName + "\\errlog.txt";
                int rowCount = 0;
                int len = roadpart.Count - 1, dlen = arrdis.Length;
                object[,] datas = new object[len, 19];
                int typeidx = 0;
                bool res = false;

                int colcnt = 1;
                for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
                {
                    int smile = roadpart[i].mile;
                    int emile = roadpart[i + 1].mile;
                    //统计位于这个区域的病害
                    RoadDiseaseTypes.Clear();
                    while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                          || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                                arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                        ++j;
                    }
                    int[] MidLevel = new int[] { 3, 6, 12 };

                    if (roadpart[i].roadtype == roadType && (startMileD * 1000 <= roadpart[i].mile && roadpart[i].mile <= endMileStrD * 1000))
                    {
                        //有病害
                        has = true;
                        smile = roadpart[i].mile;
                        emile = roadpart[i + 1].mile;
                        int milelength = Math.Abs(smile - emile);

                        //病害汇总表
                        datas[rowCount, 0] = prjinfo._RoadCode;
                        datas[rowCount, 1] = (prjinfo._Direction > 0) ? "上行" : "下行";
                        datas[rowCount, 2] = prjinfo._RoadNum;
                        if (prjinfo._Direction > 0)
                        {
                            datas[rowCount, 3] = (smile * 0.001).ToString("f3");
                            datas[rowCount, 4] = (emile * 0.001).ToString("f3");

                        }
                        else
                        {
                            datas[rowCount, 3] = (emile * 0.001).ToString("f3");
                            datas[rowCount, 4] = (smile * 0.001).ToString("f3");
                        }

                        colcnt = 5;
                        for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                        {
                            double val = 0;
                            // 合并同一病害类型下的轻重等级
                            if (dis < 16 && dis != 10)
                            {// 轻等级
                                val += RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                dis += 1;
                            }
                            if (MidLevel.Contains(dis))
                            {// 中等级
                                val += RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                dis += 1;
                            }
                            val += RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                            datas[rowCount, colcnt++] = val.ToString("f3");
                        }

                        datas[rowCount, 17] = roadConfig.DetectWidth;
                        datas[rowCount, 18] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength); // 结果示例表中 未格式化输出
                        rowCount++;
                    }
                }
                if (has)
                {
                    int StartRow = 5;
                    MSExcel.Range destrange = _Worksheet.get_Range($"A{StartRow}:S{StartRow + rowCount - 1}");
                    destrange.Value2 = datas;
                    GlobalExcel.SetBorderLine(destrange, 53);
                }

                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
            }
        }

        public static void OutputChongQingDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\四川公路院\病害明细表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_重庆病害统计明细表_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["病害明细"] as MSExcel.Worksheet;
            #region 写入数据
            int len = _RoadPart.Count - 1, dlen = _RoadDisList.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 13];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = _RoadPart[i].mile;
                int emile = _RoadPart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && _RoadDisList[j].m_mile >= smile && _RoadDisList[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && _RoadDisList[j].m_mile <= smile && _RoadDisList[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[_RoadPart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            _RoadDisList[j].RoadType, _RoadDisList[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = _RoadDisList[j].RoadDisType.Split('.');


                        vallist[rowcnt, 0] = rowcnt;
                        vallist[rowcnt, 1] = prjinfo._RoadCode;
                        vallist[rowcnt, 2] = prjinfo._Direction == 1 ? "上行" : "下行";
                        vallist[rowcnt, 3] = (_RoadDisList[j].m_mile / 1000.0).ToString("g3");
                        vallist[rowcnt, 4] = prjinfo._RoadNum;
                        vallist[rowcnt, 5] = GlobalExcel._RoadTypeStr[_RoadPart[i].roadtype];
                        vallist[rowcnt, 6] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 7] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 7] = "无";
                        }
                        vallist[rowcnt, 8] = _RoadDisList[j].calcheight;
                        vallist[rowcnt, 9] = _RoadDisList[j].calcwidth;
                        vallist[rowcnt, 10] = _RoadDisList[j].Area;

                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", _RoadDisList[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[_RoadPart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A3:M{0}", dlen + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);
            #endregion
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面平整度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_IRI_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'H', "平整度", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }
            if (_Setting.ExcelType != 19)
            {
                MSExcel.Worksheet _WorksheetPrjInfo = null;
                _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
                WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            }
            else
            {
                MSExcel.Worksheet _WorksheetPrjInfo = null;
                _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
                _WorksheetPrjInfo.Delete();
            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            if (_Setting.RQIJudgeType == 1)
            {
                _Worksheet.Cells[2, 6] = "最大IRI";
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 0 || _Setting.IRIExcelSide == 2)
                    {
                        vallist[i, 3] = LIRIVal[i];
                    }
                    if (_Setting.IRIExcelSide == 1 || _Setting.IRIExcelSide == 2)
                    {
                        vallist[i, 4] = RIRIVal[i];
                    }

                    if (_Setting.RQIJudgeType == 0)
                    {
                        vallist[i, 5] = String.Format("=ROUND(AVERAGE(D{0}:E{0}),5)", i + DataStartXlsxRow);
                    }
                    else if (_Setting.RQIJudgeType == 1)
                    {
                        vallist[i, 5] = String.Format("=ROUND(MAX(D{0}, E{0}),5)", i + DataStartXlsxRow);
                    }
                }
                else
                {
                    vallist[i, 3] = LIRIVal[i];
                    vallist[i, 4] = 0;
                    vallist[i, 5] = LIRIVal[i];

                } 
                vallist[i, 6] = String.Format("=ROUND(100/(1+{0}*EXP({1}*F{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + DataStartXlsxRow);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }
        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, double disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\车辙深度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_Rut_{2}m.xlsx", path, prjdir.Name, disval);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteRut2Xls_orirut(_Worksheet, prjinfo, _RoadPartF, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'H', "车辙深度", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            //_Worksheet = _Workbook.Sheets["Sheet2"] as MSExcel.Worksheet;
            //WriteMaxRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMaxVal, _RRutMaxVal, _SRutMaxVal, _SpeedVal, _MarkVal);
            //WriteRutStatistics(_Worksheet);
            //if (_Setting.Out_roadinfo == 0)
            //{
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            //}

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
         List<MilePartD> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal, double[] SpeedVal, string[] MarkVal,
         int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsRut)
                return;

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = Math.Round(roadpart[i].mile, 1);
                vallist[i, 1] = Math.Round(roadpart[i + 1].mile, 1);
                vallist[i, 2] = prjinfo._RoadNum;

                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];

                vallist[i, 5] = SRutVal[i];
                vallist[i, 6] = string.Format("=IF(F{0}<={1},{2}-{3}*F{0},IF(F{0}<={4},{5}-{6}*(F{0}-{1}),0))",
                        i + DataStartXlsxRow, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                //vallist[i, 10] = MarkVal[i];
            }
            destrange = _Worksheet.get_Range(String.Format("A{0}:B{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.ClearFormats();

            destrange = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;

            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }


        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\车辙深度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_Rut_{2}m.xlsx", path, prjdir.Name, disval);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'H', "车辙深度", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            //_Worksheet = _Workbook.Sheets["Sheet2"] as MSExcel.Worksheet;
            //WriteMaxRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMaxVal, _RRutMaxVal, _SRutMaxVal, _SpeedVal, _MarkVal);
            //WriteRutStatistics(_Worksheet);
            //if (_Setting.Out_roadinfo == 0)
            //{
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            //}

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsRut)
                return;

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];
                 
                   vallist[i, 5] = SRutVal[i]; 
                vallist[i, 6] = string.Format("=IF(F{0}<={1},{2}-{3}*F{0},IF(F{0}<={4},{5}-{6}*(F{0}-{1}),0))",
                        i + DataStartXlsxRow, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        private static void WriteMaxRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
          List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsRut)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];
                vallist[i, 5] = SRutVal[i];

                vallist[i, 6] = string.Format("=IF(F{0}<={1},{2}-{3}*F{0},IF(F{0}<={4},{5}-{6}*(F{0}-{1}),0))",
                        i + DataStartXlsxRow, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);

                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputPWI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面磨耗评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PWI_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePWI2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'I', "磨耗", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 12]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePWI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 12];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LMTDVal[i];
                vallist[i, 4] = RMTDVal[i];
                vallist[i, 5] = CMTDVal[i];

                if (CMTDVal[i] == 0)
                {
                    vallist[i, 6] = 0;
                }
                else
                {
                    vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ", i + DataStartXlsxRow);
                }
                vallist[i, 7] = string.Format("=100-{0}*POWER(G{1},{2})", _PWIa[0], i + DataStartXlsxRow, _PWIa[1]);
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow,
                    _PWIGrade[roadpart[i].roaddegree][0],
                    _PWIGrade[roadpart[i].roaddegree][1],
                    _PWIGrade[roadpart[i].roaddegree][2],
                    _PWIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 10] = SpeedVal[i];
                }
                vallist[i, 11] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:L{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 12, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面构造深度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_SMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _SpeedVal, _MarkVal, 4, 53);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 9]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 8]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 9];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LMTDVal[i];

                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RMTDVal[i];
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + DataStartXlsxRow);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND(D{0},5)", i + DataStartXlsxRow);
                }
                vallist[i, 6] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 7] = SpeedVal[i];
                }
                vallist[i, 8] = MarkVal[i];
            }
            destrange = _Worksheet.get_Range(String.Format("A{0}:I{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 9, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputMPD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面构造深度MPD评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_MPD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteMPD2Xls(_Worksheet, prjinfo, _RoadPart, _LMPDMeanVal, _RMPDMeanVal, _CMPDMeanVal, _SpeedVal, _MarkVal, 4);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'I', "磨耗", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 12]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteMPD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMPDVal, double[] RMPDVal, double[] CMPDVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 12];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LMPDVal[i];
                vallist[i, 4] = RMPDVal[i];
                vallist[i, 5] = CMPDVal[i];

                if (CMPDVal[i] == 0)
                {
                    vallist[i, 6] = 0;
                }
                else
                {
                    vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ", i + DataStartXlsxRow);
                }
                vallist[i, 7] = string.Format("=100-{0}*POWER(G{1},{2})", _PWIa[0], i + DataStartXlsxRow, _PWIa[1]);
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow,
                    _PWIGrade[roadpart[i].roaddegree][0],
                    _PWIGrade[roadpart[i].roaddegree][1],
                    _PWIGrade[roadpart[i].roaddegree][2],
                    _PWIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 10] = SpeedVal[i];
                }
                vallist[i, 11] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:L{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 12, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputCPMSDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            disval *= 10;
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CPMS路面病害调查表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_CPMS调查表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lqdc = _Workbook.Sheets["沥青路面损坏调查表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sndc = _Workbook.Sheets["水泥路面损坏调查表"] as MSExcel.Worksheet;

            WritePrj2CPMSXls(_Worksheet_lqdc, prjinfo);
            WritePrj2CPMSXls(_Worksheet_sndc, prjinfo);
            WriteZJGTDisHZTJ2Xls(_Worksheet_sndc, _Worksheet_lqdc, prjinfo, prjdir, _RoadPart, _RoadDisList, disval, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            if (_Setting.ExcelType == 19)
            {
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\甘肃\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            }
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (_Setting.ExcelType == 19)
            {
                bool Haslqflag = false;
                bool Hassnflag = false;
                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);
                 
            }
            else
            {
                bool Haslqflag = false;
                bool Hassnflag = false;
                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);

                MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
                WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);


                MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjdir, _RoadPart, Haslqflag, Hassnflag);

                MSExcel.Worksheet _WorksheetPrjInfo = null;
                _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
                WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);
            }
           

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        private static void WriteDisLB2Xls_roadpart(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
          DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 14];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        vallist[rowcnt, 1] = prjinfo._RoadNum;
                        vallist[rowcnt, 2] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 3] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 3] = "无";
                        }
                        vallist[rowcnt, 4] = arrdis[j].rect.Height * _RoadConfig.HeightScale;
                        vallist[rowcnt, 5] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 6] = (arrdis[j].rect.Width / 2 + arrdis[j].rect.X) * _RoadConfig.WidthScale;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        vallist[rowcnt, 8] = arrdis[j].calcheight;
                        vallist[rowcnt, 9] = arrdis[j].calcwidth;
                        vallist[rowcnt, 10] = arrdis[j].imgname;
                        vallist[rowcnt, 11] = arrdis[j].imgpath;
                        vallist[rowcnt, 12] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                        vallist[rowcnt, 13] = arrdis[j].remarks;
                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A3:N{0}", dlen + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Qufen_dis_degree == 1)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 4]).EntireColumn.Delete();
                if (_Setting.Out_roadimg == 0)
                {
                    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
                }
            }
            else
            {
                if (_Setting.Out_roadimg == 0)
                {
                    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 12]).EntireColumn.Delete();
                    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                }
            }
            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, 3, 1, 14, true);
            //}
        }

        public static void OutputDis_JiangXi2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
           
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);

            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (_Setting.ExcelType == 19)
            {
                bool Haslqflag = false;
                bool Hassnflag = false;
                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);

            }
            else
            {
                bool Haslqflag = false;
                bool Hassnflag = false;
                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);

                MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
                WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);


                MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjdir, _RoadPart, Haslqflag, Hassnflag);

                MSExcel.Worksheet _WorksheetPrjInfo = null;
                _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
                WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);
            }


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        #region 病害调绘表
        /// <summary>
        /// 病害调绘表 
        /// 病害表基础上 删去了后   病害中心位置（距路面图像左边距离）（m）后几列
        //  20230110
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void OutputDis_TH(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            //调绘表需要进行分表
            WriteDisLB2Xls_roadpart_TH(path, prjinfo, prjdir, _RoadDisList, _RoadPart, _GPSInfo, path, disval);
        }

        private static void WriteDisLB2Xls_roadpart_TH(string outPath, ProjectInfo prjinfo,
    DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart, ExcelGPS[] gpsInfos, string path, int disval)
        {
            int splitValue = 1000;
            MSExcel.Application excelApp = new MSExcel.Application()
            {
                Visible = true,
                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                AlertBeforeOverwriting = false
            };
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\调绘报表模板.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害调绘统计_{2}m_0.xlsx", path, prjdir.Name, disval);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

             

            MSExcel.Range destrange  ;
            int generation = 0;
            int splitNum = 0; 
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 11];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;
            string thDisPicPath = outPath + "\\调绘病害图片";
            if (Directory.Exists(thDisPicPath))
            {
                Directory.Delete(thDisPicPath, true);
            }
            Directory.CreateDirectory(thDisPicPath);
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        if (j!=0&&j % splitValue == 0)
                        {
                            splitNum++;
                            rowcnt = 0;
                           
                            #region 存入数据 关闭表格

                            destrange = _Worksheet.get_Range(String.Format("A3:K{0}", splitValue+ 2));
                            destrange.Value2 = vallist;
                            GlobalExcel.SetBorderLine(destrange, 53);

                            if (_Setting.Qufen_dis_degree == 1)
                            {
                                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
                            }

                            _Workbook.Save();
                            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);


                            generation = System.GC.GetGeneration(excelApp);
                            System.GC.Collect(generation);//垃圾回收
                            System.GC.WaitForPendingFinalizers();
                            excelApp.Quit();
                        
                            #endregion

                            //统计每次记录1000条新开一个表格

                            vallist = new object[dlen, 11];

                            excelApp = new MSExcel.Application()
                            {
                                Visible = true,
                                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                                AlertBeforeOverwriting = false
                            };
                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\调绘报表模板.xlsx",
            System.Windows.Forms.Application.StartupPath);
                            Destxls = string.Format(@"{0}\{1}_病害调绘统计_{2}m_{3}.xlsx", path, prjdir.Name, disval, splitNum);

                            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

                        }
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        if (string.IsNullOrEmpty(prjinfo._RoadNum))
                        {
                            vallist[rowcnt, 1] = "1";
                        }
                        else
                        {
                            vallist[rowcnt, 1] = prjinfo._RoadNum;
                        }

                        vallist[rowcnt, 2] = roadpart[i].roadtype == 0 ? "沥青" : roadpart[i].roadtype == 1 ? "水泥" : "砂石";
                        vallist[rowcnt, 3] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 4] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 4] = "无";
                        }
                        vallist[rowcnt, 5] = arrdis[j].rect.Height * _RoadConfig.HeightScale;

                        vallist[rowcnt, 6] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        vallist[rowcnt, 8] = gpsInfos[i]._latitude;
                        vallist[rowcnt, 9] = gpsInfos[i]._longitude;
                        vallist[rowcnt, 10] = gpsInfos[i]._elevation;
                        var nowDis = arrdis[j];
                        string picPath = prjinfo._PrjPath + Path.Combine(arrdis[j].imgpath, arrdis[j].imgname);

                        var picRange = _Worksheet.Range[$"L{3 + rowcnt}"];
                        // picRange.RowHeight = 27.682 * 18;
                        float widthC = 27.682f * 1.05f;

                        double ratio;
                        using (System.Drawing.Bitmap map = new System.Drawing.Bitmap(picPath))
                        {
                            ratio = (double)map.Height / (double)map.Width;
                        }

                        picRange.RowHeight = widthC * 3.112f;
                        picRange.ColumnWidth = widthC;
                        Framework.Office.Excel.CWB_ExcelHelper.InsertPicture(picRange, _Worksheet, picPath, ratio);
                        var hyperRange = _Worksheet.Range[$"M{3 + rowcnt}"];

                        string tempStr = "\\" + arrdis[j].imgpath.Split('\\').Last();
                        string thPicPath = thDisPicPath;
                        string hyperPath = "调绘病害图片\\" + arrdis[j].imgname;


                        Directory.CreateDirectory(thPicPath);
                        thPicPath += "\\" + arrdis[j].imgname;

                        File.Copy(picPath, thPicPath, true);

                        hyperRange.ColumnWidth = widthC * 4;
                        //var o= hyperRange.Select();
                        _Worksheet.Hyperlinks.Add(hyperRange, hyperPath);
                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            #region 存入数据 关闭表格

              destrange = _Worksheet.get_Range(String.Format("A3:K{0}",(dlen- splitNum*splitValue) + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Qufen_dis_degree == 1)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers(); 
            excelApp.Quit();
            #endregion

        }

        #endregion



        /// <summary>
        /// 比WriteDisLB2Xls_roadpart多了是否在轮迹带和距右侧边线距离，不输出图像路径和材质
        /// </summary>
        /// <param name="_Worksheet"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="arrdis"></param>
        /// <param name="roadpart"></param>
        private static void WriteDisLB2Xls_roadpart2(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 11];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        vallist[rowcnt, 1] = prjinfo._RoadNum;
                        vallist[rowcnt, 2] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 3] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 3] = "无";
                        }

                        double leftx = arrdis[j].rect.X * _RoadConfig.WidthScale;
                        double rightx = (arrdis[j].rect.Width + arrdis[j].rect.X) * _RoadConfig.WidthScale;

                        vallist[rowcnt, 4] = arrdis[j].rect.Height * _RoadConfig.HeightScale;
                        vallist[rowcnt, 5] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 6] = (rightx + leftx - _RoadConfig.DetectWidth) / 2;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        vallist[rowcnt, 8] = arrdis[j].calcheight;
                        vallist[rowcnt, 9] = arrdis[j].calcwidth;

                        if (leftx >= _RoadConfig.DetectWidth / 2 + 1.4
                            || rightx <= _RoadConfig.DetectWidth / 2 - 1.4
                            || leftx >= _RoadConfig.DetectWidth / 2 - 0.7 && rightx <= _RoadConfig.DetectWidth / 2 + 0.7)
                        {
                            vallist[rowcnt, 10] = "否";
                        }
                        else
                        {
                            vallist[rowcnt, 10] = "是";
                        }

                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A3:K{0}", dlen + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 11, true);
            }
        }

        private static void WriteDisHZ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            ref bool Haslqflag, ref bool Hassnflag,
            int DataStartXlsxRow, int borderType)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志

            int rowcnt_sn_s = DataStartXlsxRow;
            int rowcnt_sn_e = DataStartXlsxRow;//小计起始的计算范围
            int rowcnt_lq_s = DataStartXlsxRow;
            int rowcnt_lq_e = DataStartXlsxRow;

            int totalsnlen = 0;//水泥路段总长度
            int totallqlen = 0;//沥青路段总长度

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[j].calcheight;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 1;
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;

                    if (prjinfo._Direction == -1)
                    {

                        if (_Setting.IsExcelSort)
                        {
                            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                        }
                        else
                        {
                            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                        }
                    }
                    else
                    {
                        worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                        worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                    }


                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = prjinfo._RoadNum; 
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        if (_Setting.ExcelType == 19)
                        {
                            if (RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains("裂缝"))
                            {
                                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totallength;

                            }
                            else
                                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;

                        }
                        else
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER(Y{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_sn_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    disval[0, disnum + 2] = string.Format("=IF(Z{0}>={1},\"优\",IF(Z{0}>={2},\"良\",IF(Z{0}>={3},\"中\",IF(Z{0}>={4},\"次\",\"差\"))))",
                        rowcnt_sn_s,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3]);

                    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
                    destrange.Value2 = disval;

                    totalsnlen += milelength;
                    rowcnt_sn_s++;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    if (prjinfo._Direction == -1)
                    {

                        if (_Setting.IsExcelSort)
                        {
                            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                        }
                        else
                        {
                            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                        }
                    }
                    else
                    {
                        worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                        worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                    }
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        if (_Setting.ExcelType == 19)
                        {
                            if (RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains("纵向裂缝")
                                || RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains("横向裂缝"))
                            {
                                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totallength;
                            }
                            else
                                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;

                        }
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER(Z{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_lq_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    disval[0, disnum + 2] = string.Format("=IF(AA{0}>={1},\"优\",IF(AA{0}>={2},\"良\",IF(AA{0}>={3},\"中\",IF(AA{0}>={4},\"次\",\"差\"))))",
                        rowcnt_lq_s,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3]);

                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
                    destrange.Value2 = disval;

                    totallqlen += milelength;
                    rowcnt_lq_s++;
                }

                if (_Setting.IsOutputDisAreaSubtotal)
                {
                    if (emile % 1000 == 0)
                    {
                        if (roadpart[i].roadtype == 1)
                        {
                            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                            worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                            }
                            destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_sn_s++;
                            rowcnt_sn_e = rowcnt_sn_s;

                            if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                            {
                                GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                                worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                                disval = new object[1, disnum];
                                for (int di = 0; di < disnum; di++)
                                {
                                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                                }
                                destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                                destrange.Value2 = disval;
                                rowcnt_lq_s++;
                                rowcnt_lq_e = rowcnt_lq_s;
                            }
                        }
                        else if (roadpart[i].roadtype == 0)
                        {
                            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                            worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                            }
                            destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_lq_s++;
                            rowcnt_lq_e = rowcnt_lq_s;

                            if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                            {
                                GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                                worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                                disval = new object[1, disnum];
                                for (int di = 0; di < disnum; di++)
                                {
                                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                                }
                                destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                                destrange.Value2 = disval;
                                rowcnt_sn_s++;
                                rowcnt_sn_e = rowcnt_sn_s;
                            }
                        }
                    }
                }
            } 
            //最后的一个小计
            if (_Setting.IsOutputDisAreaSubtotal)
            {
                if (roadpart[len].mile % 1000 != 0)
                {
                    if (roadpart[len].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                        }
                        destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_sn_s++;
                        rowcnt_sn_e = rowcnt_sn_s;

                        if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                            worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                            }
                            destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_lq_s++;
                            rowcnt_lq_e = rowcnt_lq_s;
                        }
                    }
                    else if (roadpart[len].roadtype == 0)
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                        }
                        destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_lq_s++;
                        rowcnt_lq_e = rowcnt_lq_s;

                        if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                            worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                            }
                            destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_sn_s++;
                            rowcnt_sn_e = rowcnt_sn_s;
                        }
                    }
                }
            }
            if (_Setting.ExcelType != 19)
            {
                //总计
                //水泥
                GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0);
                worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                    {
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
                    }
                    else
                    {
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
                    }

                }
                destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                destrange.Value2 = disval;

                //沥青
                GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
                worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
                    else
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
                }
                destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                destrange.Value2 = disval;
            }

       
            destrange = worksheet_lqhz.get_Range(String.Format("A1:AB{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_snhz.get_Range(String.Format("A1:AA{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            RoadDiseaseTypes.Clear();

          


            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }
            else
            {
                if (!_Setting.IsOutputDisAreaSubtotal)
                {
                    if (prjinfo._Direction == -1)
                    {

                        if (_Setting.IsExcelSort)
                        {
                            MSExcel.Range destrange1 = worksheet_lqhz.get_Range(string.Format("A5:AB{0}", rowcnt_lq_s - 1));
                            MSExcel.Range sortrange = worksheet_lqhz.get_Range(string.Format("A5:A{0}", rowcnt_lq_s - 1));
                            GlobalExcel.ReflectionColnum(worksheet_lqhz, destrange1, sortrange);
                        } 

                    }


                }
            }

            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }
            else
            {
                if (!_Setting.IsOutputDisAreaSubtotal)
                {
                    if (prjinfo._Direction == -1)
                    {

                       
                        if (_Setting.IsExcelSort)
                        {
                            MSExcel.Range destrange1 = worksheet_snhz.get_Range(string.Format("A5:AA{0}", rowcnt_sn_s - 1));
                            MSExcel.Range sortrange = worksheet_snhz.get_Range(string.Format("A5:A{0}", rowcnt_sn_s - 1));
                            GlobalExcel.ReflectionColnum(worksheet_snhz, destrange1, sortrange);
                        }

                    }


                }
            }

        }
        private static void WriteDisTJ2Xls(MSExcel.Worksheet worksheet_sntj, MSExcel.Worksheet worksheet_lqtj,
            DirectoryInfo prjdir, List<MilePart> roadpart, bool Haslqflag, bool Hassnflag)
        {
            int disnum = 0;
            MSExcel.Range destrange = null;
            object[,] disval = null;
            int len = roadpart.Count - 1;

            RoadDiseaseTypes.Clear();
            if (Haslqflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                worksheet_lqtj.Cells[2, 2] = _RoadConfig.DetectWidth;
                worksheet_lqtj.Cells[2, 6] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    else
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/2", Convert.ToChar('D' + i));
                    //disval[i, 0] = string.Format("=沥青病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_lq_s);

                }
                destrange = worksheet_lqtj.get_Range("C4:C" + (disnum + 3).ToString());
                destrange.Value2 = disval;
            }
            else
            {
                worksheet_lqtj.Delete();
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                worksheet_sntj.Cells[2, 2] = _RoadConfig.DetectWidth;
                worksheet_sntj.Cells[2, 6] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    else
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/2", Convert.ToChar('D' + i));
                    //disval[i, 0] = string.Format("=水泥病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_sn_s);

                }
                destrange = worksheet_sntj.get_Range("C4:C" + (disnum + 3).ToString());
                destrange.Value2 = disval;
            }
            else
            {
                worksheet_sntj.Delete();
            }
        }

        private static double ComputPCI(RoadDiseaseType[][] disarea, int roadtype, double partarea)
        {
            double sumarea = 0;
            int len = _RoadSocre[roadtype].Keys.Count;

            for (int i = 0; i < len; i++)
            {
                sumarea += disarea[roadtype][i].totalarea * disarea[roadtype][i].weight;
            }
            return 100 * sumarea / partarea;
        }

        public static void OutputPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面破损评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PCI_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WritePCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _SpeedVal, _MarkVal, 3);
            WriteStatistics_XMJH(_Worksheet, 3, 3, 22, 'F', "破损", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 9]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 8]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 9];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                vallist[i, 0] = smile;
                vallist[i, 1] = emile;
                vallist[i, 2] = prjinfo._RoadNum;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 3] = drval;

                vallist[i, 4] = string.Format("=IF(G{0}=\"沥青\",IF(100-{1}*POWER(D{0},{2})>0,100-{1}*POWER(D{0},{2}),100-{1}*POWER(D{0},{2})),IF(100-{3}*POWER(D{0},{4})>0,100-{3}*POWER(D{0},{4}),0))",
                    i + DataStartXlsxRow,
                    _PCIa[roadpart[i].roaddegree][0][0],
                    _PCIa[roadpart[i].roaddegree][0][1],
                    _PCIa[roadpart[i].roaddegree][1][0],
                    _PCIa[roadpart[i].roaddegree][1][1]);

                vallist[i, 5] = string.Format("=IF(E{0}>={1},\"优\",IF(E{0}>={2},\"良\",IF(E{0}>={3},\"中\",IF(E{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);

                vallist[i, 6] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 7] = SpeedVal[i];
                }
                vallist[i, 8] = MarkVal[i];
            }

            destrange = worksheet.get_Range(String.Format("A{0}:I{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 9, true);
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面综合评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PQI_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _SpeedVal, _MarkVal);
            WriteStatistics_XMJH(_Worksheet, 3, 3, 28, 'O', "PQI", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 18]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 17]).EntireColumn.Delete();
            }
            if (prjinfo._RoadGrade != "高速公路" && prjinfo._RoadGrade != "一级公路")
            {
                for (int i = 0; i < 6; ++i)
                {
                    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 8]).EntireColumn.Delete();
                }
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal, double[] SpeedVal, string[] MarkVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0, wrval = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 18];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 0;
                vallist[rowcnt, colcnt++] = smile;
                vallist[rowcnt, colcnt++] = emile;
                vallist[rowcnt, colcnt++] = prjinfo._RoadNum;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt++] = Math.Round(pcival, 5);
                vallist[rowcnt, colcnt++] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);
                #region 判断是否没有平整度文件夹
        
                //判断是否存在IRIMTD文件夹 
                string iriDirPath = prjinfo._PrjPath + "\\IRIMTD";
                if (Directory.Exists(iriDirPath))
                {
                    //IRI
                    if (prjinfo._IsDIRIMTD)
                    {
                        if (LIRIVal == null)
                        {
                            throw new Exception("目前该工程不具备左平整度数据请检查！");
                        }
                        if (_Setting.IRIExcelSide == 2)
                        {
                            if (RIRIVal == null)
                            {
                                throw new Exception("目前该工程不具备右平整度数据请检查！");
                            }
                            if (_Setting.RQIJudgeType == 0)
                            {
                                irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                            }
                            else if (_Setting.RQIJudgeType == 1)
                            {
                                irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                            }
                        }
                        else if (_Setting.IRIExcelSide == 0)
                        {
                            irival = Math.Round(LIRIVal[i], 5);
                        }
                        else if (_Setting.IRIExcelSide == 1)
                        {
                            irival = Math.Round(RIRIVal[i], 5);
                        }
                    }
                    else
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                    vallist[rowcnt, colcnt] = trqival;
                    colcnt++;
                    vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                        rowcnt + 3,
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    //Rut
                    if (prjinfo._IsRut && SRutVal != null && SRutVal.Length > 0)
                    {
                     
                            rutval = SRutVal[i];
                            rutval = Math.Round(rutval, 5);
                        
                        double rdival = 0;
                        if (rutval <= _RDIRD[0][1])
                        {
                            rdival = _RDIRD[0][0] - _RDIa[0] * rutval;
                        }
                        else if (rutval <= _RDIRD[1][1])
                        {
                            rdival = _RDIRD[1][0] - _RDIa[1] * (rutval - _RDIRD[0][1]);
                        }
                        else
                        {
                            rdival = 0;
                        }
                        // if(roadpart[i].roadtype==)
                        vallist[rowcnt, colcnt++] = rdival;
                        vallist[rowcnt, colcnt++] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                            rowcnt + 3,
                            _RDIGrade[roadpart[i].roaddegree][0],
                            _RDIGrade[roadpart[i].roaddegree][1],
                            _RDIGrade[roadpart[i].roaddegree][2],
                            _RDIGrade[roadpart[i].roaddegree][3]);
                    }
                    else
                    {
                        colcnt = colcnt + 2;
                    }
                    //PBI
                    if (prjinfo._IsIRIMTD && PBVal != null && PBVal.Length > 0)
                    {
                        vallist[rowcnt, colcnt++] = string.Format("=IF((100-{0}*{1}-{2}*{3}-{4}*{5})>0,(100-{0}*{1}-{2}*{3}-{4}*{5}),0)",
                            PBVal[i][1], _PBIScore[1],
                            PBVal[i][2], _PBIScore[2],
                            PBVal[i][3], _PBIScore[3]);
                        vallist[rowcnt, colcnt++] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                            rowcnt + 3,
                            _PBIGrade[roadpart[i].roaddegree][0],
                            _PBIGrade[roadpart[i].roaddegree][1],
                            _PBIGrade[roadpart[i].roaddegree][2],
                            _PBIGrade[roadpart[i].roaddegree][3]);
                    }
                    else
                    {
                        colcnt = colcnt + 2;
                    }

                    //PWI
                    if (prjinfo._IsIRIMTD && CMTDVal != null && CMTDVal.Length > 0)
                    {

                        wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                        wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);

                        if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))
                        {
                            wrval = 0;
                        }
                        vallist[rowcnt, colcnt++] = string.Format("=IF((100-{0}*POWER({1},{2}))>0,(100-{0}*POWER({1},{2})),0)", _PWIa[0], wrval, _PWIa[1]);
                        vallist[rowcnt, colcnt++] = string.Format("=IF(L{0}>={1},\"优\",IF(L{0}>={2},\"良\",IF(L{0}>={3},\"中\",IF(L{0}>={4},\"次\",\"差\"))))",
                            rowcnt + 3,
                            _PWIGrade[roadpart[i].roaddegree][0],
                            _PWIGrade[roadpart[i].roaddegree][1],
                            _PWIGrade[roadpart[i].roaddegree][2],
                            _PWIGrade[roadpart[i].roaddegree][3]);
                    }
                    else
                    {
                        colcnt = colcnt + 2;
                    }
                }
                else
                {
                    vallist[rowcnt, colcnt] =100;
                    colcnt=13;

                    vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                        rowcnt + 3,
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    //Rut
                    if (prjinfo._IsRut && SRutVal != null && SRutVal.Length > 0)
                    {
                        //rutval = Math.Max(LRutVal[i], RRutVal[i]);
                        rutval = SRutVal[i];
                        rutval = Math.Round(rutval, 5);

                        double rdival = 0;
                        if (rutval <= _RDIRD[0][1])
                        {
                            rdival = _RDIRD[0][0] - _RDIa[0] * rutval;
                        }
                        else if (rutval <= _RDIRD[1][1])
                        {
                            rdival = _RDIRD[1][0] - _RDIa[1] * (rutval - _RDIRD[0][1]);
                        }
                        else
                        {
                            rdival = 0;
                        }
                        // if(roadpart[i].roadtype==)
                        vallist[rowcnt, colcnt++] = rdival;
                        vallist[rowcnt, colcnt++] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                            rowcnt + 3,
                            _RDIGrade[roadpart[i].roaddegree][0],
                            _RDIGrade[roadpart[i].roaddegree][1],
                            _RDIGrade[roadpart[i].roaddegree][2],
                            _RDIGrade[roadpart[i].roaddegree][3]);
                    }
                    else
                    {
                        colcnt = colcnt + 2;
                    }
                    vallist[rowcnt, colcnt++] = 100;
                    vallist[rowcnt, colcnt++] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                          rowcnt + 3,
                          _PBIGrade[roadpart[i].roaddegree][0],
                          _PBIGrade[roadpart[i].roaddegree][1],
                          _PBIGrade[roadpart[i].roaddegree][2],
                          _PBIGrade[roadpart[i].roaddegree][3]);

                    vallist[rowcnt, colcnt++] =100;
                    vallist[rowcnt, colcnt++] = string.Format("=IF(L{0}>={1},\"优\",IF(L{0}>={2},\"良\",IF(L{0}>={3},\"中\",IF(L{0}>={4},\"次\",\"差\"))))",
                        rowcnt + 3,
                        _PWIGrade[roadpart[i].roaddegree][0],
                        _PWIGrade[roadpart[i].roaddegree][1],
                        _PWIGrade[roadpart[i].roaddegree][2],
                        _PWIGrade[roadpart[i].roaddegree][3]);

                }
                #endregion

                //pqi  =IF(P54="沥青",ROUND((0.35*D54+0.3*F54+0.15*IF(EXACT(H54,"-"),0,H54)+0.1*IF(EXACT(J54,"-"),0,J54)+0.1*IF(EXACT(L54,"-"),0,L54))/(0.35+0.3+0.15+0.1+0.1),5),ROUND((0.5*D54+0.3*F54+0*IF(EXACT(H54,"-"),0,H54)+0.1*IF(EXACT(J54,"-"),0,J54)+0.1*IF(EXACT(L54,"-"),0,L54))/(0.5+0.3+0+0.1+0.1),5))
                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(J{0},\"-\"),0,J{0})+{5}*IF(EXACT(L{0},\"-\"),0,L{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*D{0}+{7}*F{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(J{0},\"-\"),0,J{0}))/({6}+{7}+{8}+{9}),5))", rowcnt + 3,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][0][2],
                        _PQIW[roadpart[i].roaddegree][0][3],
                        _PQIW[roadpart[i].roaddegree][0][4],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1],
                        _PQIW[roadpart[i].roaddegree][1][2],
                        _PQIW[roadpart[i].roaddegree][1][3]);
                    }
                    else
                    {
                        vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(J{0},\"-\"),0,J{0})+{5}*IF(EXACT(L{0},\"-\"),0,L{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*D{0}+{7}*F{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(J{0},\"-\"),0,J{0})+{10}*IF(EXACT(L{0},\"-\"),0,L{0}))/({6}+{7}+{8}+{9}+{10}),5))", rowcnt + 3,
                          _PQIW[roadpart[i].roaddegree][0][0],
                          _PQIW[roadpart[i].roaddegree][0][1],
                          _PQIW[roadpart[i].roaddegree][0][2],
                          _PQIW[roadpart[i].roaddegree][0][3],
                          _PQIW[roadpart[i].roaddegree][0][4],
                          _PQIW[roadpart[i].roaddegree][1][0],
                          _PQIW[roadpart[i].roaddegree][1][1],
                          _PQIW[roadpart[i].roaddegree][1][2],
                          _PQIW[roadpart[i].roaddegree][1][3],
                          _PQIW[roadpart[i].roaddegree][1][4]);
                    }
                }
                else
                {
                    vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),5),ROUND(({3}*D{0}+{4}*F{0})/({3}+{4}),5))", rowcnt + 3,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1]);
                }

                vallist[rowcnt, colcnt++] = string.Format("=IF(N{0}>={1},\"优\",IF(N{0}>={2},\"良\",IF(N{0}>={3},\"中\",IF(N{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                vallist[rowcnt, colcnt++] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[rowcnt, colcnt++] = SpeedVal[i];
                }
                vallist[rowcnt, colcnt++] = MarkVal[i];
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:R{0}", rowcnt + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 18, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }

        //技术状况评定明细表
        public static void OutputPDMX(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\技术状况评定明细表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_技术状况评定明细表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePDMX2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal);
            if (_StreetDisRecord.Count > 0)
            {
                WriteStreetTCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray());
            }
            if (_StreetDisRecord_RoadBed.Count > 0)
            {
                WriteRoadBedSCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray());
            }

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 5, 1, 15, true);
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePDMX2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0, wrval = 0, pbival = 0, rdival = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 15];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                if (prjinfo._Direction > 0)
                    vallist[rowcnt, 0] = smile;
                else
                    vallist[rowcnt, 0] = emile;

                vallist[rowcnt, 1] = milelength;
                vallist[rowcnt, 2] = string.Format("=D{0}*{1}+E{0}*{2}+M{0}*{3}+N{0}*{4}", rowcnt + 5, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);
                vallist[rowcnt, 3] = 100;
                vallist[rowcnt, 10] = 100;
                vallist[rowcnt, 11] = 100;
                vallist[rowcnt, 12] = 100;
                vallist[rowcnt, 13] = 100;

                vallist[rowcnt, 14] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, 5] = Math.Round(pcival, 5);
                #region  判断是否存在IRIMTD文件夹 
        

                string iriDirPath = prjinfo._PrjPath + "\\IRIMTD";
                if (Directory.Exists(iriDirPath))
                {
                    //IRI
                    if (prjinfo._IsDIRIMTD)
                    {
                        if (_Setting.IRIExcelSide == 2)
                        {
                            if (_Setting.RQIJudgeType == 0)
                            {
                                irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                            }
                            else if (_Setting.RQIJudgeType == 1)
                            {
                                irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                            }
                        }
                        else if (_Setting.IRIExcelSide == 0)
                        {
                            irival = Math.Round(LIRIVal[i], 5);
                        }
                        else if (_Setting.IRIExcelSide == 1)
                        {
                            irival = Math.Round(RIRIVal[i], 5);
                        }
                    }
                    else
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                    vallist[rowcnt, 6] = Math.Round(trqival, 5);

                    //PWI
                    if (prjinfo._IsIRIMTD && CMTDVal != null && CMTDVal.Length > 0)
                    {

                        wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                        wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);

                        if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))
                        {
                            wrval = 0;
                        }
                        wrval = 100 - _PWIa[0] * Math.Pow(wrval, _PWIa[1]);
                        wrval = wrval > 0 ? wrval : 0;
                        vallist[rowcnt, 9] = Math.Round(wrval, 5);
                    }
                    else
                    {
                        vallist[rowcnt, 9] = "-";
                    }

                    //PBI
                    if (prjinfo._IsIRIMTD && PBVal != null && PBVal.Length > 0)
                    {
                        pbival = 100 - PBVal[i][1] * _PBIScore[1] - PBVal[i][2] * _PBIScore[2] - PBVal[i][3] * _PBIScore[3];
                        pbival = pbival > 0 ? pbival : 0;
                        vallist[rowcnt, 8] = Math.Round(pbival, 5);
                    }
                    else
                    {
                        vallist[rowcnt, 8] = "-";
                    }
                }
                else
                {
                    irival = 100;
                    //IRI
                    trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                    vallist[rowcnt, 6] =100;

                    //PWI
                       vallist[rowcnt, 9] = 100;

                   
                        vallist[rowcnt, 8] = 100;
                     
                }

                #endregion
                //Rut
                if (prjinfo._IsRut && SRutVal != null && SRutVal.Length > 0)
                {
                    //rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    rdival = 0;
                    if (rutval <= _RDIRD[0][1])
                    {
                        rdival = _RDIRD[0][0] - _RDIa[0] * rutval;
                    }
                    else if (rutval <= _RDIRD[1][1])
                    {
                        rdival = _RDIRD[1][0] - _RDIa[1] * (rutval - _RDIRD[0][1]);
                    }
                    else
                    {
                        rdival = 0;
                    }
                    // if(roadpart[i].roadtype==)
                    vallist[rowcnt, 7] = Math.Round(rdival, 5);
                }
                else
                {
                    vallist[rowcnt, 7] = "-";
                }

             

               

                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(I{0},\"-\"),0,I{0})+{5}*IF(EXACT(J{0},\"-\"),0,J{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*F{0}+{7}*G{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(I{0},\"-\"),0,I{0}))/({6}+{7}+{8}+{9}),5))", rowcnt + 5,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][0][2],
                        _PQIW[roadpart[i].roaddegree][0][3],
                        _PQIW[roadpart[i].roaddegree][0][4],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1],
                        _PQIW[roadpart[i].roaddegree][1][2],
                        _PQIW[roadpart[i].roaddegree][1][3]);
                    }
                    else
                    {
                        vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(I{0},\"-\"),0,I{0})+{5}*IF(EXACT(J{0},\"-\"),0,J{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*F{0}+{7}*G{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(I{0},\"-\"),0,I{0})+{10}*IF(EXACT(J{0},\"-\"),0,J{0}))/({6}+{7}+{8}+{9}+{10}),5))", rowcnt + 5,
                          _PQIW[roadpart[i].roaddegree][0][0],
                          _PQIW[roadpart[i].roaddegree][0][1],
                          _PQIW[roadpart[i].roaddegree][0][2],
                          _PQIW[roadpart[i].roaddegree][0][3],
                          _PQIW[roadpart[i].roaddegree][0][4],
                          _PQIW[roadpart[i].roaddegree][1][0],
                          _PQIW[roadpart[i].roaddegree][1][1],
                          _PQIW[roadpart[i].roaddegree][1][2],
                          _PQIW[roadpart[i].roaddegree][1][3],
                          _PQIW[roadpart[i].roaddegree][1][4]);
                    }
                }
                else
                {
                    vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0})/({1}+{2}),5),ROUND(({3}*F{0}+{4}*G{0})/({3}+{4}),5))", rowcnt + 5,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1]);
                }

                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:O{0}", rowcnt + 4));
            destrange.Value2 = vallist;
            destrange = worksheet.get_Range(String.Format("A5:O{0}", rowcnt + 5));
            GlobalExcel.SetBorderLine(destrange, 53);

            worksheet.Cells[2, 2] = prjinfo._Province + prjinfo._City + prjinfo._District;
            worksheet.Cells[2, 5] = prjinfo._RoadCode + prjinfo._RoadName;
            worksheet.Cells[2, 7] = prjinfo._RoadGrade;
            worksheet.Cells[2, 10] = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            worksheet.Cells[2, 13] = prjinfo._Direction > 0 ? "上行" : "下行";
            worksheet.Cells[2, 14] = prjinfo._DataDate;
            worksheet.Cells[rowcnt + 5, 1] = "合计";
            worksheet.Cells[rowcnt + 5, 2] = String.Format("=SUM(B5:B{0})", rowcnt + 4);
            for (int i = 0; i < 12; ++i)
            {
                if (_Setting.JSAverageType)
                {
                    worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=AVERAGE({1}5:{1}{0})", rowcnt + 4, (char)('C' + i));
                }
                else
                {
                    worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=SUMPRODUCT(B5:B{0},{1}5:{1}{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));
                }
                //

              
            }
        }

        //输出空的技术状况评定明细表--用于提供给别人填充设备无法检测的技术指标数值
        public static void OutputPDMX_Empty(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\技术状况评定明细表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_技术状况评定明细表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePDMX2Xls_Empty(_Worksheet, prjinfo, prjdir, _RoadPart);

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePDMX2Xls_Empty(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart)
        {
            int len = roadpart.Count - 1;

            int rowcnt = 0;
            object[,] vallist = new object[len, 15];
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                if (prjinfo._Direction > 0)
                    vallist[rowcnt, 0] = smile;
                else
                    vallist[rowcnt, 0] = emile;

                vallist[rowcnt, 1] = milelength;
                for (int j = 3; j < 14; ++j)
                {
                    vallist[rowcnt, j] = 100;
                }
                vallist[rowcnt, 14] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(I{0},\"-\"),0,I{0})+{5}*IF(EXACT(J{0},\"-\"),0,J{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*F{0}+{7}*G{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(I{0},\"-\"),0,I{0}))/({6}+{7}+{8}+{9}),5))", rowcnt + 5,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][0][2],
                        _PQIW[roadpart[i].roaddegree][0][3],
                        _PQIW[roadpart[i].roaddegree][0][4],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1],
                        _PQIW[roadpart[i].roaddegree][1][2],
                        _PQIW[roadpart[i].roaddegree][1][3]);
                    }
                    else
                    {
                        vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(I{0},\"-\"),0,I{0})+{5}*IF(EXACT(J{0},\"-\"),0,J{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*F{0}+{7}*G{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(I{0},\"-\"),0,I{0})+{10}*IF(EXACT(J{0},\"-\"),0,J{0}))/({6}+{7}+{8}+{9}+{10}),5))", rowcnt + 5,
                          _PQIW[roadpart[i].roaddegree][0][0],
                          _PQIW[roadpart[i].roaddegree][0][1],
                          _PQIW[roadpart[i].roaddegree][0][2],
                          _PQIW[roadpart[i].roaddegree][0][3],
                          _PQIW[roadpart[i].roaddegree][0][4],
                          _PQIW[roadpart[i].roaddegree][1][0],
                          _PQIW[roadpart[i].roaddegree][1][1],
                          _PQIW[roadpart[i].roaddegree][1][2],
                          _PQIW[roadpart[i].roaddegree][1][3],
                          _PQIW[roadpart[i].roaddegree][1][4]);
                    }
                }
                else
                {
                    vallist[rowcnt, 4] = string.Format("=IF(O{0}=\"沥青\",ROUND(({1}*F{0}+{2}*G{0})/({1}+{2}),5),ROUND(({3}*F{0}+{4}*G{0})/({3}+{4}),5))", rowcnt + 5,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1]);
                }

                vallist[rowcnt, 2] = string.Format("=D{0}*{1}+E{0}*{2}+M{0}*{3}+N{0}*{4}", rowcnt + 5, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);

                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:O{0}", rowcnt + 5));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            worksheet.Cells[2, 2] = prjinfo._Province + prjinfo._City + prjinfo._District;
            worksheet.Cells[2, 5] = prjinfo._RoadCode + prjinfo._RoadName;
            worksheet.Cells[2, 7] = prjinfo._RoadGrade;
            worksheet.Cells[2, 10] = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            worksheet.Cells[2, 13] = prjinfo._Direction > 0 ? "上行" : "下行";
            worksheet.Cells[2, 14] = prjinfo._DataDate;

            worksheet.Cells[rowcnt + 5, 1] = "合计";
            worksheet.Cells[rowcnt + 5, 2] = String.Format("=SUM(B5:B{0})", rowcnt + 4);
            for (int i = 0; i < 12; ++i)
            {
                worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=SUMPRODUCT(B5:B{0},{1}5:{1}{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));
            }

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 5, 1, 15, true);
            }
        }

        private static void WritePrj2CPMSXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[3, 2] = prjinfo._RoadCode;
            if (prjinfo._Direction > 0)
            {
                _Worksheet.Cells[3, 4] = "上行" + prjinfo._RoadNum;
            }
            else
            {
                _Worksheet.Cells[3, 4] = "下行" + prjinfo._RoadNum;
            }
            _Worksheet.Cells[3, 8] = prjinfo._DataDate;
            _Worksheet.Cells[4, 8] = prjinfo._StartMile.ToString("K0+000");
            _Worksheet.Cells[4, 13] = prjinfo._EndMile.ToString("K0+000");
            _Worksheet.Cells[5, 13] = _RoadConfig.DetectWidth;
        }

        public static void OutputPBI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面跳车评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PBI_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet = _Workbook.Sheets["Δh"] as MSExcel.Worksheet;
            WritePB2Xls(_Worksheet, prjinfo, prjdir, _RoadPart10, _LDeltaHVal, _RDeltaHVal, _SpeedVal10, _MarkVal10, 4, 53);
            WritePBStatistics(_Worksheet);

            _Worksheet = _Workbook.Sheets["PBI"] as MSExcel.Worksheet;
            WritePBI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _PBIVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'H', "跳车", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePBI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, int[][] PBIVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = PBIVal[i][1];
                vallist[i, 4] = PBIVal[i][2];
                vallist[i, 5] = PBIVal[i][3];
                vallist[i, 6] = string.Format("=IF((100-D{0}*{1}-E{0}*{2}-F{0}*{3})>0,(100- D{0}*{1}-E{0}*{2}-F{0}*{3}),0)",
                    i + DataStartXlsxRow, _PBIScore[1], _PBIScore[2], _PBIScore[3]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + DataStartXlsxRow,
                    _PBIGrade[roadpart[i].roaddegree][0],
                    _PBIGrade[roadpart[i].roaddegree][1],
                    _PBIGrade[roadpart[i].roaddegree][2],
                    _PBIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        private static void WritePB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, double[] LDeltaHVal, double[] RDeltaHVal, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            int i = 0, len = roadpart.Count - 1;
            object[,] vallist = new object[len, 10];
            if (LDeltaHVal == null && RDeltaHVal == null)
            {
                MessageBox.Show("缺少路面左右两侧纵断面高程数据,请检查数据完整性");
                return;
            }
            for (i = 0; i < len; ++i)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                if (i < LDeltaHVal.Length)
                {
                    vallist[i, 3] = LDeltaHVal[i];
                }
                if (prjinfo._IsDIRIMTD)
                {
                    if (i < RDeltaHVal.Length)
                    {
                        vallist[i, 4] = RDeltaHVal[i];
                    }
                }
                vallist[i, 5] = string.Format("=MAX(D{0},E{0})", i + DataStartXlsxRow);
                vallist[i, 6] = string.Format("=IF(F{0}<{1},\"无跳车\",IF(F{0}<{2},\"轻度跳车\",IF(F{0}<{3},\"中度跳车\",\"重度跳车\")))",
                    i + DataStartXlsxRow, _PBIThresh[0], _PBIThresh[1], _PBIThresh[2]);
                vallist[i, 7] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 8] = SpeedVal[i];
                }
                vallist[i, 9] = MarkVal[i];
            }

            MSExcel.Range destrange;
            destrange = _Worksheet.get_Range(String.Format("A{0}:J{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;

            destrange = _Worksheet.get_Range(String.Format("A1:J{0}", len + DataStartXlsxRow - 1));
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 10, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }
        private static void WritePBStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "无跳车", "轻度跳车", "中度跳车", "重度跳车" };
            MSExcel.Range destrange = _Worksheet.get_Range("V3:Y5");
            object[,] val = new object[3, 4];
            for (int i = 0; i < 4; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(G:G,\"{0}\",A:A)-SUMIF(G:G,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 21 + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 9] = "=CONCATENATE(\"路面跳车评价等级“无跳车”率占路段总数\",ROUND(V4,4)*100,\"%，“轻度跳车”率占路段总数\",ROUND(W4,4)*100,\"%，“中度跳车”率占路段总数\",ROUND(X4,4)*100,\"%，“重度跳车”率占路段总数\",ROUND(Y4,4)*100,\"%。\")";
        }

        public static void OutputGeoAlig(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\路面几何状况检测数据统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_Geo_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteGeoAlig2Xls(_Worksheet, prjinfo, _RoadPart, _Curvature, _CrossSlope, _HeightSlope, _SpeedVal, _MarkVal, 4, 53);

            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteGeoAlig2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] Curvature, double[] CrossSlope, double[] HeightSlope, double[] SpeedVal, string[] MarkVal,
            int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsRut)
            {
                return;
            }
            else if (prjinfo._GeoAlig != 1)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 9];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = Curvature[i];
                vallist[i, 4] = HeightSlope[i];
                vallist[i, 5] = CrossSlope[i];
                vallist[i, 6] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (SpeedVal != null)
                {
                    vallist[i, 7] = SpeedVal[i];
                }
                vallist[i, 8] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:I{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 9, true);
                GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        #region 沿线设施导出病害
        private static List<StreetDisRecord> _StreetDisRecord = new List<StreetDisRecord>();
        private static List<StreetDisRecord> _StreetDisRecord_RoadBed = new List<StreetDisRecord>();
        public static void InitStreetData(ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            GlobalExcel.LoadStreetAllRecInfo(prjinfo, prjdir, ref _StreetDisRecord);
            GlobalExcel.LoadStreetAllRecInfo_RoadBed(prjinfo, prjdir, ref _StreetDisRecord_RoadBed);
        }
        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputStreetDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\沿线设施损坏汇总表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_沿线设施损坏汇总表_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["沿线设施损坏汇总表 "] as MSExcel.Worksheet;
            WriteStreetDisDC2Xls(_Worksheet_hz, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(), xlslen);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteStreetDisDC2Xls(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.streetdislist.Count;
            object[,] disval = new object[len, temp + 4];
            worksheet_hz.Cells[2, 2] = prjdir.Name;
            double tclval = 0;
            double ttclval = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int unitlen = Math.Abs(smile - emile);

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.streetdisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.streetdislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                disval[i, 0] = smile;
                disval[i, 1] = emile;

                tclval = 0;
                ttclval = 0;
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.streetdislist[k].sumval;
                    if (k > 0)
                    {
                        if (DiseaseTypes.streetdislist[k - 1].distype != DiseaseTypes.streetdislist[k].distype)
                        {
                            ttclval = ttclval * 1000 / unitlen;
                            ttclval = ttclval > 100 ? 100 : ttclval;
                            tclval += DiseaseTypes.streetdislist[k - 1].weight * (100 - ttclval);
                            ttclval = 0;
                        }
                    }
                    ttclval = ttclval + DiseaseTypes.streetdislist[k].unitscore * DiseaseTypes.streetdislist[k].sumval;
                }
                ttclval = ttclval * 1000 / unitlen;
                ttclval = ttclval > 100 ? 100 : ttclval;
                tclval += DiseaseTypes.streetdislist[temp - 1].weight * (100 - ttclval);

                disval[i, temp + 2] = tclval;
                disval[i, temp + 3] = string.Format("=IF({1}{0}>=90,\"优\",IF({1}{0}>=80,\"良\",IF({1}{0}>=70,\"中\",IF({1}{0}>=60,\"次\",\"差\"))))",
                    i + 6, GlobalExcel.GetCol((char)(temp + 2 + 'A')));

                smile = emile;
                DiseaseTypes.Clear();
            }
            destrange = worksheet_hz.get_Range(string.Format("A6:{1}{0}", len + 5, GlobalExcel.GetCol((char)('A' + temp + 3))));
            destrange.Value2 = disval;
            GlobalExcel.SetBorderLine(destrange, 53);

            disval = new object[temp, 1];
            for (int i = 0; i < temp; ++i)
            {
                disval[i, 0] = string.Format("=SUM({0}6:{0}{1})", GlobalExcel.GetCol((char)('C' + i)), len + 5);
            }
            destrange = worksheet_hz.get_Range(string.Format("{0}3:{0}{1}", GlobalExcel.GetCol((char)('A' + temp + 5)), temp + 2));
            destrange.Value2 = disval;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet_hz, 6, 1, temp + 4, true);
                GlobalExcel.Reflection(worksheet_hz, 6, 1, 2, false);
            }
        }

        private static void WriteStreetTCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.streetdislist.Count;
            object[,] disval = new object[len, 1];

            double tclval = 0;
            double ttclval = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int unitlen = Math.Abs(smile - emile);

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.streetdisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.streetdislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                tclval = 0;
                ttclval = 0;
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    if (k > 0)
                    {
                        if (DiseaseTypes.streetdislist[k - 1].distype != DiseaseTypes.streetdislist[k].distype)
                        {
                            ttclval = ttclval * 1000 / unitlen;
                            ttclval = ttclval > 100 ? 100 : ttclval;
                            tclval += DiseaseTypes.streetdislist[k - 1].weight * (100 - ttclval);
                            ttclval = 0;
                        }
                    }
                    ttclval = ttclval + DiseaseTypes.streetdislist[k].unitscore * DiseaseTypes.streetdislist[k].sumval;
                }
                ttclval = ttclval * 1000 / unitlen;
                ttclval = ttclval > 100 ? 100 : ttclval;
                tclval += DiseaseTypes.streetdislist[temp - 1].weight * (100 - ttclval);

                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            }
            destrange = worksheet.get_Range(string.Format("N5:N{0}", len + 4));
            destrange.Value2 = disval;
        }

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSStreetDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\\CPMS_沿线设施损坏.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_CPMS沿线设施损坏.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_dc = _Workbook.Sheets["沿线设施损坏调查表"] as MSExcel.Worksheet;
            WriteCPMSStreetDisDC2Xls(_Worksheet_dc, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(), _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteCPMSStreetDisDC2Xls(MSExcel.Worksheet worksheet_dc, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, StreetDisRecord[] arrdis, string[] MarkVal)
        {
            MSExcel.Range srcrange, destrange;
            object[,] disval;

            const int tablerow = 19;
            int tcnt = 0;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int smile = roadpart[0].mile;
            int emile = roadpart[len].mile;
            int csmile = smile, cemile = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                smile = roadpart[i].mile;
                emile = roadpart[i + 1].mile;

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.streetdisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.streetdislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.streetdislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                //病害汇总表
                disval = new object[DiseaseTypes.streetdislist.Count, 1];
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    disval[k, 0] = DiseaseTypes.streetdislist[k].sumval;
                }
                destrange = worksheet_dc.get_Range(string.Format("{0}{1}:{0}{2}",
                    GlobalExcel.GetCol((char)('F' + (Math.Min(smile, emile) % 1000) * 10 / 1000)),
                    tablerow * tcnt + 7,
                    tablerow * tcnt + 6 + DiseaseTypes.streetdislist.Count));
                destrange.Value2 = disval;

                if (emile % 1000 == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")) || roadpart[i].roadtype != roadpart[i + 1].roadtype)
                {
                    cemile = emile;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 1] = "路线名称：" + prjinfo._RoadName;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                    worksheet_dc.Cells[tablerow * tcnt + 3, 8] = prjinfo._DataDate;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 13] = prjinfo._DataPerson;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 8] = csmile;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 13] = cemile;
                    worksheet_dc.Cells[tablerow * tcnt + 5, 8] = Math.Abs(csmile - cemile);
                    worksheet_dc.Cells[tablerow * tcnt + 5, 13] = _RoadConfig.DetectWidth;
                    if (cemile != prjinfo._EndMile)
                    {
                        srcrange = worksheet_dc.get_Range(String.Format("A{0}:T{1}", tablerow * tcnt + 1, tablerow * (tcnt + 1) - 1));
                        ++tcnt;
                        destrange = worksheet_dc.get_Range(String.Format("A{0}", tablerow * tcnt + 1));
                        srcrange.Copy(destrange);
                        destrange = worksheet_dc.get_Range(String.Format("F{0}:O{1}", tablerow * tcnt + 7, tablerow * tcnt + 6 + DiseaseTypes.streetdislist.Count));
                        destrange.ClearContents();
                    }
                    csmile = cemile;
                }
                smile = emile;
                DiseaseTypes.Clear();
            }
            if (prjinfo._EndMile % 1000 != 0 || (MarkVal[len] != null && MarkVal[len].Contains("路面单元")))
            {
                worksheet_dc.Cells[tablerow * tcnt + 3, 1] = "路线名称：" + prjinfo._RoadName;
                worksheet_dc.Cells[tablerow * tcnt + 3, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                worksheet_dc.Cells[tablerow * tcnt + 3, 8] = prjinfo._DataDate;
                worksheet_dc.Cells[tablerow * tcnt + 3, 13] = prjinfo._DataPerson;
                worksheet_dc.Cells[tablerow * tcnt + 4, 8] = csmile;
                worksheet_dc.Cells[tablerow * tcnt + 4, 13] = prjinfo._EndMile;
                worksheet_dc.Cells[tablerow * tcnt + 5, 8] = Math.Abs(csmile - prjinfo._EndMile);
                worksheet_dc.Cells[tablerow * tcnt + 5, 13] = _RoadConfig.DetectWidth;
            }
        }
        #endregion

        #region 路基损坏报表
        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\路基损坏汇总表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_路基损坏汇总表_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["路基损坏汇总表"] as MSExcel.Worksheet;
            WriteRoadBedDisDC2Xls(_Worksheet_hz, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), xlslen);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteRoadBedSCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.roadbeddislist.Count;
            object[,] disval = new object[len, 1];

            double tclval = 0;
            double ttclval = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int unitlen = Math.Abs(smile - emile);

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.roadbeddisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.roadbeddislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                tclval = 0;
                ttclval = 0;
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    if (k > 0)
                    {
                        if (DiseaseTypes.roadbeddislist[k - 1].distype != DiseaseTypes.roadbeddislist[k].distype)
                        {
                            ttclval = ttclval * 1000 / unitlen;
                            ttclval = ttclval > 100 ? 100 : ttclval;
                            tclval += DiseaseTypes.roadbeddislist[k - 1].weight * (100 - ttclval);
                            ttclval = 0;
                        }
                    }
                    ttclval = ttclval + DiseaseTypes.roadbeddislist[k].unitscore * DiseaseTypes.roadbeddislist[k].sumval;
                }
                int ttypeidx = DiseaseTypes.roadbeddisIdx["路基构造物损坏.重"];
                if (DiseaseTypes.roadbeddislist[ttypeidx].sumval > 0)
                {
                    tclval = 0;
                }
                else
                {
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.roadbeddislist[temp - 1].weight * (100 - ttclval);
                }

                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            }
            destrange = worksheet.get_Range(string.Format("D5:D{0}", len + 4));
            destrange.Value2 = disval;
        }


        /// <summary>
        /// 图片地址  病害信息
        /// </summary>
        private static Dictionary<(string, int), List<MyStreetMile2DisInfo>> curProjectStreetDic = new Dictionary<(string, int), List<MyStreetMile2DisInfo>>();
        public static void OutputStreetAllDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            curProjectStreetDic.Clear();
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\景观病害明细表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_景观病害明细表.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_dc = _Workbook.Sheets[1] as MSExcel.Worksheet;
            ExcelGPS[] GPSInfos = null;
            string[] gpsinfostrs;
            if (File.Exists(prjinfo._PrjPath + "\\GPS2Mile.txt"))
            {
                gpsinfostrs = File.ReadAllLines(prjinfo._PrjPath + "\\GPS2Mile.txt");
                GPSInfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    GPSInfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }
            else
            {
                MessageBox.Show($"工程{prjinfo._PrjPath}获取经纬度信息失败！");
                return;
            }
            ExcelGPS[] tempinfos = null;
            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
            {
                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
                tempinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }
            //自定义病害
            string streetPath = prjinfo._PrjPath + "\\StreetImg";
            DirectoryInfo dir = new DirectoryInfo(streetPath);
            FileInfo[] files = dir.GetFiles("*_UserSign.txt", SearchOption.AllDirectories);
            List<UserSignMsg> dataInfos = new List<UserSignMsg>();
            foreach (var item in files)
            {
                dataInfos.AddRange(File.ReadAllLines(item.FullName).Select(line => new UserSignMsg(line)));
            }
            List<StreetDisRecord> allStreetDis = new List<StreetDisRecord>();
            allStreetDis.AddRange(_StreetDisRecord);
            allStreetDis.AddRange(_StreetDisRecord_RoadBed);
            int len = dataInfos.Count + allStreetDis.Count;

            List<MyStreetDisRecord> disRecords = new List<MyStreetDisRecord>();
            int gi = 0;
            ExcelGPS tempgpsinfo;
            foreach (var item in dataInfos)
            {
                MyStreetDisRecord dis = new MyStreetDisRecord();
                dis.RoadCode = prjinfo._RoadCode;
                dis.RoadNum = prjinfo._RoadNum;
                dis.Direction = prjinfo._Direction > 0 ? "上行" : "下行";
                dis.StartMile = int.Parse(item.Mile.Replace("K", "").Replace("+", ""));
                int roadLen = 0;
                if (item.isHasRect())
                {
                    roadLen = (int)(prjinfo._StreetImgDis_Left * item.SignRect.Height / 2048.0);
                }
                dis.EndMile = dis.StartMile + roadLen;
                dis.DisName = item.DisName;
                dis.DisGrad = "";
                dis.Area = item.DisCnt;
                dis.Mark = item.Info;
                for (; gi < tempinfos.Length; ++gi)
                {
                    if (prjinfo._Direction > 0)
                    {
                        if (tempinfos[gi]._mile >= dis.StartMile)
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (tempinfos[gi]._mile <= dis.StartMile)
                        {
                            break;
                        }
                    }
                }
                if (gi < tempinfos.Length)
                {
                    tempgpsinfo = tempinfos[gi];
                }
                else
                {
                    tempgpsinfo = tempinfos[tempinfos.Length - 1];
                }
                dis.Latitude = double.Parse(tempgpsinfo._latitude);
                dis.Longitude = double.Parse(tempgpsinfo._longitude);
                dis.Height = double.Parse(tempgpsinfo._elevation);
                disRecords.Add(dis);
            }
            gi = 0;
            foreach (var item in allStreetDis)
            {
                MyStreetDisRecord dis = new MyStreetDisRecord();
                dis.RoadCode = prjinfo._RoadCode;
                dis.RoadNum = prjinfo._RoadNum;
                dis.Direction = prjinfo._Direction > 0 ? "上行" : "下行";
                dis.StartMile = int.Parse(item._mile.Replace("K", "").Replace("+", ""));
                int roadLen = 0;
                if (item.isHasRect())
                {
                    roadLen = (int)(prjinfo._StreetImgDis_Left * item.SignRect.Height / 2048.0);
                }
                dis.EndMile = dis.StartMile + roadLen;
                dis.DisName = item._disname;
                
                dis.DisGrad = "";
                string[] nameSplit = dis.DisName.Split('.');
                if (nameSplit.Length > 1)
                {
                    dis.DisName = nameSplit[0];
                    dis.DisGrad = nameSplit[1];
                }
                dis.Area = item._ndislen == 0 ? item._ndisnum : item._ndislen;
                dis.Mark = "";
                for (; gi < tempinfos.Length; ++gi)
                {
                    if (prjinfo._Direction > 0)
                    {
                        if (tempinfos[gi]._mile >= dis.StartMile)
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (tempinfos[gi]._mile <= dis.StartMile)
                        {
                            break;
                        }
                    }
                }
                if (gi < tempinfos.Length)
                {
                    tempgpsinfo = tempinfos[gi];
                }
                else
                {
                    tempgpsinfo = tempinfos[tempinfos.Length - 1];
                }
                dis.Latitude = double.Parse(tempgpsinfo._latitude);
                dis.Longitude = double.Parse(tempgpsinfo._longitude);
                dis.Height = double.Parse(tempgpsinfo._elevation);
                disRecords.Add(dis);

            }


            disRecords.Sort((a, b) => a.StartMile.CompareTo(b.StartMile));
            // 1. 定义列数（根据你的 MyStreetDisRecord 属性数量，这里假设有 12 列）
            int colCount = 12;

            // 2. 创建二维数组 (行数 = 记录数, 列数 = 字段数)
            // 注意：C#数组下标从0开始，但Excel单元格下标从1开始，这里我们只负责准备数据
            object[,] dataArray = new object[disRecords.Count, colCount];

            for (int i = 0; i < disRecords.Count; i++)
            {
                MyStreetDisRecord item = disRecords[i];

                // 按 Excel 模板的列顺序填充数据
                // 假设 Excel 列顺序为：路线代码, 路线编号, 上下行, 开始桩号, 结束桩号, 病害名称, 等级, 面积, 备注, 经度, 纬度, 高程
                // 请根据你实际的 Excel 模板调整索引顺序 [i, 0] 对应第1列
                dataArray[i, 0] = item.RoadCode;    // 第1列
                dataArray[i, 1] = item.RoadNum;     // 第2列
                dataArray[i, 2] = item.Direction;   // 第3列
                dataArray[i, 3] = item.StartMile;   // 第4列
                dataArray[i, 4] = item.EndMile;     // 第5列
                dataArray[i, 5] = item.DisName;     // 第6列
                dataArray[i, 6] = item.DisGrad;     // 第7列
                dataArray[i, 7] = item.Area;        // 第8列
                dataArray[i, 8] = item.Mark;        // 第9列
                dataArray[i, 9] = item.Longitude;   // 第10列 (注意经纬度顺序)
                dataArray[i, 10] = item.Latitude;   // 第11列
                dataArray[i, 11] = item.Height;     // 第12列
            }

            // 3. 一次性写入 Excel
            if (disRecords.Count > 0)
            {
                // 假设模板只有表头，数据从第 2 行开始写入 (如果表头有多行，请修改这里的 2)
                int startRow = 2;

                // 获取开始单元格 (A2)
                MSExcel.Range startCell = _Worksheet_dc.Cells[startRow, 1];
                // 获取结束单元格 (最后一行的最后一列)
                MSExcel.Range endCell = _Worksheet_dc.Cells[startRow + disRecords.Count - 1, colCount];

                // 获取写入区域
                MSExcel.Range writeRange = _Worksheet_dc.Range[startCell, endCell];

                // 核心：一次性赋值，速度极快
                writeRange.Value2 = dataArray;

                // 4. (可选) 设置边框 - 如果模板自带边框格式，这段可以省略
                // writeRange.Borders.LineStyle = MSExcel.XlLineStyle.xlContinuous;
            }
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();



            foreach (var item in dataInfos)
            {
                //获得图片地址
                if (item.isHasRect())
                {
                    string streetPciturePath = getStreetPcitureFilePath(prjdir.FullName, int.Parse(item.Mile.Replace("K", "").Replace("+", "")), item.Side);
                    MyStreetMile2DisInfo info = new MyStreetMile2DisInfo();
                    info.Rect = item.SignRect;
                    info.Mile = int.Parse(item.Mile.Replace("K", "").Replace("+", ""));
                    info.DisInfo = item.DisName + " " + item.Info;
                    if (curProjectStreetDic.ContainsKey((streetPciturePath, item.Side)))
                    {
                        curProjectStreetDic[(streetPciturePath, item.Side)].Add(info);
                    }
                    else
                    {
                        curProjectStreetDic[(streetPciturePath, item.Side)] = new List<MyStreetMile2DisInfo> { info };
                    }
                }
            }

            foreach (var item in allStreetDis)
            {
                //获得图片地址
                if (item.isHasRect())
                {
                    string streetPciturePath = getStreetPcitureFilePath(prjdir.FullName, int.Parse(item._mile.Replace("K", "").Replace("+", "")), item.Side);
                    MyStreetMile2DisInfo info = new MyStreetMile2DisInfo();
                    info.Rect = item.SignRect;
                    info.DisInfo = item._disname;
                    info.Mile = int.Parse(item._mile.Replace("K", "").Replace("+", ""));
                    if (curProjectStreetDic.ContainsKey((streetPciturePath, item.Side)))
                    {
                        curProjectStreetDic[(streetPciturePath, item.Side)].Add(info);
                    }
                    else
                    {
                        curProjectStreetDic[(streetPciturePath, item.Side)] = new List<MyStreetMile2DisInfo> { info };
                    }
                }
            }
            //用于记录已经清空过的文件夹，防止循环中重复删除
            HashSet<string> cleanedDirectories = new HashSet<string>();
            for (int i = 0; i < curProjectStreetDic.Count; i++)
            {
                var pathInfos = curProjectStreetDic.Keys.ElementAt(i);
                string picturePath = pathInfos.Item1;
                int side = pathInfos.Item2;
                List<MyStreetMile2DisInfo> disInfos = curProjectStreetDic.Values.ElementAt(i);

                // 1. 检查源文件是否存在
                if (!File.Exists(picturePath)) continue;

                string outName = prjdir.Name;
                if (!string.IsNullOrEmpty(prjinfo._RoadCode.Trim()))
                {
                    outName = prjinfo._RoadCode;
                }

                // 2. 构建输出路径 (注意：这里去掉了原代码中多余的空格)
                // 建议使用 Path.Combine 兼容性更好
                string outPath = System.IO.Path.Combine(path, outName, side.ToString());

                // 【关键修复】文件夹清理逻辑
                // 如果这个输出文件夹还没被清理过，并且存在，则清理一次。
                // 之后同一个文件夹下的其他图片进来时，就不会再删除了。
                if (!cleanedDirectories.Contains(outPath))
                {
                    if (Directory.Exists(outPath))
                    {
                        // 如果你想每次运行都清空旧数据，保留这行
                        // Directory.Delete(outPath, true); 
                    }
                    if (!Directory.Exists(outPath))
                    {
                        Directory.CreateDirectory(outPath);
                    }
                    cleanedDirectories.Add(outPath);
                }

                // 3. 开始绘图处理
                try
                {
                    // 使用流读取，避免文件锁死
                    using (FileStream fs = new FileStream(picturePath, FileMode.Open, FileAccess.Read))
                    using (Image originalImg = Image.FromStream(fs))
                    using (Bitmap bmp = new Bitmap(originalImg)) // 创建副本用于修改
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        // 设置绘图质量，让框和字更清晰
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                        // 定义画笔（红色边框，线宽5）
                        using (Pen pen = new Pen(Color.Red, 5))
                        // 定义字体（宋体/Arial，大小30，粗体）根据图片分辨率调整大小
                        using (System.Drawing.Font font = new System.Drawing.Font("Arial", 10, FontStyle.Bold))
                        // 定义文字画刷（黄色文字，红色描边通常比较显眼，这里简单用黄色）
                        using (SolidBrush textBrush = new SolidBrush(Color.White))
                        {
                            for (int t = 0; t < disInfos.Count; t++)
                            {
                                MyStreetMile2DisInfo disInfo = disInfos[t];
                                string disName = prjinfo._RoadName + " " + disInfo.Mile / 1000.0 + "km\n" + disInfo.DisInfo;
                                System.Drawing.Rectangle rect = disInfo.Rect;

                                // --- 将矩形框绘制到图片上 ---
                                g.DrawRectangle(pen, rect);

                                // --- 将文字写到图片上 ---
                                // 计算文字位置，通常放在矩形框上方。如果框在最顶端，则放在框内部。
                                float textY = rect.Y - 40;
                                if (textY < 0) textY = rect.Y + 5; // 防止文字跑出图片上边界

                                // 绘制文字背景条（可选，为了让文字看不清楚时能看清背景）
                                // SizeF textSize = g.MeasureString(disName, font);
                                // g.FillRectangle(Brushes.Black, rect.X, textY, textSize.Width, textSize.Height);

                                g.DrawString(disName, font, textBrush, rect.X, textY);
                            }
                        }

                        // --- 保存图片到outPath ---
                        string fileName = System.IO.Path.GetFileName(picturePath);
                        string saveFilePath = System.IO.Path.Combine(outPath, fileName);

                        // 保存为 JPG 格式
                        bmp.Save(saveFilePath, ImageFormat.Jpeg);
                    }
                }
                catch (Exception ex)
                {
                    // 记录日志或跳过损坏的图片
                    Console.WriteLine($"处理图片失败: {picturePath}, 错误: {ex.Message}");
                }
            }
        }

        private static string getStreetPcitureFilePath(string prjdir, int mile, int side)
        {
            string path = prjdir + $"\\StreetImg\\Camera{side}\\Street2Mile.txt";
            List<string> pictureFileTxt = File.ReadAllLines(path).ToList();

            List<MyStreetMile2Path> mile2Paths = new List<MyStreetMile2Path>();
            for (int i = 0; i < pictureFileTxt.Count; i++)
            {
                string line = pictureFileTxt[i];
                try
                {
                    string[] info = line.Split(' ');
                    MyStreetMile2Path mile2Path = new MyStreetMile2Path();
                    mile2Path.Mile = int.Parse(info.First());
                    mile2Path.FilePath = prjdir + $"\\StreetImg\\Camera{side}" + info.Last();
                    mile2Paths.Add(mile2Path);
                }
                catch (Exception)
                {

                    continue;
                }
            }

            string filePath = mile2Paths.Where(t => t.Mile == mile).FirstOrDefault().FilePath;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                throw new Exception($"未在{path}下找到所在桩号的景观图片，导致无法输出图片,请检查!");
            }
            return filePath;
        }

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\CPMS_路基损坏.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_CPMS路基损坏.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_dc = _Workbook.Sheets["路基损坏调查表"] as MSExcel.Worksheet;
            WriteCPMSRoadBedDisDC2Xls(_Worksheet_dc, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteCPMSRoadBedDisDC2Xls(MSExcel.Worksheet worksheet_dc, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, StreetDisRecord[] arrdis, string[] MarkVal)
        {
            MSExcel.Range srcrange, destrange;
            object[,] disval;

            const int tablerow = 32;
            int tcnt = 0;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int smile = roadpart[0].mile;
            int emile = roadpart[len].mile;
            int csmile = smile, cemile = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                smile = roadpart[i].mile;
                emile = roadpart[i + 1].mile;

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.roadbeddisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.roadbeddislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                //病害汇总表
                disval = new object[DiseaseTypes.roadbeddislist.Count, 1];
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    disval[k, 0] = DiseaseTypes.roadbeddislist[k].sumval;
                }
                destrange = worksheet_dc.get_Range(string.Format("{0}{1}:{0}{2}",
                    GlobalExcel.GetCol((char)('F' + (Math.Min(smile, emile) % 1000) * 10 / 1000)),
                    tablerow * tcnt + 7,
                    tablerow * tcnt + 6 + DiseaseTypes.roadbeddislist.Count));
                destrange.Value2 = disval;

                if (emile % 1000 == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")) || roadpart[i].roadtype != roadpart[i + 1].roadtype)
                {
                    cemile = emile;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 1] = "路线名称：" + prjinfo._RoadName;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                    worksheet_dc.Cells[tablerow * tcnt + 3, 8] = prjinfo._DataDate;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 13] = prjinfo._DataPerson;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 8] = csmile;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 13] = cemile;
                    worksheet_dc.Cells[tablerow * tcnt + 5, 8] = Math.Abs(csmile - cemile);
                    worksheet_dc.Cells[tablerow * tcnt + 5, 13] = _RoadConfig.DetectWidth;
                    if (cemile != prjinfo._EndMile)
                    {
                        srcrange = worksheet_dc.get_Range(String.Format("A{0}:T{1}", tablerow * tcnt + 1, tablerow * (tcnt + 1) - 1));
                        ++tcnt;
                        destrange = worksheet_dc.get_Range(String.Format("A{0}", tablerow * tcnt + 1));
                        srcrange.Copy(destrange);
                        destrange = worksheet_dc.get_Range(String.Format("F{0}:O{1}", tablerow * tcnt + 7, tablerow * tcnt + 6 + DiseaseTypes.roadbeddislist.Count));
                        destrange.ClearContents();
                    }
                    csmile = cemile;
                }
                smile = emile;
                DiseaseTypes.Clear();
            }
            if (prjinfo._EndMile % 1000 != 0 || (MarkVal[len] != null && MarkVal[len].Contains("路面单元")))
            {
                worksheet_dc.Cells[tablerow * tcnt + 3, 1] = "路线名称：" + prjinfo._RoadName;
                worksheet_dc.Cells[tablerow * tcnt + 3, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                worksheet_dc.Cells[tablerow * tcnt + 3, 8] = prjinfo._DataDate;
                worksheet_dc.Cells[tablerow * tcnt + 3, 13] = prjinfo._DataPerson;
                worksheet_dc.Cells[tablerow * tcnt + 4, 8] = csmile;
                worksheet_dc.Cells[tablerow * tcnt + 4, 13] = prjinfo._EndMile;
                worksheet_dc.Cells[tablerow * tcnt + 5, 8] = Math.Abs(csmile - prjinfo._EndMile);
                worksheet_dc.Cells[tablerow * tcnt + 5, 13] = _RoadConfig.DetectWidth;
            }
        }
        #endregion


        #region  导出水泥板块病害
        public static void OutputBkDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\板块病害列表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_板块病害列表.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteBkDisLB2Xls(_Worksheet_lb, prjinfo, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteBkDisLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            int len = dislist.Length, i = 0, temp = 1;
            if (len < 1)
                return;

            MSExcel.Range destrange;
            string[] s;
            int bknum = SnbkSetForm.bknum;
            double bklength = _Setting.PlateLength;
            int smile = prjinfo._Direction > 0 ? prjinfo._StartMile : prjinfo._EndMile;
            int emile = prjinfo._Direction < 0 ? prjinfo._StartMile : prjinfo._EndMile;
            int cnt = Convert.ToInt32((emile - smile) / bklength) + 1;
            len = len + Convert.ToInt32((emile - smile) / bklength);
            object[,] val = new object[len, 11];
            List<Disease> tdis = new List<Disease>();
            List<Disease> udis = new List<Disease>();
            for (int j = temp; j < cnt; ++j) //遍历每个板块
            {
                val[i, 0] = bknum;
                val[i, 1] = i + 1;
                val[i, 2] = smile + (j - 1) * bklength;
                val[i, 3] = smile + j * bklength;
                if ((smile + j * bklength) > prjinfo._EndMile)
                {
                    val[i, 3] = prjinfo._EndMile;
                }
                val[i, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                temp++; bknum++;
                double ssmile = smile + (j - 1) * bklength;
                if (judgeroadtype(ssmile, _RoadPart, prjinfo) == 0)
                {
                    val[i, 10] = "沥青路面";
                    ++i;
                    continue;
                }
                else
                {
                    for (int k = 0; k < dislist.Length; ++k) //遍历所有病害
                    {
                        if (dislist[k].m_mile >= smile + (j - 1) * bklength && dislist[k].m_mile < smile + j * bklength && dislist[k].RoadType == "水泥")//统计一个板块内的水泥病害,然后再计算
                        {
                            tdis.Add(dislist[k]);   //一个板块内所有病害tdis
                        }
                    }
                    if (tdis.Count == 0)//如果板块没有病害，下一行
                    {
                        val[i, 10] = "水泥路面";
                        ++i;
                        continue;
                    }
                    foreach (Disease t in tdis)
                    {
                        if (t.computetype == 5)
                        {
                            udis.Add(t);//添加破碎版病害列表
                            break;
                        }
                    }
                    if (udis.Count > 0)//如果该板块内病害含有破碎板
                    {
                        s = udis[0].RoadDisType.Split('.');

                        if (udis[0].RoadType == "水泥")
                        {
                            val[i, 5] = s[0];
                            if (s.Length > 1)
                            {
                                val[i, 6] = s[1];
                            }
                            else
                            {
                                val[i, 6] = " ";
                            }
                        }
                        val[i, 7] = Math.Round(udis[0].Area, 5);
                        val[i, 8] = udis[0].calcheight;
                        val[i, 9] = udis[0].calcwidth;
                        val[i, 10] = udis[0].RoadType + "路面";
                        ++i;
                    }
                    else //如果板块内没有破碎版病害
                    {
                        foreach (Disease t in tdis)
                        {
                            val[i, 0] = bknum;
                            val[i, 1] = i + 1;
                            val[i, 2] = smile + (j - 1) * bklength;
                            val[i, 3] = smile + j * bklength;
                            val[i, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                            s = t.RoadDisType.Split('.');
                            if (t.RoadType == "水泥")
                            {
                                val[i, 5] = s[0];
                                if (s.Length > 1)
                                {
                                    val[i, 6] = s[1];
                                }
                                else
                                {
                                    val[i, 6] = " ";
                                }
                                val[i, 7] = Math.Round(t.Area, 5);
                                val[i, 8] = Math.Round(t.calcheight, 5);
                                val[i, 9] = Math.Round(t.calcwidth, 5);
                            }
                            val[i, 10] = t.RoadType + "路面";
                            ++i;
                        }
                    }
                    tdis.Clear();
                    udis.Clear();


                }
            }

            destrange = _Worksheet.get_Range(String.Format("A2:K{0}", len + 2));
            destrange.Value2 = val;
            int tlen = 0;
            for (int k = 0; k < len; ++k)
            {
                if (val[k, 0] == null)
                {
                    break;
                }
                tlen++;
            }
            destrange = _Worksheet.get_Range(String.Format("A2:K{0}", tlen + 1));
            GlobalExcel.SetBorderLine(destrange, 53);
        }
        #endregion

        private static int judgeroadtype(double mile, List<MilePart> roadpart, ProjectInfo prjinfo)
        {
            for (int i = 0; i < roadpart.Count - 1; ++i)
            {
                if ((prjinfo._Direction > 0 && mile >= roadpart[i].mile && mile < roadpart[i + 1].mile) || (prjinfo._Direction < 0 && mile < roadpart[i].mile && mile >= roadpart[i + 1].mile))
                {
                    return roadpart[i].roadtype;
                }

            }
            return -1;
        }

        public static void OutputGPSRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\综合报表模板GPS.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_DR = _Workbook.Sheets["DR"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_IRI = _Workbook.Sheets["IRI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_MTD = _Workbook.Sheets["WR"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_Rut = _Workbook.Sheets["RD"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PQI = _Workbook.Sheets["PQI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PCI = _Workbook.Sheets["PCI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_RQI = _Workbook.Sheets["RQI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_RDI = _Workbook.Sheets["RDI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PBI = _Workbook.Sheets["PBI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PWI = _Workbook.Sheets["PWI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sn = _Workbook.Sheets["水泥病害"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lq = _Workbook.Sheets["沥青病害"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PB = _Workbook.Sheets["PB"] as MSExcel.Worksheet;

            if (prjinfo._Direction > 0)
            {
                WriteGPSAll2XlsUp(_Worksheet_PQI, _Worksheet_PCI, _Worksheet_RQI, _Worksheet_RDI,
                    _Worksheet_MTD, _Worksheet_DR, _Worksheet_IRI, _Worksheet_Rut, _Worksheet_PBI,
                    _Worksheet_PWI, _Worksheet_PB, prjinfo, prjdir, _RoadPart, _RoadDisList,
                    _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                    _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _GPSInfo);
            }
            else
            {
                WriteGPSAll2XlsDown(_Worksheet_PQI, _Worksheet_PCI, _Worksheet_RQI, _Worksheet_RDI,
                                   _Worksheet_MTD, _Worksheet_DR, _Worksheet_IRI, _Worksheet_Rut, _Worksheet_PBI,
                                   _Worksheet_PWI, _Worksheet_PB, prjinfo, prjdir, _RoadPart, _RoadDisList,
                                   _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                                   _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _GPSInfo);
            }
            WriteGPSDis2Xls(_Worksheet_lq, _Worksheet_sn, prjinfo, prjdir, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }




        //上行
        private static void WriteGPSAll2XlsUp(
            MSExcel.Worksheet worksheetPQI, MSExcel.Worksheet worksheetPCI,
            MSExcel.Worksheet worksheetRQI, MSExcel.Worksheet worksheetRDI,
            MSExcel.Worksheet worksheetMTD, MSExcel.Worksheet worksheetDR,
            MSExcel.Worksheet worksheetIRI, MSExcel.Worksheet worksheetRut,
            MSExcel.Worksheet worksheetPBI, MSExcel.Worksheet worksheetPWI,
            MSExcel.Worksheet worksheetPB, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal,
            ExcelGPS[] GPSInfo)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] GPSStartObj = new object[len, 5];
            object[,] GPSEndObj = new object[len, 4];
            object[,] IDObj = new object[len, 1];
            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 4];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];
            object[,] PWIObj = new object[len, 1];
            object[,] PBObj = new object[len, 3];
            object[,] PBIObj = new object[len, 1];

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //GPS
                IDObj[rowcnt, 0] = i + 1; ;
                GPSStartObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;
                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                //PCI
                double drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                DRObj[rowcnt, 0] = drval;
                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    PCIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(DR!J{1},{2})",
                                _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                len - rowcnt + 1, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                else
                {
                    PCIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(DR!J{1},{2})",
                                _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                rowcnt + 2, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }

                //平整度相关
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 0 || _Setting.IRIExcelSide == 2)
                    {
                        IRIObj[rowcnt, 0] = LIRIVal[i];
                    }
                    if (_Setting.IRIExcelSide == 1 || _Setting.IRIExcelSide == 2)
                    {
                        IRIObj[rowcnt, 1] = RIRIVal[i];
                    }
                }
                else
                {
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                }

                if (_Setting.RQIJudgeType == 0)
                {
                    IRIObj[rowcnt, 2] = string.Format("=AVERAGE(J{0}:K{0})", rowcnt + 2);
                }
                else if (_Setting.RQIJudgeType == 1)
                {
                    IRIObj[rowcnt, 2] = string.Format("=MAX(J{0}, K{0})", rowcnt + 2);
                }

                RQIObj[rowcnt, 0] = string.Format("=100/(1+{0}*EXP(IRI!L{2}*{1}))",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], rowcnt + 2);

                //跳车相关
                if (prjinfo._IsIRIMTD)
                {
                    PBObj[rowcnt, 0] = PBVal[i][1];
                    PBObj[rowcnt, 1] = PBVal[i][2];
                    PBObj[rowcnt, 2] = PBVal[i][3];
                    PBIObj[rowcnt, 0] = string.Format("=IF((100-PB!J{0}*{1}-PB!K{0}*{2}-PB!L{0}*{3})>0,(100-PB!J{0}*{1}-PB!K{0}*{2}-PB!L{0}*{3}),0)",
                    rowcnt + 2, _PBIScore[1], _PBIScore[2], _PBIScore[3]);
                }

                //构造深度相关              
                if (prjinfo._IsDIRIMTD)
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 1] = RMTDVal[i];

                    //wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                    if (CMTDVal != null)
                    {
                        MTDObj[rowcnt, 2] = CMTDVal[i];
                        //计算磨耗WR
                        if (CMTDVal[i] == 0)
                        {
                            MTDObj[rowcnt, 3] = 0;
                        }
                        else
                        {
                            MTDObj[rowcnt, 3] = string.Format("=IF(WR!L{0}-MIN(WR!J{0},WR!K{0})>0, 100*(WR!L{0}-MIN(WR!J{0},WR!K{0}))/WR!L{0},0)", rowcnt + 2);
                        }
                    }
                }
                else
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                }
                PWIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(WR!M{1},{2})", _PWIa[0], rowcnt + 2, _PWIa[1]);

                //车辙相关
                if (prjinfo._IsRut)
                {
                    RutObj[rowcnt, 0] = LRutVal[i];
                    RutObj[rowcnt, 1] = RRutVal[i];
                    RutObj[rowcnt, 2] = SRutVal[i];
                    //RutObj[rowcnt, 2] = string.Format("=MAX(J{0},K{0})", rowcnt + 2);
                    RDIObj[rowcnt, 0] = string.Format("=IF(RD！L{0}<{1},{2}-{3}*RD！L{0},IF(RD！L{0}<{4},{5}-{6}*(RD！L{0}-{1}),0))",
                        rowcnt + 2, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                }

                //PQI
                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})+{3}*(RDI!J{0})+{4}*(PBI!J{0}))/({1}+{2}+{3}+{4}),5)",
                                rowcnt + 2,
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})+{3}*(RDI!J{0})+{4}*(PBI!J{0})+{5}*(PWI!J{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                            rowcnt + 2,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);

                    }

                }
                else
                {
                    PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0}))/({1}+{2}),5)",
                                rowcnt + 2,
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                }

                ++rowcnt;
            }

            //将结果复制进Excel
            MSExcel.Range destrange, sortrange;
            MSExcel.Worksheet[] tsheet = { worksheetPB,worksheetDR, worksheetIRI, worksheetRut,  worksheetMTD,
                                             worksheetPCI, worksheetRQI,worksheetRDI, worksheetPQI, worksheetPBI, worksheetPWI };
            object[] tobj = { PBObj, DRObj, IRIObj, RutObj, MTDObj, PCIObj, RQIObj, RDIObj, PQIObj, PBIObj, PWIObj };
            char[] valnum = { 'L', 'J', 'L', 'L', 'M', 'J', 'J', 'J', 'J', 'J', 'J' };
            for (int i = 0; i < tsheet.Length; ++i)
            {
                destrange = tsheet[i].get_Range(String.Format("A2:A{0}", len + 1));
                destrange.Value2 = IDObj;
                if (i <= 5)
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }

                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = tsheet[i].get_Range(String.Format("B2:E{0}", len + 1));
                    destrange.Value2 = GPSEndObj;
                    destrange = tsheet[i].get_Range(String.Format("F2:I{0}", len + 1));
                    destrange.Value2 = GPSStartObj;

                    destrange = tsheet[i].get_Range(String.Format("B2:{0}{1}", valnum[i], len + 1));
                    sortrange = tsheet[i].get_Range(String.Format("C2:C{0}", len + 1));

                    GlobalExcel.ReflectionColnum(tsheet[i], destrange, sortrange);
                }
                else
                {
                    destrange = tsheet[i].get_Range(String.Format("B2:E{0}", len + 1));
                    destrange.Value2 = GPSStartObj;
                    destrange = tsheet[i].get_Range(String.Format("F2:I{0}", len + 1));
                    destrange.Value2 = GPSEndObj;
                }
                if (i > 5)
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }

            }
        }

        //下行
        private static void WriteGPSAll2XlsDown(
          MSExcel.Worksheet worksheetPQI, MSExcel.Worksheet worksheetPCI,
          MSExcel.Worksheet worksheetRQI, MSExcel.Worksheet worksheetRDI,
          MSExcel.Worksheet worksheetMTD, MSExcel.Worksheet worksheetDR,
          MSExcel.Worksheet worksheetIRI, MSExcel.Worksheet worksheetRut,
          MSExcel.Worksheet worksheetPBI, MSExcel.Worksheet worksheetPWI,
          MSExcel.Worksheet worksheetPB, ProjectInfo prjinfo, DirectoryInfo prjdir,
          List<MilePart> roadpart, Disease[] arrdis,
          double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
          double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal,
          ExcelGPS[] GPSInfo)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] GPSStartObj = new object[len, 5];
            object[,] GPSEndObj = new object[len, 4];
            object[,] IDObj = new object[len, 1];
            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 4];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];
            object[,] PWIObj = new object[len, 1];
            object[,] PBObj = new object[len, 3];
            object[,] PBIObj = new object[len, 1];

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = len - 1, j = dlen - 1; i >= 0; i--)//i区间索引，j病害索引
            {
                int smile = roadpart[i + 1].mile;   //340750
                int emile = roadpart[i].mile;    //341000
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j >= 0 && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    --j;
                }

                //GPS
                IDObj[rowcnt, 0] = len - i; ;
                GPSStartObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;
                GPSEndObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i]._latitude;

                //PCI
                double drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                DRObj[rowcnt, 0] = drval;
                PCIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(DR!J{1},{2})",
                            _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            rowcnt + 2, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                //平整度相关
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 0 || _Setting.IRIExcelSide == 2)
                    {
                        IRIObj[rowcnt, 0] = LIRIVal[i];
                    }
                    if (_Setting.IRIExcelSide == 1 || _Setting.IRIExcelSide == 2)
                    {
                        IRIObj[rowcnt, 1] = RIRIVal[i];
                    }
                }
                else
                {
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                }

                if (_Setting.RQIJudgeType == 0)
                {
                    IRIObj[rowcnt, 2] = string.Format("=AVERAGE(J{0}:K{0})", rowcnt + 2);
                }
                else if (_Setting.RQIJudgeType == 1)
                {
                    IRIObj[rowcnt, 2] = string.Format("=MAX(J{0}, K{0})", rowcnt + 2);
                }

                RQIObj[rowcnt, 0] = string.Format("=100/(1+{0}*EXP(IRI!L{2}*{1}))",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], rowcnt + 2);

                //跳车相关
                if (prjinfo._IsIRIMTD)
                {
                    PBObj[rowcnt, 0] = PBVal[i][1];
                    PBObj[rowcnt, 1] = PBVal[i][2];
                    PBObj[rowcnt, 2] = PBVal[i][3];
                    PBIObj[rowcnt, 0] = string.Format("=IF((100-PB!J{0}*{1}-PB!K{0}*{2}-PB!L{0}*{3})>0,(100-PB!J{0}*{1}-PB!K{0}*{2}-PB!L{0}*{3}),0)",
                    rowcnt + 2, _PBIScore[1], _PBIScore[2], _PBIScore[3]);
                }

                //构造深度相关              
                if (prjinfo._IsDIRIMTD)
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 1] = RMTDVal[i];

                    //wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                    if (CMTDVal != null)
                    {
                        MTDObj[rowcnt, 2] = CMTDVal[i];
                        //计算磨耗WR
                        if (CMTDVal[i] == 0)
                        {
                            MTDObj[rowcnt, 3] = 0;
                        }
                        else
                        {
                            MTDObj[rowcnt, 3] = string.Format("=IF(WR!L{0}-MIN(WR!J{0},WR!K{0})>0, 100*(WR!L{0}-MIN(WR!J{0},WR!K{0}))/WR!L{0},0)", rowcnt + 2);
                        }
                    }
                }
                else
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                }
                PWIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(WR!M{1},{2})", _PWIa[0], rowcnt + 2, _PWIa[1]);

                //车辙相关
                if (prjinfo._IsRut)
                {
                    RutObj[rowcnt, 0] = LRutVal[i];
                    RutObj[rowcnt, 1] = RRutVal[i];
                    RutObj[rowcnt, 2] = SRutVal[i];
                    //RutObj[rowcnt, 2] = string.Format("=MAX(J{0},K{0})", rowcnt + 2);
                    RDIObj[rowcnt, 0] = string.Format("=IF(RD！L{0}<{1},{2}-{3}*RD！L{0},IF(RD！L{0}<{4},{5}-{6}*(RD！L{0}-{1}),0))",
                        rowcnt + 2, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                }

                //PQI
                if (roadpart[i].roaddegree <= 1)
                {
                    PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})+{3}*(RDI!J{0})+{4}*(PBI!J{0})+{5}*(PWI!J{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                            rowcnt + 2,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                }
                else
                {
                    PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0}))/({1}+{2}),5)",
                            rowcnt + 2,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }

                ++rowcnt;
            }

            //将结果复制进Excel
            MSExcel.Range destrange, sortrange;
            MSExcel.Worksheet[] tsheet = { worksheetPB,worksheetDR, worksheetIRI, worksheetRut,  worksheetMTD,
                                             worksheetPCI, worksheetRQI,worksheetRDI, worksheetPQI, worksheetPBI, worksheetPWI };
            object[] tobj = { PBObj, DRObj, IRIObj, RutObj, MTDObj, PCIObj, RQIObj, RDIObj, PQIObj, PBIObj, PWIObj };
            char[] valnum = { 'L', 'J', 'L', 'L', 'M', 'J', 'J', 'J', 'J', 'J', 'J' };
            for (int i = 0; i < tsheet.Length; ++i)
            {
                destrange = tsheet[i].get_Range(String.Format("A2:A{0}", len + 1));
                destrange.Value2 = IDObj;
                if (i <= 5)
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }

                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = tsheet[i].get_Range(String.Format("B2:E{0}", len + 1));
                    destrange.Value2 = GPSStartObj;  //  GPSStartObj
                    destrange = tsheet[i].get_Range(String.Format("F2:I{0}", len + 1));
                    destrange.Value2 = GPSEndObj;  //  GPSEndObj

                    destrange = tsheet[i].get_Range(String.Format("B2:{0}{1}", valnum[i], len + 1));
                    sortrange = tsheet[i].get_Range(String.Format("C2:C{0}", len + 1));

                    GlobalExcel.ReflectionColnum(tsheet[i], destrange, sortrange);
                }

                if (i > 5)
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }

            }
        }

        private static void WriteGPSDis2Xls(MSExcel.Worksheet worksheet_lq, MSExcel.Worksheet worksheet_sn,
            ProjectInfo prjinfo, DirectoryInfo prjdir, Disease[] dislist)
        {
            object[,] disinfo = new object[1, 17];

            string[] gpsinfostrs = null;
            ExcelGPS[] tempinfos = null;
            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
            {
                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
                tempinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }
            string disname = "";
            string disgrade = "";
            ExcelGPS tempgpsinfo;
            int len = dislist.Length;
            int gi = 0, colcnt = 0;
            int rowcnt_lq = 2, rowcnt_sn = 2;
            MSExcel.Range destrange = null, sortrange = null;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                for (; gi < tempinfos.Length; ++gi)
                {
                    if (prjinfo._Direction > 0)
                    {
                        if (tempinfos[gi]._mile >= dislist[i].m_mile)
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (tempinfos[gi]._mile <= dislist[i].m_mile)
                        {
                            break;
                        }
                    }
                }
                if (gi < tempinfos.Length)
                {
                    tempgpsinfo = tempinfos[gi];
                }
                else
                {
                    tempgpsinfo = tempinfos[tempinfos.Length - 1];
                }

                string[] temp = dislist[i].RoadDisType.Split('.');
                if (temp.Length > 1)
                {
                    disname = temp[0];
                    disgrade = temp[1];
                }
                else if (temp.Length == 1)
                {
                    disname = temp[0];
                    disgrade = "";
                }
                else
                {
                    disname = "";
                    disgrade = "";
                }
                colcnt = 0;
                disinfo[0, colcnt++] = i + 1;
                disinfo[0, colcnt++] = tempgpsinfo._utctime;
                disinfo[0, colcnt++] = dislist[i].m_mile;
                disinfo[0, colcnt++] = tempgpsinfo._longitude;
                disinfo[0, colcnt++] = tempgpsinfo._latitude;

                disinfo[0, colcnt++] = tempgpsinfo._utctime;
                disinfo[0, colcnt++] = dislist[i].m_mile;
                disinfo[0, colcnt++] = tempgpsinfo._longitude;
                disinfo[0, colcnt++] = tempgpsinfo._latitude;

                disinfo[0, colcnt++] = disgrade;
                disinfo[0, colcnt++] = disname;
                disinfo[0, colcnt++] = dislist[i].calcheight;
                disinfo[0, colcnt++] = dislist[i].calcwidth;
                disinfo[0, colcnt++] = dislist[i].Area;
                disinfo[0, colcnt++] = prjinfo._RoadNum;
                disinfo[0, colcnt++] = dislist[i].imgname;
                disinfo[0, colcnt++] = dislist[i].imgpath;

                if (dislist[i].RoadType == "沥青")
                {
                    destrange = worksheet_lq.get_Range(string.Format("A{0}:Q{0}", rowcnt_lq));
                    destrange.Value2 = disinfo;
                    rowcnt_lq++;
                }
                else if (dislist[i].RoadType == "水泥")
                {
                    destrange = worksheet_sn.get_Range(string.Format("A{0}:Q{0}", rowcnt_sn));
                    destrange.Value2 = disinfo;
                    rowcnt_sn++;
                }
            }
            if (_Setting.Qufen_dis_degree == 1)
            {

                ((MSExcel.Range)worksheet_lq.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();

                ((MSExcel.Range)worksheet_sn.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();

            }

            if (rowcnt_lq < 3)
            {
                worksheet_lq.Delete();
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_lq.get_Range(String.Format("B2:Q{0}", rowcnt_lq - 1));
                sortrange = worksheet_lq.get_Range(String.Format("C2:C{0}", len + 1));//按桩号排序
                GlobalExcel.ReflectionColnum(worksheet_lq, destrange, sortrange);

            }
            if (rowcnt_sn < 3)
            {
                worksheet_sn.Delete();
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_sn.get_Range(String.Format("B2:Q{0}", rowcnt_sn - 1));
                sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);
            }
        }

        public static void OutputGPSStreetImg(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string fnamemile0 = prjdir.FullName + "\\StreetImg\\Camera0\\Street2Mile.txt";
            string fnamemile1 = prjdir.FullName + "\\StreetImg\\Camera1\\Street2Mile.txt";
            if (!(File.Exists(fnamemile0) || File.Exists(fnamemile1)))
            {
                MessageBox.Show("工程文件缺少景观图像数据");
                return;
            }
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\综合报表模板GPS _景观图像.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}.xlsx", path, prjdir.Name, "景观图像");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["景观图像"] as MSExcel.Worksheet;

            WriteGPSImg2Xls(_Worksheet, prjinfo, prjdir, "Street", prjinfo._StreetImgDis_Left);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputGPSPanoImg(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string fnamemile0 = prjdir.FullName + "\\PanoImg\\Camera0\\Pano2Mile.txt";
            if (!File.Exists(fnamemile0))
            {
                MessageBox.Show("工程文件缺少全景图像数据，请检查是否拼接全景图像，或重新加载工程");
                return;
            }
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\综合报表模板GPS _全景图像.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}.xlsx", path, prjdir.Name, "全景图像");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["全景图像"] as MSExcel.Worksheet;

            WriteGPSImg2Xls(_Worksheet, prjinfo, prjdir, "Pano", prjinfo._PanoImgDis);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputGPSRoadImg(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string fname = prjdir.FullName + "\\RoadImg\\Camera0\\Road2Mile.txt";
            if (!File.Exists(fname))
            {
                MessageBox.Show("工程文件缺少路面图像数据");
                return;
            }
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\综合报表模板GPS _路面图像.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}.xlsx", path, prjdir.Name, "路面图像");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["路面图像"] as MSExcel.Worksheet;

            WriteGPSImg2Xls(_Worksheet, prjinfo, prjdir, "Road", prjinfo._RoadImgDis);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteGPSImg2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, string ImgType, int ImgDis)
        {
            string fname = prjdir.FullName + "\\GPS2Mile.txt";
            string fnamemile0 = string.Format("{0}\\{1}Img\\Camera0\\{1}2Mile.txt", prjdir.FullName, ImgType);
            string fnamemile1 = string.Format("{0}\\{1}Img\\Camera1\\{1}2Mile.txt", prjdir.FullName, ImgType);
            string[] gpsinfostrs = null;
            ExcelGPS[] tempinfos = null;
            if (File.Exists(fname))
            {
                gpsinfostrs = File.ReadAllLines(fname);
                tempinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }

            string[] leftimgsinfo = null;
            string[] rightimgsinfo = null;
            int[] leftidx = null;
            int[] rightidx = null;
            string[] tstrs = null;

            if (File.Exists(fnamemile0))
            {
                leftimgsinfo = File.ReadAllLines(fnamemile0);
                leftidx = new int[leftimgsinfo.Length];
                for (int i = 0; i < leftimgsinfo.Length; ++i)
                {
                    tstrs = leftimgsinfo[i].Split('_');
                    if (tstrs.Length > 3)
                    {
                        leftidx[i] = int.Parse(tstrs[tstrs.Length - 3].Remove(4, 1));
                    }
                    else
                    {
                        leftidx[i] = int.Parse(tstrs[tstrs.Length - 2].Remove(4, 1));
                    }
                }
            }
            if (File.Exists(fnamemile1))
            {
                rightimgsinfo = File.ReadAllLines(fnamemile1);
                rightidx = new int[rightimgsinfo.Length];
                for (int i = 0; i < rightimgsinfo.Length; ++i)
                {
                    tstrs = rightimgsinfo[i].Split('_');
                    if (tstrs.Length > 3)
                    {
                        rightidx[i] = int.Parse(tstrs[tstrs.Length - 3].Remove(4, 1));
                    }
                    else
                    {
                        rightidx[i] = int.Parse(tstrs[tstrs.Length - 2].Remove(4, 1));
                    }
                }
            }

            int tmpmile;
            int len = prjinfo._EndDmi / ImgDis + 1;
            int temp = 0;
            object[,] dataobj = new object[len, 9];
            int gi = 0, tmile = 0, tdmi = 0;
            ExcelGPS tempgpsinfo = null;
            int colcnt = 0;
            if (leftidx != null)
            {
                for (int i = 0; i < leftidx.Length; i++)//i区间索引，j病害索引
                {
                    if (leftidx[i] < len)
                    {
                        //  tdmi = leftidx[i] * ImgDis;
                        tdmi = i * ImgDis;
                        tmile = prjinfo.Dmi2Mile(tdmi);
                        for (; gi < tempinfos.Length; ++gi)
                        {
                            if (prjinfo._Direction > 0)
                            {
                                if (tempinfos[gi]._mile >= tmile)
                                {
                                    break;
                                }
                            }
                            else
                            {
                                if (tempinfos[gi]._mile <= tmile)
                                {
                                    break;
                                }
                            }
                        }
                        if (gi < tempinfos.Length)
                        {
                            tempgpsinfo = tempinfos[gi];
                        }
                        else
                        {
                            tempgpsinfo = tempinfos[tempinfos.Length - 1];
                        }
                        colcnt = 0;
                        if (leftidx[i] != i)
                        {
                            leftidx[i] = i;
                            dataobj[leftidx[i], colcnt++] = i + 1;
                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._utctime;

                            //tmpmile = tempinfos[gi+1]._mile - prjinfo._Direction * ImgDis;
                            //if (tmpmile < 0)
                            //    tmpmile = 0;
                            //dataobj[leftidx[i], colcnt++] = tmpmile;
                            dataobj[leftidx[i], colcnt++] = tmile;

                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._longitude;
                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._latitude;
                            dataobj[leftidx[i], colcnt++] = "图像丢帧";
                            dataobj[leftidx[i], colcnt++] = "图像丢帧";

                        }
                        else
                        {
                            dataobj[leftidx[i], colcnt++] = i + 1;
                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._utctime;

                            //tmpmile = tempgpsinfo._mile;
                            //if (tmpmile < 0)
                            //    tmpmile = 0;
                            //dataobj[leftidx[i], colcnt++] = tmpmile;
                            dataobj[leftidx[i], colcnt++] = tmile;

                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._longitude;
                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._latitude;
                            temp = leftimgsinfo[i].LastIndexOf('\\');
                            dataobj[leftidx[i], colcnt++] = leftimgsinfo[i].Substring(temp + 1);
                            int temp2 = leftimgsinfo[i].IndexOf(' ') + 2;
                            dataobj[leftidx[i], colcnt++] = string.Format("\\{0}Img\\Camera0\\{1}", ImgType, leftimgsinfo[i].Substring(temp2, temp - temp2));
                        }
                    }
                }
            }

            gi = 0;
            if (rightidx != null)
            {
                for (int i = 0; i < rightidx.Length; i++)//i区间索引，j病害索引
                {
                    if (rightidx[i] < len)
                    {
                        if (dataobj[rightidx[i], 0] != null)
                        {
                            tdmi = rightidx[i] * ImgDis;
                            tmile = prjinfo.Dmi2Mile(tdmi);
                            for (; gi < tempinfos.Length; ++gi)
                            {
                                if (prjinfo._Direction > 0)
                                {
                                    if (tempinfos[gi]._mile >= tmile)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    if (tempinfos[gi]._mile <= tmile)
                                    {
                                        break;
                                    }
                                }
                            }
                            if (gi < tempinfos.Length)
                            {
                                tempgpsinfo = tempinfos[gi];
                            }
                            else
                            {
                                tempgpsinfo = tempinfos[tempinfos.Length - 1];
                            }
                            colcnt = 0;
                            if (rightidx[i] != i)
                            {
                                rightidx[i] = i;
                                dataobj[rightidx[i], colcnt++] = i + 1;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._utctime;

                                //tmpmile = tempinfos[gi + 1]._mile - prjinfo._Direction * ImgDis;
                                //if (tmpmile < 0)
                                //    tmpmile = 0;
                                //dataobj[rightidx[i], colcnt++] = tmpmile;
                                dataobj[rightidx[i], colcnt++] = tmile;

                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._longitude;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._latitude;

                                colcnt = 7;
                                rightidx[i] = i;
                                dataobj[rightidx[i], colcnt++] = "图像丢帧";
                                dataobj[rightidx[i], colcnt++] = "图像丢帧";
                            }
                            else
                            {
                                dataobj[rightidx[i], colcnt++] = i + 1;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._utctime;

                                //tmpmile = tempgpsinfo._mile;
                                //if (tmpmile < 0)
                                //    tmpmile = 0;
                                //dataobj[rightidx[i], colcnt++] = tmpmile;
                                dataobj[rightidx[i], colcnt++] = tmile;

                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._longitude;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._latitude;
                                colcnt = 7;
                                temp = rightimgsinfo[i].LastIndexOf('\\');
                                dataobj[rightidx[i], colcnt++] = rightimgsinfo[i].Substring(temp + 1);
                                int temp2 = rightimgsinfo[i].IndexOf(' ') + 2;
                                dataobj[rightidx[i], colcnt++] = string.Format("\\{0}Img\\Camera1\\{1}", ImgType, rightimgsinfo[i].Substring(temp2, temp - temp2));
                            }  
                        }
                    }
                }
            }

            MSExcel.Range destrange = worksheet.get_Range(string.Format("A2:I{0}", len + 1));
            destrange.Value2 = dataobj;
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(string.Format("B2:I{0}", len + 1));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("C2:C{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
        }

        #region 中交国通报表模板

        public static void OutputZJGTIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Worksheet _Worksheet = null;
            MSExcel.Workbook _Workbook = null;
            string Destxls = string.Format(@"{0}\{1}_路面平整度.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 20)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\001-路面平整度报告模板.xlsx",
                    System.Windows.Forms.Application.StartupPath);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }

            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 6;
            }
            else
            {
                datacol = 6;
                nodatacol = 3;
            }

            object[,] vallist = new object[len, 9];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadCode;
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;

                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            vallist[i, datacol] = Math.Round((LIRIVal[i] + RIRIVal[i]) / 2, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            vallist[i, datacol] = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        vallist[i, datacol] = Math.Round((LIRIVal[i]), 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        vallist[i, datacol] = Math.Round((RIRIVal[i]), 5);
                    }
                }
                else
                {
                    vallist[i, datacol] = Math.Round(LIRIVal[i], 5);
                }

                vallist[i, datacol + 1] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{3}{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 11, (char)('A' + datacol));
                vallist[i, datacol + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 11,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                    (char)('A' + datacol + 1));

                vallist[i, nodatacol] = null;
                vallist[i, nodatacol + 1] = null;
                vallist[i, nodatacol + 2] = null;
            }

            destrange = _Worksheet.get_Range(String.Format("A11:I{0}", len + 10));
            destrange.Value2 = vallist;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 11, 2, 9, true);
                GlobalExcel.Reflection(_Worksheet, 11, 2, 2, false);
            }
            if (_Setting.IsStatistics)
            {
                len += 11;
                GlobalExcel.WriteExcel(len, 1, 1, 3, "平均值", _Worksheet, 63);
                _Worksheet.Cells[len, datacol + 1] = string.Format("=ROUND(AVERAGE({1}11:{1}{0}),5)", len - 1, (char)('A' + datacol));
                _Worksheet.Cells[len, datacol + 2] = string.Format("=ROUND(100/(1+{0}*EXP({1}*{3}{2})),5)",
                        _RQIa[roadpart[0].roaddegree][roadpart[0].roadtype][0], _RQIa[roadpart[0].roaddegree][roadpart[0].roadtype][1], len, (char)('A' + datacol));
                _Worksheet.Cells[len, datacol + 3] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    len,
                    _RQIGrade[roadpart[0].roaddegree][roadpart[0].roadtype][0],
                    _RQIGrade[roadpart[0].roaddegree][roadpart[0].roadtype][1],
                    _RQIGrade[roadpart[0].roaddegree][roadpart[0].roadtype][2],
                    _RQIGrade[roadpart[0].roaddegree][roadpart[0].roadtype][3],
                     (char)('A' + datacol + 1));

                GlobalExcel.WriteExcel(++len, 1, 1, 3, "备注", _Worksheet, 63);
                GlobalExcel.WriteExcel(len, 4, 1, 6, "---", _Worksheet, 63);
                destrange = _Worksheet.get_Range(String.Format("A11:I{0}", len));
            }
            _Worksheet.Cells[11, 10] = "起点";
            _Worksheet.Cells[len + 10, 10] = "终点";
            GlobalExcel.SetBorderLine(destrange, 63);
        }

        public static void OutputZJGTDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _Workbook = null;
            MSExcel.Worksheet _Worksheet_lqpci = null;
            MSExcel.Worksheet _Worksheet_snpci = null;
            MSExcel.Worksheet _Worksheet_lqhz = null;
            MSExcel.Worksheet _Worksheet_snhz = null;

            string Destxls = string.Format(@"{0}\{1}_路面损坏.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 100)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\002-路面损坏报告模板.xlsx",
                    System.Windows.Forms.Application.StartupPath);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
                WriteZJGTPrj2DisXls(_WorksheetPrj, prjinfo);
            }
            if (disval == 100)
            {
                if (prjinfo._Direction > 0)
                {
                    _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏下行100m"] as MSExcel.Worksheet;
                    _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏下行100m"] as MSExcel.Worksheet;
                    _Worksheet_lqhz.Delete();
                    _Worksheet_snhz.Delete();

                    _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏上行100m"] as MSExcel.Worksheet;
                    _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏上行100m"] as MSExcel.Worksheet;
                }
                else
                {
                    _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏上行100m"] as MSExcel.Worksheet;
                    _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏上行100m"] as MSExcel.Worksheet;
                    _Worksheet_lqhz.Delete();
                    _Worksheet_snhz.Delete();

                    _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏下行100m"] as MSExcel.Worksheet;
                    _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏下行100m"] as MSExcel.Worksheet;
                }
                WriteZJGTDisHZTJ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, disval * 10, _MarkVal);
            }
            else if (disval == 1000)
            {
                _Worksheet_lqpci = _Workbook.Sheets["沥青路面损坏1000m"] as MSExcel.Worksheet;
                _Worksheet_snpci = _Workbook.Sheets["水泥路面损坏1000m"] as MSExcel.Worksheet;
                WriteZJGTPCI2Xls(_Worksheet_snpci, _Worksheet_lqpci, prjinfo, prjdir, _RoadPart, _RoadDisList);
            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTDisHZTJ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen, string[] MarkVal)
        {
            MSExcel.Range srcrange, destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int sn_tablerow = _Setting.cmop_rows;
            int lq_tablerow = _Setting.cmop_rows;

            int tcnt_sn = 0;
            int tcnt_lq = 0;

            int sn_csmile = 0, sn_cemile = 0;
            int lq_csmile = 0, lq_cemile = 0;
            bool sn_flag = false, lq_flag = false;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[j].calcheight;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[disnum, 1];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            disval[kk, 0] = type.totallength;
                        }
                        else
                        {
                            disval[kk, 0] = type.totalarea;
                        }
                    }
                    destrange = worksheet_snhz.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('F' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        sn_tablerow * tcnt_sn + 7,
                        sn_tablerow * tcnt_sn + 6 + disnum));
                    destrange.Value2 = disval;

                    sn_cemile = emile;
                    if (!sn_flag)
                    {
                        sn_flag = true;
                        sn_csmile = smile;
                    }
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[disnum, 1];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            disval[kk, 0] = type.totallength;
                        }
                        else
                        {
                            disval[kk, 0] = type.totalarea;
                        }
                    }
                    destrange = worksheet_lqhz.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('F' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        lq_tablerow * tcnt_lq + 7,
                        lq_tablerow * tcnt_lq + 6 + disnum));
                    destrange.Value2 = disval;

                    lq_cemile = emile;
                    if (!lq_flag)
                    {
                        lq_flag = true;
                        lq_csmile = smile;
                    }
                }
                //cwb
                //道路材质切换
                if (roadpart[i].roadtype!= roadpart[i+1].roadtype)
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 13] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("F{0}:O{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 13] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("F{0}:O{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }
                }

                if (emile % xlslen == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")))
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 13] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("F{0}:O{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 13] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("F{0}:O{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }
                }
            }
            if (roadpart[len].mile % xlslen != 0 || (MarkVal[len] != null && MarkVal[len].Contains("路面单元")))
            {
                if (sn_csmile != sn_cemile)
                {
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 13] = roadpart[len].mile;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 13] = roadpart[len].mile;
                }
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (tcnt_sn + 1) + 1));
                destrange = worksheet_snhz.get_Range(String.Format("F{0}:O{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                object[,] ttobj = (object[,])destrange.Value2;
                bool hasdata = false;
                for (int i = 1; i <= disnum; ++i)
                {
                    for (int j = 1; j <= 10; ++j)
                    {
                        if (ttobj[i, j] != null)
                        {
                            hasdata = true;
                            break;
                        }
                    }
                    if (hasdata)
                    {
                        break;
                    }
                }
                if (!hasdata)
                {
                    srcrange.EntireRow.Delete();
                }
            }

            if (Haslqflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (tcnt_lq + 1) + 1));
                destrange = worksheet_lqhz.get_Range(String.Format("F{0}:O{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
                object[,] ttobj = (object[,])destrange.Value2;
                bool hasdata = false;
                for (int i = 1; i <= disnum; ++i)
                {
                    for (int j = 1; j <= 10; ++j)
                    {
                        if (ttobj[i, j] != null)
                        {
                            hasdata = true;
                            break;
                        }
                    }
                    if (hasdata)
                    {
                        break;
                    }
                }
                if (!hasdata)
                {
                    srcrange.EntireRow.Delete();
                }
            }

            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }
            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }

            RoadDiseaseTypes.Clear();
        }
        private static void WriteZJGTPCI2Xls(MSExcel.Worksheet worksheet_sn, MSExcel.Worksheet worksheet_lq,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            object[,] vallist_sn = new object[len, 9];
            object[,] vallist_lq = new object[len, 9];
            object[,] vallist;
            int snidx = -1, lqidx = -1, tidx = 0;

            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 6;
            }
            else
            {
                datacol = 6;
                nodatacol = 3;
            }

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
                if (roadpart[i].roadtype == 1)
                {
                    tidx = ++snidx;
                    vallist = vallist_sn;
                }
                else
                {
                    tidx = ++lqidx;
                    vallist = vallist_lq;
                }
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);

                vallist[tidx, 0] = prjinfo._RoadCode;
                vallist[tidx, 1] = roadpart[i].mile;
                vallist[tidx, 2] = roadpart[i + 1].mile;
                vallist[tidx, datacol] = drval;
                vallist[tidx, datacol + 1] = string.Format("=100-{0}*POWER({3}{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], tidx + 10, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], (char)('A' + datacol));
                vallist[tidx, datacol + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    tidx + 10,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3],
                    (char)('A' + datacol + 1));
                vallist[tidx, nodatacol] = null;
                vallist[tidx, nodatacol + 1] = null;
                vallist[tidx, nodatacol + 2] = null;
            }
            if (lqidx >= 0)
            {
                destrange = worksheet_lq.get_Range(String.Format("A10:I{0}", lqidx + 10));
                destrange.Value2 = vallist_lq;
                GlobalExcel.SetBorderLine(destrange, 53);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_lq, 10, 2, 9, true);
                    GlobalExcel.Reflection(worksheet_lq, 10, 2, 2, false);
                }
                worksheet_lq.Cells[10, 10] = "起点";
                worksheet_lq.Cells[lqidx + 10, 10] = "终点";
            }
            else
            {
                worksheet_lq.Delete();
            }
            if (snidx >= 0)
            {
                destrange = worksheet_sn.get_Range(String.Format("A10:I{0}", snidx + 10));
                destrange.Value2 = vallist_sn;
                GlobalExcel.SetBorderLine(destrange, 53);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_sn, 10, 2, 9, true);
                    GlobalExcel.Reflection(worksheet_sn, 10, 2, 2, false);
                }
                worksheet_sn.Cells[10, 10] = "起点";
                worksheet_sn.Cells[snidx + 10, 10] = "终点";
            }
            else
            {
                worksheet_sn.Delete();
            }
        }

        public static void OutputZJGTRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Worksheet _Worksheet = null;
            MSExcel.Workbook _Workbook = null;
            string srcxls;
            string Destxls = string.Format(@"{0}\{1}_路面车辙.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 10)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\003-路面车辙报告模板.xlsx",
                    System.Windows.Forms.Application.StartupPath, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }
            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            if (disval == 10)
            {
                Destxls = string.Format(@"{0}\{1}_原始车辙.xlsx", path, prjdir.Name);
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\013-路面车辙原始.xlsx",
                    System.Windows.Forms.Application.StartupPath, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Worksheet = _Workbook.Sheets["10米"] as MSExcel.Worksheet;
                WriteZJGTRut2Xls(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
            }
        }
        private static void WriteZJGTRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 6;
            }
            else
            {
                datacol = 6;
                nodatacol = 3;
            }

            object[,] vallist = new object[len, 9];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadCode;
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;

                //if (!(roadpart[i].roadtype == 0 && roadpart[i].roaddegree <= 1))
                //{
                //    vallist[i, datacol] = "-";
                //    vallist[i, datacol + 1] = "-";
                //    vallist[i, datacol + 2] = "-";
                //    continue;
                //}
                //vallist[i, datacol] = Math.Round(Math.Max(LRutVal[i], RRutVal[i]), 2);
                vallist[i, datacol] = Math.Round(SRutVal[i], 5);
                vallist[i, datacol + 1] = string.Format("=IF({7}{0}<={1},{2}-{3}*{7}{0},IF({7}{0}<={4},{5}-{6}*({7}{0}-{1}),0))",
                    i + 10, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1], (char)('A' + datacol));
                vallist[i, datacol + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 10,
                    _RDIGrade[roadpart[i].roaddegree][0],
                    _RDIGrade[roadpart[i].roaddegree][1],
                    _RDIGrade[roadpart[i].roaddegree][2],
                    _RDIGrade[roadpart[i].roaddegree][3],
                    (char)('A' + datacol + 1));

                vallist[i, nodatacol] = null;
                vallist[i, nodatacol + 1] = null;
                vallist[i, nodatacol + 2] = null;
            }

            destrange = _Worksheet.get_Range(String.Format("A10:I{0}", len + 9));
            destrange.Value2 = vallist;
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 10, 2, 9, true);
                GlobalExcel.Reflection(_Worksheet, 10, 2, 2, false);
            }
            if (_Setting.IsStatistics)
            {
                len += 10;
                GlobalExcel.WriteExcel(len, 1, 1, 3, "平均值", _Worksheet, 63);

                if (prjinfo._Direction > 0)
                {
                    _Worksheet.Cells[len, 4] = string.Format("=ROUND(AVERAGE(D10:D{0}),5)", len - 1);
                    _Worksheet.Cells[len, 5] = string.Format("=IF(D{0}<{1},{2}-{3}*D{0},IF(D{0}<{4},{5}-{6}*(D{0}-{1}),0))",
                        len, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                    _Worksheet.Cells[len, 6] = string.Format("=IF(E{0}>={1},\"优\",IF(E{0}>={2},\"良\",IF(E{0}>={3},\"中\",IF(E{0}>={4},\"次\",\"差\"))))",
                        len,
                        _RDIGrade[roadpart[0].roaddegree][0],
                        _RDIGrade[roadpart[0].roaddegree][1],
                        _RDIGrade[roadpart[0].roaddegree][2],
                        _RDIGrade[roadpart[0].roaddegree][3]);
                }
                else
                {
                    _Worksheet.Cells[len, 7] = string.Format("=ROUND(AVERAGE(G10:G{0}),5)", len - 1);
                    _Worksheet.Cells[len, 8] = string.Format("=IF(G{0}<={1},{2}-{3}*G{0},IF(G{0}<={4},{5}-{6}*(G{0}-{1}),0))",
                        len, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                    _Worksheet.Cells[len, 9] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                        len,
                        _RDIGrade[roadpart[0].roaddegree][0],
                        _RDIGrade[roadpart[0].roaddegree][1],
                        _RDIGrade[roadpart[0].roaddegree][2],
                        _RDIGrade[roadpart[0].roaddegree][3]);
                }

                GlobalExcel.WriteExcel(++len, 1, 1, 3, "备注", _Worksheet, 63);
                GlobalExcel.WriteExcel(len, 4, 1, 6, "---", _Worksheet, 63);

                destrange = _Worksheet.get_Range(String.Format("A10:I{0}", len));
            }
            GlobalExcel.SetBorderLine(destrange, 63);
            _Worksheet.Cells[10, 10] = "起点";
            _Worksheet.Cells[len + 9, 10] = "终点";
        }
        private static void WriteZJGTRut2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 2;
                nodatacol = 5;
            }
            else
            {
                datacol = 5;
                nodatacol = 2;
            }

            object[,] vallist = new object[len, 8];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                if (!(roadpart[i].roaddegree <= 1))
                {
                    vallist[i, datacol] = "--";
                    vallist[i, datacol + 1] = "--";
                    vallist[i, datacol + 2] = "--";
                }
                else
                {
                    vallist[i, datacol] = LRutVal[i];
                    vallist[i, datacol + 1] = RRutVal[i];
                    //vallist[i, datacol + 2] = string.Format("=MAX({0}{2}:{1}{2})", (char)('A' + datacol), (char)('A' + datacol + 1), i + 3);
                    vallist[i, datacol + 2] = SRutVal[i];
                }
            }

            destrange = _Worksheet.get_Range(String.Format("A3:H{0}", len + 2));
            destrange.Value2 = vallist;
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 8, true);
                GlobalExcel.Reflection(_Worksheet, 3, 1, 2, false);
            }
            GlobalExcel.SetBorderLine(destrange, 63);
            _Worksheet.Cells[3, 9] = "起点";
            _Worksheet.Cells[len + 2, 9] = "终点";
        }

        public static void OutputZJGTMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Worksheet _Worksheet = null;
            MSExcel.Workbook _Workbook = null;
            string Destxls = string.Format(@"{0}\{1}_路面抗滑性能.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 10)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\004-路面抗滑性能报告模板.xlsx",
                    System.Windows.Forms.Application.StartupPath, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }
            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 5;
            }
            else
            {
                datacol = 5;
                nodatacol = 3;
            }

            object[,] vallist = new object[len, 7];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadCode;
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;

                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, datacol] = Math.Round((LMTDVal[i] + RMTDVal[i]) * 0.5, 5);
                }
                else
                {
                    vallist[i, datacol] = Math.Round(LMTDVal[i], 5);
                }

                vallist[i, datacol + 1] = string.Format("=IF({0}{1}>='sheet1 (2)'!$O$11,\"合格\",\"不合格\")", (char)('A' + datacol), i + 10);
                vallist[i, nodatacol] = null;
                vallist[i, nodatacol + 1] = null;
            }
            destrange = _Worksheet.get_Range(String.Format("A10:G{0}", len + 9));
            destrange.Value2 = vallist;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 10, 2, 7, true);
                GlobalExcel.Reflection(_Worksheet, 10, 2, 2, false);
            }
            if (_Setting.IsStatistics)
            {
                len += 10;
                GlobalExcel.WriteExcel(len, 1, 3, 3, "备注", _Worksheet, 63);
                _Worksheet.Cells[len, datacol + 1] = "检测总数：";
                _Worksheet.Cells[len + 1, datacol + 1] = "合格点数：";
                _Worksheet.Cells[len + 2, datacol + 1] = "合格率：";
                _Worksheet.Cells[len, datacol + 2] = string.Format("=COUNT({0}10:{0}{1})", (char)('A' + datacol), len - 1);
                _Worksheet.Cells[len + 1, datacol + 2] = string.Format("=COUNTIF({0}10:{0}{1},\"合格\")", (char)('A' + datacol + 1), len - 1);
                _Worksheet.Cells[len + 2, datacol + 2] = string.Format("={0}{1}/{0}{2}", (char)('A' + datacol + 1), len + 1, len);
                destrange = _Worksheet.get_Range(String.Format("{0}{1}", (char)('A' + datacol + 1), len + 2));
                destrange.NumberFormat = "0.00%";
                destrange = _Worksheet.get_Range(String.Format("A10:G{0}", len + 2));
            }
            GlobalExcel.SetBorderLine(destrange, 63);
            _Worksheet.Cells[10, 8] = "起点";
            _Worksheet.Cells[len + 9, 8] = "终点";
        }

        public static void OutputZJGTGPS(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Worksheet _Worksheet = null;
            MSExcel.Workbook _Workbook = null;
            string Destxls = string.Format(@"{0}\{1}_路面高程GPS.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 5)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\005-路面高程、GPS.xlsx",
                    System.Windows.Forms.Application.StartupPath);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }

            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteGPS2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, disval);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteGPS2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, int xlslen)
        {
            string gpsfname = string.Format(@"{0}\GPS2Mile.txt", prjdir.FullName);
            if (!File.Exists(gpsfname))
            {
                MessageBox.Show("不存在GPS2Mile.txt，请进行GPS桩号匹配操作！");
                return;
            }

            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 2;
                nodatacol = 5;
            }
            else
            {
                datacol = 5;
                nodatacol = 2;
            }

            string[] gpsinfostrs = null;
            ExcelGPS[] gpsinfos = null;
            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
            {
                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
                gpsinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    gpsinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }

            int len = roadpart.Count - 1;
            object[,] mxlist = new object[len, 8];
            int gi = 0;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                mxlist[i, 0] = prjinfo._RoadCode;//国省道
                mxlist[i, 1] = smile;
                for (; gi < gpsinfos.Length - 1; gi++)
                {
                    if (prjinfo._Direction > 0)
                    {
                        if (gpsinfos[gi + 1]._mile >= smile && gpsinfos[gi]._mile < emile)
                        {
                            break;
                        }
                    }
                    else
                    {
                        if (gpsinfos[gi + 1]._mile <= smile && gpsinfos[gi]._mile > emile)
                        {
                            break;
                        }
                    }
                }
                mxlist[i, datacol] = gpsinfos[gi]._longitude;//经度
                mxlist[i, datacol + 1] = gpsinfos[gi]._latitude;//纬度
                mxlist[i, datacol + 2] = gpsinfos[gi]._elevation;//高程
                mxlist[i, nodatacol] = null;//经度
                mxlist[i, nodatacol + 1] = null;//纬度
                mxlist[i, nodatacol + 2] = null;//高程
            }

            len += 9;
            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A10:H{0}", len));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 63);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 10, 2, 8, true);
            }
            if (_Setting.IsStatistics)
            {
                GlobalExcel.WriteExcel(++len, 1, 1, 3, "备注", _Worksheet, 63);
                GlobalExcel.WriteExcel(len, 4, 1, 5, "---", _Worksheet, 63);
            }
            _Worksheet.Cells[10, 9] = "起点";
            _Worksheet.Cells[len, 9] = "终点";
        }

        public static void OutputZJGTPBI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _Workbook = null;
            MSExcel.Worksheet _Worksheet = null;
            string srcxls;
            string Destxls = string.Format(@"{0}\{1}_路面跳车报告模板.xlsx", path, prjdir.Name);

            if (File.Exists(Destxls) && disval != 10)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\006-路面跳车报告模板.xlsx",
                    System.Windows.Forms.Application.StartupPath, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }

            //_Worksheet = _Workbook.Sheets["Δh"] as MSExcel.Worksheet;
            //WritePB2Xls(_Worksheet, prjinfo, prjdir);
            //_Worksheet = _Workbook.Sheets["PBI"] as MSExcel.Worksheet;

            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTPBI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _PBIVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTPBI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, int[][] PBIVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 8;
            }
            else
            {
                datacol = 8;
                nodatacol = 3;
            }


            object[,] vallist = new object[len, 13];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadName;
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;
                vallist[i, datacol] = PBIVal[i][1];
                vallist[i, datacol + 1] = PBIVal[i][2];
                vallist[i, datacol + 2] = PBIVal[i][3];
                vallist[i, datacol + 3] = string.Format("=IF((100-{4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3})>0,(100- {4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3}),0)",
                    i + 11, _PBIScore[1], _PBIScore[2], _PBIScore[3], (char)('A' + datacol), (char)('A' + datacol + 1), (char)('A' + datacol + 2));
                vallist[i, datacol + 4] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 11,
                    _PBIGrade[roadpart[i].roaddegree][0],
                    _PBIGrade[roadpart[i].roaddegree][1],
                    _PBIGrade[roadpart[i].roaddegree][2],
                    _PBIGrade[roadpart[i].roaddegree][3],
                    (char)('A' + datacol + 3));
            }

            destrange = _Worksheet.get_Range(String.Format("A11:M{0}", len + 10));
            destrange.Value2 = vallist;

            // WritePBIStatistics(_Worksheet);
            GlobalExcel.SetBorderLine(destrange, 53);

            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, 11, 1, 8, true);
            //    GlobalExcel.Reflection(_Worksheet, 11, 1, 8, false);
            //}
        }

        public static void OutputZJGTPWI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _Workbook = null;
            MSExcel.Worksheet _Worksheet = null;
            string srcxls;
            string Destxls = string.Format(@"{0}\{1}_路面磨耗报告模板.xlsx.xlsx", path, prjdir.Name);

            if (File.Exists(Destxls) && disval != 10)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\007-路面磨耗报告模板.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);

                Destxls = string.Format(@"{0}\{1}_PWI_{2}m.xlsx", path, prjdir.Name, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _WorksheetPrj = _Workbook.Sheets["sheet1 (2)"] as MSExcel.Worksheet;
                WriteZJGTPrj2Xls(_WorksheetPrj, prjinfo);
            }

            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTPWI2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTPWI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int datacol = 0, nodatacol = 0;
            if (prjinfo._Direction > 0)
            {
                datacol = 3;
                nodatacol = 9;
            }
            else
            {
                datacol = 9;
                nodatacol = 3;
            }

            object[,] vallist = new object[len, 15];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadName;
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;

                vallist[i, datacol] = LMTDVal[i];
                vallist[i, datacol + 1] = RMTDVal[i];
                vallist[i, datacol + 2] = CMTDVal[i];
                if (CMTDVal[i] == 0)
                {
                    vallist[i, datacol + 3] = 0;
                }
                else  // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                {
                    vallist[i, datacol + 3] = string.Format("=IF({3}{0}-MIN({1}{0},{2}{0}), 100*({3}{0}-MIN({1}{0},{2}{0}))/{3}{0},0)", i + 10,
                       (char)('A' + datacol), (char)('A' + datacol + 1), (char)('A' + datacol + 2));
                }
                vallist[i, datacol + 4] = string.Format("=100-{0}*POWER({3}{1},{2})", _PWIa[0], i + 10, _PWIa[1], (char)('A' + datacol + 3));
                vallist[i, datacol + 5] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 10,
                    _PWIGrade[roadpart[i].roaddegree][0],
                    _PWIGrade[roadpart[i].roaddegree][1],
                    _PWIGrade[roadpart[i].roaddegree][2],
                    _PWIGrade[roadpart[i].roaddegree][3],
                (char)('A' + datacol + 4));
            }

            destrange = _Worksheet.get_Range(String.Format("A10:O{0}", len + 9));
            destrange.Value2 = vallist;

            GlobalExcel.SetBorderLine(destrange, 53);

            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, 4, 1, 9, true);
            //    GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            //}
        }

        public static void OutputZJGTPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\011-综合大表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_大表.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["PQI (2)"] as MSExcel.Worksheet;
            WriteZJGTAll2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTAll2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
           List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal,
           double[] LRutVal, double[] RRutVal, double[] SRutVal, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal)
        {
            bool IsExistGPSInfo = false;
            string[] gpsinfostrs = null;
            ExcelGPS[] gpsinfos = null;
            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
            {
                IsExistGPSInfo = true;
                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
                gpsinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    gpsinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }

            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double irival = 0, tpcival = 0;

            object[,] mxlist = new object[len, 34];
            string errlog = prjdir.FullName + "\\errlog.txt";
            int gi = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                mxlist[i, 0] = null;//国省道
                mxlist[i, 1] = string.Format("{0}{1}交通局", prjinfo._City, prjinfo._District);
                mxlist[i, 2] = prjinfo._RoadCode;
                mxlist[i, 3] = GlobalExcel._RoadDegreeStr[roadpart[i].roaddegree];
                mxlist[i, 4] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                mxlist[i, 5] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 6] = smile;
                mxlist[i, 7] = emile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                if (!IsExistGPSInfo)
                {
                    mxlist[i, 8] = null;//经度
                    mxlist[i, 9] = null;//纬度
                    mxlist[i, 10] = null;//高程
                }
                else
                {
                    for (; gi < gpsinfos.Length - 1; gi++)
                    {
                        if (prjinfo._Direction > 0)
                        {
                            if (gpsinfos[gi + 1]._mile >= smile && gpsinfos[gi]._mile < emile)
                            {
                                mxlist[i, 8] = gpsinfos[gi]._longitude;//经度
                                mxlist[i, 9] = gpsinfos[gi]._latitude;//纬度
                                mxlist[i, 10] = gpsinfos[gi]._elevation;//高程
                                break;
                            }
                        }
                        else
                        {
                            if (gpsinfos[gi + 1]._mile <= smile && gpsinfos[gi]._mile > emile)
                            {
                                mxlist[i, 8] = gpsinfos[gi]._longitude;//经度
                                mxlist[i, 9] = gpsinfos[gi]._latitude;//纬度
                                mxlist[i, 10] = gpsinfos[gi]._elevation;//高程
                                break;
                            }
                        }
                    }
                }

                //PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                mxlist[i, 17] = drval;
                mxlist[i, 18] = string.Format("=100-{0}*POWER(R{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 19] = string.Format("=IF(S{0}>={1},\"优\",IF(S{0}>={2},\"良\",IF(S{0}>={3},\"中\",IF(S{0}>={4},\"次\",\"差\"))))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI                
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                // trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                mxlist[i, 11] = irival;

                mxlist[i, 12] = string.Format("=ROUND(100/(1+{0}*EXP({1}*L{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 3);
                mxlist[i, 13] = string.Format("=IF(M{0}>={1},\"优\",IF(M{0}>={2},\"良\",IF(M{0}>={3},\"中\",IF(M{0}>={4},\"次\",\"差\"))))",
                    i + 3, _RQIGrade[roadpart[i].roaddegree][1][0], _RQIGrade[roadpart[i].roaddegree][1][1], _RQIGrade[roadpart[i].roaddegree][1][2], _RQIGrade[roadpart[i].roaddegree][1][3]);

                #region

                //Rut
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    //  double rutval = SRutVal[i];
                    mxlist[i, 14] = Math.Round(SRutVal[i], 5);

                    mxlist[i, 15] = string.Format("=IF(O{0}<{1},{2}-{3}*O{0},IF(O{0}<{4},{5}-{6}*(O{0}-{1}),0))",
                        i + 3,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                    mxlist[i, 16] = string.Format("=IF(P{0}>={1},\"优\",IF(P{0}>={2},\"良\",IF(P{0}>={3},\"中\",IF(P{0}>={4},\"次\",\"差\"))))",
                        i + 3,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);
                }

                #region PBI
                //各程度跳车数
                mxlist[i, 20] = _PBIVal[i][1];
                mxlist[i, 21] = _PBIVal[i][2];
                mxlist[i, 22] = _PBIVal[i][3];
                //PBI得分与评价
                int datacol = 20;
                mxlist[i, 23] = string.Format("=IF((100-{4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3})>0,(100- {4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3}),0)",
                i + 3, _PBIScore[1], _PBIScore[2], _PBIScore[3], (char)('A' + datacol), (char)('A' + datacol + 1), (char)('A' + datacol + 2));
                mxlist[i, 24] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                i + 3,
                _PBIGrade[roadpart[i].roaddegree][0],
                _PBIGrade[roadpart[i].roaddegree][1],
                _PBIGrade[roadpart[i].roaddegree][2],
                _PBIGrade[roadpart[i].roaddegree][3],
                (char)('A' + datacol + 3));
                #endregion

                #region PWI
                double wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                //if (CMTDVal[i] != 0)  缺少数据如何处理
                //{
                //    mxlist[i, 26] = wrval;

                //}
                //else
                //{
                //    mxlist[i, 26] = null;
                //    //mxlist[i, 27] = null;
                //    //mxlist[i, 28] = null;
                //}
                wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                mxlist[i, 25] = wrval;
                if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))
                {
                    mxlist[i, 25] = 0;
                }
                mxlist[i, 26] = string.Format("=100-{0}*POWER(Z{1},{2})", _PWIa[0], i + 3, _PWIa[1]);

                mxlist[i, 27] = string.Format("=IF(AA{0}>={1},\"优\",IF(AA{0}>={2},\"良\",IF(AA{0}>={3},\"中\",IF(AA{0}>={4},\"次\",\"差\"))))",
                    i + 3,
                    _PWIGrade[roadpart[i].roaddegree][0],
                    _PWIGrade[roadpart[i].roaddegree][1],
                    _PWIGrade[roadpart[i].roaddegree][2],
                    _PWIGrade[roadpart[i].roaddegree][3]
                );
                #endregion
                mxlist[i, 28] = milelength;
                //PQI
                if (roadpart[i].roaddegree < 2)
                {
                    //=ROUND(({1}*T{0}+{2}*N{0}+{3}*IF(EXACT(Q{0},\"-\"),0,Q{0}))/({1}+{2}+{3}),2)
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        mxlist[i, 29] = string.Format("=ROUND(({1}*S{0}+{2}*M{0}+{3}*IF(EXACT(P{0},\"-\"),0,P{0})+{4}*X{0})/({1}+{2}+{3}+{4}),5)",
                            i + 3,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        mxlist[i, 29] = string.Format("=ROUND(({1}*S{0}+{2}*M{0}+{3}*IF(EXACT(P{0},\"-\"),0,P{0})+{4}*X{0}+{5}*AA{0})/({1}+{2}+{3}+{4}+{5}),5)",
                                           i + 3,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                    }
                }
                else
                {
                    mxlist[i, 29] = string.Format("=ROUND(({1}*S{0}+{2}*M{0})/({1}+{2}),5)",
                        i + 3,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 30] = string.Format("=IF(AD{0}>={1},\"优\",IF(AD{0}>={2},\"良\",IF(AD{0}>={3},\"中\",IF(AD{0}>={4},\"次\",\"差\"))))",
                    i + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
            }
            #endregion
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:AE{0}", len + 2));
            destrange.Value2 = mxlist;

            GlobalExcel.SetBorderLine(destrange, 63);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("A3:AE{0}", len + 2));
                MSExcel.Range sortrange = worksheet.get_Range(String.Format("G3:G{0}", len + 2));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }

            worksheet.Cells[3, 32] = "起点";
            worksheet.Cells[len + 2, 32] = "终点";

        }

        public static void OutputZJGTRoadType(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Worksheet _Worksheet = null;
            MSExcel.Workbook _Workbook = null;
            string Destxls = string.Format(@"{0}\{1}_路面材质.xlsx", path, prjdir.Name);
            if (File.Exists(Destxls) && disval != 5)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\014-路面材质.xlsx",
                    System.Windows.Forms.Application.StartupPath, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            string sheetname = disval.ToString() + "米";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteZJGTRoadType2Xls(_Worksheet, _RoadPart);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTRoadType2Xls(MSExcel.Worksheet _Worksheet, List<MilePart> roadpart)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 3];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
            }

            destrange = _Worksheet.get_Range(String.Format("A2:C{0}", len + 1));
            destrange.Value2 = vallist;
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 2, 1, 3, true);
                GlobalExcel.Reflection(_Worksheet, 2, 1, 2, false);
            }
            GlobalExcel.SetBorderLine(destrange, 63);
            _Worksheet.Cells[2, 4] = "起点";
            _Worksheet.Cells[len + 1, 4] = "终点";
        }

        private static void WriteZJGTPrj2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[6, 3] = prjinfo._Province;
            _Worksheet.Cells[9, 3] = prjinfo._DataDate;
            _Worksheet.Cells[13, 3] = string.Format("{0} {1:K0+000}-{2:K0+000}", prjinfo._RoadCode, prjinfo._StartMile, prjinfo._EndMile);
            _Worksheet.Cells[17, 3] = prjinfo._DataWeather;
            _Worksheet.Cells[6, 15] = prjinfo._RoadCode;
        }
        private static void WriteZJGTPrj2DisXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[5, 3] = prjinfo._Province;
            _Worksheet.Cells[8, 3] = prjinfo._DataDate;
            _Worksheet.Cells[12, 3] = string.Format("{0} {1:K0+000}-{2:K0+000}", prjinfo._RoadCode, prjinfo._StartMile, prjinfo._EndMile);
            _Worksheet.Cells[16, 3] = prjinfo._DataWeather;
            int rowidx = 0;
            if (prjinfo._Direction > 0)
            {
                rowidx = 2;
            }
            else
            {
                rowidx = 3;
            }
            _Worksheet.Cells[rowidx, 3] = 3.75;
            _Worksheet.Cells[rowidx, 5] = prjinfo._RoadCode;
            _Worksheet.Cells[rowidx, 6] = prjinfo._StartMile.ToString("K0+000");
            _Worksheet.Cells[rowidx, 7] = prjinfo._EndMile.ToString("K0+000");
        }
        private static void WriteZJGTPrj2CPMSXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[2, 2] = prjinfo._RoadCode;
            if (prjinfo._Direction > 0)
            {
                _Worksheet.Cells[2, 4] = "上行";
            }
            else
            {
                _Worksheet.Cells[2, 4] = "下行";
            }
            _Worksheet.Cells[2, 8] = prjinfo._DataDate;
            _Worksheet.Cells[3, 8] = prjinfo._StartMile.ToString("K0+000");
            _Worksheet.Cells[3, 13] = prjinfo._EndMile.ToString("K0+000");
        }
        //沥青和水泥分为两个xlxs表出
        private static void OutLqDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中交国通\015-沥青破损.xlsx",
              System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_沥青破损_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritelqDis2Xls(_Worksheet, prjinfo, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritelqDis2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            int len = dislist.Length, i = 0;
            if (len < 1)
                return;

            MSExcel.Range destrange;
            object[,] val = new object[len, 22];
            string[] s;
            string[] lqdis = { "龟裂.轻", "龟裂.中", "龟裂.重", "块状裂缝.轻", "块状裂缝.重", "纵向裂缝.轻", "纵向裂缝.重", "横向裂缝.轻", "横向裂缝.重", "坑槽.轻", "坑槽.重", "松散.轻", "松散.重", "沉陷.轻", "沉陷.重", "车辙.轻", "车辙.重", "波浪拥包.轻", "波浪拥包.重" };
            foreach (Disease tdis in dislist)
            {
                if (tdis.RoadType == "沥青")
                {
                    s = tdis.RoadDisType.Split('.');
                    val[i, 0] = tdis.m_mile;
                    val[i, 1] = prjinfo._RoadNum;
                    val[i, 2] = s[0];
                    if (s.Length > 1)
                    {
                        val[i, 3] = s[1];
                    }
                    else
                    {
                        val[i, 3] = "无";
                    }
                    val[i, 4] = tdis.rect.Height * _RoadConfig.HeightScale;
                    val[i, 5] = tdis.rect.Width * _RoadConfig.WidthScale;
                    val[i, 6] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                    val[i, 7] = tdis.Area;
                    val[i, 8] = tdis.calcheight;
                    val[i, 9] = tdis.calcwidth;
                    val[i, 10] = tdis.imgname;
                    val[i, 11] = tdis.imgpath;
                    ++i;
                }
            }
            destrange = _Worksheet.get_Range(String.Format("A3:L{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:L{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);
        }

        private static void WriteDis2xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            MSExcel.Range srcrange, destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int sn_tablerow = _Setting.cmop_rows;
            int lq_tablerow = _Setting.cmop_rows;

            int tcnt_sn = 0;
            int tcnt_lq = 0;

            int sn_csmile = 0, sn_cemile = 0;
            int lq_csmile = 0, lq_cemile = 0;
            bool sn_flag = false, lq_flag = false;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count - 1;
                    disval = new object[disnum, 1];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        disval[kk, 0] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        if (kk == disnum - 1)
                        {
                            disval[kk, 0] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di + 1].totalarea;
                        }
                    }
                    destrange = worksheet_snhz.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('F' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        sn_tablerow * tcnt_sn + 7,
                        sn_tablerow * tcnt_sn + 6 + disnum));
                    destrange.Value2 = disval;

                    sn_cemile = emile;
                    if (!sn_flag)
                    {
                        sn_flag = true;
                        sn_csmile = smile;
                    }
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count - 1;
                    disval = new object[disnum, 1];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        disval[kk, 0] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        if (kk == disnum - 1)
                        {
                            disval[kk, 0] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di + 1].totalarea;
                        }
                    }
                    destrange = worksheet_lqhz.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('F' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        lq_tablerow * tcnt_lq + 7,
                        lq_tablerow * tcnt_lq + 6 + disnum));
                    destrange.Value2 = disval;

                    lq_cemile = emile;
                    if (!lq_flag)
                    {
                        lq_flag = true;
                        lq_csmile = smile;
                    }
                }

                if (emile % xlslen == 0)
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 13] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:R{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("F{0}:O{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 13] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:R{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("F{0}:O{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }
                }
            }
            if (roadpart[len].mile % xlslen != 0)
            {
                if (sn_csmile != sn_cemile)
                {
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 13] = roadpart[len].mile;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 13] = roadpart[len].mile;
                }
            }
            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }
            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }

            RoadDiseaseTypes.Clear();
        }

        #endregion

        #region  奥路通报表模板
        public static void OutputALTDIS(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _Workbook = null;
            MSExcel.Worksheet _Worksheet = null;
            string srcxls;
            string Destxls = string.Format(@"{0}\{1}_路面病害模板.xls.xls", path, prjdir.Name);
            if (File.Exists(Destxls) && (disval != 10))
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\001-路面病害模板.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);

                Destxls = string.Format(@"{0}\{1}_路面病害_{2}m.xlsx", path, prjdir.Name, disval);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥路面病害汇总表"] as MSExcel.Worksheet;
            WriteALTDisHZTJ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList);
            string sheetname = "路面技术状况明细表";
            _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            WriteALTPQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\004-平台技术指标模板.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}-七项指标.xlsx", path, prjinfo._RoadCode);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteALT_PT_PQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, disval);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook _Workbook_LQ = null;
            MSExcel.Workbook _Workbook_SN = null;
            MSExcel.Worksheet _Worksheet_LQ = null;
            MSExcel.Worksheet _Worksheet_SN = null;

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\003-平台病害模板（沥青）",
                System.Windows.Forms.Application.StartupPath, disval);
            string LQDestxls = string.Format(@"{0}\{1}-病害（沥青）.xlsx", path, prjinfo._RoadCode);
            _Workbook_LQ = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook_LQ.SaveAs(LQDestxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_LQ = _Workbook_LQ.Sheets["Sheet1"] as MSExcel.Worksheet;

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\003-平台病害模板（水泥）",
                System.Windows.Forms.Application.StartupPath, disval);
            string SNDestxls = string.Format(@"{0}\{1}-病害（水泥）.xlsx", path, prjinfo._RoadCode);
            _Workbook_SN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook_SN.SaveAs(SNDestxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_SN = _Workbook_SN.Sheets["Sheet1"] as MSExcel.Worksheet;

            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            WriteALTDis_PT_HZTJ2Xls(_Worksheet_SN, _Worksheet_LQ, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, disval);

            _Workbook_LQ.Save();
            _Workbook_LQ.Close(Type.Missing, Type.Missing, Type.Missing);

            _Workbook_SN.Save();
            _Workbook_SN.Close(Type.Missing, Type.Missing, Type.Missing);

            if (!Haslqflag)
            {
                File.Delete(LQDestxls);
            }

            if (!Hassnflag)
            {
                File.Delete(SNDestxls);
            }

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteALTDisHZTJ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
           ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int rowcnt_sn_s = 5;
            int rowcnt_sn_e = 5;//小计起始的计算范围
            int rowcnt_lq_s = 5;
            int rowcnt_lq_e = 5;

            int totalsnlen = 0;//水泥路段总长度
            int totallqlen = 0;//沥青路段总长度

            worksheet_snhz.Cells[2, 20] = prjinfo._Direction > 0 ? "上行" : "下行";
            worksheet_lqhz.Cells[2, 20] = prjinfo._Direction > 0 ? "上行" : "下行";
            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 1;
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt] = string.Format("K{000:000+000}-K{1:000+000}", roadpart[i].mile, roadpart[i + 1].mile);

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 1];
                    for (int di = 0, kk = 1; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    disval[0, disnum - 1] = Convert.ToDouble(disval[0, disnum - 1]) + Convert.ToDouble(disval[0, disnum]);
                    disval[0, disnum] = null;
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);

                    disval[0, 0] = string.Format("=100-{0}*POWER({1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    destrange = worksheet_snhz.get_Range(string.Format("B{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('B' + disnum))));
                    destrange.Value2 = disval;

                    totalsnlen += milelength;
                    rowcnt_sn_s++;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt] = string.Format("K{000:000+000}-K{1:000+000}", roadpart[i].mile, roadpart[i + 1].mile);
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 1];
                    for (int di = 0, kk = 1; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    disval[0, disnum - 1] = Convert.ToDouble(disval[0, disnum - 1]) + Convert.ToDouble(disval[0, disnum]);
                    disval[0, disnum] = null;
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, 0] = string.Format("=100-{0}*POWER({1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);


                    destrange = worksheet_lqhz.get_Range(string.Format("B{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('B' + disnum))));
                    destrange.Value2 = disval;

                    totallqlen += milelength;
                    rowcnt_lq_s++;
                }

                if (emile % 1000 == 0)
                {
                    if (roadpart[i].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum - 1; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                        }
                        destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_sn_s++;
                        rowcnt_sn_e = rowcnt_sn_s;

                        if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                            worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum - 1; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                            }
                            destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_lq_s++;
                            rowcnt_lq_e = rowcnt_lq_s;
                        }
                    }
                    else if (roadpart[i].roadtype == 0)//沥青
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum - 1; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                        }
                        destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_lq_s++;
                        rowcnt_lq_e = rowcnt_lq_s;

                        if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                            worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum - 1; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                            }
                            destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_sn_s++;
                            rowcnt_sn_e = rowcnt_sn_s;
                        }
                    }
                }
            }

            //最后的一个小计
            if (roadpart[len].mile % 1000 != 0)
            {
                if (roadpart[len].roadtype == 1)
                {
                    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                    disval = new object[1, disnum];
                    for (int di = 0; di < disnum - 1; di++)
                    {
                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                    }
                    destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                    destrange.Value2 = disval;
                    rowcnt_sn_s++;
                    rowcnt_sn_e = rowcnt_sn_s;

                    if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum - 1; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                        }
                        destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_lq_s++;
                        rowcnt_lq_e = rowcnt_lq_s;
                    }
                }
                else if (roadpart[len].roadtype == 0)
                {
                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                    disval = new object[1, disnum];
                    for (int di = 0; di < disnum - 1; di++)
                    {
                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                    }
                    destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                    destrange.Value2 = disval;
                    rowcnt_lq_s++;
                    rowcnt_lq_e = rowcnt_lq_s;

                    if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum - 1; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                        }
                        destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_sn_s++;
                        rowcnt_sn_e = rowcnt_sn_s;
                    }
                }
            }

            //总计
            //水泥
            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0);
            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum - 1; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_s - 1);
            }
            destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
            destrange.Value2 = disval;

            //沥青
            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum - 1; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_s - 1);
            }
            destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
            destrange.Value2 = disval;

            destrange = worksheet_lqhz.get_Range(String.Format("A1:W{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet_snhz.get_Range(String.Format("A1:V{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, 53);

            RoadDiseaseTypes.Clear();
            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }

            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }
        }

        private static void WriteALTDis_PT_HZTJ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
           ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, ref bool Haslqflag, ref bool Hassnflag, int unitlen)
        {
            MSExcel.Range destrange;
            object[,] disval = null;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志

            int rowcnt_sn_s = 1;
            int rowcnt_lq_s = 1;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            string[] XlsDisNamesLQ = { "车辙.轻", "车辙.重",
                                       "横向裂缝.轻", "横向裂缝.重",
                                       "纵向裂缝.轻", "纵向裂缝.重",
                                       "坑槽.轻", "坑槽.重",
                                       "松散.轻", "松散.重",
                                       "沉陷.轻", "沉陷.重",
                                       "波浪拥包.轻", "波浪拥包.重",
                                       "泛油", "修补",
                                       "龟裂.轻", "龟裂.中", "龟裂.重",
                                       "块状裂缝.轻", "块状裂缝.重"};

            string[] XlsDisNamesSN = { "破碎板.轻", "破碎板.重",
                                       "裂缝.轻", "裂缝.中", "裂缝.重",
                                       "板角断裂.轻", "板角断裂.中", "板角断裂.重",
                                       "错台.轻", "错台.重",
                                       "拱起",
                                       "边角剥落.轻", "边角剥落.中", "边角剥落.重",
                                       "接缝料损坏.轻", "接缝料损坏.重",
                                       "坑洞",
                                       "唧泥",
                                       "露骨",
                                       "修补"};

            int typeidx = 0;
            bool res = false;

            int disnum = 0;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0, pcival = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    j++;
                }

                //病害汇总表
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;

                    disval = new object[1, 26];
                    disval[0, 0] = rowcnt_sn_s;
                    disval[0, 1] = string.Format("K{0:0+000}-K{1:0+000}", roadpart[i].mile, roadpart[i + 1].mile);

                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disval[0, 2] = pcival;

                    for (int kk = 0; kk < XlsDisNamesSN.Length; ++kk)
                    {
                        disval[0, 3 + kk] = 0;
                        for (int di = 0; di < disnum; ++di)
                        {
                            if (RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains(XlsDisNamesSN[kk]))
                            {
                                disval[0, 3 + kk] = Convert.ToDouble(disval[0, 3 + kk]) + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                            }
                        }
                    }
                    disval[0, 3 + XlsDisNamesSN.Length] = milelength;
                    disval[0, 3 + XlsDisNamesSN.Length + 1] = _Setting.DistrictCode;
                    disval[0, 3 + XlsDisNamesSN.Length + 2] = roadpart[i].degreestr;

                    destrange = worksheet_snhz.get_Range(string.Format("A{0}:Z{0}", rowcnt_sn_s + 7));
                    destrange.Value2 = disval;

                    rowcnt_sn_s++;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;

                    disval = new object[1, 27];
                    disval[0, 0] = rowcnt_lq_s;
                    disval[0, 1] = string.Format("K{0:0+000}-K{1:0+000}", roadpart[i].mile, roadpart[i + 1].mile);

                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disval[0, 2] = pcival;

                    for (int kk = 0; kk < XlsDisNamesLQ.Length; ++kk)
                    {
                        disval[0, 3 + kk] = 0;
                        for (int di = 0; di < disnum; ++di)
                        {
                            if (RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains(XlsDisNamesLQ[kk]))
                            {
                                disval[0, 3 + kk] = Convert.ToDouble(disval[0, 3 + kk]) + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                            }
                        }
                    }
                    disval[0, 3 + XlsDisNamesLQ.Length] = milelength;
                    disval[0, 3 + XlsDisNamesLQ.Length + 1] = _Setting.DistrictCode;
                    disval[0, 3 + XlsDisNamesLQ.Length + 2] = roadpart[i].degreestr;

                    destrange = worksheet_lqhz.get_Range(string.Format("A{0}:AA{0}", rowcnt_lq_s + 7));
                    destrange.Value2 = disval;

                    rowcnt_lq_s++;
                }
            }

            destrange = worksheet_lqhz.get_Range(String.Format("A8:Z{0}", rowcnt_lq_s + 6));
            GlobalExcel.SetBorderLine(destrange, 53);

            destrange = worksheet_snhz.get_Range(String.Format("A8:AA{0}", rowcnt_sn_s + 6));
            GlobalExcel.SetBorderLine(destrange, 53);

            worksheet_lqhz.Cells[1, 2] = unitlen;
            worksheet_lqhz.Cells[1, 4] = prjinfo._RoadCode.Substring(0, 4);
            worksheet_lqhz.Cells[1, 6] = _Setting.DetectYear;
            worksheet_lqhz.Cells[2, 2] = "沥青路面";
            worksheet_lqhz.Cells[2, 4] = prjinfo._Direction > 0 ? "1-上行" : "2-下行";
            worksheet_lqhz.Cells[2, 6] = _Setting.DetectNum;
            worksheet_lqhz.Cells[3, 2] = prjinfo._RoadNum.Replace("车道", "") + "车道";
            worksheet_lqhz.Cells[3, 4] = prjinfo._DataDate.Insert(6, "/").Insert(4, "/");

            worksheet_snhz.Cells[1, 2] = unitlen;
            worksheet_snhz.Cells[1, 4] = prjinfo._RoadCode.Substring(0, 4);
            worksheet_snhz.Cells[1, 6] = _Setting.DetectYear;
            worksheet_snhz.Cells[2, 2] = "水泥路面";
            worksheet_snhz.Cells[2, 4] = prjinfo._Direction > 0 ? "1-上行" : "2-下行";
            worksheet_snhz.Cells[2, 6] = _Setting.DetectNum;
            worksheet_snhz.Cells[3, 2] = prjinfo._RoadNum.Replace("车道", "") + "车道";
            worksheet_snhz.Cells[3, 4] = prjinfo._DataDate.Insert(6, "/").Insert(4, "/");
        }

        private static void WriteALT_PT_PQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal, int disval)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0;
            bool haslq = false; //是否存在沥青路面
            bool hassn = false; //是否存在水泥路面

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 13];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                vallist[rowcnt, 0] = i + 1;
                vallist[rowcnt, 1] = smile;
                vallist[rowcnt, 2] = emile;
                vallist[rowcnt, 10] = string.Format("=ABS(B{0}-C{0})", i + 6);

                // DR PCI
                if (roadpart[i].roadtype == 1)
                {
                    hassn = true;
                }
                else
                {
                    haslq = true;
                }
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[rowcnt, 3] = drval;
                vallist[rowcnt, 6] = string.Format("=100-{0}*POWER(D{1},{2})",
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    i + 6,
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                //IRI RQI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, 4] = Math.Round(irival, 5);
                vallist[rowcnt, 7] = string.Format("=ROUND(100/(1+{0}*EXP({1}*E{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    i + 6);

                //Rut
                if (prjinfo._IsRut)
                {
                    //rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    vallist[rowcnt, 5] = Math.Round(rutval, 5);
                    vallist[rowcnt, 8] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
                        i + 6,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                }

                if (roadpart[i].roaddegree <= 1)
                {
                    vallist[rowcnt, 9] = string.Format("=ROUND(({1}*G{0}+{2}*H{0}+{3}*I{0})/({1}+{2}+{3}),5)",
                           i + 6,
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                }
                else
                {
                    vallist[rowcnt, 9] = string.Format("=ROUND(({1}*G{0}+{2}*H{0})/({1}+{2}),5)",
                            i + 6,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }

                vallist[rowcnt, 11] = _Setting.DistrictCode;
                vallist[rowcnt, 12] = roadpart[i].degreestr;
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A6:M{0}", rowcnt + 5));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            worksheet.Cells[1, 2] = disval;
            worksheet.Cells[1, 4] = prjinfo._RoadCode.Substring(0, 4);
            worksheet.Cells[1, 6] = _Setting.DetectYear;

            if (haslq & hassn) worksheet.Cells[2, 2] = "混合路面";
            else if (haslq) worksheet.Cells[2, 2] = "沥青路面";
            else worksheet.Cells[2, 2] = "水泥路面";

            worksheet.Cells[2, 4] = prjinfo._Direction > 0 ? "1-上行" : "2-下行";
            worksheet.Cells[2, 6] = _Setting.DetectNum;

            worksheet.Cells[3, 2] = prjinfo._RoadNum.Replace("车道", "") + "车道";
            worksheet.Cells[3, 4] = prjinfo._DataDate.Insert(6, "/").Insert(4, "/");

            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(worksheet, 6, 2, 12, true);
            //    GlobalExcel.Reflection(worksheet, 6, 2, 2, false);
            //}
        }

        private static void WriteALTPQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0, wrval = 0;
            worksheet.Cells[2, 12] = prjinfo._RoadNum;
            worksheet.Cells[2, 2] = prjinfo._RoadName;
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 16];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 0;
                //  
                vallist[rowcnt, colcnt] = string.Format("K{0:0000+000}-K{1:0000+000}", roadpart[i].mile, roadpart[i + 1].mile);
                vallist[rowcnt, colcnt + 1] = milelength;

                // DR PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt + 2] = drval;
                vallist[rowcnt, colcnt + 9] = string.Format("=100-{0}*POWER(C{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 5, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                //IRI RQI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, colcnt + 3] = Math.Round(irival, 5);
                vallist[rowcnt, colcnt + 10] = string.Format("=ROUND(100/(1+{0}*EXP({1}*D{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 5);

                //Rut
                if (prjinfo._IsRut)
                {
                    //rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    vallist[rowcnt, colcnt + 4] = Math.Round(rutval, 5);
                    vallist[rowcnt, colcnt + 11] = string.Format("=IF(E{0}<{1},{2}-{3}*E{0},IF(E{0}<{4},{5}-{6}*(E{0}-{1}),0))",
                        i + 5,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                }

                #region  PBI
                //各程度跳车数
                vallist[rowcnt, colcnt + 5] = _PBIVal[i][1];
                vallist[rowcnt, colcnt + 6] = _PBIVal[i][2];
                vallist[rowcnt, colcnt + 7] = _PBIVal[i][3];
                //PBI得分与评价
                int datacol = 5;
                vallist[rowcnt, colcnt + 12] = string.Format("=IF((100-{4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3})>0,(100- {4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3}),0)",
                i + 5, _PBIScore[1], _PBIScore[2], _PBIScore[3], (char)('A' + datacol), (char)('A' + datacol + 1), (char)('A' + datacol + 2));

                //PWI
                wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                if (!((wrval >= 0) && (wrval <= 100)))
                {
                    wrval = 0;
                }
                if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))   // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                {
                    vallist[rowcnt, colcnt + 8] = 0;
                }
                else
                {
                    vallist[rowcnt, colcnt + 8] = wrval;
                }
                vallist[rowcnt, colcnt + 13] = string.Format("=IF(100-{0}*POWER(I{1},{2})>100,100,100-{0}*POWER(I{1},{2}))", _PWIa[0], i + 5, _PWIa[1]);
                #endregion

                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        vallist[rowcnt, colcnt + 14] = string.Format("=ROUND(({1}*J{0}+{2}*K{0}+{3}*IF(EXACT(L{0},\"-\"),0,L{0})+{4}*IF(EXACT(M{0},\"-\"),0,M{0})))/({1}+{2}+{3}+{4}),5)",
                                i + 5,
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        vallist[rowcnt, colcnt + 14] = string.Format("=ROUND(({1}*J{0}+{2}*K{0}+{3}*IF(EXACT(L{0},\"-\"),0,L{0})+{4}*IF(EXACT(M{0},\"-\"),0,M{0})+{5}*IF(EXACT(N{0},\"-\"),0,N{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                                i + 5,
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                    }
                }
                else
                {
                    vallist[rowcnt, colcnt + 14] = string.Format("=ROUND(({1}*J{0}+{2}*K{0})/({1}+{2}),5)",
                            i + 5,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                vallist[rowcnt, colcnt + 15] = roadpart[i].degreestr;
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:P{0}", rowcnt + 4));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            //if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            //{
            //    destrange = worksheet.get_Range(string.Format("A5:P{0}", rowcnt + 4));
            //    MSExcel.Range sortrange = worksheet.get_Range(string.Format("A5:A{0}", rowcnt + 4));
            //    GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            //}
            int chartlen = len + 4;
            MSExcel.ChartObject chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            MSExcel.Chart chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0},K3:K{0}, J3:J{0}, L3:L{0}, M3:M{0}, N3:N{0}, O3:O{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, true, "", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(2);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, C3:C{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "破损率DR(%)", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(3);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, D3:D{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "IRI", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(4);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, F4:F{0},G4:G{0},H4:H{0}", len + 2));  //F -H
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "PB", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(5);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, E3:E{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "RD", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(6);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, I3:I{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "WR", Type.Missing, Type.Missing, Type.Missing);
        }


        #endregion

        #region 中南安环
        public static void OutputZNRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中南安环\路况信息综合表2.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{3}_{2}m.xlsx", path, prjdir.Name, disval, "路况信息综合表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_tj = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetIRI = _Workbook.Sheets["平整度"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetRUT = _Workbook.Sheets["车辙"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetDR = _Workbook.Sheets["破损率"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetPB = _Workbook.Sheets["跳车"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetPW = _Workbook.Sheets["磨耗"] as MSExcel.Worksheet;
            WriteZNHZ2Xls2(_Worksheet_hz, _Worksheet_tj, sheetIRI, sheetRUT, sheetDR, sheetPB, sheetPW,
                prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _PBIVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZNHZ2Xls(MSExcel.Worksheet worksheet, MSExcel.Worksheet worksheet2,
            MSExcel.Worksheet sheetIRI, MSExcel.Worksheet sheetRUT, MSExcel.Worksheet sheetDR, MSExcel.Worksheet sheetPB, MSExcel.Worksheet sheetPW,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal, int[][] PBIVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal)
        {
            worksheet.Cells[1, 1] = prjinfo._RoadCode + prjinfo._RoadName
                + prjinfo._StartMile.ToString("K0+000") + "~"
                + prjinfo._EndMile.ToString("K0+000") + "段\r\n路面使用性能指数评定汇总表";

            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, tpcival = 0;


            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 4];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];
            object[,] PWIObj = new object[len, 1];
            object[,] PBObj = new object[len, 3];
            object[,] PBIObj = new object[len, 1];

            object[,] mxlist = new object[len, 21];
            string errlog = prjdir.FullName + "\\errlog.txt";

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double repair = 0, ksumarea = 0;
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                mxlist[i, 0] = smile;
                mxlist[i, 1] = emile;
                mxlist[i, 2] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                mxlist[i, 3] = String.Format("=ABS(A{0}-B{0})", i + 4);
                mxlist[i, 19] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 20] = roadpart[i].degreestr.Replace("公路", "");

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;

                        if (arrdis[j].RoadDisType.Contains("修补"))
                        {
                            repair += arrdis[j].Area;
                        }
                        ksumarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                mxlist[i, 6] = drval;
                mxlist[i, 7] = string.Format("=100-{0}*POWER(G{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 4, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 17] = repair > 0 ? Math.Round(repair * 100 / ksumarea, 5) : 0;

                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                mxlist[i, 4] = irival;
                mxlist[i, 5] = String.Format("=ROUND(100/(1+{0}*EXP({1}*E{2})),5)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 4);

                //Rut
                if (prjinfo._IsRut)
                {
                    double rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    mxlist[i, 8] = rutval;
                    mxlist[i, 9] = string.Format("=IF(I{0}<{1},{2}-{3}*I{0},IF(I{0}<{4},{5}-{6}*(I{0}-{1}),0))",
                        i + 4,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                }

                mxlist[i, 10] = PBIVal[i][1];
                mxlist[i, 11] = PBIVal[i][2];
                mxlist[i, 12] = PBIVal[i][3];
                mxlist[i, 13] = string.Format("=IF((100-K{0}*{1}-L{0}*{2}-M{0}*{3})>0,(100-K{0}*{1}-L{0}*{2}-M{0}*{3}),0)",
                    i + 4, _PBIScore[1], _PBIScore[2], _PBIScore[3]);

                //构造深度相关         
                double wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];

                wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))  // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                {
                    mxlist[i, 14] = 0;
                }
                else
                {
                    mxlist[i, 14] = wrval;
                }
                mxlist[i, 15] = string.Format("=100-{0}*POWER(O{1},{2})", _PWIa[0], i + 4, _PWIa[1]);

                //PQI
                if (roadpart[i].roaddegree < 2)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        mxlist[i, 16] = string.Format("=ROUND(({1}*H{0}+{2}*F{0}+{3}*IF(EXACT(J{0},\" \"),0,J{0})+{4}*IF(EXACT(N{0},\"-\"),0,N{0}))/({1}+{2}+{3}+{4}),5)",
                                         i + 4,
                                         _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        mxlist[i, 16] = string.Format("=ROUND(({1}*H{0}+{2}*F{0}+{3}*IF(EXACT(J{0},\" \"),0,J{0})+{4}*IF(EXACT(N{0},\"-\"),0,N{0})+{5}*IF(EXACT(P{0},\"-\"),0,P{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                                        i + 4,
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                    }
                }
                else
                {
                    mxlist[i, 16] = string.Format("=ROUND(({1}*H{0}+{2}*F{0})/({1}+{2}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 18] = string.Format("=IF(Q{0}>={1},\"优\",IF(Q{0}>={2},\"良\",IF(Q{0}>={3},\"中\",IF(Q{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
            }

            int datarow = len + 3;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:U{0}", datarow));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 63);

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("A4:U{0}", datarow));
                MSExcel.Range sortrange = worksheet.get_Range(String.Format("A4:A{0}", datarow));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
                GlobalExcel.Reflection(worksheet, 4, 1, 2, false);
            }

            //汇总部分
            MSExcel.Range srcrange = worksheet2.get_Range("A1:Z10");
            destrange = worksheet.get_Range(String.Format("A{0}", datarow + 1));
            srcrange.Copy(destrange);

            worksheet.Cells[datarow + 1, 4] = string.Format("=SUM(D4:D{0})", datarow);
            for (int i = 0; i < 14; i++)
            {
                worksheet.Cells[datarow + 1, 5 + i] = string.Format("=SUMPRODUCT(D4:D{1},{0}4:{0}{1})/SUM(D4:D{1})", (char)('E' + i), datarow);
            }

            worksheet.Cells[datarow + 1, 19] = string.Format("=IF(Q{0}>={1},\"优\",IF(Q{0}>={2},\"良\",IF(Q{0}>={3},\"中\",IF(Q{0}>={4},\"次\",\"差\"))))",
                    datarow + 1,
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            // 优良中次差 个数统计
            //RQI   _RQIGrade;//道路等级 路面材质 等级区间
            string roadtype = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            worksheet.Cells[datarow + 3, 5] = string.Format("=SUMIF(F4:F{0},\">={1}\",D4:D{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][0]);
            worksheet.Cells[datarow + 4, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][1], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][0]);
            worksheet.Cells[datarow + 5, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][2], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][1]);
            worksheet.Cells[datarow + 6, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][3], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][2]);
            worksheet.Cells[datarow + 7, 5] = string.Format("=SUMIF(F4:F{0},\"<{1}\",D4:D{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][3]);
            //PCI
            worksheet.Cells[datarow + 3, 9] = string.Format("=SUMIF(H4:H{0},\">={1}\",D4:D{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 9] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 9] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 9] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 9] = string.Format("=SUMIF(H4:H{0},\"<{1}\",D4:D{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //RDI
            worksheet.Cells[datarow + 3, 13] = string.Format("=SUMIF(J4:J{0},\">={1}\",D4:D{0})/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 13] = string.Format("=SUMIFS(D4:D{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 13] = string.Format("=SUMIFS(D4:D{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 13] = string.Format("=SUMIFS(D4:D{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 13] = string.Format("=SUMIF(J4:J{0},\"<{1}\",D4:D{0})/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //PB
            worksheet.Cells[datarow + 3, 17] = string.Format("=SUMIF(N4:N{0},\">={1}\",D4:D{0})/1000", datarow,
                 _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 17] = string.Format("=SUMIFS(D4:D{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 17] = string.Format("=SUMIFS(D4:D{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 17] = string.Format("=SUMIFS(D4:D{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 17] = string.Format("=SUMIF(N4:N{0},\"<{1}\",D4:D{0})/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //PW
            worksheet.Cells[datarow + 3, 21] = string.Format("=SUMIF(P4:P{0},\">={1}\",D4:D{0})/1000", datarow,
                 _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 21] = string.Format("=SUMIFS(D4:D{0},P4:P{0},\">={1}\",P4:P{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 21] = string.Format("=SUMIFS(D4:D{0},P4:P{0},\">={1}\",P4:P{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 21] = string.Format("=SUMIFS(D4:D{0},P4:P{0},\">={1}\",P4:P{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 21] = string.Format("=SUMIF(P4:P{0},\"<{1}\",D4:D{0})/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);
            //PQI
            worksheet.Cells[datarow + 3, 25] = string.Format("=SUMIF(Q4:Q{0},\">={1}\",D4:D{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 25] = string.Format("=SUMIFS(D4:D{0},Q4:Q{0},\">={1}\",Q4:Q{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 25] = string.Format("=SUMIFS(D4:D{0},Q4:Q{0},\">={1}\",Q4:Q{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 25] = string.Format("=SUMIFS(D4:D{0},Q4:Q{0},\">={1}\",Q4:Q{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 25] = string.Format("=SUMIF(Q4:Q{0},\"<{1}\",Q4:Q{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //计算百分比
            for (int i = 0; i < 5; ++i)
            {
                worksheet.Cells[datarow + 3 + i, 6] = string.Format("=E{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 10] = string.Format("=I{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 14] = string.Format("=M{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 18] = string.Format("=Q{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 22] = string.Format("=U{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 26] = string.Format("=Y{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
            }

            //将数值复制进去
            srcrange = worksheet.get_Range(string.Format("A4:B{0}", datarow));
            destrange = sheetIRI.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetPB.get_Range("A11");
            srcrange.Copy(destrange);
            destrange = sheetPW.get_Range("A10");
            srcrange.Copy(destrange);

            //工程名称
            string prjname = string.Format("{0}-{1}", prjinfo._RoadCode, prjinfo._RoadName);
            sheetIRI.Cells[7, 3] = prjname;
            sheetRUT.Cells[7, 3] = prjname;
            sheetDR.Cells[7, 3] = prjname;
            sheetPB.Cells[7, 3] = prjname;
            sheetPW.Cells[7, 3] = prjname;

            //IRI
            srcrange = worksheet.get_Range(string.Format("E4:F{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetIRI.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetIRI.get_Range("C10");
            }
            srcrange.Copy(destrange);
            destrange = sheetIRI.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //RUT
            srcrange = worksheet.get_Range(string.Format("I4:J{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetRUT.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetRUT.get_Range("C10");
            }
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PCI
            srcrange = worksheet.get_Range(string.Format("G4:H{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetDR.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetDR.get_Range("C10");
            }
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PB
            srcrange = worksheet.get_Range(string.Format("K4:N{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPB.get_Range("G11");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPB.get_Range("C11");
            }
            srcrange.Copy(destrange);
            destrange = sheetPB.get_Range(string.Format("A11:K{0}", datarow + 7));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PW
            srcrange = worksheet.get_Range(string.Format("O4:P{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPW.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPW.get_Range("C10");
            }
            srcrange.Copy(destrange);
            destrange = sheetPW.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);
        }

        private static void WriteZNHZ2Xls2(MSExcel.Worksheet worksheet, MSExcel.Worksheet worksheet2,
            MSExcel.Worksheet sheetIRI, MSExcel.Worksheet sheetRUT, MSExcel.Worksheet sheetDR, MSExcel.Worksheet sheetPB, MSExcel.Worksheet sheetPW,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal, int[][] PBIVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal)
        {
            worksheet.Cells[1, 1] = prjinfo._RoadCode + prjinfo._RoadName
                + prjinfo._StartMile.ToString("K0+000") + "~"
                + prjinfo._EndMile.ToString("K0+000") + "段\r\n路面使用性能指数评定汇总表";

            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, tpcival = 0;

            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 4];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];
            object[,] PWIObj = new object[len, 1];
            object[,] PBObj = new object[len, 3];
            object[,] PBIObj = new object[len, 1];

            object[,] mxlist = new object[len, 24];
            string errlog = prjdir.FullName + "\\errlog.txt";

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double repair = 0, ksumarea = 0;
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                mxlist[i, 0] = _Setting.DutyUnit;
                mxlist[i, 1] = prjinfo._RoadCode;
                mxlist[i, 2] = prjinfo._RoadName;
                mxlist[i, 3] = smile;
                mxlist[i, 4] = emile;
                mxlist[i, 5] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 6] = roadpart[i].degreestr.Replace("公路", "");
                mxlist[i, 7] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                mxlist[i, 8] = String.Format("=ABS(D{0}-E{0})", i + 4);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;

                        if (arrdis[j].RoadDisType.Contains("修补"))
                        {
                            repair += arrdis[j].Area;
                        }
                        ksumarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                mxlist[i, 15] = drval;
                mxlist[i, 10] = string.Format("=100-{0}*POWER(P{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 4, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 18] = repair > 0 ? Math.Round(repair * 100 / ksumarea, 5) : 0;

                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                mxlist[i, 16] = irival;
                mxlist[i, 11] = String.Format("=ROUND(100/(1+{0}*EXP({1}*Q{2})),5)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 4);

                //Rut
                if (prjinfo._IsRut && roadpart[i].roaddegree < 2)
                {
                    double rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    mxlist[i, 17] = rutval;
                    mxlist[i, 12] = string.Format("=IF(R{0}<{1},{2}-{3}*R{0},IF(R{0}<{4},{5}-{6}*(R{0}-{1}),0))",
                        i + 4,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                }
                else
                {
                    mxlist[i, 17] = "/";
                    mxlist[i, 12] = "/";
                }

                //PQI
                if (roadpart[i].roaddegree < 2)
                {
                    mxlist[i, 19] = PBIVal[i][1];
                    mxlist[i, 20] = PBIVal[i][2];
                    mxlist[i, 21] = PBIVal[i][3];
                    mxlist[i, 13] = string.Format("=IF((100-T{0}*{1}-U{0}*{2}-V{0}*{3})>0,(100-T{0}*{1}-U{0}*{2}-V{0}*{3}),0)",
                        i + 4, _PBIScore[1], _PBIScore[2], _PBIScore[3]);

                    //构造深度相关         
                    double wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];

                    wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                    if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))  // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                    {
                        mxlist[i, 22] = 0;
                    }
                    else
                    {
                        mxlist[i, 22] = wrval;
                    }
                    mxlist[i, 14] = string.Format("=100-{0}*POWER(W{1},{2})", _PWIa[0], i + 4, _PWIa[1]);

                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        mxlist[i, 9] = string.Format("=ROUND(({1}*K{0}+{2}*L{0}+{3}*IF(EXACT(M{0},\" \"),0,M{0})+{4}*IF(EXACT(N{0},\"-\"),0,N{0}))/({1}+{2}+{3}+{4}),5)",
                                         i + 4,
                                         _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        mxlist[i, 9] = string.Format("=ROUND(({1}*K{0}+{2}*L{0}+{3}*IF(EXACT(M{0},\" \"),0,M{0})+{4}*IF(EXACT(N{0},\"-\"),0,N{0})+{5}*IF(EXACT(O{0},\"-\"),0,O{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                                        i + 4,
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                    }
                }
                else
                {
                    mxlist[i, 19] = "/";
                    mxlist[i, 20] = "/";
                    mxlist[i, 21] = "/";
                    mxlist[i, 13] = "/";
                    mxlist[i, 22] = "/";
                    mxlist[i, 14] = "/";

                    mxlist[i, 9] = string.Format("=ROUND(({1}*K{0}+{2}*L{0})/({1}+{2}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 23] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
            }

            int datarow = len + 3;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:X{0}", datarow));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 63);

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("A4:X{0}", datarow));
                MSExcel.Range sortrange = worksheet.get_Range(String.Format("D4:D{0}", datarow));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
                GlobalExcel.Reflection(worksheet, 4, 4, 2, false);
            }

            //汇总部分
            MSExcel.Range srcrange = worksheet2.get_Range("A1:Z10");
            destrange = worksheet.get_Range(String.Format("B{0}", datarow + 1));
            srcrange.Copy(destrange);

            worksheet.Cells[datarow + 1, 9] = string.Format("=SUM(I4:I{0})", datarow);
            for (int i = 0; i < 14; i++)
            {
                worksheet.Cells[datarow + 1, 10 + i] = string.Format("=SUMPRODUCT(I4:I{1},{0}4:{0}{1})/SUM(I4:I{1})", (char)('J' + i), datarow);
            }

            worksheet.Cells[datarow + 1, 24] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    datarow + 1,
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            // 优良中次差 个数统计
            //PQI
            worksheet.Cells[datarow + 3, 6] = string.Format("=SUMIF(J4:J{0},\">={1}\",I4:I{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 6] = string.Format("=SUMIFS(I4:I{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 6] = string.Format("=SUMIFS(I4:I{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 6] = string.Format("=SUMIFS(I4:I{0},J4:J{0},\">={1}\",J4:J{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 6] = string.Format("=SUMIF(J4:J{0},\"<{1}\",I4:I{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //PCI
            worksheet.Cells[datarow + 3, 10] = string.Format("=SUMIF(K4:K{0},\">={1}\",I4:I{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 10] = string.Format("=SUMIFS(I4:I{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 10] = string.Format("=SUMIFS(I4:I{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 10] = string.Format("=SUMIFS(I4:I{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 10] = string.Format("=SUMIF(K4:K{0},\"<{1}\",I4:I{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //RQI   _RQIGrade;//道路等级 路面材质 等级区间
            string roadtype = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            worksheet.Cells[datarow + 3, 14] = string.Format("=SUMIF(L4:L{0},\">={1}\",I4:I{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][0]);
            worksheet.Cells[datarow + 4, 14] = string.Format("=SUMIFS(I4:I{0},L4:L{0},\">={1}\",L4:L{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][1], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][0]);
            worksheet.Cells[datarow + 5, 14] = string.Format("=SUMIFS(I4:I{0},L4:L{0},\">={1}\",L4:L{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][2], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][1]);
            worksheet.Cells[datarow + 6, 14] = string.Format("=SUMIFS(I4:I{0},L4:L{0},\">={1}\",L4:L{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][3], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][2]);
            worksheet.Cells[datarow + 7, 14] = string.Format("=SUMIF(L4:L{0},\"<{1}\",I4:I{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][RoadDiseaseTypes.roadtypedict[roadtype]][3]);

            //RDI
            worksheet.Cells[datarow + 3, 18] = string.Format("=SUMIF(M4:M{0},\">={1}\",I4:I{0})/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 18] = string.Format("=SUMIFS(I4:I{0},M4:M{0},\">={1}\",M4:M{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 18] = string.Format("=SUMIFS(I4:I{0},M4:M{0},\">={1}\",M4:M{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 18] = string.Format("=SUMIFS(I4:I{0},M4:M{0},\">={1}\",M4:M{0},\"<{2}\")/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 18] = string.Format("=SUMIF(M4:M{0},\"<{1}\",I4:I{0})/1000", datarow,
                _RDIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //PBI
            worksheet.Cells[datarow + 3, 22] = string.Format("=SUMIF(N4:N{0},\">={1}\",I4:I{0})/1000", datarow,
                 _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 22] = string.Format("=SUMIFS(I4:I{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 22] = string.Format("=SUMIFS(I4:I{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 22] = string.Format("=SUMIFS(I4:I{0},N4:N{0},\">={1}\",N4:N{0},\"<{2}\")/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 22] = string.Format("=SUMIF(N4:N{0},\"<{1}\",I4:I{0})/1000", datarow,
                _PBIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //PWI
            worksheet.Cells[datarow + 3, 26] = string.Format("=SUMIF(O4:O{0},\">={1}\",I4:I{0})/1000", datarow,
                 _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 26] = string.Format("=SUMIFS(I4:I{0},O4:O{0},\">={1}\",O4:O{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 26] = string.Format("=SUMIFS(I4:I{0},O4:O{0},\">={1}\",O4:O{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 26] = string.Format("=SUMIFS(I4:I{0},O4:O{0},\">={1}\",O4:O{0},\"<{2}\")/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 26] = string.Format("=SUMIF(O4:O{0},\"<{1}\",I4:I{0})/1000", datarow,
                _PWIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            //计算百分比
            for (int i = 0; i < 5; ++i)
            {
                worksheet.Cells[datarow + 3 + i, 7] = string.Format("=F{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 11] = string.Format("=J{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 15] = string.Format("=N{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 19] = string.Format("=R{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 23] = string.Format("=V{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 27] = string.Format("=Z{0}/I{1}*100000", datarow + 3 + i, datarow + 1);
            }

            //将数值复制进去
            srcrange = worksheet.get_Range(string.Format("D4:E{0}", datarow));
            destrange = sheetIRI.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetPB.get_Range("A11");
            srcrange.Copy(destrange);
            destrange = sheetPW.get_Range("A10");
            srcrange.Copy(destrange);

            //工程名称
            string prjname = string.Format("{0}-{1}", prjinfo._RoadCode, prjinfo._RoadName);
            sheetIRI.Cells[7, 3] = prjname;
            sheetRUT.Cells[7, 3] = prjname;
            sheetDR.Cells[7, 3] = prjname;
            sheetPB.Cells[7, 3] = prjname;
            sheetPW.Cells[7, 3] = prjname;

            //IRI
            srcrange = worksheet.get_Range(string.Format("Q4:Q{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetIRI.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetIRI.get_Range("C10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            srcrange = worksheet.get_Range(string.Format("L4:L{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetIRI.get_Range("F10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetIRI.get_Range("D10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            destrange = sheetIRI.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //RUT
            srcrange = worksheet.get_Range(string.Format("R4:R{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetRUT.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetRUT.get_Range("C10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            srcrange = worksheet.get_Range(string.Format("M4:M{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetRUT.get_Range("F10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetRUT.get_Range("D10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            destrange = sheetRUT.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PCI
            srcrange = worksheet.get_Range(string.Format("P4:P{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetDR.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetDR.get_Range("C10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            srcrange = worksheet.get_Range(string.Format("K4:K{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetDR.get_Range("F10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetDR.get_Range("D10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            destrange = sheetDR.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PB
            srcrange = worksheet.get_Range(string.Format("T4:V{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPB.get_Range("G11");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPB.get_Range("C11");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            srcrange = worksheet.get_Range(string.Format("N4:N{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPB.get_Range("J11");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPB.get_Range("F11");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            destrange = sheetPB.get_Range(string.Format("A11:K{0}", datarow + 7));
            GlobalExcel.SetBorderLine(destrange, 63);

            //PW
            srcrange = worksheet.get_Range(string.Format("W4:W{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPW.get_Range("E10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPW.get_Range("C10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            srcrange = worksheet.get_Range(string.Format("O4:O{0}", datarow));
            if (prjinfo._Direction > 0)
            {
                destrange = sheetPW.get_Range("F10");
            }
            else if (prjinfo._Direction < 0)
            {
                destrange = sheetPW.get_Range("D10");
            }
            srcrange.Copy();
            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
            destrange = sheetPW.get_Range(string.Format("A10:G{0}", datarow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);
        }

        //原始数据 RQI PCI RDI PBI PWI
        public static void OutputZNDataRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中南安环\原始数据记录表.xlsx",
                           System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}_{3}m.xlsx", path, prjdir.Name, "原始数据记录表", disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_RQI = _Workbook.Sheets["RQI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PCI = _Workbook.Sheets["PCI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_RDI = _Workbook.Sheets["RDI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PBI = _Workbook.Sheets["PBI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PWI = _Workbook.Sheets["PWI"] as MSExcel.Worksheet;

            WriteIRI2Xls(_Worksheet_RQI, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet_RQI, 4, 3, 22, 'H', "平整度", 1);

            WriteRut2Xls_orirut(_Worksheet_RDI, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet_RDI, 4, 3, 22, 'H', "车辙深度", 1);

            WritePCI2Xls(_Worksheet_PCI, prjinfo, prjdir, _RoadPart, _RoadDisList, _SpeedVal, _MarkVal, 3);
            WriteStatistics_XMJH(_Worksheet_PCI, 3, 3, 22, 'F', "破损", 1);

            WritePBI2Xls(_Worksheet_PBI, prjinfo, prjdir, _RoadPart, _PBIVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet_PBI, 4, 3, 22, 'H', "跳车", 1);

            WritePWI2Xls(_Worksheet_PWI, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet_PWI, 4, 3, 22, 'I', "磨耗", 1);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        // 病害明细
        public static void OutputZNRoadDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\中南安环\病害统计表.xlsx",
                           System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}.xlsx", path, prjdir.Name, "病害统计表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_mx = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;

            WriteZNDis2Xls(_Worksheet_mx, _Worksheet_hz, prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZNDis2Xls(MSExcel.Worksheet _Worksheet_mx, MSExcel.Worksheet _Worksheet_hz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] dislist)
        {
            string disname = null;
            string disgrade = null;

            int colidx = 0;
            int colcnt = 0;
            MSExcel.Range destrange = null, sortrange = null;

            if (prjinfo._Direction > 0)
            {
                colidx = 3;
            }
            else
            {
                colidx = 5;
            }

            string errlog = prjdir.FullName + "\\errlog.txt";
            int dlen = dislist.Length;
            object[,] disinfo = new object[dlen, 8];
            for (int i = 0; i < dlen; ++i)//i区间索引，j病害索引
            {
                string[] temp = dislist[i].RoadDisType.Split('.');
                disname = temp[0];
                try
                {
                    disgrade = temp[1];
                }
                catch { }

                colcnt = 0;
                disinfo[i, colcnt++] = dislist[i].m_mile;

                disinfo[i, colcnt++] = dislist[i].RoadType;
                disinfo[i, colcnt++] = prjinfo._RoadNum;

                disinfo[i, colcnt++] = disname;
                disinfo[i, colcnt++] = disgrade;
                disinfo[i, colcnt++] = dislist[i].calcheight;
                disinfo[i, colcnt++] = dislist[i].calcwidth;
                disinfo[i, colcnt++] = dislist[i].Area;
            }

            destrange = _Worksheet_mx.get_Range(String.Format("A3:H{0}", dlen + 2));
            destrange.Value2 = disinfo;
            GlobalExcel.SetBorderLine(destrange, 63);
            //上行 沥青
            int m = 0;
            string[] distype = { "龟裂", "块状裂缝", "纵向裂缝", "横向裂缝", "坑槽", "松散", "沉陷", "车辙", "波浪拥包", "泛油", "修补", "破碎板", "裂缝", "板角断裂", "错台", "唧泥", "边角剥落", "接缝料损坏", "坑洞", "拱起", "露骨", "修补" };
            for (int n = 0; n < distype.Length; n++)
            {
                m = n + 4;
                if (n == 19)
                {
                    //  =SUMIFS(病害列表!H:H,病害列表!B:B,"水泥",病害列表!D:D,"修补")        =COUNTIFS(病害列表!D:D,"修补",病害列表!B:B,"水泥")
                    _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIF(病害列表！D:D,\"{0}\")+", distype[n]) + string.Format("COUNTIF(病害列表！D:D,\"{0}\")", distype[n + 1]);
                    _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIF(病害列表!D:D,\"{0}\",病害列表!H:H)+", distype[n]) + string.Format("SUMIF(病害列表!D:D,\"{0}\",病害列表!H:H)", distype[n + 1]);
                    n++;
                }
                else
                {
                    if (n > 19)
                    {
                        m--;
                        _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIF(病害列表！D:D,\"{0}\")", distype[n]);
                        _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIF(病害列表!D:D,\"{0}\",病害列表!H:H)", distype[n]);
                        if (n == 21)
                        {
                            _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIFS(病害列表！D:D,\"{0}\",病害列表！B:B,\"{1}\")", distype[n], "水泥");
                            _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIFS(病害列表!H:H,病害列表!B:B,\"水泥\",病害列表!D:D,\"{0}\")", distype[n]);
                        }
                    }
                    else
                    {
                        _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIF(病害列表！D:D,\"{0}\")", distype[n]);
                        _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIF(病害列表!D:D,\"{0}\",病害列表!H:H)", distype[n]);
                        if (n == 10)
                        {
                            _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIFS(病害列表！D:D,\"{0}\",病害列表！B:B,\"{1}\")", distype[n], "沥青");
                            _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIFS(病害列表!H:H,病害列表!B:B,\"沥青\",病害列表!D:D,\"{0}\")", distype[n]);
                        }
                    }
                }
            }

            if (_Setting.IsExcelSort && prjinfo._Direction < 0 && dlen > 0)
            {
                destrange = _Worksheet_mx.get_Range(String.Format("A3:H{0}", dlen + 2));
                sortrange = _Worksheet_mx.get_Range(String.Format("A3:A{0}", dlen + 2));
                GlobalExcel.ReflectionColnum(_Worksheet_mx, destrange, sortrange);
            }
        }

        #endregion

        #region 广西桂兴达
        /////////////////////////////////////////////////////////////////////////////////////////////////
        public static void OutputRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\综合报表模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
            // WriteDisLB2Xls(_Worksheet_lb, prjinfo, _RoadDisList);
            WriteDisLB2Xls_roadpart2(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);

            bool Haslqflag = false;
            bool Hassnflag = false;

            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
            WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjdir, _RoadPart, Haslqflag, Hassnflag);

            WriteAll2Xls(_Workbook, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _DeltaHVal, disval);

            MSExcel.Worksheet _worksheet_RoadInfo = _Workbook.Sheets["路线信息表"] as MSExcel.Worksheet;
            WriteRoadInfo(_worksheet_RoadInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        //养护需求建议表、技术状况明细表、分项指标统计表
        private static void WriteAll2Xls(MSExcel.Workbook workbook, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
            Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] MMTDVal, int[][] PBVal, double[] deltahVal, int disval)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, mtdval = 0, tpcival = 0;

            object[,] mxlist = new object[len, 32];
            object[,] yhlist = new object[len, 21];
            int yhi = 0;
            string lenstr = "0";
            int tlen = len;
            while ((tlen = tlen / 10) > 0)
            {
                lenstr += "0";
            }

            string errlog = prjdir.FullName + "\\errlog.txt";

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                mxlist[i, 0] = prjinfo._RoadCode;
                mxlist[i, 1] = prjinfo._District + "交通运输局";
                mxlist[i, 2] = _Setting.RoadSideType;
                mxlist[i, 3] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 4] = smile;
                mxlist[i, 5] = emile;
                mxlist[i, 6] = milelength;
                mxlist[i, 13] = GlobalExcel._RoadTypeExcelStr[roadpart[i].roadtype];

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }


                //PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                tpcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 14] = drval;
                mxlist[i, 8] = string.Format("=100-{0}*POWER(O{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 4, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 25] = string.Format("=IF(I{0}>={1},\"优\",IF(I{0}>={2},\"良\",IF(I{0}>={3},\"中\",IF(I{0}>={4},\"次\",\"差\"))))",
                    i + 4, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI                
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                //trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                //mxlist[i, 5] = String.Format("=ROUND(100/(1+{0}*EXP({1}*E{2})),2)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 4);

                mxlist[i, 15] = irival;
                mxlist[i, 9] = String.Format("=ROUND(100/(1+{0}*EXP({1}*P{2})),5)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 4);
                mxlist[i, 26] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);

                //Rut
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    double rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);
                    mxlist[i, 16] = rutval;

                    mxlist[i, 10] = string.Format("=IF(Q{0}<{1},{2}-{3}*Q{0},IF(Q{0}<{4},{5}-{6}*(Q{0}-{1}),0))",
                        i + 4,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                    mxlist[i, 27] = string.Format("=IF(K{0}>={1},\"优\",IF(K{0}>={2},\"良\",IF(K{0}>={3},\"中\",IF(K{0}>={4},\"次\",\"差\"))))",
                        i + 4,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);
                }

                if (prjinfo._IsIRIMTD)
                {
                    mxlist[i, 17] = PBVal[i][1];
                    mxlist[i, 18] = PBVal[i][2];
                    mxlist[i, 19] = PBVal[i][3];
                    mxlist[i, 11] = string.Format("=IF((100-R{0}*{1}-S{0}*{2}-T{0}*{3})>0,(100-R{0}*{1}-S{0}*{2}-T{0}*{3}),0)",
                    i + 4, _PBIScore[1], _PBIScore[2], _PBIScore[3]);

                    mxlist[i, 28] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _PBIGrade[roadpart[i].roaddegree][0],
                    _PBIGrade[roadpart[i].roaddegree][1],
                    _PBIGrade[roadpart[i].roaddegree][2],
                    _PBIGrade[roadpart[i].roaddegree][3],
                    (char)('A' + 11));

                    if (disval == 10)
                    {
                        mxlist[i, 31] = deltahVal[i];
                    }
                }

                //构造深度相关 
                mxlist[i, 20] = LMTDVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    mxlist[i, 21] = RMTDVal[i];
                    if (prjinfo._IsMMTD)
                    {
                        mxlist[i, 22] = MMTDVal[i];
                        //wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                        if (MMTDVal[i] == 0)
                        {
                            mxlist[i, 23] = 0;
                        }
                        else  // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                        {
                            mxlist[i, 23] = string.Format("=IF(W{0}-MIN(U{0},V{0})>0, 100*(W{0}-MIN(U{0},V{0}))/W{0},0)", i + 4);
                        }
                    }
                    else
                    {
                        mxlist[i, 23] = 0;
                    }
                }
                mxlist[i, 12] = string.Format("=100-{0}*POWER(X{1},{2})", _PWIa[0], i + 4, _PWIa[1]);
                mxlist[i, 29] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                i + 4,
                _PWIGrade[roadpart[i].roaddegree][0],
                _PWIGrade[roadpart[i].roaddegree][1],
                _PWIGrade[roadpart[i].roaddegree][2],
                _PWIGrade[roadpart[i].roaddegree][3],
                (char)('A' + 12));

                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        mxlist[i, 30] = string.Format("=ROUND(({1}*I{0}+{2}*J{0}+{3}*IF(EXACT(K{0},\"-\"),0,K{0})+{4}*IF(EXACT(L{0},\"-\"),0,L{0}))/({1}+{2}+{3}+{4}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                    }
                    else
                    {
                        mxlist[i, 30] = string.Format("=ROUND(({1}*I{0}+{2}*J{0}+{3}*IF(EXACT(K{0},\"-\"),0,K{0})+{4}*IF(EXACT(L{0},\"-\"),0,L{0})+{5}*IF(EXACT(M{0},\"-\"),0,M{0}))/({1}+{2}+{3}+{4}+{5}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                    }
                }
                else
                {
                    mxlist[i, 30] = string.Format("=ROUND(({1}*I{0}+{2}*J{0})/({1}+{2}),5)",
                    i + 4,
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }

                mxlist[i, 24] = string.Format("=IF(AE{0}>={1},\"优\",IF(AE{0}>={2},\"良\",IF(AE{0}>={3},\"中\",IF(AE{0}>={4},\"次\",\"差\"))))",
                i + 4,
                _PQIGrade[roadpart[i].roaddegree][0],
                _PQIGrade[roadpart[i].roaddegree][1],
                _PQIGrade[roadpart[i].roaddegree][2],
                _PQIGrade[roadpart[i].roaddegree][3]);
                mxlist[i, 7] = String.Format("=CONCATENATE(TEXT(AE{0},\"0.00\"),\"(\",Y{0},\")\")", i + 4);
                if (_Setting.YHType == 0)  //广西标准
                {
                    if (!(trqival > 85 && tpcival > 85))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!J{0}<=70,\"大修\",IF(技术状况明细表!I{0}>85,IF(技术状况明细表!J{0}>85,\"日常养护\",\"中修\"),\"中修\"))", i + 4);
                        yhi++;
                    }
                }
                else if (_Setting.YHType == 1) //辽宁标准
                {
                    if (!(trqival > 85 && tpcival > 70))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!I{0}>=70,IF(技术状况明细表!I{0}>=85,\"日常养护\",IF(技术状况明细表!I{0}>=75,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\"))),IF(技术状况明细表!I{0}>=60,IF(技术状况明细表!I{0}>=85,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\")),IF(技术状况明细表!I{0}>=40,IF(技术状况明细表!I{0}>=75,\"中修\",\"大修\"),\"大修\")))", i + 4);
                        yhi++;
                    }
                }
                else if (_Setting.YHType == 2) //  广西PCI标准  养护标准各参数阈值是多少？
                {
                    if ((roadpart[i].roaddegree < 2 && tpcival < 80)
                        || (roadpart[i].roaddegree >= 2 && tpcival < 70))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!I{0}<60,\"大修\",\"中修\")", i + 4);
                        yhi++;
                    }
                }
            }

            MSExcel.Worksheet worksheet = workbook.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:AF{0}", len + 3));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (disval != 10)
            {
                ((MSExcel.Range)worksheet.Cells[System.Reflection.Missing.Value, 32]).EntireColumn.Delete();
            }
            else
            {
                GlobalExcel.WriteExcel(2, 32, 2, 1, "跳车值\nH", worksheet, 15);
            }
            MSExcel.ChartObject chartobj = null;
            MSExcel.Chart chart = null;

            //PCI RQI RDI PBI PWI
            if (prjinfo._IsRut) destrange = worksheet.get_Range(string.Format("F2:F{0}, I2:M{0},AE2:AE{0}", len + 3));
            else destrange = worksheet.get_Range(string.Format("F2:F{0}, I2:J{0}, AE2:AE{0}", len + 3));
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, true, "", Type.Missing, Type.Missing, Type.Missing);
            chart.Legend.Position = MSExcel.XlLegendPosition.xlLegendPositionTop;

            //DR
            destrange = worksheet.get_Range(string.Format("F3:F{0}, O3:O{0}", len + 3));
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(2);
            chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "破损率DR(%)", Type.Missing, Type.Missing, Type.Missing);

            destrange = worksheet.get_Range(string.Format("F3:F{0}, P3:P{0}", len + 3));
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(3);
            chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "平整度IRI", Type.Missing, Type.Missing, Type.Missing);

            if (prjinfo._IsRut)
            {
                destrange = worksheet.get_Range(string.Format("F3:F{0}, Q3:Q{0}", len + 3));
                chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(4);
                chart = chartobj.Chart;
                chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "车辙Rut", Type.Missing, Type.Missing, Type.Missing);
            }

            //destrange = worksheet.get_Range(string.Format("F3:F{0}, R3:T{0}", len + 3));
            //chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(5);
            //chart = chartobj.Chart;
            //chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "跳车PB", Type.Missing, Type.Missing, Type.Missing);

            //destrange = worksheet.get_Range(string.Format("F3:F{0}, U3:W{0}", len + 3));
            //chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(6);
            //chart = chartobj.Chart;
            //chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "构造深度MTD", Type.Missing, Type.Missing, Type.Missing);

            worksheet = workbook.Sheets["养护需求建议表"] as MSExcel.Worksheet;
            destrange = worksheet.get_Range(String.Format("A3:G{0}", yhi + 2));
            destrange.Value2 = yhlist;
            GlobalExcel.SetBorderLine(destrange, 53);

            object[,] tjlist = new object[6, 1];
            worksheet = workbook.Sheets["分项指标统计表"] as MSExcel.Worksheet;
            tjlist[0, 0] = String.Format("=SUMPRODUCT(技术状况明细表!G4:G{1}, 技术状况明细表!{0}4:{0}{1})/SUM(技术状况明细表!G4:G{1})", "AE", len + 3);
            for (int i = 1; i < 6; ++i)
            {
                tjlist[i, 0] = String.Format("=SUMPRODUCT(技术状况明细表!G4:G{1}, 技术状况明细表!{0}4:{0}{1})/SUM(技术状况明细表!G4:G{1})",
                    GlobalExcel.GetCol(((char)('H' + i))), len + 3);
            }
            destrange = worksheet.get_Range("B3:B8");
            destrange.Value2 = tjlist;

            //if (prjinfo._IsRut) destrange = worksheet.get_Range("L2:Q8");
            //else destrange = worksheet.get_Range("L2:Q5,L7:Q8");
            //chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            //chart = chartobj.Chart;
            //chart.SetSourceData(destrange);
        }
        private static void WriteRoadInfo(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string degreeinfo = File.ReadAllText(prjdir.FullName + "\\DegreeInfo.txt").Replace(" ", Environment.NewLine);
            string roadtypeinfo = File.ReadAllText(prjdir.FullName + "\\RoadTypeInfo.txt").Replace(" ", Environment.NewLine);
            object[,] valobj = new object[7, 1];
            valobj[0, 0] = prjinfo._DataDate;

            MSExcel.Range destrange = worksheet.get_Range("B2");
            destrange.Value2 = valobj;

            destrange = worksheet.get_Range("B4:B9");
            valobj[0, 0] = prjinfo._DataPerson;
            valobj[1, 0] = prjinfo._DataWeather;
            valobj[2, 0] = prjinfo._RoadCode;
            valobj[3, 0] = prjinfo._RoadName;
            valobj[4, 0] = prjinfo._City;
            valobj[5, 0] = prjinfo._District + "交通运输局";
            destrange.Value2 = valobj;

            destrange = worksheet.get_Range("B15:B20");
            valobj[0, 0] = _RoadConfig.DetectWidth;
            valobj[1, 0] = prjinfo._Direction > 0 ? "上行" : "下行";
            valobj[2, 0] = "=IF(分项指标统计表!B3>=90,\"优\",IF(分项指标统计表!B3>=80,\"良\",IF(分项指标统计表!B3>=70,\"中\",IF(分项指标统计表!B3>=60,\"次\",\"差\"))))";
            valobj[3, 0] = degreeinfo;
            valobj[4, 0] = roadtypeinfo;
            valobj[5, 0] = Math.Abs(prjinfo._StartMile * 0.001 - prjinfo._EndMile * 0.001);
            destrange.Value2 = valobj;
        }

        public static void OutputGXDIRIMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\报表模板5.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteIRIMTD2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _LMTDMeanVal, _RMTDMeanVal, _SpeedVal, disval);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteIRIMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
            double[] LIRIVal, double[] RIRIVal, double[] LMTDVal, double[] RMTDVal, double[] SpeedVal, int disval)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            int len = roadpart.Count - 1;
            const int startidx = 14;

            object[,] vallist = new object[len, 12];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = String.Format("{0:0000}+{1:000} - {2:0000}+{3:000}",
                    roadpart[i].mile / 1000, roadpart[i].mile % 1000, roadpart[i + 1].mile / 1000, roadpart[i + 1].mile % 1000);
                if (SpeedVal != null)
                {
                    vallist[i, 1] = SpeedVal[i];
                }

                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 3] = String.Format("=H{0}*0.6", i + startidx);
                    vallist[i, 5] = RMTDVal[i];
                    if (_Setting.IRIExcelSide == 1 || _Setting.IRIExcelSide == 2)
                    {
                        vallist[i, 7] = RIRIVal[i];
                    }
                    vallist[i, 10] = String.Format("=ROUND(100/(1+{0}*EXP({1}*H{2})),5)",
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + startidx);
                }

                if (_Setting.RQIJudgeType == 0)
                {
                    vallist[i, 8] = String.Format("=AVERAGE(G{0}:H{0})", i + startidx);
                }
                else if (_Setting.RQIJudgeType == 1)
                {
                    vallist[i, 8] = String.Format("=MAX(G{0}, H{0})", i + startidx);
                }

                vallist[i, 2] = String.Format("=G{0}*0.6", i + startidx);
                vallist[i, 4] = LMTDVal[i];
                if (_Setting.IRIExcelSide == 0 || _Setting.IRIExcelSide == 2)
                {
                    vallist[i, 6] = LIRIVal[i];
                }
                vallist[i, 9] = String.Format("=ROUND(100/(1+{0}*EXP({1}*G{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + startidx);
                vallist[i, 11] = String.Format("=ROUND(100/(1+{0}*EXP({1}*I{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + startidx);
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A{0}:L{1}", startidx, len + startidx - 1));
            destrange.Value2 = vallist;

            GlobalExcel.SetBorderLine(destrange, 53);


            object[,] vallist2 = new object[8, 1];
            vallist2[0, 0] = string.Format(@"{0}_IRIMTD_{1}m.xlsx", prjdir.Name, disval);
            vallist2[1, 0] = prjinfo._DataPerson;
            vallist2[2, 0] = string.Format("{0}_{1}_{2}_{3}_{4}_{5}", prjinfo._Province, prjinfo._City, prjinfo._District,
                prjinfo._RoadCode, prjinfo._RoadName, prjinfo._RoadNum);
            vallist2[3, 0] = prjinfo._DataDate;
            vallist2[4, 0] = prjinfo._EndDmi.ToString() + "米";
            vallist2[5, 0] = prjinfo._RoadGrade;
            vallist2[6, 0] = string.Format("从K{0:0000}+{1:000}到K{2:0000}+{3:000}",
                prjinfo._StartMile / 1000, prjinfo._StartMile % 1000, prjinfo._EndMile / 1000, prjinfo._EndMile % 1000);
            vallist2[7, 0] = disval.ToString() + "米";

            MSExcel.Range destrange2 = _Worksheet.get_Range("C4:C11");
            destrange2.Value2 = vallist2;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, startidx, 1, 12, true);
            }
        }

        #endregion
        #region  孝感定制
        public static void OutputRoadBedDis_XG(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\路基损坏汇总表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_路基损坏汇总表_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["路基损坏汇总表"] as MSExcel.Worksheet;
            WriteRoadBedDisDC2Xls(_Worksheet_hz, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), xlslen);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputStreetDis_XG(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\景观报表模板\孝感定制\景观病害统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_路基破损及沿线设施病害_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //路基
            MSExcel.Worksheet _Worksheet_ljbh = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            WriteStreetDisDC2Xls_XG(_Worksheet_ljbh, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(), xlslen);


            //沿线设施
            MSExcel.Worksheet _Worksheet_yxss = _Workbook.Sheets["sheet2"] as MSExcel.Worksheet;

            WriteRoadBedDisDC2Xls_XG(_Worksheet_yxss, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), xlslen);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteRoadBedDisDC2Xls(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.roadbeddislist.Count;
            object[,] disval = new object[len, temp + 4];
            worksheet_hz.Cells[2, 2] = prjdir.Name;
            double tclval = 0;
            double ttclval = 0;
            for (int i = 0, j = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int unitlen = Math.Abs(smile - emile);

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j]._nmile >= smile && arrdis[j]._nmile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j]._nmile <= smile && arrdis[j]._nmile > emile)))
                {
                    int typeidx = DiseaseTypes.roadbeddisIdx[arrdis[j]._disname];
                    if (DiseaseTypes.roadbeddislist[typeidx].unitval != 0)
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndisnum;
                    }
                    else
                    {
                        DiseaseTypes.roadbeddislist[typeidx].sumval += arrdis[j]._ndislen;
                    }
                    j++;
                }

                disval[i, 0] = smile;
                disval[i, 1] = emile;

                tclval = 0;
                ttclval = 0;
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.roadbeddislist[k].sumval;
                    if (k > 0)
                    {
                        if (DiseaseTypes.roadbeddislist[k - 1].distype != DiseaseTypes.roadbeddislist[k].distype)
                        {
                            ttclval = ttclval * 1000 / unitlen;
                            ttclval = ttclval > 100 ? 100 : ttclval;
                            tclval += DiseaseTypes.roadbeddislist[k - 1].weight * (100 - ttclval);
                            ttclval = 0;
                        }
                    }
                    ttclval = ttclval + DiseaseTypes.roadbeddislist[k].unitscore * DiseaseTypes.roadbeddislist[k].sumval;
                }
                int ttypeidx = DiseaseTypes.roadbeddisIdx["路基构造物损坏.重"];
                if (DiseaseTypes.roadbeddislist[ttypeidx].sumval > 0)
                {
                    tclval = 0;
                }
                else
                {
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.roadbeddislist[temp - 1].weight * (100 - ttclval);
                }

                disval[i, temp + 2] = tclval;
                disval[i, temp + 3] = string.Format("=IF({1}{0}>=90,\"优\",IF({1}{0}>=80,\"良\",IF({1}{0}>=70,\"中\",IF({1}{0}>=60,\"次\",\"差\"))))",
                    i + 5, GlobalExcel.GetCol((char)(temp + 2 + 'A')));

                smile = emile;
                DiseaseTypes.Clear();
            }
            destrange = worksheet_hz.get_Range(string.Format("A5:{1}{0}", len + 4, GlobalExcel.GetCol((char)('A' + temp + 3))));
            destrange.Value2 = disval;
            GlobalExcel.SetBorderLine(destrange, 53);

            disval = new object[temp, 1];
            for (int i = 0; i < temp; ++i)
            {
                disval[i, 0] = string.Format("=SUM({0}6:{0}{1})", GlobalExcel.GetCol((char)('C' + i)), len + 4);
            }
            destrange = worksheet_hz.get_Range(string.Format("{0}3:{0}{1}", GlobalExcel.GetCol((char)('A' + temp + 6)), temp + 2));
            destrange.Value2 = disval;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet_hz, 5, 1, temp + 4, true);
                GlobalExcel.Reflection(worksheet_hz, 5, 1, 2, false);
            }
        }
        private static void WriteStreetDisDC2Xls_XG(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;
            object[,] disLJval = new object[dlen, 6]; //路基
            string city = prjinfo._City;
            string town = prjinfo._District;
            string roadName = prjinfo._RoadName;
            string code = prjinfo._RoadCode;
            for (int i = 0; i < dlen; i++)
            {
                disLJval[i, 0] = city;
                disLJval[i, 1] = town;
                disLJval[i, 2] = roadName;
                disLJval[i, 3] = code;
                string mile = arrdis[i]._mile.Substring(1, arrdis[i]._mile.Length - 1);
                string[] sp = mile.Split('+');
                int mileInt = int.Parse(sp[0]) * 1000 + int.Parse(sp[1]);


                disLJval[i, 4] = mileInt;
                if (arrdis[i]._dislen == "0")
                {
                    disLJval[i, 5] = arrdis[i]._disname + arrdis[i]._disnum + "处";
                }
                else
                {
                    disLJval[i, 5] = arrdis[i]._disname + arrdis[i]._dislen + "m";
                }

            }
            destrange = worksheet_hz.get_Range(string.Format("A3:{1}{0}", dlen + 2, GlobalExcel.GetCol((char)('A' + 5))));
            destrange.Value2 = disLJval;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                destrange = worksheet_hz.get_Range(string.Format("A3:F{0}", dlen + 2));
                MSExcel.Range sortrange = worksheet_hz.get_Range(string.Format("E3:E{0}", dlen + 2));
                GlobalExcel.ReflectionColnum(worksheet_hz, destrange, sortrange);
            }
        }
        private static void WriteRoadBedDisDC2Xls_XG(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;
            object[,] disLJval = new object[dlen, 6]; //沿线
            string city = prjinfo._City;
            string town = prjinfo._District;
            string roadName = prjinfo._RoadName;
            string code = prjinfo._RoadCode;
            for (int i = 0; i < dlen; i++)
            {
                disLJval[i, 0] = city;
                disLJval[i, 1] = town;
                disLJval[i, 2] = roadName;
                disLJval[i, 3] = code;
                string mile = arrdis[i]._mile.Substring(1, arrdis[i]._mile.Length - 1);
                string[] sp = mile.Split('+');
                int mileInt = int.Parse(sp[0]) * 1000 + int.Parse(sp[1]);


                disLJval[i, 4] = mileInt;
                if (arrdis[i]._dislen == "0")
                {
                    disLJval[i, 5] = arrdis[i]._disname + arrdis[i]._disnum + "处";
                }
                else
                {
                    disLJval[i, 5] = arrdis[i]._disname + arrdis[i]._dislen + "m";
                }

            }
            destrange = worksheet_hz.get_Range(string.Format("A3:{1}{0}", dlen + 2, GlobalExcel.GetCol((char)('A' + 5))));
            destrange.Value2 = disLJval;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                destrange = worksheet_hz.get_Range(string.Format("A3:F{0}", dlen + 2));
                MSExcel.Range sortrange = worksheet_hz.get_Range(string.Format("E3:E{0}", dlen + 2));
                GlobalExcel.ReflectionColnum(worksheet_hz, destrange, sortrange);
            }
        }

        #endregion
        #region 多车道病害面积统计
        public static void OutputAreaStatistics(MSExcel.Application excelApp, List<string> BHTJxlslist, List<string> JSMXxlslist, string outpath)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet_mqi = null;
            MSExcel.Worksheet destsheet = null;
            destsheet = _Workbook.Sheets["病害面积"] as MSExcel.Worksheet;
            WriteBHMJ2Xlsx(excelApp, destsheet, BHTJxlslist);

            destsheet = _Workbook.Sheets["技术指标"] as MSExcel.Worksheet;
            destsheet_mqi = _Workbook.Sheets["MQI"] as MSExcel.Worksheet;
            WriteJSZB2Xlsx(excelApp, destsheet, destsheet_mqi, JSMXxlslist);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteBHMJ2Xlsx(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> BHTJxlslist)
        {
            int lqdisnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
            int sndisnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;

            int rowidx = 4;
            int userow = 0;
            object[,] lqobj = new object[1, lqdisnum];
            object[,] snobj = new object[1, sndisnum];
            object[,] infoobj = new object[1, 12];
            foreach (string tlane in BHTJxlslist)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(tlane, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet tsheet = null;
                MSExcel.Range trange = null;
                object[,] tobj = null;
                bool islq = false;
                bool issn = false;

                GetPrjInfo(tbook, ref infoobj);

                //沥青路面的病害统计
                try
                {
                    tsheet = tbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                }
                catch (Exception) { }

                if (tsheet != null)
                {
                    islq = true;
                    userow = GlobalExcel.judegeusedrow(tsheet, 4, 5);
                    trange = tsheet.get_Range("A1:AB" + userow.ToString());
                    tobj = (object[,])trange.Value2;
                    for (int i = 0; i < lqdisnum; ++i)
                    {
                        lqobj[0, i] = tobj[userow, i + 4];
                    }
                    for (int i = 0; i < sndisnum; ++i)
                    {
                        snobj[0, i] = "─";
                    }
                }

                if (islq)
                {
                    trange = destsheet.get_Range(string.Format("A{0}:L{0}", rowidx));
                    trange.Value2 = infoobj;
                    trange = destsheet.get_Range(string.Format("M{0}:AH{0}", rowidx));
                    trange.Value2 = lqobj;
                    trange = destsheet.get_Range(string.Format("AI{0}:BC{0}", rowidx));
                    trange.Value2 = snobj;
                    ++rowidx;
                }

                //水泥路面的病害统计
                try
                {
                    tsheet = tbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }

                if (tsheet != null)
                {
                    issn = true;
                    userow = GlobalExcel.judegeusedrow(tsheet, 4, 5);
                    trange = tsheet.get_Range("A1:AA" + userow.ToString());
                    tobj = (object[,])trange.Value2;
                    for (int i = 0; i < sndisnum; ++i)
                    {
                        snobj[0, i] = tobj[userow, i + 4];
                    }
                    for (int i = 0; i < lqdisnum; ++i)
                    {
                        lqobj[0, i] = "─";
                    }
                }
                if (issn)
                {
                    trange = destsheet.get_Range(string.Format("A{0}:L{0}", rowidx));
                    trange.Value2 = infoobj;
                    trange = destsheet.get_Range(string.Format("M{0}:AH{0}", rowidx));
                    trange.Value2 = lqobj;
                    trange = destsheet.get_Range(string.Format("AI{0}:BC{0}", rowidx));
                    trange.Value2 = snobj;
                    ++rowidx;
                }

                tbook.Close();
            }
        }

        //技术状况评定明细表
        private static void WriteJSZB2Xlsx(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, MSExcel.Worksheet destsheet_mqi, List<string> JSMXxlslist)
        {
            double tmp = 0;
            int rowidx = 2;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] mqiobj = new object[1, 16];

            MSExcel.Range trange = null;
            foreach (string tlane in JSMXxlslist)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(tlane, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet tsheet = null;
                object[,] tobj = null;
                object[,] gradeobj = null;

                GetPrjInfo(tbook, ref infoobj);
                trange = destsheet.get_Range(string.Format("A{0}:L{0}", rowidx));
                trange.Value2 = infoobj;

                try
                {
                    tsheet = tbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }

                if (tsheet != null)
                {
                    userow = GlobalExcel.judegeusedrow(tsheet, 1, 5);

                    //统计分级区间
                    tobj = new object[6, 6];
                    tobj[0, 1] = "优";
                    tobj[0, 2] = "良";
                    tobj[0, 3] = "中";
                    tobj[0, 4] = "次";
                    tobj[0, 5] = "差";
                    tobj[1, 0] = "MQI";
                    tobj[2, 0] = "SCI";
                    tobj[3, 0] = "PQI";
                    tobj[4, 0] = "BCI";
                    tobj[5, 0] = "TCI";
                    string[] colstr = { "C", "D", "E", "M", "N" };
                    for (int i = 1; i < 6; ++i)
                    {
                        tobj[i, 1] = string.Format("=(SUMIF({0}5:{0}{1},\">={2}\",B5:B{1}))*0.001", colstr[i - 1], userow - 1, _MQIGrade[0]);
                        for (int j = 2; j < 6; ++j)
                        {
                            tobj[i, j] = string.Format("=(SUMIF({0}5:{0}{1},\">={2}\",B5:B{1})-SUMIF({0}5:{0}{1},\">={3}\",B5:B{1}))*0.001", colstr[i - 1], userow - 1, _MQIGrade[j - 1], _MQIGrade[j - 2]);
                        }
                    }
                    trange = tsheet.get_Range(string.Format("R4:W9"));
                    trange.Value2 = tobj;
                    GlobalExcel.SetBorderLine(trange, 63);

                    //写入技术指标表单
                    trange = tsheet.get_Range(string.Format("C{0}:N{0}", userow));
                    tobj = (object[,])trange.Value2;
                    for (int k = 1; k <= tobj.Length; ++k)
                    {
                        try
                        {
                            tmp = Convert.ToDouble(tobj[1, k]);
                            if (tmp < 0)
                                tobj[1, k] = 0;
                            else if (tmp > 100)
                                tobj[1, k] = 100;
                        }
                        catch (System.Exception )
                        {
                            tobj[1, k] = 0;
                        }
                    }

                    trange = destsheet.get_Range(string.Format("M{0}:V{0}", rowidx));
                    trange.Value2 = tobj;

                    //写入MQI表单
                    trange = tsheet.get_Range(string.Format("S5:W9"));
                    gradeobj = (object[,])trange.Value2;
                    mqiobj[0, 0] = infoobj[0, 1];
                    mqiobj[0, 1] = infoobj[0, 3];
                    mqiobj[0, 2] = infoobj[0, 10];

                    mqiobj[0, 3] = tobj[1, 1];
                    mqiobj[0, 4] = tobj[1, 2];
                    mqiobj[0, 5] = tobj[1, 3];
                    mqiobj[0, 6] = tobj[1, 9];
                    mqiobj[0, 7] = tobj[1, 12];

                    mqiobj[0, 8] = gradeobj[1, 1];
                    mqiobj[0, 9] = gradeobj[1, 2];
                    mqiobj[0, 10] = gradeobj[1, 3];
                    mqiobj[0, 11] = gradeobj[1, 4];
                    mqiobj[0, 12] = gradeobj[1, 5];

                    mqiobj[0, 13] = string.Format("=I{0}/C{0}*100", rowidx + 1);
                    mqiobj[0, 14] = string.Format("=(I{0}+J{0})/C{0}*100", rowidx + 1);
                    mqiobj[0, 15] = string.Format("=(L{0}+M{0})/C{0}*100", rowidx + 1);

                    trange = destsheet_mqi.get_Range(string.Format("A{0}:P{0}", rowidx + 1));
                    trange.Value2 = mqiobj;

                    ++rowidx;
                }
                tbook.Save();
                tbook.Close();
            }

            trange = destsheet.get_Range(string.Format("A1:V{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);

            trange = destsheet_mqi.get_Range(string.Format("A1:P{0}", rowidx));
            GlobalExcel.SetBorderLine(trange, 63);
        }

        private static void GetPrjInfo(MSExcel.Workbook tbook, ref object[,] infoobj)
        {
            MSExcel.Worksheet tsheet = null;
            MSExcel.Range trange = null;
            object[,] tobj = null;

            //工程信息
            try
            {
                tsheet = tbook.Sheets["工程信息"] as MSExcel.Worksheet;
                trange = tsheet.get_Range("B2:B17");
                tobj = (object[,])trange.Value2;

                infoobj[0, 0] = tobj[2, 1];
                infoobj[0, 1] = tobj[3, 1];

                string roadcode = tobj[4, 1].ToString();
                if (roadcode.StartsWith("C") || roadcode.StartsWith("c"))
                {
                    infoobj[0, 2] = "村道";
                }
                else if (roadcode.StartsWith("X") || roadcode.StartsWith("x"))
                {
                    infoobj[0, 2] = "县道";
                }
                else if (roadcode.StartsWith("Y") || roadcode.StartsWith("y"))
                {
                    infoobj[0, 2] = "乡道";
                }
                else if (roadcode.StartsWith("G") || roadcode.StartsWith("g"))
                {
                    infoobj[0, 2] = "国道";
                }
                else if (roadcode.StartsWith("S") || roadcode.StartsWith("s"))
                {
                    infoobj[0, 2] = "省道";
                }
                infoobj[0, 3] = tobj[4, 1];
                infoobj[0, 4] = tobj[5, 1];
                infoobj[0, 5] = tobj[8, 1];
                infoobj[0, 6] = tobj[7, 1];
                infoobj[0, 7] = tobj[9, 1];
                infoobj[0, 8] = tobj[6, 1];
                infoobj[0, 9] = tobj[15, 1];
                infoobj[0, 10] = tobj[16, 1];

                string roadtype = tobj[14, 1].ToString();
                if (roadtype.Contains("沥青"))
                {
                    if (roadtype.Contains("水泥"))
                    {
                        infoobj[0, 11] = "混合";
                    }
                    else
                    {
                        infoobj[0, 11] = "沥青";
                    }
                }
                else if (roadtype.Contains("水泥"))
                {
                    infoobj[0, 11] = "水泥";
                }
            }
            catch (Exception ) { }
        }

        public static void OutputAllRoadStatistics(MSExcel.Application excelApp, string outpath,
            List<string> ExcelBHTJList, List<string> ExcelJSMXList,
            List<string> ExcelDRList, List<string> ExcelIRIList)
        {
            string srcxls = string.Format(@"{0}\报表模板\湖南农村公路\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计-HN.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;
            destsheet = _Workbook.Sheets["路面破损病害明细表"] as MSExcel.Worksheet;
            WriteAllRoadDis2Xls(excelApp, destsheet, ExcelBHTJList);

            destsheet = _Workbook.Sheets["公里路线技术状况评定明细表"] as MSExcel.Worksheet;
            WriteAllRoadPQI2Xls(excelApp, destsheet, ExcelJSMXList);

            destsheet = _Workbook.Sheets["公里单元技术状况评定明细表"] as MSExcel.Worksheet;
            WriteAllUnitPQI2Xls(excelApp, destsheet, ExcelDRList, ExcelIRIList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// 惠普 生成病害统计表
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="outpath"></param>
        /// <param name="ExcelBHTJList"></param>
        /// <param name="ExcelJSMXList"></param>
        /// <param name="ExcelDRList"></param>
        /// <param name="ExcelIRIList"></param>
        public static void OutputAllRoadDisease_HP(MSExcel.Application excelApp, string outpath,
          List<string> ExcelList)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\道路病害统计模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\道路病害统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;
            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteAllRoadDis2Xls_HP(excelApp, destsheet, ExcelList);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteAllRoadDis2Xls(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> xlslist)
        {
            int rowidx = 2;
            int userow = 0;
            object[,] infoobj = new object[1, 12];

            MSExcel.Range trange = null;
            foreach (string tlane in xlslist)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(tlane, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet tsheet = null;
                object[,] tobj = null;

                GetPrjInfo(tbook, ref infoobj);

                try
                {
                    tsheet = tbook.Sheets["病害列表"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }

                if (tsheet != null)
                {
                    userow = GlobalExcel.judegeusedrow(tsheet, 1, 2);
                    trange = tsheet.get_Range(string.Format("A3:N{0}", userow));
                    tobj = (object[,])trange.Value2;
                    userow = userow - 2;

                    object[,] dataobj = new object[userow, 15];
                    for (int i = 1; i <= userow; ++i)
                    {
                        for (int j = 0; j < 8; ++j)
                        {
                            dataobj[i - 1, j] = infoobj[0, j];
                        }
                        dataobj[i - 1, 8] = tobj[i, 1];
                        dataobj[i - 1, 9] = tobj[i, 3];
                        dataobj[i - 1, 10] = tobj[i, 8];
                        dataobj[i - 1, 11] = tobj[i, 9];
                        dataobj[i - 1, 12] = tobj[i, 10];
                        dataobj[i - 1, 14] = tobj[i, 13];
                    }
                    trange = destsheet.get_Range(string.Format("A{0}:O{1}", rowidx, rowidx + userow - 1));
                    trange.Value2 = dataobj;
                    rowidx = rowidx + userow;
                }
                tbook.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:O{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        private static void WriteAllRoadDis2Xls_HP(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> xlslist)
        {


            int rowidx = 3;
            int userow = 0;


            MSExcel.Range trange = null;
            foreach (string tlane in xlslist)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(tlane, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MSExcel.Worksheet tsheet = null;
                object[,] tobj = null;
                object[,] infoobj = null;
                tsheet = tbook.Sheets["工程信息"] as MSExcel.Worksheet;
                trange = tsheet.get_Range("B2:B17");
                infoobj = (object[,])trange.Value2;


                try
                {
                    tsheet = tbook.Sheets["病害列表"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }
                if (tsheet != null)
                {
                    userow = GlobalExcel.judegeusedrow(tsheet, 1, 2);
                    trange = tsheet.get_Range(string.Format("A3:N{0}", userow));
                    tobj = (object[,])trange.Value2;
                    userow = userow - 2;

                    object[,] dataobj = new object[1, 51];
                    dataobj[0, 0] = infoobj[5, 1];  //道路名
                    dataobj[0, 1] = infoobj[3, 1];  //镇
                    dataobj[0, 2] = infoobj[10, 1];  //日期*
                    string roadGradStr = infoobj[8, 1].ToString();

                    int roadGrad = 0;
                    if (roadGradStr.Contains("一"))
                    {
                        roadGrad = 1;
                    }
                    else if (roadGradStr.Contains("二"))
                    {
                        roadGrad = 2;
                    }
                    else if (roadGradStr.Contains("三"))
                    {
                        roadGrad = 3;
                    }
                    else if (roadGradStr.Contains("四"))
                    {
                        roadGrad = 4;
                    }
                    string roadLenStr = Math.Abs((double.Parse(infoobj[15, 1].ToString()) - double.Parse(infoobj[6, 1].ToString())) / 1000).ToString();
                    dataobj[0, 3] = roadLenStr;  //道路长度


                    dataobj[0, 4] = userow;  //病害总数

                    int qCount = 0;
                    int zCount = 0;
                    int zhCount = 0;
                    int[] datas = new int[38];  //名称对应的个数   
                    double[] score = new double[38];  //分值
                    double roadArea = _RoadConfig.DetectWidth * double.Parse(roadLenStr) * 1000;
                    bool isLq = true;
                    int roadType = 0;
                    for (int t = 1; t <= userow; t++)
                    {
                        for (int c = 1; c <= 14; c++)
                        {

                            string qCell = tobj[t, 4].ToString(); //病害程度
                            double disArea = double.Parse(tobj[t, 8].ToString());
                            isLq = tobj[t, 13].ToString().Contains("沥青") ? true : false; //是否是沥青
                            string cellDisData = ""; //病害名称
                            if (qCell.Contains("无"))
                            {
                                qCount++;
                            }
                            else if (qCell.Contains("轻"))
                            {
                                qCount++;
                            }
                            else if (qCell.Contains("中"))
                            {
                                zCount++;
                            }
                            else
                            {
                                zhCount++;
                            }

                            #region 病害个数获取
                            if (tobj[t, 3] == null)
                            {
                                cellDisData = "";
                            }
                            else
                            {
                                cellDisData = tobj[t, 3].ToString();

                            }
                            if (cellDisData == "纵向裂缝")
                            {
                                datas[0]++;

                                if (qCell.Contains("中"))
                                {
                                    score[0] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[0] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[0] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData == "横向裂缝")
                            {
                                datas[1]++;
                                if (qCell.Contains("中"))
                                {
                                    score[0] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[0] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[0] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("龟裂"))
                            {
                                datas[2]++;
                                if (qCell.Contains("中"))
                                {
                                    score[2] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[2] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[2] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData == "块状裂缝")
                            {
                                datas[3]++;
                                if (qCell.Contains("中"))
                                {
                                    score[3] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[3] += double.Parse((disArea * 0.8).ToString());
                                }
                                else
                                {
                                    score[3] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("坑槽"))
                            {
                                datas[4]++;
                                if (qCell.Contains("中"))
                                {
                                    score[4] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[4] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[4] += double.Parse((disArea * 0.8).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("沉陷") && isLq)
                            {

                                datas[5]++;
                                if (qCell.Contains("中"))
                                {
                                    score[5] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[5] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[5] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("车辙"))
                            {
                                datas[6]++;
                                if (qCell.Contains("中"))
                                {
                                    score[6] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[6] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[6] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("松散"))
                            {
                                datas[7]++;
                                if (qCell.Contains("中"))
                                {
                                    score[7] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[7] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[7] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("泛油"))
                            {
                                datas[8]++;
                                score[8] += double.Parse((disArea * 0.2).ToString());

                                break;
                            }
                            else if (cellDisData.Contains("波浪拥包"))
                            {
                                datas[9]++;
                                if (qCell.Contains("中"))
                                {
                                    score[9] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[9] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[9] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("翻浆"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("剥落"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("啃边"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("路框差"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("唧浆"))
                            {

                                break;
                            }
                            else if (cellDisData.Contains("线裂"))
                            {

                                break;
                            }
                            else if (cellDisData.Contains("修补") && isLq)
                            {
                                datas[16]++;
                                score[16] += double.Parse((disArea * 0.1).ToString());
                                break;
                            }
                            else if (cellDisData.Contains("路面保洁差"))
                            {

                                break;
                            }
                            else if (cellDisData.Contains("综合病害"))
                            {
                                break;
                            }
                            else if (cellDisData == "裂缝" && !isLq)
                            {
                                datas[19]++;
                                if (qCell.Contains("中"))
                                {
                                    score[19] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[19] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[19] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("坑洞"))
                            {
                                datas[20]++;
                                score[20] += double.Parse((disArea * 1).ToString());

                                break;
                            }
                            else if (cellDisData.Contains("边角剥落"))
                            {
                                datas[21]++;
                                if (qCell.Contains("中"))
                                {
                                    score[21] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[21] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[21] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("接缝料损坏"))
                            {
                                datas[22]++;
                                if (qCell.Contains("中"))
                                {
                                    score[22] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[22] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[22] += double.Parse((disArea * 0.4).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("露骨"))
                            {
                                datas[23]++;
                                score[23] += double.Parse((disArea * 0.3).ToString());


                                break;
                            }
                            else if (cellDisData.Contains("板角断裂"))
                            {
                                datas[24]++;
                                if (qCell.Contains("中"))
                                {
                                    score[24] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[24] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[24] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("破碎板"))
                            {
                                datas[25]++;
                                if (qCell.Contains("中"))
                                {
                                    score[25] += double.Parse((disArea * 0.8).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[25] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[25] += double.Parse((disArea * 0.8).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("错台"))
                            {
                                datas[26]++;
                                if (qCell.Contains("中"))
                                {
                                    score[26] += double.Parse((disArea * 0.6).ToString());
                                }
                                else if (qCell.Contains("重"))
                                {
                                    score[26] += double.Parse((disArea * 1).ToString());
                                }
                                else
                                {
                                    score[26] += double.Parse((disArea * 0.6).ToString());
                                }
                                break;
                            }
                            else if (cellDisData.Contains("唧泥"))
                            {
                                datas[27]++;
                                score[27] += double.Parse((disArea * 1).ToString());

                                break;
                            }
                            else if (cellDisData.Contains("拱起"))
                            {
                                datas[28]++;
                                score[28] += double.Parse((disArea * 1).ToString());
                                break;
                            }
                            else if (cellDisData.Contains("边角裂缝"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("表面纹裂"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("层状剥落"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("路框差"))
                            {
                                break;
                            }
                            else if (cellDisData.Contains("沉陷"))
                            {

                                break;
                            }
                            else if (cellDisData.Contains("修补") && !isLq)
                            {
                                datas[34]++;
                                score[34] += double.Parse((disArea * 0.1).ToString());

                                break;
                            }
                            #endregion
                        }

                    }
                    double areaDiseaseSum = 0; //病害总
                    roadType = isLq ? 0 : 1;
                    for (int i = 0; i < 38; i++)
                    {
                        areaDiseaseSum += double.Parse((score[i]).ToString());
                    }
                    double dr = 100 * areaDiseaseSum / roadArea;
                    double v1 = _PCIa[roadGrad][roadType][0];
                    double v2 = _PCIa[roadGrad][roadType][1];
                    double pci = 100 - v1 * Math.Pow(dr, v2);
                    dataobj[0, 5] = qCount;  //轻总数
                    dataobj[0, 6] = zCount;  //中总数
                    dataobj[0, 7] = zhCount;  //重总数
                    dataobj[0, 8] = (int)(zhCount / double.Parse(roadLenStr));  //重度病害密度（处/公里）
                    dataobj[0, 9] = zhCount;  //严重
                    dataobj[0, 10] = "0";  //危险

                    dataobj[0, 11] = pci.ToString("f2");  //PCI

                    dataobj[0, 12] = pci >= 90 ? "优" : dr >= 80 ? "良" : dr >= 70 ? "重" : dr >= 60 ? "次" : "差";

                    dataobj[0, 13] = datas[0];  //纵向裂缝
                    dataobj[0, 14] = datas[1];  //横向裂缝
                    dataobj[0, 15] = datas[2];  //龟裂
                    dataobj[0, 16] = datas[3];  //块状裂缝
                    dataobj[0, 17] = datas[4];  //坑槽
                    dataobj[0, 18] = datas[5];  //沉陷
                    dataobj[0, 19] = datas[6];  //车辙
                    dataobj[0, 20] = datas[7];    //松散 
                    dataobj[0, 21] = datas[8];   //泛油 
                    dataobj[0, 22] = datas[9];  //波浪拥包 
                    dataobj[0, 23] = datas[10];   // 翻浆 
                    dataobj[0, 24] = datas[11];  //剥落
                    dataobj[0, 25] = datas[12];  //啃边 
                    dataobj[0, 26] = datas[13];  //路框差 
                    dataobj[0, 27] = datas[14];  //唧浆 
                    dataobj[0, 28] = datas[15];  //线裂 
                    dataobj[0, 29] = datas[16];  //修补未达标 
                    dataobj[0, 30] = datas[17];  //路面保洁差 
                    dataobj[0, 31] = datas[18];  //综合病害   
                    dataobj[0, 32] = datas[19];  //裂缝 
                    dataobj[0, 33] = datas[20];  //坑洞 
                    dataobj[0, 34] = datas[21];  //边角剥落  
                    dataobj[0, 35] = datas[22];  //接缝料损坏 
                    dataobj[0, 36] = datas[23];  //露骨 
                    dataobj[0, 37] = datas[24];  //板角断裂   
                    dataobj[0, 38] = datas[25];  //破碎板 
                    dataobj[0, 39] = datas[26];  //错台 
                    dataobj[0, 40] = datas[27];  //唧泥 
                    dataobj[0, 41] = datas[28];  //拱起  
                    dataobj[0, 42] = datas[29];  //边角裂缝 
                    dataobj[0, 43] = datas[30];  //表面纹裂 
                    dataobj[0, 44] = datas[31];  //层状剥落 
                    dataobj[0, 45] = datas[32];  //路框差 
                    dataobj[0, 46] = datas[33];  //沉陷 
                    dataobj[0, 47] = datas[34];  //修补未达标 
                    dataobj[0, 48] = datas[35];  //路面保洁差 
                    dataobj[0, 49] = datas[36];  //其他  
                    dataobj[0, 50] = datas[37];  //综合病害 
                    trange = destsheet.get_Range(string.Format("A{0}:AY{1}", rowidx, rowidx + 1 - 1));
                    trange.Value2 = dataobj;
                    rowidx = rowidx + 1;

                }

                tbook.Close();
            }
            trange = destsheet.get_Range(string.Format("A4:AY{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        private static void WriteAllUnitPQI2Xls(MSExcel.Application excelApp, MSExcel.Worksheet destsheet,
            List<string> ExcelDRList, List<string> ExcelIRIList)
        {
            if (ExcelDRList.Count != ExcelIRIList.Count)
            {
                MessageBox.Show("IRI和PCI的报表数量不一致，请检查！");
            }

            int rowidx = 3;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int roadlinenum = ExcelDRList.Count;
            for (int ri = 0; ri < roadlinenum; ++ri)
            {
                MSExcel.Workbook tbook_dr = excelApp.Workbooks.Open(ExcelDRList[ri], Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MSExcel.Workbook tbook_iri = excelApp.Workbooks.Open(ExcelIRIList[ri], Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                object[,] tobj_dr = null;
                object[,] tobj_iri = null;

                GetPrjInfo(tbook_iri, ref infoobj);

                MSExcel.Worksheet tsheet_dr = tbook_dr.Sheets["Sheet1"] as MSExcel.Worksheet;
                MSExcel.Worksheet tsheet_iri = tbook_iri.Sheets["Sheet1"] as MSExcel.Worksheet;

                int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                userow = GlobalExcel.judegeusedrow(tsheet_dr, 1, 3);

                trange = tsheet_dr.get_Range(string.Format("A3:I{0}", userow));
                tobj_dr = (object[,])trange.Value2;
                userow = userow - 2;
                outobj = new object[userow, 21];

                trange = tsheet_iri.get_Range(string.Format("A4:I{0}", userow + 3));
                tobj_iri = (object[,])trange.Value2;

                for (int i = 0; i < userow; ++i)
                {
                    for (int ii = 0; ii < 8; ++ii)
                    {
                        outobj[i, ii] = infoobj[0, ii];
                    }

                    outobj[i, 8] = tobj_dr[i + 1, 1];
                    outobj[i, 9] = tobj_dr[i + 1, 2];
                    outobj[i, 10] = string.Format("=ABS(I{0}-J{0})", rowidx + i);

                    int roadtype = RoadDiseaseTypes.roadtypedict[tobj_dr[i + 1, 7].ToString()];
                    outobj[i, 11 + roadtype] = string.Format("=K{0}", rowidx + i);

                    outobj[i, 13] = tobj_dr[i + 1, 4];
                    outobj[i, 14] = string.Format("=100-{1}*POWER(N{0},{2})", rowidx + i, _PCIa[roadgrade][roadtype][0], _PCIa[roadgrade][roadtype][1]);

                    outobj[i, 15] = tobj_iri[i + 1, 6];
                    outobj[i, 16] = string.Format("=ROUND(100/(1+{0}*EXP({1}*P{2})),5)", _RQIa[roadgrade][roadtype][0], _RQIa[roadgrade][roadtype][1], rowidx + i);

                    outobj[i, 17] = string.Format("=ROUND(({1}*O{0}+{2}*Q{0})/({1}+{2}),5)", rowidx + i, _PQIW[roadgrade][roadtype][0], _PQIW[roadgrade][roadtype][1]);

                    outobj[i, 18] = string.Format("=IF(O{0}>={1},\"优\",IF(O{0}>={2},\"良\",IF(O{0}>={3},\"中\",IF(O{0}>={4},\"次\",\"差\"))))",
                        rowidx + i, _PCIGrade[roadgrade][0], _PCIGrade[roadgrade][1], _PCIGrade[roadgrade][2], _PCIGrade[roadgrade][3]);

                    outobj[i, 19] = string.Format("=IF(Q{0}>={1},\"优\",IF(Q{0}>={2},\"良\",IF(Q{0}>={3},\"中\",IF(Q{0}>={4},\"次\",\"差\"))))",
                        rowidx + i, _RQIGrade[roadgrade][roadtype][0], _RQIGrade[roadgrade][roadtype][1], _RQIGrade[roadgrade][roadtype][2], _RQIGrade[roadgrade][roadtype][3]);

                    outobj[i, 20] = string.Format("=IF(R{0}>={1},\"优\",IF(R{0}>={2},\"良\",IF(R{0}>={3},\"中\",IF(R{0}>={4},\"次\",\"差\"))))",
                        rowidx + i, _PQIGrade[roadtype][0], _PQIGrade[roadtype][1], _PQIGrade[roadtype][2], _PQIGrade[roadtype][3]);
                }

                trange = destsheet.get_Range(string.Format("A{0}:U{1}", rowidx, rowidx + userow - 1));
                trange.Value2 = outobj;
                rowidx += userow;

                tbook_dr.Close();
                tbook_iri.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:U{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }

        private static void WriteAllRoadPQI2Xls(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> xlslist)
        {
            int rowidx = 2;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = new object[1, 21];

            MSExcel.Range trange = null;
            foreach (string tlane in xlslist)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(tlane, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet tsheet = null;
                object[,] tobj = null;

                GetPrjInfo(tbook, ref infoobj);
                try
                {
                    tsheet = tbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }

                if (tsheet != null)
                {
                    int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                    userow = GlobalExcel.judegeusedrow(tsheet, 1, 5) - 1;
                    trange = tsheet.get_Range(string.Format("A5:O{0}", userow));
                    tobj = (object[,])trange.Value2;

                    double[] sumidxval = new double[3];
                    double[] tidxval = new double[3];
                    for (int i = 0; i < 3; ++i)
                    {
                        sumidxval[i] = 0;
                        tidxval[i] = 0;
                    }

                    double dmivalLQ = 0;
                    double dmivalSN = 0;
                    int[] PQIGradeDmi = new int[5];
                    for (int i = 0; i < 5; ++i)
                    {
                        PQIGradeDmi[i] = 0;
                    }

                    string roadtypestr = null;
                    double sumlen = 0;
                    userow = userow - 4;
                    for (int i = 1; i <= userow; ++i)
                    {
                        int dmival = Convert.ToInt32(tobj[i, 2]);
                        for (int j = 0; j < 3; ++j)
                        {
                            tidxval[j] = Convert.ToDouble(tobj[i, 5 + j]);
                            sumidxval[j] += tidxval[j] * dmival;
                        }
                        sumlen += dmival;

                        roadtypestr = tobj[i, 15].ToString();
                        if (roadtypestr == "水泥")
                        {
                            dmivalSN += dmival;
                        }
                        else if (roadtypestr == "沥青")
                        {
                            dmivalLQ += dmival;
                        }

                        int k = 0;
                        for (k = 0; k < 5; ++k)
                        {
                            if (_PQIGrade[roadgrade][k] <= tidxval[0])
                            {
                                break;
                            }
                        }
                        PQIGradeDmi[k] += dmival;
                    }
                    for (int i = 0; i < 3; ++i)
                    {
                        sumidxval[i] = sumidxval[i] / sumlen;
                    }

                    for (int i = 0; i < 11; ++i)
                    {
                        outobj[0, i] = infoobj[0, i];
                    }
                    outobj[0, 11] = dmivalSN * 0.001;
                    outobj[0, 12] = dmivalLQ * 0.001;
                    for (int i = 0; i < 3; ++i)
                    {
                        outobj[0, 13 + i] = sumidxval[i];
                    }
                    for (int i = 0; i < 5; ++i)
                    {
                        outobj[0, 16 + i] = PQIGradeDmi[i];
                    }

                    trange = destsheet.get_Range(string.Format("A{0}:U{0}", rowidx));
                    trange.Value2 = outobj;
                    ++rowidx;
                }
                tbook.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:X{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }

        #endregion

        #region 厦门捷航
        /// <summary>
        /// 不同等级汇总文字
        /// </summary>
        /// <param name="_Worksheet"></param>
        /// <param name="startrow">数据开始的行</param>
        /// <param name="statisticsrow">数据开始的行</param>
        /// <param name="statisticscol">汇总开始的列</param>
        /// <param name="gradecol">等级所在的列</param>
        /// <param name="indextype">指数类型文字</param>
        /// <param name="strposition">0-汇总说明文字在下侧，1-汇总说明文字在左侧</param>
        private static void WriteStatistics_XMJH(MSExcel.Worksheet _Worksheet, int startrow, int statisticsrow, int statisticscol, char gradecol, string indextype, int strposition)
        {
            int userow = GlobalExcel.judegeusedrow(_Worksheet, 1, startrow);
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range(string.Format("{0}{1}:{2}{3}",
                GlobalExcel.GetCol((char)('A' + statisticscol - 1)),
                statisticsrow,
                GlobalExcel.GetCol((char)('A' + statisticscol + 4 - 1)),
                statisticsrow + 2));

            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF({3}{1}:{3}{2},\"{0}\",A{1}:A{2})-SUMIF({3}{1}:{3}{2},\"{0}\",B{1}:B{2}))", degstr[i], statisticsrow, userow, gradecol);
                val[2, i] = string.Format("=ABS(SUM(A{0}:A{1})-SUM(B{0}:B{1}))", statisticsrow, userow);
                val[1, i] = string.Format("={0}{1}/{0}{2}", GlobalExcel.GetCol((char)('A' + statisticscol - 1 + i)), statisticsrow, statisticsrow + 2);
            }

            destrange.Value2 = val;

            int strrow = 0;
            int strcol = 0;
            if (strposition == 0)
            {
                strrow = startrow + 6;
                strcol = statisticscol;
            }
            else if (strposition == 1)
            {
                strrow = statisticsrow - 1;
                strcol = statisticscol - 7;
            }

            _Worksheet.Cells[strrow, strcol] = string.Format("=CONCATENATE(\"沥青路面{6}评价等级“优”率占路段总数\",ROUND({1}{0},4)*100,\"%，“良”率占路段总数\",ROUND({2}{0},4)*100,\"%，“中”率占路段总数\",ROUND({3}{0},4)*100,\"%，“次”率占路段总数\",ROUND({4}{0},4)*100,\"%，“差”率占路段总数\",ROUND({5}{0},4)*100,\"%。\")",
                statisticsrow + 1,
                GlobalExcel.GetCol((char)('A' + statisticscol - 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 2)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 3)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 4)),
                indextype);
        }

        public static void OutputIRI_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\激光平整度试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_IRI_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面平整度"] as MSExcel.Worksheet;
            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 19, 63);
            WriteStatistics_XMJH(_Worksheet, 19, 19, 18, 'H', "平整度", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\激光路面平整度试验检测记录.xlsx",
                System.Windows.Forms.Application.StartupPath);
            Destxls = string.Format(@"{0}\{1}_IRI_{2}m_检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面平整度记录"] as MSExcel.Worksheet;
            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 13, 63);
            WriteStatistics_XMJH(_Worksheet, 13, 13, 18, 'H', "平整度", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputRut_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\激光车辙试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_Rut_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面车辙"] as MSExcel.Worksheet;
            WriteRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _SpeedVal, _MarkVal, 19, 63);
            WriteStatistics_XMJH(_Worksheet, 19, 19, 18, 'H', "车辙深度", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\激光路面车辙试验检测记录.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}_Rut_{2}m_检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面车辙记录"] as MSExcel.Worksheet;
            WriteRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _SpeedVal, _MarkVal, 13, 63);
            WriteStatistics_XMJH(_Worksheet, 13, 13, 18, 'H', "车辙深度", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputMTD_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\激光构造深度试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_MTD_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面构造深度"] as MSExcel.Worksheet;
            WriteMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _SpeedVal, _MarkVal, 17, 63);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\激光路面构造试验检测记录.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}_MTD_{2}m_检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面构造深度记录"] as MSExcel.Worksheet;
            WriteMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _SpeedVal, _MarkVal, 12, 63);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputPWI_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\激光磨耗试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PWI_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面磨耗"] as MSExcel.Worksheet;
            WritePWI2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _SpeedVal, _MarkVal, 18, 63);
            WriteStatistics_XMJH(_Worksheet, 18, 18, 18, 'I', "磨耗", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\激光路面磨耗试验检测记录.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}_PWI_{2}m_检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面磨耗记录"] as MSExcel.Worksheet;
            WritePWI2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _SpeedVal, _MarkVal, 11, 63);
            WriteStatistics_XMJH(_Worksheet, 11, 11, 18, 'I', "磨耗", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputPBI_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\激光跳车试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PBI_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面跳车"] as MSExcel.Worksheet;
            WritePBI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _PBIVal, _SpeedVal, _MarkVal, 20, 63);
            WriteStatistics_XMJH(_Worksheet, 20, 20, 18, 'H', "跳车", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\激光路面跳车试验检测记录 .xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}_PB_10m_检测记录.xlsx", path, prjdir.Name);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面跳车记录"] as MSExcel.Worksheet;
            WritePB2Xls(_Worksheet, prjinfo, prjdir, _RoadPart10, _LDeltaHVal, _RDeltaHVal, _SpeedVal10, _MarkVal10, 11, 63);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputPQI_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\路面技术状况指数PQI.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PQI_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面PQI"] as MSExcel.Worksheet;
            WritePQI2Xls_XMJH(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _MarkVal, 19);
            WriteStatistics_XMJH(_Worksheet, 19, 19, 24, 'O', "综合评价指数", 0);
            WriteStatistics_XMJH(_Worksheet, 19, 19, 31, 'E', "破损", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\路面技术状况指数PQI记录.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            Destxls = string.Format(@"{0}\{1}_PQI_{2}m_检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["路面PQI记录"] as MSExcel.Worksheet;
            WritePQI2Xls_XMJH(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _MarkVal, 12);
            WriteStatistics_XMJH(_Worksheet, 12, 12, 24, 'O', "综合评价指数", 0);
            WriteStatistics_XMJH(_Worksheet, 12, 12, 31, 'E', "破损", 0);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePQI2Xls_XMJH(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal, string[] MarkVal,
            int DataStartXlsxRow)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0, wrval = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 17];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 0;
                vallist[rowcnt, colcnt++] = smile;
                vallist[rowcnt, colcnt++] = emile;
                vallist[rowcnt, colcnt++] = prjinfo._RoadNum;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt++] = Math.Round(pcival, 5);
                vallist[rowcnt, colcnt++] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={4},\"次\",\"差\"))))",
                    rowcnt + DataStartXlsxRow,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, colcnt] = trqival;

                colcnt++;
                vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                    rowcnt + DataStartXlsxRow,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);

                //Rut
                if (prjinfo._IsRut)
                {
                    //rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    double rdival = 0;
                    if (rutval <= _RDIRD[0][1])
                    {
                        rdival = _RDIRD[0][0] - _RDIa[0] * rutval;
                    }
                    else if (rutval <= _RDIRD[1][1])
                    {
                        rdival = _RDIRD[1][0] - _RDIa[1] * (rutval - _RDIRD[0][1]);
                    }
                    else
                    {
                        rdival = 0;
                    }
                    // if(roadpart[i].roadtype==)
                    vallist[rowcnt, colcnt++] = rdival;
                    vallist[rowcnt, colcnt++] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                        rowcnt + DataStartXlsxRow,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);
                }
                else
                {
                    colcnt = colcnt + 2;
                }


                //PBI
                if (prjinfo._IsIRIMTD)
                {
                    vallist[rowcnt, colcnt++] = string.Format("=IF((100-{0}*{1}-{2}*{3}-{4}*{5})>0,(100-{0}*{1}-{2}*{3}-{4}*{5}),0)",
                        PBVal[i][1], _PBIScore[1],
                        PBVal[i][2], _PBIScore[2],
                        PBVal[i][3], _PBIScore[3]);
                    vallist[rowcnt, colcnt++] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                        rowcnt + DataStartXlsxRow,
                        _PBIGrade[roadpart[i].roaddegree][0],
                        _PBIGrade[roadpart[i].roaddegree][1],
                        _PBIGrade[roadpart[i].roaddegree][2],
                        _PBIGrade[roadpart[i].roaddegree][3]);
                }

                //PWI
                if (prjinfo._IsIRIMTD)
                {

                    wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                    wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);

                    if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))
                    {
                        wrval = 0;
                    }
                    vallist[rowcnt, colcnt++] = string.Format("=IF((100-{0}*POWER({1},{2}))>0,(100-{0}*POWER({1},{2})),0)", _PWIa[0], wrval, _PWIa[1]);
                    vallist[rowcnt, colcnt++] = string.Format("=IF(L{0}>={1},\"优\",IF(L{0}>={2},\"良\",IF(L{0}>={3},\"中\",IF(L{0}>={4},\"次\",\"差\"))))",
                        rowcnt + DataStartXlsxRow,
                        _PWIGrade[roadpart[i].roaddegree][0],
                        _PWIGrade[roadpart[i].roaddegree][1],
                        _PWIGrade[roadpart[i].roaddegree][2],
                        _PWIGrade[roadpart[i].roaddegree][3]);
                }
                //pqi  =IF(P54="沥青",ROUND((0.35*D54+0.3*F54+0.15*IF(EXACT(H54,"-"),0,H54)+0.1*IF(EXACT(J54,"-"),0,J54)+0.1*IF(EXACT(L54,"-"),0,L54))/(0.35+0.3+0.15+0.1+0.1),5),ROUND((0.5*D54+0.3*F54+0*IF(EXACT(H54,"-"),0,H54)+0.1*IF(EXACT(J54,"-"),0,J54)+0.1*IF(EXACT(L54,"-"),0,L54))/(0.5+0.3+0+0.1+0.1),5))
                if (roadpart[i].roaddegree <= 1)
                {
                    if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                    {
                        vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(J{0},\"-\"),0,J{0})+{5}*IF(EXACT(L{0},\"-\"),0,L{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*D{0}+{7}*F{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(J{0},\"-\"),0,J{0}))/({6}+{7}+{8}+{9}),5))",
                            rowcnt + DataStartXlsxRow,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][0][2],
                        _PQIW[roadpart[i].roaddegree][0][3],
                        _PQIW[roadpart[i].roaddegree][0][4],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1],
                        _PQIW[roadpart[i].roaddegree][1][2],
                        _PQIW[roadpart[i].roaddegree][1][3]);
                    }
                    else
                    {
                        vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0})+{4}*IF(EXACT(J{0},\"-\"),0,J{0})+{5}*IF(EXACT(L{0},\"-\"),0,L{0}))/({1}+{2}+{3}+{4}+{5}),5),ROUND(({6}*D{0}+{7}*F{0}+{8}*IF(EXACT(H{0},\"-\"),0,H{0})+{9}*IF(EXACT(J{0},\"-\"),0,J{0})+{10}*IF(EXACT(L{0},\"-\"),0,L{0}))/({6}+{7}+{8}+{9}+{10}),5))",
                            rowcnt + DataStartXlsxRow,
                          _PQIW[roadpart[i].roaddegree][0][0],
                          _PQIW[roadpart[i].roaddegree][0][1],
                          _PQIW[roadpart[i].roaddegree][0][2],
                          _PQIW[roadpart[i].roaddegree][0][3],
                          _PQIW[roadpart[i].roaddegree][0][4],
                          _PQIW[roadpart[i].roaddegree][1][0],
                          _PQIW[roadpart[i].roaddegree][1][1],
                          _PQIW[roadpart[i].roaddegree][1][2],
                          _PQIW[roadpart[i].roaddegree][1][3],
                          _PQIW[roadpart[i].roaddegree][1][4]);
                    }
                }
                else
                {
                    vallist[rowcnt, colcnt++] = string.Format("=IF(P{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),5),ROUND(({3}*D{0}+{4}*F{0})/({3}+{4}),5))",
                        rowcnt + DataStartXlsxRow,
                        _PQIW[roadpart[i].roaddegree][0][0],
                        _PQIW[roadpart[i].roaddegree][0][1],
                        _PQIW[roadpart[i].roaddegree][1][0],
                        _PQIW[roadpart[i].roaddegree][1][1]);
                }

                vallist[rowcnt, colcnt++] = string.Format("=IF(N{0}>={1},\"优\",IF(N{0}>={2},\"良\",IF(N{0}>={3},\"中\",IF(N{0}>={4},\"次\",\"差\"))))",
                    rowcnt + DataStartXlsxRow,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                vallist[rowcnt, colcnt++] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[rowcnt, colcnt++] = MarkVal[i];
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A{0}:Q{1}", DataStartXlsxRow, rowcnt + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 63);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 18, true);
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }

        public static void OutputPCI_XMJH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\报告模板\路面破损（图像法-人工勾画法）PCI试验检测报告.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_PCI_{2}m_检测报告.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            bool Haslqflag = false;
            bool Hassnflag = false;
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面破损"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥路面破损"] as MSExcel.Worksheet;
            WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 20, 63);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\厦门捷航\记录模板\路面破损（图像法-人工勾画法）试验检测记录.xlsx",
                System.Windows.Forms.Application.StartupPath);
            Destxls = string.Format(@"{0}\{1}_路面破损检测记录.xlsx", path, prjdir.Name, disval);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet_DisLB = _Workbook.Sheets["破损"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart_XMJH(_Worksheet_DisLB, prjinfo, prjdir, _RoadDisList, _RoadPart);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteDisLB2Xls_roadpart_XMJH(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 12];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        vallist[rowcnt, 1] = prjinfo._RoadNum;
                        vallist[rowcnt, 2] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 3] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 3] = "无";
                        }
                        vallist[rowcnt, 4] = arrdis[j].rect.Height * _RoadConfig.HeightScale;
                        vallist[rowcnt, 5] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 6] = (arrdis[j].rect.Width / 2 + arrdis[j].rect.X) * _RoadConfig.WidthScale;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        vallist[rowcnt, 8] = arrdis[j].calcheight;
                        vallist[rowcnt, 9] = arrdis[j].calcwidth;
                        vallist[rowcnt, 10] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A12:L{0}", dlen + 11));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 63);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 12, 1, 12, true);
            }
        }
        #endregion

        #region 河南焦作
        /// <summary>
        /// 综合PQI评定表格.xlsx，河南焦作综合评定
        /// </summary>
        public static void OutputPQI_HNJZ_ZHPD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\河南焦作\综合PQI评定表格.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}_{3}_综合PQI评定表格_{4}m.xlsx",
                path, prjinfo._RoadCode, prjinfo._RoadName, prjinfo._RoadNum, disval);

            MSExcel.Workbook _Workbook = null;

            //上下行，先导入下行的数据到Excel
            //如果是下行，原来存在这个文件，就将这个文件删掉，重新生成
            //如果是上行，原来存在这个文件，说明已经将下行的数据导入了表中，只需要将上行的数据追加写入表中
            //如果是上行，原来不存在这个文件，说明没有下行数据，重新写文件
            if (prjinfo._Direction > 0)
            {
                if (File.Exists(Destxls))
                {
                    _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing,
                        false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                else
                {
                    _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                        true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            else
            {
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

            MSExcel.Worksheet worksheet_pqi = null;
            if (prjinfo._Direction > 0)
                worksheet_pqi = _Workbook.Sheets["PQI汇总-上行"] as MSExcel.Worksheet;
            else
                worksheet_pqi = _Workbook.Sheets["PQI汇总-下行"] as MSExcel.Worksheet;

            MSExcel.Worksheet worksheet_pci = _Workbook.Sheets["PCI明细"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_rqi = _Workbook.Sheets["RQI明细"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_rdi = _Workbook.Sheets["RDI明细"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_pbi = _Workbook.Sheets["PBI明细"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_pwi = _Workbook.Sheets["PWI明细"] as MSExcel.Worksheet;

            Write_HNJZ_ZHPD(worksheet_pqi, worksheet_pci, worksheet_rqi, worksheet_rdi, worksheet_pbi, worksheet_pwi,
                prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _PBIVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
        }

        private static void Write_HNJZ_ZHPD(MSExcel.Worksheet worksheet_pqi, MSExcel.Worksheet worksheet_pci, MSExcel.Worksheet worksheet_rqi,
            MSExcel.Worksheet worksheet_rdi, MSExcel.Worksheet worksheet_pbi, MSExcel.Worksheet worksheet_pwi,
            ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal,
            double[] LRutVal, double[] RRutVal, double[] SRutVal,
            int[][] PBVal,
            double[] LMTDVal, double[] RMTDVal, double[] CMTDVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double irival = 0, rutval = 0, wrval = 0, drval = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;

            object[,] objmile = new object[len, 2];
            object[,] objpqi = new object[len, 9];
            object[,] objpci = new object[len, 4];
            object[,] objrqi = new object[len, 3];
            object[,] objrdi = new object[len, 3];
            object[,] objpbi = new object[len, 5];
            object[,] objpwi = new object[len, 3];

            char pci_area_col;
            char pci_dr_col;
            char pci_pci_col;

            char rqi_iri_col;
            char rqi_rqi_col;

            char rdi_rd_col;
            char rdi_rdi_col;

            char[] pbi_pb_col = new char[3];
            char pbi_pbi_col;

            char pwi_wr_col;
            char pwi_pwi_col;

            char pqi_pqi_col = 'C';

            if (prjinfo._Direction > 0)
            {
                pci_area_col = 'C';
                pci_dr_col = 'D';
                pci_pci_col = 'E';

                rqi_iri_col = 'C';
                rqi_rqi_col = 'D';

                rdi_rd_col = 'C';
                rdi_rdi_col = 'D';

                pbi_pb_col[0] = 'C';
                pbi_pb_col[1] = 'D';
                pbi_pb_col[2] = 'E';
                pbi_pbi_col = 'F';

                pwi_wr_col = 'C';
                pwi_pwi_col = 'D';
            }
            else
            {
                pci_area_col = 'G';
                pci_dr_col = 'H';
                pci_pci_col = 'I';

                rqi_iri_col = 'F';
                rqi_rqi_col = 'G';

                rdi_rd_col = 'F';
                rdi_rdi_col = 'G';

                pbi_pb_col[0] = 'H';
                pbi_pb_col[1] = 'I';
                pbi_pb_col[2] = 'J';
                pbi_pbi_col = 'K';

                pwi_wr_col = 'F';
                pwi_pwi_col = 'G';
            }

            int pci_start_row = 5;
            int rqi_start_row = 5;
            int rdi_start_row = 5;
            int pbi_start_row = 6;
            int pwi_start_row = 5;
            int pqi_start_row = 5;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                double sumdisarea = 0;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        sumdisarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                objmile[rowcnt, 0] = smile;
                objmile[rowcnt, 1] = emile;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                objpci[rowcnt, 0] = sumdisarea;
                objpci[rowcnt, 1] = drval;
                objpci[rowcnt, 2] = string.Format("=IF(100-{1}*POWER({3}{0},{2})>0,100-{1}*POWER({3}{0},{2}),100-{1}*POWER({3}{0},{2}))",
                    rowcnt + pci_start_row,
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    pci_dr_col);
                objpci[rowcnt, 3] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    rowcnt + pci_start_row,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3],
                    pci_pci_col);

                //IRI
                if (prjinfo._IsIRIMTD)
                {
                    if (prjinfo._IsDIRIMTD)
                    {
                        if (_Setting.IRIExcelSide == 2)
                        {
                            if (_Setting.RQIJudgeType == 0)
                            {
                                irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                            }
                            else if (_Setting.RQIJudgeType == 1)
                            {
                                irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                            }
                        }
                        else if (_Setting.IRIExcelSide == 0)
                        {
                            irival = Math.Round(LIRIVal[i], 5);
                        }
                        else if (_Setting.IRIExcelSide == 1)
                        {
                            irival = Math.Round(RIRIVal[i], 5);
                        }
                    }
                    else
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    objrqi[rowcnt, 0] = irival;
                    objrqi[rowcnt, 1] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{3}{2})),5)",
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        rowcnt + rqi_start_row,
                        rqi_iri_col);
                    objrqi[rowcnt, 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt + rqi_start_row,
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                        rqi_rqi_col);
                }

                //Rut
                if (prjinfo._IsRut)
                {
                    rutval = Math.Round(SRutVal[i], 5);
                    objrdi[rowcnt, 0] = rutval;
                    objrdi[rowcnt, 1] = string.Format("=IF({7}{0}<{1},{2}-{3}*{7}{0},IF({7}{0}<{4},{5}-{6}*({7}{0}-{1}),0))",
                        rowcnt + rdi_start_row,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1],
                        rdi_rd_col);
                    objrdi[rowcnt, 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt + rdi_start_row,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3],
                        rdi_rdi_col);
                }

                //PBI
                if (prjinfo._IsIRIMTD)
                {
                    objpbi[rowcnt, 0] = PBVal[i][1];
                    objpbi[rowcnt, 1] = PBVal[i][2];
                    objpbi[rowcnt, 2] = PBVal[i][3];
                    objpbi[rowcnt, 3] = string.Format("=IF((100-{4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3})>0,(100- {4}{0}*{1}-{5}{0}*{2}-{6}{0}*{3}),0)",
                        rowcnt + pbi_start_row,
                        _PBIScore[1],
                        _PBIScore[2],
                        _PBIScore[3],
                        pbi_pb_col[0],
                        pbi_pb_col[1],
                        pbi_pb_col[2]);
                    objpbi[rowcnt, 4] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt + pbi_start_row,
                        _PBIGrade[roadpart[i].roaddegree][0],
                        _PBIGrade[roadpart[i].roaddegree][1],
                        _PBIGrade[roadpart[i].roaddegree][2],
                        _PBIGrade[roadpart[i].roaddegree][3],
                        pbi_pbi_col);
                }

                //PWI
                if (prjinfo._IsIRIMTD)
                {
                    wrval = 100 * (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i])) / CMTDVal[i];
                    wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                    if (CMTDVal[i] == 0 || (CMTDVal[i] - Math.Min(LMTDVal[i], RMTDVal[i]) < 0))
                    {
                        wrval = 0;
                    }
                    objpwi[rowcnt, 0] = wrval;
                    objpwi[rowcnt, 1] = string.Format("=IF((100-{0}*POWER({1}{3},{2}))>0,(100-{0}*POWER({1}{3},{2})),0)",
                        _PWIa[0],
                        pwi_wr_col,
                        _PWIa[1],
                        rowcnt + pwi_start_row);
                    objpwi[rowcnt, 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt + pwi_start_row,
                        _PWIGrade[roadpart[i].roaddegree][0],
                        _PWIGrade[roadpart[i].roaddegree][1],
                        _PWIGrade[roadpart[i].roaddegree][2],
                        _PWIGrade[roadpart[i].roaddegree][3],
                        pwi_pwi_col);
                }

                objpqi[rowcnt, 1] = string.Format("=PCI明细!{0}{1}", pci_pci_col, rowcnt + pci_start_row);
                objpqi[rowcnt, 2] = string.Format("=RQI明细!{0}{1}", rqi_rqi_col, rowcnt + rqi_start_row);
                objpqi[rowcnt, 3] = string.Format("=RDI明细!{0}{1}", rdi_rdi_col, rowcnt + rdi_start_row);
                objpqi[rowcnt, 4] = string.Format("=PBI明细!{0}{1}", pbi_pbi_col, rowcnt + pbi_start_row);
                objpqi[rowcnt, 5] = string.Format("=PWI明细!{0}{1}", pwi_pwi_col, rowcnt + pwi_start_row);

                if (_Setting.Is_SnCarve == 1 && roadpart[i].roadtype == 1)//有刻槽,并且是水泥，pwi不参与计算
                {
                    objpqi[rowcnt, 0] = string.Format("=ROUND(({1}*D{0}+{2}*E{0}+{3}*F{0}+{4}*G{0})/({1}+{2}+{3}+{4}),5)",
                        rowcnt + pqi_start_row,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3]);
                }
                else
                {
                    objpqi[rowcnt, 0] = string.Format("=ROUND(({1}*D{0}+{2}*E{0}+{3}*F{0}+{4}*G{0}+{5}*H{0})/({1}+{2}+{3}+{4}+{5}),5)",
                        rowcnt + pqi_start_row,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][3],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][4]);
                }

                objpqi[rowcnt, 8] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    rowcnt + pqi_start_row,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3],
                    pqi_pqi_col);

                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet_pci.get_Range(String.Format("A{0}:B{1}", pci_start_row, rowcnt + pci_start_row - 1));
            destrange.Value2 = objmile;
            destrange = worksheet_pci.get_Range(String.Format("{2}{0}:{3}{1}", pci_start_row, rowcnt + pci_start_row - 1, pci_area_col, (char)(pci_area_col + 3)));
            destrange.Value2 = objpci;
            destrange = worksheet_pci.get_Range(String.Format("A{0}:J{1}", pci_start_row, rowcnt + pci_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet_rqi.get_Range(String.Format("A{0}:B{1}", rqi_start_row, rowcnt + rqi_start_row - 1));
            destrange.Value2 = objmile;
            destrange = worksheet_rqi.get_Range(String.Format("{2}{0}:{3}{1}", rqi_start_row, rowcnt + rqi_start_row - 1, rqi_iri_col, (char)(rqi_iri_col + 2)));
            destrange.Value2 = objrqi;
            destrange = worksheet_rqi.get_Range(String.Format("A{0}:H{1}", rqi_start_row, rowcnt + rqi_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet_rdi.get_Range(String.Format("A{0}:B{1}", rdi_start_row, rowcnt + rdi_start_row - 1));
            destrange.Value2 = objmile;
            destrange = worksheet_rdi.get_Range(String.Format("{2}{0}:{3}{1}", rdi_start_row, rowcnt + rdi_start_row - 1, rdi_rd_col, (char)(rdi_rd_col + 2)));
            destrange.Value2 = objrdi;
            destrange = worksheet_rdi.get_Range(String.Format("A{0}:H{1}", rdi_start_row, rowcnt + rdi_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet_pbi.get_Range(String.Format("A{0}:B{1}", pbi_start_row, rowcnt + pbi_start_row - 1));
            destrange.Value2 = objmile;
            destrange = worksheet_pbi.get_Range(String.Format("{2}{0}:{3}{1}", pbi_start_row, rowcnt + pbi_start_row - 1, pbi_pb_col[0], (char)(pbi_pb_col[0] + 4)));
            destrange.Value2 = objpbi;
            destrange = worksheet_pbi.get_Range(String.Format("A{0}:L{1}", pbi_start_row, rowcnt + pbi_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet_pwi.get_Range(String.Format("A{0}:B{1}", pwi_start_row, rowcnt + pwi_start_row - 1));
            destrange.Value2 = objmile;
            destrange = worksheet_pwi.get_Range(String.Format("{2}{0}:{3}{1}", pwi_start_row, rowcnt + pwi_start_row - 1, pwi_wr_col, (char)(pwi_wr_col + 2)));
            destrange.Value2 = objpwi;
            destrange = worksheet_pwi.get_Range(String.Format("A{0}:H{1}", pwi_start_row, rowcnt + pwi_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet_pqi.get_Range(String.Format("A{0}:B{1}", pqi_start_row, rowcnt + pqi_start_row - 1));
            destrange.Value2 = objmile;

            //下行-数据要逆序
            if (prjinfo._Direction < 0)
            {
                GlobalExcel.Reflection(worksheet_pci, pci_start_row, 1, 10, true);
                GlobalExcel.Reflection(worksheet_pci, pci_start_row, 1, 2, false);

                GlobalExcel.Reflection(worksheet_rqi, rqi_start_row, 1, 8, true);
                GlobalExcel.Reflection(worksheet_rqi, rqi_start_row, 1, 2, false);

                GlobalExcel.Reflection(worksheet_rdi, rdi_start_row, 1, 8, true);
                GlobalExcel.Reflection(worksheet_rdi, rdi_start_row, 1, 2, false);

                GlobalExcel.Reflection(worksheet_pbi, pbi_start_row, 1, 12, true);
                GlobalExcel.Reflection(worksheet_pbi, pbi_start_row, 1, 2, false);

                GlobalExcel.Reflection(worksheet_pwi, pwi_start_row, 1, 8, true);
                GlobalExcel.Reflection(worksheet_pwi, pwi_start_row, 1, 2, false);

                GlobalExcel.Reflection(worksheet_pqi, pqi_start_row, 1, 2, true);
                GlobalExcel.Reflection(worksheet_pqi, pqi_start_row, 1, 2, false);
            }

            destrange = worksheet_pqi.get_Range(String.Format("C{0}:K{1}", pqi_start_row, rowcnt + pqi_start_row - 1));
            destrange.Value2 = objpqi;
            destrange = worksheet_pqi.get_Range(String.Format("A{0}:K{1}", pqi_start_row, rowcnt + pqi_start_row - 1));
            GlobalExcel.SetBorderLine(destrange, 63);
        }

        #endregion

        #region 广东华路
        public static void OutputDis_GDHL(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            MSExcel.Worksheet _Worksheet_lqhz = null;

            string subdname = null;

            if (disval == 10)
            {
                subdname = "十米";
            }
            else if (disval == 100)
            {
                subdname = "百米";
            }
            else if (disval == 1000)
            {
                subdname = "公里";
            }
            else
            {
                return;
            }

            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\广东华路\明细表\路面病害明细表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string destxls = string.Format(@"{0}\{1}_路面病害明细表.xlsx", path, prjdir.Name);
            MSExcel.Workbook _WorkbookMX = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _WorkbookMX.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet_lb = _WorkbookMX.Sheets["路面病害明细表"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart_GDHL(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);
            _WorkbookMX.Save();
            _WorkbookMX.Close(Type.Missing, Type.Missing, Type.Missing);

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\广东华路\{1}\水泥混凝土路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            destxls = string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.xlsx", path, prjdir.Name, subdname);
            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _WorkbookSN.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("水泥混凝土路面破损{0}汇总-0-0", subdname)] as MSExcel.Worksheet;

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\广东华路\{1}\沥青路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            destxls = string.Format(@"{0}\{1}_沥青路面破损{2}汇总.xlsx", path, prjdir.Name, subdname);
            _WorkbookLQ = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _WorkbookLQ.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_lqhz = _WorkbookLQ.Sheets[string.Format("沥青路面破损{0}汇总-0-0", subdname)] as MSExcel.Worksheet;

            bool Haslqflag = false;
            bool Hassnflag = false;
            WriteDisHZ2Xls_GDHL(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 53);

            if (!Haslqflag)
            {
                _WorkbookLQ.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{1}_沥青路面破损{2}汇总.xlsx", path, prjdir.Name, subdname));
            }
            else
            {
                _WorkbookLQ.Save();
                _WorkbookLQ.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            if (!Hassnflag)
            {
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.xlsx", path, prjdir.Name, subdname));
            }
            else
            {
                _WorkbookSN.Save();
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteDisLB2Xls_roadpart_GDHL(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 10];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        vallist[rowcnt, 1] = (int)(arrdis[j].m_mile + arrdis[j].realheight);
                        vallist[rowcnt, 2] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype] + "路面";
                        vallist[rowcnt, 3] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 4] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 4] = "无";
                        }
                        vallist[rowcnt, 5] = arrdis[j].calcheight;
                        vallist[rowcnt, 6] = arrdis[j].calcwidth;
                        vallist[rowcnt, 7] = arrdis[j].depth;
                        vallist[rowcnt, 8] = arrdis[j].Area;
                        vallist[rowcnt, 9] = arrdis[j].remarks;
                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            _Worksheet.Cells[2, 2] = string.Format("{0}-{1}-{2}-{3}",
                prjinfo._RoadCode,
                prjinfo._RoadName,
                prjinfo._Direction > 0 ? "上行" : "下行",
                prjinfo._RoadNum.Replace("车道", ""));
            _Worksheet.Cells[2, 8] = string.Format("A{0:K0+000}", prjinfo._StartMile);

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A4:J{0}", dlen + 3));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 10, true);
            }
        }

        private static void WriteDisHZ2Xls_GDHL(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            ref bool Haslqflag, ref bool Hassnflag, int borderType)
        {
            MSExcel.Range destrange;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志

            int rowcnt_sn = 0;
            int rowcnt_lq = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            object[,] disvallq = new object[len, RoadDiseaseTypes.DiseaseTypeDict[0].Count + 5];
            object[,] disvalsn = new object[len, RoadDiseaseTypes.DiseaseTypeDict[1].Count + 5];

            string[] sndis_name = { "水泥.破碎板.轻", "水泥.破碎板.重",
                                      "水泥.裂缝.轻", "水泥.裂缝.中", "水泥.裂缝.重",
                                      "水泥.板角断裂.轻", "水泥.板角断裂.中", "水泥.板角断裂.重",
                                      "水泥.错台.轻", "水泥.错台.重",
                                      "水泥.唧泥",
                                      "水泥.边角剥落.轻", "水泥.边角剥落.中", "水泥.边角剥落.重",
                                      "水泥.接缝料损坏.轻", "水泥.接缝料损坏.重",
                                      "水泥.坑洞",
                                      "水泥.拱起",
                                      "水泥.露骨",
                                      "水泥.修补.条状", "水泥.修补.块状"};

            string[] lqdis_name = { "沥青.龟裂.轻", "沥青.龟裂.中", "沥青.龟裂.重",
                                      "沥青.块状裂缝.轻", "沥青.块状裂缝.重",
                                      "沥青.纵向裂缝.轻", "沥青.纵向裂缝.重",
                                      "沥青.横向裂缝.轻", "沥青.横向裂缝.重",
                                      "沥青.坑槽.轻", "沥青.坑槽.重",
                                      "沥青.松散.轻", "沥青.松散.重",
                                      "沥青.沉陷.轻", "沥青.沉陷.重",
                                      "沥青.车辙.轻", "沥青.车辙.重",
                                      "沥青.波浪拥包.轻", "沥青.波浪拥包.重",
                                      "沥青.泛油",
                                      "沥青.修补.条状", "沥青.修补.块状"};

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[j].calcheight;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 0;
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;
                    disvalsn[rowcnt_sn, colcnt++] = smile;
                    disvalsn[rowcnt_sn, colcnt++] = emile;
                    for (int kk = 0; kk < sndis_name.Length; ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][RoadDiseaseTypes.DiseaseTypeDict[1][sndis_name[kk]]];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            disvalsn[rowcnt_sn, colcnt++] = type.totallength;
                        }
                        else
                        {
                            disvalsn[rowcnt_sn, colcnt++] = type.totalarea;
                        }
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disvalsn[rowcnt_sn, colcnt++] = drval;
                    disvalsn[rowcnt_sn, colcnt++] = string.Format("=100-{0}*POWER(X{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_sn + 5, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disvalsn[rowcnt_sn, colcnt++] = string.Format("=IF(Y{0}>={1},\"优\",IF(Y{0}>={2},\"良\",IF(Y{0}>={3},\"中\",IF(Y{0}>={4},\"次\",\"差\"))))",
                        rowcnt_sn + 5,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3]);
                    ++rowcnt_sn;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    disvallq[rowcnt_lq, colcnt++] = smile;
                    disvallq[rowcnt_lq, colcnt++] = emile;
                    for (int kk = 0; kk < lqdis_name.Length; ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][RoadDiseaseTypes.DiseaseTypeDict[0][lqdis_name[kk]]];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            disvallq[rowcnt_lq, colcnt++] = type.totallength;
                        }
                        else
                        {
                            disvallq[rowcnt_lq, colcnt++] = type.totalarea;
                        }
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disvallq[rowcnt_lq, colcnt++] = drval;
                    disvallq[rowcnt_lq, colcnt++] = string.Format("=100-{0}*POWER(Y{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_lq + 5, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disvallq[rowcnt_lq, colcnt++] = string.Format("=IF(Z{0}>={1},\"优\",IF(Z{0}>={2},\"良\",IF(Z{0}>={3},\"中\",IF(Z{0}>={4},\"次\",\"差\"))))",
                        rowcnt_lq + 5,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3]);
                    ++rowcnt_lq;
                }
            }

            if (Haslqflag)
            {
                destrange = worksheet_lqhz.get_Range(String.Format("A5:AA{0}", len + 4));
                destrange.Value2 = disvallq;
                destrange = worksheet_lqhz.get_Range(String.Format("A5:AA{0}", rowcnt_lq + 4));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_lqhz, 5, 1, 27, true);
                    GlobalExcel.Reflection(worksheet_lqhz, 5, 1, 2, false);
                }
            }

            if (Hassnflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A5:Z{0}", len + 4));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A5:Z{0}", rowcnt_sn + 4));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_snhz, 5, 1, 26, true);
                    GlobalExcel.Reflection(worksheet_snhz, 5, 1, 2, false);
                }
            }
        }
        #endregion

        #region CICS
        public static void OutputDis_CICS(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            MSExcel.Worksheet _Worksheet_lqhz = null;

            string subdname = null;

            if (disval == 10)
            {
                subdname = "十米";
            }
            else
            {
                return;
            }

            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CICS\{1}\水泥混凝土路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            string destxls = string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.xlsx", path, prjdir.Name, subdname);
            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _WorkbookSN.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("水泥混凝土路面破损{0}汇总", subdname)] as MSExcel.Worksheet;

            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CICS\{1}\沥青路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            destxls = string.Format(@"{0}\{1}_沥青路面破损{2}汇总.xlsx", path, prjdir.Name, subdname);
            _WorkbookLQ = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _WorkbookLQ.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet_lqhz = _WorkbookLQ.Sheets[string.Format("沥青路面破损{0}汇总", subdname)] as MSExcel.Worksheet;

            bool Haslqflag = false;
            bool Hassnflag = false;
            WriteDisHZ2Xls_CICS(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 53);

            if (!Haslqflag)
            {
                _WorkbookLQ.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{1}_沥青路面破损{2}汇总.xlsx", path, prjdir.Name, subdname));
            }
            else
            {
                _WorkbookLQ.Save();
                _WorkbookLQ.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            if (!Hassnflag)
            {
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.xlsx", path, prjdir.Name, subdname));
            }
            else
            {
                _WorkbookSN.Save();
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteDisHZ2Xls_CICS(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            ref bool Haslqflag, ref bool Hassnflag, int borderType)
        {
            MSExcel.Range destrange;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志

            int rowcnt_sn = 0;
            int rowcnt_lq = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            object[,] disvallq = new object[len, 19];
            object[,] disvalsn = new object[len, 19];

            string[] sndis_name = { "水泥.破碎板.轻", "水泥.破碎板.重",
                                      "水泥.裂缝.轻", "水泥.裂缝.中", "水泥.裂缝.重",
                                      "水泥.板角断裂.轻", "水泥.板角断裂.中", "水泥.板角断裂.重",
                                      "水泥.错台.轻", "水泥.错台.重",
                                      "水泥.拱起",
                                       "水泥.边角剥落.轻", "水泥.边角剥落.中", "水泥.边角剥落.重",

                                      "水泥.接缝料损坏.轻", "水泥.接缝料损坏.重",
                                      "水泥.坑洞",
                                      "水泥.唧泥",
                                      "水泥.露骨",
                                      "水泥.修补.块状",
                                      "水泥.修补.条状"
                                      };

            string[] lqdis_name = { "沥青.龟裂.轻", "沥青.龟裂.中", "沥青.龟裂.重",
                                      "沥青.块状裂缝.轻", "沥青.块状裂缝.重",
                                      "沥青.纵向裂缝.轻", "沥青.纵向裂缝.重",
                                      "沥青.横向裂缝.轻", "沥青.横向裂缝.重",
                                       "沥青.沉陷.轻", "沥青.沉陷.重",
                                      "沥青.车辙.轻", "沥青.车辙.重",
                                        "沥青.波浪拥包.轻", "沥青.波浪拥包.重",
                                        "沥青.坑槽.轻", "沥青.坑槽.重",
                                        "沥青.松散.轻", "沥青.松散.重",
                                        "沥青.泛油",
                                      "沥青.修补.块状",
                                      "沥青.修补.条状"};

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[j].calcheight;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                //病害汇总表
                int colcnt = 0;
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;
                    disvalsn[rowcnt_sn, colcnt++] = smile;
                    disvalsn[rowcnt_sn, colcnt++] = emile;
                    double tmp = 0.0;
                    for (int kk = 0; kk < sndis_name.Length; ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][RoadDiseaseTypes.DiseaseTypeDict[1][sndis_name[kk]]];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            tmp += type.totallength;

                        }
                        else
                        {
                            tmp += type.totalarea;
                        }
                        if (kk == 1 || kk == 4 || kk == 7 || kk == 9 || kk == 10 || kk == 13 || kk == 15 || kk == 16 || kk == 17 || kk == 18 || kk == 19 || kk == 20)
                        {
                            disvalsn[rowcnt_sn, colcnt++] = tmp;
                            tmp = 0.0;
                        }
                    }
                    disvallq[rowcnt_lq, colcnt++] = 0;
                    disvallq[rowcnt_lq, colcnt++] = 0;
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disvalsn[rowcnt_sn, colcnt++] = drval;
                    disvalsn[rowcnt_sn, colcnt++] = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disvallq[rowcnt_lq, colcnt++] = 2.7;
                    ++rowcnt_sn;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    disvallq[rowcnt_lq, colcnt++] = smile;
                    disvallq[rowcnt_lq, colcnt++] = emile;
                    double tmp = 0.0;
                    for (int kk = 0; kk < lqdis_name.Length; ++kk)
                    {
                        RoadDiseaseType type = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][RoadDiseaseTypes.DiseaseTypeDict[0][lqdis_name[kk]]];
                        if (type.computetype == 1 || type.computetype == 3 || type.computetype == 4)
                        {
                            tmp += type.totallength;


                        }
                        else
                        {
                            tmp += type.totalarea;

                        }
                        if (kk == 2 || kk == 4 || kk == 6 || kk == 8 || kk == 10 || kk == 12 || kk == 14 || kk == 16 || kk == 18 || kk == 19 || kk == 20 || kk == 21)
                        {
                            disvallq[rowcnt_lq, colcnt++] = tmp;
                            tmp = 0.0;
                        }


                    }
                    disvallq[rowcnt_lq, colcnt++] = 0;
                    disvallq[rowcnt_lq, colcnt++] = 0;
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disvallq[rowcnt_lq, colcnt++] = drval;
                    disvallq[rowcnt_lq, colcnt++] = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    //string.Format("=100-{0}*POWER(Y{1},{2})",
                    //_PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    //rowcnt_lq + 5, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    disvallq[rowcnt_lq, colcnt++] = 2.7;
                    ++rowcnt_lq;
                }
            }

            if (Haslqflag)
            {
                destrange = worksheet_lqhz.get_Range(String.Format("A2:S{0}", len + 1));
                destrange.Value2 = disvallq;
                destrange = worksheet_lqhz.get_Range(String.Format("A2:S{0}", rowcnt_lq + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_lqhz, 2, 1, 19, true);
                    GlobalExcel.Reflection(worksheet_lqhz, 2, 1, 2, false);
                }
            }

            if (Hassnflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:S{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:S{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort)
                {
                    GlobalExcel.Reflection(worksheet_snhz, 2, 1, 19, true);
                    GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
                }
            }
        }
        #endregion
        #region csv模块12
        public static void OutputDis_HPcsv_0(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            // MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            // MSExcel.Worksheet _Worksheet_lqhz = null;

            string  subdname = "两米";

           
            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
            string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            // string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CICS\{1}\水泥混凝土路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            // string destxls = string.Format(@"{0}\{1}{2}{3}({4}{5}{3}(已识别)-路面破损-水泥路面).csv", path, prjinfo._RoadCode, prjinfo._RoadName, strDirection, prjinfo._DataDate, prjinfo._DataTime);
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\X636310115X2黑色.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-黑色-{6}.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);
            // string destxls = string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.csv", path, prjdir.Name, subdname);
            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("Sheet1")] as MSExcel.Worksheet;

            _WorkbookSN.SaveAs(destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);



            bool Haslqflag = false;
            WriteDisHZ2Xls_modle10(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road",_MarkVal);

            if (!Haslqflag)
            {
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{2} {3} {1} {4}{5}-黑色.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime));
            }
            else
            {
                _WorkbookSN.Save();
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
            }


            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteDisHZ2Xls_modle10(MSExcel.Worksheet worksheet_snhz,
       ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart1M, Disease[] arrdis,
       ref bool Haslqflag, int borderType, int ImgDis, string ImgType, string[] markInfo)
        {
            string fname = prjdir.FullName + "\\GPS2Mile.txt";
            string fnamemile0 = string.Format("{0}\\{1}Img\\Camera0\\{1}2Mile.txt", prjdir.FullName, ImgType);
            string fnamemile1 = string.Format("{0}\\{1}Img\\Camera1\\{1}2Mile.txt", prjdir.FullName, ImgType);

            string[] gpsinfostrs = null;
            ExcelGPS[] tempinfos = null;
            Dictionary<int, ExcelGPS> dicGps = new Dictionary<int, ExcelGPS>();
            List<ExcelGPS> tempinfosHanle = new List<ExcelGPS>(); //根据图像txt桩号对gps多余数据进行剔除
            if (File.Exists(fname))
            {
                gpsinfostrs = File.ReadAllLines(fname);
                tempinfos = new ExcelGPS[gpsinfostrs.Length];
                for (int i = 0; i < gpsinfostrs.Length; ++i)
                {
                    //tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                    int mile = int.Parse(gpsinfostrs[i].Split(' ')[5]);
                    if (dicGps.Keys.Contains(mile))
                        dicGps[mile] = new ExcelGPS(gpsinfostrs[i]);
                    else
                        dicGps.Add(mile, new ExcelGPS(gpsinfostrs[i]));
                }
            }

            string[] leftimgsinfo = null;
            string[] rightimgsinfo = null;
            int[] leftidx = null;
            int[] rightidx = null;
            string[] tstrs = null;
            List<string> leftimgsinfos = new List<string>();
            if (File.Exists(fnamemile0))
            {
                leftimgsinfo = File.ReadAllLines(fnamemile0);
                foreach (string str in leftimgsinfo)
                {
                    leftimgsinfos.Add(str.Split(' ').First());
                }
            }
            var temp = from a in RoadPart1M
                       where leftimgsinfos.Contains(a.mile.ToString())
                       select a;

            List<(int, string)> markDic = new List<(int, string)>();

            for (int i = 0; i < RoadPart1M.Count; i++)
            {
                if (markInfo[i] != null)
                {
                    markDic.Add((RoadPart1M[i].dmi, markInfo[i]));
                }
            }


            RoadPart1M = temp.ToList();
            MSExcel.Range destrange;
            Haslqflag = true;//有沥青路段标志

            int rowcnt_sn = 0;
            int rowcnt_lq = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = RoadPart1M.Count - 1, dlen = arrdis.Length;
            // object[,] disvallq = new object[len, 30];
            object[,] disvalsn = new object[len, 31];
            int typeidx = 0;
            bool res = false;
            string roadSplitStr = "Start";
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;

                double smile = RoadPart1M[i].mile;
                double emile = RoadPart1M[i + 1].mile;
                int curDmi = RoadPart1M[i].dmi;
                //int emile = RoadPart1M[i + 1].mile;
                //int milelength = Math.Abs(smile - emile);
                disvalsn[i, colcnt++] = prjinfo._RoadCode;//路线代码
                disvalsn[i, colcnt++] = prjinfo._DataDate.Substring(0, 4) + "/" + prjinfo._DataDate.Substring(4, 2) + "/" + prjinfo._DataDate.Substring(6, 2) + " " + prjinfo._DataTime.Substring(0, 2) + ":" + prjinfo._DataTime.Substring(2, 2);
                disvalsn[i, colcnt++] = prjinfo._RoadName;
                disvalsn[i, colcnt++] = prjinfo._DataPerson;
                disvalsn[i, colcnt++] = "dis";
                disvalsn[i, colcnt++] = "XR-M";
                disvalsn[i, colcnt++] = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

                for (int t = 0; t < markDic.Count; t++)
                {
                    if (markDic[markDic.Count - 1].Item1 <= curDmi)
                    {
                        roadSplitStr = "End";
                        break;
                    }

                    if (markDic[t].Item1 > curDmi)
                    {
                        if (t == 0)
                        {
                            roadSplitStr = "Start";
                        }
                        else
                        {
                            roadSplitStr = "Reference" + " " +t;
                        }

                        break;
                    }
                }

                disvalsn[i, colcnt++] = roadSplitStr;


                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (smile * 0.001).ToString("f3");
                disvalsn[i, colcnt++] = s1;
                disvalsn[i, colcnt++] = s2;
                disvalsn[i, colcnt++] = leftimgsinfo[i].Split('\\').Last();//
                disvalsn[i, colcnt++] = "XR-c";
                switch (RoadPart1M[i].roadtype)
                {
                    case 0:
                        disvalsn[i, colcnt++] = "沥青路面";
                        break;
                    case 1:
                        disvalsn[i, colcnt++] = "水泥路面";
                        break;
                    case 2:
                        disvalsn[i, colcnt++] = "砂石路面";
                        break;
                    default:
                        disvalsn[i, colcnt++] = "";
                        break;

                }
                if (dicGps.Keys.Contains(RoadPart1M[i].mile))
                {
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile]._latitude;
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile]._longitude;
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile]._elevation;
                }
                else if (dicGps.Keys.Contains(RoadPart1M[i].mile + 1))
                {
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile + 1]._latitude;
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile + 1]._longitude;
                    disvalsn[i, colcnt++] = dicGps[RoadPart1M[i].mile + 1]._elevation;
                }
                else
                {
                    disvalsn[i, colcnt++] = " ";
                    disvalsn[i, colcnt++] = " ";
                    disvalsn[i, colcnt++] = " ";
                }



                disvalsn[i, colcnt++] = "FALSE";
                disvalsn[i, colcnt++] = "";//事件
                #region 病害统计

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[RoadPart1M[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    try
                    {
                        if (res)
                        {

                            //  RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                            string name = RoadDiseaseTypes.roaddis[RoadPart1M[i].roadtype][typeidx].disname;
                            int roadType = RoadPart1M[i].roadtype;
                            switch (roadType)
                            {
                                case 0:
                                    if (name.Contains("龟裂"))
                                    {
                                        name = "01" + name;
                                        break;
                                    }
                                    if (name.Contains("块状裂缝"))
                                    {
                                        name = "02" + name;
                                        break;
                                    }
                                    if (name.Contains("纵向裂缝"))
                                    {
                                        name = "03" + name;
                                        break;
                                    }
                                    if (name.Contains("横向裂缝"))
                                    {
                                        name = "04" + name;
                                        break;
                                    }
                                    if (name.Contains("坑槽"))
                                    {
                                        name = "05" + name;
                                        break;
                                    }
                                    if (name.Contains("松散"))
                                    {
                                        name = "06" + name;
                                        break;
                                    }
                                    if (name.Contains("沉陷"))
                                    {
                                        name = "07" + name;
                                        break;
                                    }
                                    if (name.Contains("波浪拥包"))
                                    {
                                        name = "08" + name;
                                        break;
                                    }
                                    if (name.Contains("翻浆"))
                                    {
                                        name = "09" + name;
                                        break;
                                    }
                                    if (name.Contains("泛油"))
                                    {
                                        name = "10" + name;
                                        break;
                                    }
                                    if (name.Contains("修补"))
                                    {
                                        name = "11" + name;
                                        break;
                                    }
                                    break;
                                case 1:
                                    if (name.Contains("破碎板"))
                                    {
                                        name = "21" + name;
                                        break;
                                    }
                                    if (name.Contains("裂缝"))
                                    {
                                        name = "22" + name;
                                        break;
                                    }
                                    if (name.Contains("板角断裂"))
                                    {
                                        name = "23" + name;
                                        break;
                                    }
                                    if (name.Contains("错台"))
                                    {
                                        name = "24" + name;
                                        break;
                                    }
                                    if (name.Contains("唧泥"))
                                    {
                                        name = "25" + name;
                                        break;
                                    }
                                    if (name.Contains("边角剥落"))
                                    {
                                        name = "26" + name;
                                        break;
                                    }
                                    if (name.Contains("接缝料损坏"))
                                    {
                                        name = "27" + name;
                                        break;
                                    }
                                    if (name.Contains("坑洞"))
                                    {
                                        name = "28" + name;
                                        break;
                                    }
                                    if (name.Contains("拱起"))
                                    {
                                        name = "29" + name;
                                        break;
                                    }
                                    if (name.Contains("露骨"))
                                    {
                                        name = "30" + name;
                                        break;
                                    }
                                    if (name.Contains("修补"))
                                    {
                                        name = "31" + name;
                                        break;
                                    }
                                    break;
                            }
                            disvalsn[i, colcnt++] = name;

                            switch (arrdis[j].degree)
                            {
                                case "重":
                                    disvalsn[i, colcnt++] = "H";
                                    break;

                                case "中":
                                    disvalsn[i, colcnt++] = "M";
                                    break;
                                case "轻":
                                    disvalsn[i, colcnt++] = "L";
                                    break;
                                default:
                                    disvalsn[i, colcnt++] = "";
                                    break;
                            }
                            if (name.Contains("块状裂缝"))
                            {
                                disvalsn[i, colcnt++] = "";
                                disvalsn[i, colcnt++] = (double)arrdis[j].Area;
                            }
                            else if (name.Contains("裂缝"))
                            {
                                disvalsn[i, colcnt++] = arrdis[j].calcheight;
                                disvalsn[i, colcnt++] = "";
                            }

                            else
                            {
                                disvalsn[i, colcnt++] = arrdis[j].calcheight;
                                disvalsn[i, colcnt++] = (double)arrdis[j].Area;
                            }

                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[RoadPart1M[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        //用户需要吞掉这个报错  大于三个的直接过滤

                        //throw new Exception(arrdis[j].m_mile+"桩号一个里程范围内出现了大于三处的病害不符合报表格式！");
                    }

                    ++j;
                }
                disvalsn[i, 30] = "1";
            }
            #endregion
            if (Haslqflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:AE{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:AE{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);

            }
        }
        /// <summary>
        /// iri.csv与 iri.txt
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void OutputDis_HPcsv_IRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            // MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            // MSExcel.Worksheet _Worksheet_lqhz = null;

            //string subdname = null;

            //if (disval == 10)
            //{
            //    subdname = "两米";
            //}
            //else
            //{
            //    return;
            //}
            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
            string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\x640310115s2-iri-1.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-iri-{6}m.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);
            string destxlsTxt = string.Format(@"{0}\{2}-{3} {1} {4}{5}-iri-{6}m.txt", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);

            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("Sheet1")] as MSExcel.Worksheet;

            _WorkbookSN.SaveAs(destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            bool Haslqflag = false;
            //WriteDisHZ2Xls_modle1_1(_Worksheet_snhz, prjinfo, prjdir, _RoadPart1M, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road", destxlsTxt);
            //   WriteDisHZ2Xls_modle1_1(_Worksheet_snhz, prjinfo, prjdir, _RoadPart20M, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road", destxlsTxt);
            WriteDisHZ2Xls_modle1_1(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road", destxlsTxt, _MarkVal);

            if (!Haslqflag)
            {
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{2} {3} {1} {4}{5}-iri.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime));
                File.Delete(destxlsTxt);
            }
            else
            {
                _WorkbookSN.Save();
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
            }


            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputDis_HPcsv_Rut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            MSExcel.Worksheet _Worksheet_snhz = null;

            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
           // string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\X552310115S2-RUT.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-rut-{6}m.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);
            string destxlsTxt = string.Format(@"{0}\{2}-{3} {1} {4}{5}-rut-{6}m.txt", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);

            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("Sheet1")] as MSExcel.Worksheet;

            _WorkbookSN.SaveAs(destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            bool Haslqflag = false;
            //_LRutMaxVal
            //_RRutMaxVal
            //_SRutMaxVal
            WriteDisHZ2Xls_modle1_rut(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road", destxlsTxt);


            _WorkbookSN.Save();
            _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputDis_HPcsv_2(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            // MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            // MSExcel.Worksheet _Worksheet_lqhz = null;

            string subdname = null;

            if (disval == 10)
            {
                subdname = "两米";
            }
            else
            {
                return;
            }
            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
            string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            // string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CICS\{1}\水泥混凝土路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            // string destxls = string.Format(@"{0}\{1}{2}{3}({4}{5}{3}(已识别)-路面破损-水泥路面).csv", path, prjinfo._RoadCode, prjinfo._RoadName, strDirection, prjinfo._DataDate, prjinfo._DataTime);
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\X552310115S2-PBI.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-PBI-{6}.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime,disval);
            // string destxls = string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.csv", path, prjdir.Name, subdname);
            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("Sheet1")] as MSExcel.Worksheet;

            _WorkbookSN.SaveAs(destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);



            bool Haslqflag = false;
            WriteDisHZ2Xls_modle1_2(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road");

            if (!Haslqflag)
            {
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
                File.Delete(string.Format(@"{0}\{2} {3} {1} {4}{5}-PBI.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime));
            }
            else
            {
                _WorkbookSN.Save();
                _WorkbookSN.Close(Type.Missing, Type.Missing, Type.Missing);
            }


            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        //iri
        private static void WriteDisHZ2Xls_modle1_1(MSExcel.Worksheet worksheet_snhz,
          ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart, Disease[] arrdis,
          ref bool Haslqflag, int borderType, int ImgDis, string ImgType, string textPath, string[] marks)
        {

            ExcelGPS[] dicGps = null;
            int rowcnt_sn = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart, ref dicGps);
            MSExcel.Range destrange;
            Haslqflag = true;//有沥青路段标志
            string errlog = prjdir.FullName + "\\errlog.txt";

            double[] IRI_LIRIVal = null;
            double[] IRI_RIRIVal = null;
            if (prjinfo._IsIRIMTD)
            {
                GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, RoadPart, ref _SpeedVal);
                GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, RoadPart, ref IRI_LIRIVal, ref IRI_RIRIVal, _Setting.IsWarning);

            }
            int len = RoadPart.Count - 1, dlen = arrdis.Length;
            // object[,] disvallq = new object[len, 30];
            object[,] disvalsn = new object[len, 28];
            StreamWriter sw = new StreamWriter(textPath);
            string strTxt = string.Empty;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;
                double smile = RoadPart[i].mile;
                double emile = RoadPart[i + 1].mile;
                //int emile = RoadPart1M[i + 1].mile;
                //int milelength = Math.Abs(smile - emile);
                disvalsn[i, colcnt++] = prjinfo._RoadCode;//路线代码

                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (smile * 0.001).ToString("f3");
                string s3 = (emile * 0.001).ToString("f3");
                disvalsn[i, colcnt++] = s1;
                disvalsn[i, colcnt++] = s2;

                if (prjinfo._IsIRIMTD)
                {
                    strTxt = string.Format("{0}        {1}        {2}", s1, s3, IRI_LIRIVal[i].ToString("f3"));
                    if (prjinfo._IsDIRIMTD)
                    {
                        strTxt = string.Format("{0}        {1}        {2}", s1, s3, ((IRI_LIRIVal[i] + IRI_RIRIVal[i]) / 2).ToString("f3"));
                        disvalsn[i, 3] = IRI_RIRIVal[i].ToString("f3");
                        disvalsn[i, 4] = IRI_LIRIVal[i].ToString("f3");
                        disvalsn[i, 5] = String.Format("=ROUND(AVERAGE(D{0}:E{0}),3)", i + 2);
                    }
                    else
                    {
                        disvalsn[i, 3] = "X";
                        disvalsn[i, 4] = IRI_LIRIVal[i].ToString("f3");
                
                        disvalsn[i, 5] = IRI_LIRIVal[i].ToString("f3");
                    }
                }
                sw.WriteLine(strTxt);

                //disvalsn[i, colcnt++] = "X";//右IRI
                //disvalsn[i, colcnt++] = "X";//左Iri
                //disvalsn[i, colcnt++] = "X";//平均IRI
                colcnt = 6;
                disvalsn[i, colcnt++] = "X";//车道iri
                disvalsn[i, colcnt++] = "X";//中央车道
                disvalsn[i, colcnt++] = "X";//右 HATI
                disvalsn[i, colcnt++] = "X";//左 HATI
                disvalsn[i, colcnt++] = "X";//平均 HATI
                disvalsn[i, colcnt++] = "X";//右 RN
                disvalsn[i, colcnt++] = "X";//左 RN
                disvalsn[i, colcnt++] = "X";//平均 RN
                disvalsn[i, colcnt++] = "X";//NAASRA
                disvalsn[i, colcnt++] = "X";//Bump Int
                disvalsn[i, colcnt++] = _SpeedVal[i];
                disvalsn[i, colcnt++] = dicGps[i]._latitude;
                disvalsn[i, colcnt++] = dicGps[i]._longitude;
                disvalsn[i, colcnt++] = dicGps[i]._elevation;
                disvalsn[i, colcnt++] = "FALSE";//正在计算 GPS 定位
                disvalsn[i, colcnt++] =marks[i];//事件
                string date = prjinfo._DataDate;
                string[] format = { "yyyyMMddHHmmss" };
                DateTime resDateTime;
                if (DateTime.TryParseExact(prjinfo._DataDate + prjinfo._DataTime, format, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out resDateTime))
                {
                    date = resDateTime.ToString("yyyy/MM/dd");
                    disvalsn[i, colcnt++] = date;//测量日
                    date = resDateTime.ToString("HH:mm:ss");
                    disvalsn[i, colcnt++] = date;//测量时间
                }
                else
                {
                    disvalsn[i, colcnt++] = "";
                    disvalsn[i, colcnt++] = "";
                }



                disvalsn[i, colcnt++] = "Dis";//测量名称


                disvalsn[i, colcnt++] = "XR-M";//车辆名
                disvalsn[i, colcnt++] = prjinfo._DataPerson;//操作员名
                disvalsn[i, colcnt++] = "Y";//Device Sync Flag
            }

            if (Haslqflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:AB{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:AB{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                    //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                    //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                    destrange = worksheet_snhz.get_Range(String.Format("A2:AB{0}", rowcnt_sn + 1));
                    MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("B2:B{0}", len + 1));//按桩号排序
                    GlobalExcel.ReflectionColnumDescending(worksheet_snhz, destrange, sortrange);

                }
            }
            sw.Close();
            sw.Dispose();
        }


        private static void WriteDisHZ2Xls_modle1_rut(MSExcel.Worksheet worksheet_snhz,
          ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart, double[] LRutMaxVal, double[] RRutMaxVal, double[] SRutMaxVal,
          ref bool Haslqflag, int borderType, int ImgDis, string ImgType, string textPath)
        {
            
          
            ExcelGPS[] dicGps = null;
            int rowcnt_sn = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart, ref dicGps);
            MSExcel.Range destrange;
            Haslqflag = true;//有沥青路段标志
            string errlog = prjdir.FullName + "\\errlog.txt";

           
            if (prjinfo._IsIRIMTD)
            {
                GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, RoadPart, ref _SpeedVal); 
            }
            int len = RoadPart.Count - 1 ;
            object[,] disvalsn = new object[len, 18];
            StreamWriter sw = new StreamWriter(textPath);
            string strTxt = string.Empty;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;
                double smile = RoadPart[i].mile;
                double emile = RoadPart[i + 1].mile;
                //int emile = RoadPart1M[i + 1].mile;
                //int milelength = Math.Abs(smile - emile);
                disvalsn[i, colcnt++] = prjinfo._RoadName;//路线代码

                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (smile * 0.001).ToString("f3");
                string s3 = (emile * 0.001).ToString("f3");
                disvalsn[i, colcnt++] = s1;
                disvalsn[i, colcnt++] = s2;

                
                   
                strTxt = string.Format("{0}        {1}        {2}", s1, s3, SRutMaxVal[i].ToString("f3"));
                disvalsn[i, colcnt++] = RRutMaxVal[i].ToString("f3");
                disvalsn[i, colcnt++] = LRutMaxVal[i].ToString("f3"); 
                disvalsn[i, colcnt++] = SRutMaxVal[i].ToString("f3");
                   
                 
                sw.WriteLine(strTxt);

               
                disvalsn[i, colcnt++] = _SpeedVal[i];
                disvalsn[i, colcnt++] = dicGps[i]._latitude;
                disvalsn[i, colcnt++] = dicGps[i]._longitude;
                disvalsn[i, colcnt++] = dicGps[i]._elevation;
                disvalsn[i, colcnt++] = "FALSE";//正在计算 GPS 定位
                disvalsn[i, colcnt++] = "X";//事件
                string date = prjinfo._DataDate;
                string[] format = { "yyyyMMddHHmmss" };
                DateTime resDateTime;
                if (DateTime.TryParseExact(prjinfo._DataDate + prjinfo._DataTime, format, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out resDateTime))
                {
                    date = resDateTime.ToString("yyyy/MM/dd");
                    disvalsn[i, colcnt++] = date;//测量日
                    date = resDateTime.ToString("HH:mm:ss");
                    disvalsn[i, colcnt++] = date;//测量时间
                }
                else
                {
                    disvalsn[i, colcnt++] = "";
                    disvalsn[i, colcnt++] = "";
                }



                disvalsn[i, colcnt++] = prjinfo._RoadCode;//测量名称 
                disvalsn[i, colcnt++] = "TTJ 120";//车辆名
                disvalsn[i, colcnt++] = prjinfo._DataPerson;//操作员名
                disvalsn[i, colcnt++] = "Y";//Device Sync Flag
            }

            if (Haslqflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:R{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:R{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                    //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                    //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                    destrange = worksheet_snhz.get_Range(String.Format("A2:R{0}", rowcnt_sn + 1));
                    MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("B2:B{0}", len + 1));//按桩号排序
                    GlobalExcel.ReflectionColnumDescending(worksheet_snhz, destrange, sortrange);

                }
            }
            sw.Close();
            sw.Dispose();
        }

        private static void WriteDisHZ2Xls_modle1_2(MSExcel.Worksheet worksheet_snhz,
        ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, Disease[] arrdis,
        ref bool Haslqflag, int borderType, int ImgDis, string ImgType)
        {

            ExcelGPS[] dicGps = null;
            int rowcnt_sn = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);

            MSExcel.Range destrange;
            Haslqflag = true;//有沥青路段标志
            string errlog = prjdir.FullName + "\\errlog.txt";

            double[] IRI_LIRIVal = null;
            double[] IRI_RIRIVal = null;
            double[] cA = null;//加速度
            if (prjinfo._IsIRIMTD)
            {
                //在此处理后所有都变为2M桩号对应的数据(名称未改易产生歧异特此注释)
                GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, RoadPart10M, ref IRI_LIRIVal, ref IRI_RIRIVal, _Setting.IsWarning);

                //由于速度.txt  记录为十米一次，所以当以2m为间隔时会导致   数值重复
                GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, RoadPart10M, ref _SpeedVal);


                //GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 10, prjinfo._Direction, _RoadGradeStr, ref _RoadPart10, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                GlobalExcel.GetDeltaHVal(prjinfo, prjdir, RoadPart10M, 0, ref _LDeltaHVal);
                if (prjinfo._IsDIRIMTD)
                {
                    GlobalExcel.GetDeltaHVal(prjinfo, prjdir, RoadPart10M, 1, ref _RDeltaHVal);
                }
                GlobalExcel.GetPBVal(prjinfo, prjdir, _RoadPart, RoadPart10M, ref _PBIVal, _PBIThresh, _LDeltaHVal, _RDeltaHVal, 0, ref _DeltaHVal);
                GlobalExcel.GetMarkInfo(prjinfo, prjdir, RoadPart10M, ref _MarkVal10);
                GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, RoadPart10M, ref _SpeedVal10);
                GlobalExcel.calculateAcceleratedSpeed(RoadPart10M, _SpeedVal, dicGps, ref cA);
            }
            int len = RoadPart10M.Count - 1, dlen = arrdis.Length;
            // object[,] disvallq = new object[len, 30];
            object[,] disvalsn = new object[len, 17];
            int typeidx = 0;
            bool res = false;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;
                double drval = 0;
                double smile = RoadPart10M[i].mile;
                double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (smile * 0.001).ToString("f3");
                disvalsn[i, colcnt++] = i; //唯一编号id
                disvalsn[i, colcnt++] = prjinfo._RoadCode;//检测名称
                string date = prjinfo._DataDate;
                string[] format = { "yyyyMMddHHmmss" };
                DateTime resDateTime;
                if (DateTime.TryParseExact(prjinfo._DataDate + prjinfo._DataTime, format, System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out resDateTime))
                {
                    date = resDateTime.ToString("yyyy/MM/dd");
                    disvalsn[i, colcnt++] = date;//测量日
                    date = resDateTime.ToString("HH:mm:ss");
                    disvalsn[i, colcnt++] = date;//测量时间
                }
                else
                {
                    disvalsn[i, colcnt++] = "";
                    disvalsn[i, colcnt++] = "";
                }
                disvalsn[i, colcnt++] = "LeadOut";//参考点 ?
                disvalsn[i, colcnt++] = s1;//起点
                disvalsn[i, colcnt++] = s1;//终点
                disvalsn[i, colcnt++] = _SpeedVal[i];//速度
                disvalsn[i, colcnt++] = cA[i].ToString("f3");//加速度

                disvalsn[i, colcnt++] = dicGps[i]._latitude;
                disvalsn[i, colcnt++] = dicGps[i]._longitude;
                disvalsn[i, colcnt++] = dicGps[i]._elevation;

                double LDeltaHVal = _LDeltaHVal[i];
                double RDeltaHVal = 0;
                if (prjinfo._IsDIRIMTD)
                {
                    RDeltaHVal = _RDeltaHVal[i];

                    if (RDeltaHVal < _PBIThresh[0])
                    {
                        disvalsn[i, colcnt++] = "None";
                    }
                    else if (RDeltaHVal < _PBIThresh[1])
                    {
                        disvalsn[i, colcnt++] = "Low";
                    }
                    else if (RDeltaHVal < _PBIThresh[2])
                    {
                        disvalsn[i, colcnt++] = "Moderate";
                    }
                    else
                    {
                        disvalsn[i, colcnt++] = "High";
                    }
                }


                if (LDeltaHVal < _PBIThresh[0])
                {
                    disvalsn[i, colcnt++] = "None";
                }
                else if (LDeltaHVal < _PBIThresh[1])
                {
                    disvalsn[i, colcnt++] = "Low";
                }
                else if (LDeltaHVal < _PBIThresh[2])
                {
                    disvalsn[i, colcnt++] = "Moderate";
                }
                else
                {
                    disvalsn[i, colcnt++] = "High";
                }


                disvalsn[i, colcnt++] = (RDeltaHVal * 0.1).ToString("f3");
                disvalsn[i, colcnt++] = (LDeltaHVal * 0.1).ToString("f3");
            }
            if (Haslqflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:P{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:P{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                    // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = worksheet_snhz.get_Range(String.Format("A2:P{0}", rowcnt_sn + 1));
                    MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("F2:F{0}", len + 1));//按桩号排序
                    GlobalExcel.ReflectionColnumDescending(worksheet_snhz, destrange, sortrange);

                }
            }
        }
        #endregion

        #region 自动化报表
        /// <summary>
        /// 空间定位数据
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        public static void outPutAutoTest_0(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\空间定位数据.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_空间定位数据.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["空间定位数据"] as MSExcel.Worksheet;
            writeAutoTestXls_gpsData(_Worksheet_, prjinfo, prjdir, _RoadPart);



            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// 平整度报表
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        public static void outPutAutoTest_1(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\路面平整度自动化检测数据.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_路面平整度.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["路面平整度"] as MSExcel.Worksheet;
            writeAutoTestXls_iri(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }




        public static void outPutAutoTest_2(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\路面平整度自动化检测原始数据.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_路面平整度原始数据.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["路面平整度原始数据"] as MSExcel.Worksheet;
            writeAutoTestXls_hightData(_Worksheet_iri, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// 断面高程  平整度
        /// </summary>
        private static void writeAutoTestXls_hightData(MSExcel.Worksheet workSheet, List<MilePartD> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count - 1;
            object[,] disvalsn = new object[len, 4];
            bool hasRdlta = rDeltaHVal != null ? true : false;
            MSExcel.Range destrange;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;
                double smile = roadpart10[i].mile;
                double emile = roadpart10[i + 1].mile;
                disvalsn[i, colcnt++] = (smile * 0.001).ToString("f4");
                if (lDeltaHVal != null)
                {
                    disvalsn[i, colcnt++] = lDeltaHVal[i];

                    if (hasRdlta && rDeltaHVal.Length > i)
                        disvalsn[i, colcnt++] = rDeltaHVal[i];
                    else
                        disvalsn[i, colcnt++] = "0";
                    // disvalsn[i, colcnt++] = RIRIMeanVal[i];
                    disvalsn[i, colcnt++] = (speed[i] * 1000 / 3600).ToString("f2");
                }
            }
            destrange = workSheet.get_Range(String.Format("A2:D{0}", len + 1));
            destrange.Value2 = disvalsn;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = workSheet.get_Range(String.Format("A2:C{0}", len + 1));
                MSExcel.Range sortrange = workSheet.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(workSheet, destrange, sortrange);
            }
        }

        private static void writeAutoTestXls_iri(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            object[,] disvalsn = new object[len, 3];
            MSExcel.Range destrange;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                disvalsn[i, colcnt++] = smile * 0.001;
                disvalsn[i, colcnt++] = LIRIMeanVal[i].ToString("f2");
                // disvalsn[i, colcnt++] = RIRIMeanVal[i];
                disvalsn[i, colcnt++] = (SpeedVal[i] * 1000 / 3600).ToString("f2");


            }
            destrange = _Worksheet.get_Range(String.Format("A2:C{0}", len + 1));
            destrange.Value2 = disvalsn;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = _Worksheet.get_Range(String.Format("A2:C{0}", len + 1));
                MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
            }
        }
        private static void writeAutoTestXls_iri_JiangXi(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            object[,] disvalsn = new object[len, 5];
            MSExcel.Range destrange;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                disvalsn[i, colcnt++] = smile * 0.001;
                disvalsn[i, colcnt++] = LIRIMeanVal[i].ToString("f2");
                if ( prjinfo._IsDIRIMTD)
                {
                  
                        disvalsn[i, colcnt++] = RIRIMeanVal[i].ToString("f2");
                        if (_Setting.RQIJudgeType == 0)
                        {
                            disvalsn[i, colcnt++] = String.Format("=ROUND(AVERAGE(B{0}:C{0}),5)", i + 2);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            disvalsn[i, colcnt++] = String.Format("=ROUND(MAX(B{0}, C{0}),5)", i + 2);
                        }
                    
                  
                }
                else
                {

                    disvalsn[i, colcnt++] = "";
                    disvalsn[i, colcnt++] = LIRIMeanVal[i].ToString("f2"); ;
                }
               


               
                disvalsn[i, colcnt++] =_RoadConfig.DetectWidth;

                

            }
            destrange = _Worksheet.get_Range(String.Format("A2:E{0}", len + 1));
            destrange.Value2 = disvalsn;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = _Worksheet.get_Range(String.Format("A2:E{0}", len + 1));
                MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:E{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
            }
        }
        private static void writeAutoTestXls_gpsData(MSExcel.Worksheet worksheet_snhz,
         ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;

            int rowcnt = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, 5];

            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                disvalsn[i, _colcnt++] = dicGps[i]._latitude;
                disvalsn[i, _colcnt++] = dicGps[i]._elevation;
                disvalsn[i, _colcnt++] = "√";
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", len + 1));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", rowcnt + 1));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", rowcnt + 1));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        #endregion

        #region 报送格式
        #region 报送格式2023农村路

        #endregion
        public static void outPutAutoTest_5(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\传输文件格式（巴东）\DR破损率csv检测结果数据表格.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\DR.xlsx", path);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["DR"] as MSExcel.Worksheet;
            WritePCI2Xls_5(_Worksheet, prjinfo, prjdir, _RoadPart, 2, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        private static void WritePCI2Xls_5(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
       List<MilePart> roadpart, int DataStartXlsxRow, Disease[] arrdis)
        {



            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 8];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }


                vallist[i, 0] = prjinfo._City;


                vallist[i, 1] = prjinfo._District;
                vallist[i, 2] = prjinfo._RoadCode;
                vallist[i, 3] = prjinfo._RoadName;
                vallist[i, 4] = smile;
                vallist[i, 5] = emile;
                vallist[i, 6] = milelength;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 7] = drval;
            }
            destrange = worksheet.get_Range(String.Format("A{0}:H{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 9, true);
                GlobalExcel.Reflection(worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }


        public static void outPutAutoTest_6(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\传输文件格式（巴东）\IRI平整度csv检测结果数据表格.xlsx",
              System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\IRI.xlsx", path);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["IRI平整度csv检测结果数据表格"] as MSExcel.Worksheet;
            WriteIRI2Xls_6(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, 2, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        private static void WriteIRI2Xls_6(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
        List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 8];
            for (int i = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                vallist[i, 0] = prjinfo._City;
                vallist[i, 1] = prjinfo._District;
                vallist[i, 2] = prjinfo._RoadCode;
                vallist[i, 3] = prjinfo._RoadName;
                vallist[i, 4] = smile;
                vallist[i, 5] = emile;
                vallist[i, 6] = milelength;
                vallist[i, 7] = LIRIVal[i];
            }
            destrange = _Worksheet.get_Range(String.Format("A{0}:H{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                //GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                //  GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }
        public static void outPutAutoTest_7(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\传输文件格式（巴东）\IRI平整度csv检测原始数据表格.xlsx",
            System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\平整度.xlsx", path);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["IRI平整度csv检测原始数据表格"] as MSExcel.Worksheet;
            WriteIRI2Xls_7(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, 2, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        private static void WriteIRI2Xls_7(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
  List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, int DataStartXlsxRow, int borderType)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 5];
            for (int i = 0; i < len; i++)
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                vallist[i, 0] = i + 1;
                vallist[i, 1] = smile;
                vallist[i, 2] = emile;
                vallist[i, 3] = LIRIVal[i]; ;
                if (RIRIVal != null && RIRIVal.Length > i)
                {
                    vallist[i, 4] = RIRIVal[i];
                }
                else
                {
                    vallist[i, 4] = 0;
                }
            }
            destrange = _Worksheet.get_Range(String.Format("A{0}:E{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, borderType);

            if (_Setting.IsExcelSort)
            {
                //   GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
                //   GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            }
        }
        public static void outPutAutoTest_8(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\传输文件格式（巴东）\空间定位数据csv检测原始数据表格.xlsx",
           System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\桩号与GPS.xlsx", path);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["空间定位数据csv检测原始数据表格"] as MSExcel.Worksheet;
            writeAutoTestXls_gpsData_8(_Worksheet, prjinfo, prjdir, _RoadPart);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        private static void writeAutoTestXls_gpsData_8(MSExcel.Worksheet worksheet_snhz,
     ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart)
        {
            ExcelGPS[] dicGps = null;

            int rowcnt = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart, ref dicGps);
            int len = RoadPart.Count;
            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, 6];


            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                int smile = RoadPart[i].mile;
                //double emile = RoadPart10M[i + 1].mile;

                disvalsn[i, _colcnt++] = i + 1;
                disvalsn[i, _colcnt++] = smile;
                if (dicGps != null && dicGps != null)
                {
                    disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._latitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._elevation;
                }

                disvalsn[i, _colcnt++] = 0;
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:F{0}", len + 1));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:F{0}", rowcnt + 1));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {


                //destrange = worksheet_snhz.get_Range(String.Format("A2:F{0}", rowcnt + 1));
                //MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                //GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        #endregion
        #region 国检数据转换
        public static void Convent_Rut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "车辙_RD.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_RD(_Worksheet_, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, 4);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        public static void Convent_Bump(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "跳车_PB.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_Bump(_Worksheet_, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }

        
        //public static void Convent_Acc(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        //{
        //    string excelFileName = "路面平整度原始数据_加速度.xlsx";
        //    string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
        //        System.Windows.Forms.Application.StartupPath, excelFileName);

        //    //拼接文件名称
        //    string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
        //    string time = prjinfo._DataDate + prjinfo._DataTime;
        //    string excelName = excelFileName.Split('.')[0].Split('_').Last();
        //    fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
        //    string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


        //    //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
        //    MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
        //        true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //    ExcelGPS[] dicGps = null;

        //    GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart10M, ref dicGps);

        //    double[] cA = null;//加速度
        //    if (prjinfo._IsIRIMTD)
        //    {
        //        GlobalExcel.calculateAcceleratedSpeed(_RoadPart10M, _SpeedVal, dicGps, ref cA);


        //        MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
        //        writeAutoTestXls_Acc(_Worksheet_, prjinfo, prjdir, _RoadPart, _SpeedVal, cA, 4);
        //    }
        //    _Workbook.Save();
        //    _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
        //    int generation = System.GC.GetGeneration(excelApp);
        //    System.GC.Collect(generation);//垃圾回收
        //    System.GC.WaitForPendingFinalizers();
        //    SingleProject.XlsxToCsv(Destxls);
        //}

        public static void Convent_Mpd(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "磨耗_MPD.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);
            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_Mpd(_Worksheet_, prjinfo, prjdir, _RoadPart, _LMPDMeanVal, _RMPDMeanVal, _CMPDMeanVal, 4);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        public static void Convent_Lbi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "空间定位数据_LBI.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_Lbi(_Worksheet_, prjinfo, prjdir, _RoadPart);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }

        public static void Convent_Iri_JiangXi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\江西\平整度_IRI.xlsx",
                System.Windows.Forms.Application.StartupPath);
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);
            // string Destxls = string.Format(@"{0}\{1}_路面平整度.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_iri_JiangXi(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers(); SingleProject.XlsxToCsv(Destxls);
        }

        public static void Convent_Iri(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\平整度_IRI.xlsx",
                System.Windows.Forms.Application.StartupPath);
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);
            // string Destxls = string.Format(@"{0}\{1}_路面平整度.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_iri(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers(); SingleProject.XlsxToCsv(Destxls);
        }
 
        private static void writeAutoTestXls_Lbi(MSExcel.Worksheet worksheet_snhz,
   ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;

            int rowcnt = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            MSExcel.Range destrange;
            object[,] disvalsn = new object[len,5];


            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                if (dicGps != null && dicGps != null)
                {
                    disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._latitude;
                   disvalsn[i, _colcnt++] = dicGps[i]._elevation;
                }

                disvalsn[i, _colcnt++] = "√";
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", len + 1));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", rowcnt + 1));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = worksheet_snhz.get_Range(String.Format("A2:E{0}", rowcnt + 1));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }

      
        private static void writeAutoTestXls_gc_gj(MSExcel.Worksheet _Worksheet, List<MilePartD> roadpart10, DirectoryInfo path, ProjectInfo proj, double[] speedVal)
        {
            List<float[]> allData = new List<float[]>();
            #region 获得高程数据list<float[]>
            int valnum = 0;
            switch (proj._RutMode)
            {
                case 0: valnum = 3; break;
                case 1: valnum = 1; break;
                case 2: valnum = 3; break;
                default: break;
            }
            int cameracnt = valnum == 1 ? 2 : 1;

            try
            {
                string[] dat = { string.Format("{0}\\camera0\\data", path.FullName), string.Format("{0}\\camera1\\data", path.FullName) };
                string[] process = { string.Format("{0}\\RUT\\camera0\\data", path.FullName), string.Format("{0}\\RUT\\camera1\\data", path.FullName) };
                string[] cfg = { string.Format("{0}\\camera0\\rutcfg.ini", path.FullName), string.Format("{0}\\camera1\\rutcfg.ini", path.FullName) };
                short hpix = short.Parse(IniFileOpr.ReadIniData("camera", "hpixel", "2048", cfg[0]));
                for (int i = 0; i < cameracnt; i++)
                {

                    IniFiles rutcfg = new IniFiles(cfg[i]);
                    int m, n, temp, j = 0;
                    float[] objlas = new float[hpix];
                    short[] profile = new short[hpix];
                    string _dtwname = "";
                    float[] tobjlas = new float[hpix];

                    string[] _dats = Directory.GetFiles(dat[i], "*.dat");
                    float _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
                    allData.Clear();
                    for (j = 0; j < _dats.Length; ++j)
                    {
                        _dtwname = _dats[j].Substring(_dats[j].LastIndexOf('\\') + 1);
                        _dtwname = _dtwname.Substring(0, _dtwname.IndexOf('.'));
                        using (FileStream frstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Open))
                        {
                            // fsbar = fsbar / frstream.Length;
                            temp = hpix * 2;
                            byte[] rbarr = new byte[hpix * 2];

                            while (frstream.Read(rbarr, 0, temp) > 0)
                            {

                                Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                for (m = 0, n = 0; m < hpix; ++m)
                                {
                                    // if (profile[m] != 0x7FFF)
                                    // {
                                    objlas[n] = profile[m] / _scaleval;
                                    if (proj._RutMode == 2)
                                    {
                                        objlas[n] = -objlas[n];
                                    }

                                    tobjlas[n] = objlas[m];
                                    ++n;


                                    //  }
                                }

                                allData.Add(tobjlas.Select(t => t).ToArray());
                            }

                        }
                    }
                }

                #endregion
                int len = roadpart10.Count - 1 <= allData.Count - 1 ? roadpart10.Count - 1 : allData.Count - 1;
                object[,] disvalsn = new object[len, 22];
                MSExcel.Range destrange;
                for (int i = 0; i < len; i++)
                {
                    int colcnt = 0;
                    double smile = roadpart10[i].mile;
                    double emile = roadpart10[i + 1].mile;
                    disvalsn[i, colcnt++] = smile * 0.001;
                    //这个n=100可能有问题
                    for (int n = 100; n < allData[i].Length && n < 2100; n += 100)
                    {
                        disvalsn[i, colcnt++] = (int)(allData[i][n] * 10); //原始高程单位 mm
                    }
                    disvalsn[i, colcnt++] = speedVal[i];
                }
                destrange = _Worksheet.get_Range(String.Format("A2:V{0}", len + 1));
                destrange.Value2 = disvalsn;
                GlobalExcel.SetBorderLine(destrange, 53);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("原始高程表导出失败，可能是车辙计算数据缺失！\n" + ex.Message);
            }

        }

        public static void Convent_TP(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_TP.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_gc_gj(_Worksheet_, _RoadPartF, prjdir, prjinfo, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        public static void Convent_LP(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_LP.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            //writeAutoTestXls_gc(_Worksheet_, _RoadPart, prjdir, prjinfo, _SpeedVal);
            writeAutoTestXls_hightData(_Worksheet_, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }

        public static void Convent_LP_hebei(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_LP.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            //writeAutoTestXls_gc(_Worksheet_, _RoadPart, prjdir, prjinfo, _SpeedVal);
            writeAutoTestXls_hightData(_Worksheet_, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        public static void Convent_Damage(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            string[] srcxlsArr = new string[] {
                string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\识别结果_沥青.xlsx",
                System.Windows.Forms.Application.StartupPath),
                string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\识别结果_水泥.xlsx",
                System.Windows.Forms.Application.StartupPath)

            };
            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int col;
                string srcxls;
                int[] tableCol = new[] { 13, 13 };
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    i = 1;
                }
                else
                {
                    i = 2;
                }
                col = tableCol[i];
                srcxls = srcxlsArr[i];
                string tableNamePart = Path.GetFileNameWithoutExtension(srcxls).Substring(0, 7);
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;

                string tempName;
                if (prjinfo._Direction == 1)
                {
                    tempName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, endMile, tableNamePart);
                }
                else
                {
                    tempName = string.Format("{0}-{3}-{2}-{1}", fileName, startMile, endMile, tableNamePart);
                }


                string Destxls = string.Format(@"{0}\{1}.xlsx", path, tempName);

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;


                WriteDis_Damage2Xls_gj(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, i, col, startMileD, endMileStrD);


                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                SingleProject.XlsxToCsv(Destxls);
            }
        }

        public static void Convent_Damage_JiangXi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            string[] srcxlsArr = new string[] {
                string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\江西\识别结果_沥青.xlsx",
                System.Windows.Forms.Application.StartupPath),
                string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\江西\识别结果_水泥.xlsx",
                System.Windows.Forms.Application.StartupPath)

            };
            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int col;
                string srcxls;
                int[] tableCol = new[] { 14, 14 };
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    i = 1;
                }
                else
                {
                    i = 2;
                }
                col = tableCol[i];
                srcxls = srcxlsArr[i];
                string tableNamePart = Path.GetFileNameWithoutExtension(srcxls).Substring(0, 7);
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;

                string tempName;
                if (prjinfo._Direction == 1)
                {
                    tempName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, endMile, tableNamePart);
                }
                else
                {
                    tempName = string.Format("{0}-{3}-{2}-{1}", fileName, startMile, endMile, tableNamePart);
                }


                string Destxls = string.Format(@"{0}\{1}.xlsx", path, tempName);

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;


                WriteDis_Damage2Xls_gj_JiangXi(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, i, col, startMileD, endMileStrD);


                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                SingleProject.XlsxToCsv(Destxls);
            }
        }


        public static void Convent_Damage2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, Encoding encoding)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int i = 0;
                string tableNamePart = "";
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "标准沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "标准水泥";
                    i = 1;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;

                string tempName;
                string time = prjinfo._DataDate + prjinfo._DataTime;
                // 文件命名无需tableNamePart？？？
                if (prjinfo._Direction == 1)
                {
                    tempName = string.Format("{0}-DR-{1}-{2}-{3}", fileName, startMile, endMile, time);
                }
                else
                {
                    tempName = string.Format("{0}-DR-{2}-{1}-{3}", fileName, startMile, endMile, time);

                }
                string Destxls = string.Format(@"{0}\{1}.txt", path, tempName);
                WriteDis_Damage2Xls_gj2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD,encoding);
            }
        }

        public static void Convent_DamageStandard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");
            
            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int i = 0;
                string tableNamePart = "";
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "标准沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "标准水泥";
                    i = 1;
                } 
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;

                string tempName;
                string time = prjinfo._DataDate + prjinfo._DataTime;
                // 文件命名无需tableNamePart？？？
                tempName = string.Format("{0}-DR-{1}-{2}-{3}", fileName, startMile, endMile, time);
              
                string Destxls = string.Format(@"{0}\{1}.txt", path, tempName);
                WriteDis_Damage2Xls_standard(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }
        private static void WriteDis_Damage2Xls_standard(string path,
      ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,int roadType, double sMile, double eMile)
        {
            List<string> allDatas= new List<string>();
            if (roadType==0)
            {
                allDatas.Add(GJTitles.LQ_2018_BIG); 
            }
            else
            {
                allDatas.Add(GJTitles.SN_2018_BIG);
            }     allDatas[0].Replace("\t","");
      
            int len = roadpart.Count - 1,
                dlen = arrdis.Length;
            
            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                if (prjinfo._Direction > 0)
                {
                    if (smile >= sMile && emile <= eMile)
                    {

                    }
                    else
                        continue;
                }
                else
                {
                    if (smile <= sMile && emile >= eMile)
                    {

                    }
                    else
                        continue;
                } 
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                j = 0;

                for (int d = 0; d < arrdis.Length; d++)
                {
                    if (d < dlen && ((prjinfo._Direction > 0 && arrdis[d].m_mile >= smile && arrdis[d].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[d].m_mile <= smile && arrdis[d].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[d].RoadType, arrdis[d].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[d].Area;
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[d].calcheight;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[d].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                 
                        }
                    }
                } 
                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);
                    string line;
                    //病害汇总表
                    string dr = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth.ToString(), dr);
                    
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    { 
                        if (roadType == 0)
                        {
                            if (dis == 0)
                            {
                                
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                      + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea;
                                line +=  area.ToString("f2");
                                dis++; dis++;

                            }
                            else if (dis == 3 || dis == 5 || dis == 7 || dis == 9 || dis == 11 || dis == 13 || dis == 15 || dis == 17 || dis == 20)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                if (dis ==5||dis==7||dis==11 )
                                {
                                    area = RoadDiseaseTypes.roaddis[roadType][dis].totallength + RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength;
                                }

                                line +=  area.ToString("f2");
                                dis++;
                            }
                            else
                            {
                                line += ",";
                                line += RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2"); 
                            }
                        }
                        else
                        {
                            if (dis == 0 || dis == 8 || dis == 14 || dis == 19)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                if (dis ==8 || dis == 14)
                                {
                                    area = RoadDiseaseTypes.roaddis[roadType][dis].totallength + RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength;
                                }
                                line += area.ToString("f2"); 
                                dis++;

                            }
                            else if (dis == 2 || dis == 5 || dis == 11)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea;
                                if (dis == 2 || dis == 11)
                                {
                                    area = RoadDiseaseTypes.roaddis[roadType][dis].totallength + RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totallength; 
                                }


                                line +=area.ToString("f2");
                             
                                dis++; dis++;
                            }
                            else
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                if (dis == 17)
                                {
                                    area = RoadDiseaseTypes.roaddis[roadType][dis].totallength  ;
                                }
                                line += area.ToString("f2"); 


                            }
                        }
                    }
                    allDatas.Add(line);
                }
            }
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, GlobalExcel.getConvert_GJ_TempAreaStr(12)));

            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, GlobalExcel.getConvert_GJ_TempAreaStr(12)));

            }
            File.WriteAllLines(path, allDatas);

        }


        private static void WriteDis_Damage2Xls_gj2024_ChongQing(string path,
    ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile ,Encoding encoding)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_2018_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SN_2018_BIG);
            }
            allDatas[0].Replace("\t", "");

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                if (prjinfo._Direction > 0)
                {
                    if (smile >= sMile && emile <= eMile)
                    {

                    }
                    else
                        continue;
                }
                else
                {
                    if (smile <= sMile && emile >= eMile)
                    {

                    }
                    else
                        continue;
                }
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                j = 0;

                for (int d = 0; d < arrdis.Length; d++)
                {
                    if (d < dlen && ((prjinfo._Direction > 0 && arrdis[d].m_mile >= smile && arrdis[d].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[d].m_mile <= smile && arrdis[d].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[d].RoadType, arrdis[d].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[d].Area;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[d].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);

                        }
                    }
                }
                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);
                    string line;
                    //病害汇总表
                    string dr = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth.ToString(), dr);

                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        if (roadType == 0)
                        {
                            if (dis == 0)
                            {

                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                      + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea;
                                line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                dis++; dis++;

                            }
                            else if (dis == 3 || dis == 5 || dis == 7 || dis == 9 || dis == 11 || dis == 13 || dis == 15 || dis == 17 || dis == 20)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                dis++;
                            }
                            else
                            {
                                line += ",";
                                line += RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f3");
                            }
                        }
                        else
                        {
                            if (dis == 0 || dis == 8 || dis == 14 || dis == 19)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                dis++;

                            }
                            else if (dis == 2 || dis == 5 || dis == 11)
                            {
                                line += ",";
                                double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea;
                                line += area == 0 ? area.ToString("f3") : area.ToString("f3");

                                dis++; dis++;
                            }
                            else
                            {
                                line += ",";
                                line += RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f3");
                            }
                        }
                    }
                    allDatas.Add(line);
                }
            } 
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0.000", "0.000", "0.000", "0.000", "0.000")); 
            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0.000", "0.000", "0.000", "0.000", "0.000", "0.000")); 
            }
            else
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0.000", "0.000", "0.000", "0.000")); 
            }
            File.WriteAllLines(path, allDatas, encoding);
        }

        /// <summary>
        /// 国检转换 破损
        /// </summary>
        /// <param name="_Worksheet"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="roadpart"></param>
        /// <param name="arrdis"></param>
        /// <param name="borderType"></param>
        /// <param name="roadType"></param>
        /// <param name="cluCount"></param>
        /// <param name="refDatas"></param>
        /// <param name="has"></param>
        private static void WriteDis_Damage2Xls_gj(MSExcel.Worksheet _Worksheet,
        ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType, int cluCount, double sMile, double eMile)
        {
            MSExcel.Range destrange;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowCount = 0;
            int len = roadpart.Count - 1,
                dlen = arrdis.Length;
            object[,] datas = new object[len, cluCount];
            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                if (prjinfo._Direction > 0)
                {
                    if (smile >= sMile && emile <= eMile)
                    {

                    }
                    else
                        continue;
                }
                else
                {
                    if (smile <= sMile && emile >= eMile)
                    {

                    }
                    else
                        continue;
                }



                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                j = 0;

                for (int d = 0; d < arrdis.Length; d++)
                {
                    if (d < dlen && ((prjinfo._Direction > 0 && arrdis[d].m_mile >= smile && arrdis[d].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[d].m_mile <= smile && arrdis[d].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[d].RoadType, arrdis[d].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[d].Area;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[d].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                    }
                }


                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    colcnt = 0;
                    datas[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    datas[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");


                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {


                        if (roadType == 0)
                        {
                            if (dis == 0)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                      + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                  .ToString("f2");

                                dis++; dis++;

                            }
                            else if (dis == 3 || dis == 5 || dis == 7 || dis == 9 || dis == 11 || dis == 13 || dis == 15 || dis == 17 || dis == 20)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                        else
                        {
                            if (dis == 0 || dis == 8 || dis == 14 || dis == 19)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;

                            }
                            else if (dis == 2 || dis == 5 || dis == 11)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                .ToString("f2");
                                dis++; dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                    }
                    rowCount++;
                }
            }
            if (has)
            {
                destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, SingleProject.chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {
                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, SingleProject.chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", rowCount + 1));//按桩号排序
                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }
        private static void WriteDis_Damage2Xls_gj_JiangXi(MSExcel.Worksheet _Worksheet,
     ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType, int cluCount, double sMile, double eMile)
        {
            MSExcel.Range destrange;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowCount = 0;
            int len = roadpart.Count - 1,
                dlen = arrdis.Length;
            object[,] datas = new object[len, cluCount];
            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                if (prjinfo._Direction > 0)
                {
                    if (smile >= sMile && emile <= eMile)
                    {

                    }
                    else
                        continue;
                }
                else
                {
                    if (smile <= sMile && emile >= eMile)
                    {

                    }
                    else
                        continue;
                }



                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                j = 0;

                for (int d = 0; d < arrdis.Length; d++)
                {
                    if (d < dlen && ((prjinfo._Direction > 0 && arrdis[d].m_mile >= smile && arrdis[d].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[d].m_mile <= smile && arrdis[d].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[d].RoadType, arrdis[d].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[d].Area;
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totallength += arrdis[d].calcheight;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[d].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                    }
                }


                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    colcnt = 0;
                    datas[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    datas[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");


                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {


                        if (roadType == 0)
                        {
                            if (dis == 0)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                      + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                  .ToString("f2");

                                dis++; dis++;

                            }
                            else if (dis == 3 || dis == 9 || dis == 11 || dis == 13 || dis == 15 || dis == 17 )
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;
                            }
                            else if (dis == 5 || dis == 7)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totallength + RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength)
                                  .ToString("f2");
                                dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                        else
                        {
                            if (dis == 0 || dis == 8 || dis == 14)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;

                            }
                            else if ( dis == 5 || dis == 11)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                .ToString("f2");
                                dis++; dis++;
                            }
                            else if(dis == 2 )
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totallength + RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength
                                   + RoadDiseaseTypes.roaddis[roadType][dis + 2].totallength)
                               .ToString("f2");
                                dis++; dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                    }
                    rowCount++;
                }
            }
            if (has)
            {
                destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, SingleProject.chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {
                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, SingleProject.chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", rowCount + 1));//按桩号排序
                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }
        private static void OutDisPart(string path, string fileName, object[,] datas, int rowCount, string type)
        {
            string srcXls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\识别结果_{1}.xlsx",
                System.Windows.Forms.Application.StartupPath, type);

            // 桩号 破损率  判断破损率就可以知道是否存在病害
            //1. 先分段
            List<int> splitIndex = new List<int>();
            splitIndex.Add(0);
            List<List<double>> hasDisList = new List<List<double>>();
            //过滤出所有具有病害的行
            for (int i = 0; i < rowCount; i++)
            {
                if (datas[i, 0] == null)
                {
                    break;
                }
                if (double.Parse(datas[i, 1].ToString()) != 0)
                {
                    List<double> data = new List<double>();
                    for (int t = 0; t < 13; t++)
                    {
                        data.Add(double.Parse(datas[i, t].ToString()));
                    }
                    hasDisList.Add(data);
                }

            }

            //获取索引
            for (int i = 0; i < hasDisList.Count - 1; i++)
            {
                if (i - 1 > 0)
                {
                    int currentMile = (int)(hasDisList[i][0] * 1000);
                    int frontMile = (int)(hasDisList[i - 1][0] * 1000);
                    if (Math.Abs(currentMile - frontMile) != 10) //判断是否分段了
                    {
                        splitIndex.Add(i - 1);
                        splitIndex.Add(i);
                    }
                }
            }
            splitIndex.Add(hasDisList.Count - 1);
            if (hasDisList.Count <= 0)
            {
                return;
            }
            //出表
            for (int t = 1; t < splitIndex.Count; t += 2)
            {
                //拼接文件名称
                string startMile = (hasDisList[splitIndex[t - 1]][0]).ToString("f3");
                string endMile = "";
                if (t != splitIndex.Count - 1)
                {
                    endMile = (hasDisList[splitIndex[t]][0] + 0.010).ToString("f3");
                }
                else
                {
                    endMile = (hasDisList[splitIndex[t]][0]).ToString("f3");
                }

                string tempFileName = string.Format("{0}-识别结果-{1}-{2}", fileName, startMile, endMile);
                string Destxls = string.Format(@"{0}\{1}.xlsx", path, tempFileName);
                int rows = splitIndex[t] - splitIndex[t - 1] + 1; //行数
                var tempApp = new MSExcel.Application()
                {
                    Visible = true,
                    DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                    AlertBeforeOverwriting = false
                };

                var CollectBookTemp = tempApp.Workbooks.Open(srcXls, Type.Missing,
                                      true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                CollectBookTemp.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //CollectBookTemp.Save();
                MSExcel.Worksheet sheet = CollectBookTemp.Sheets["sheet1"] as MSExcel.Worksheet;
                object[,] allData = new object[rows, 13];
                int row = -1;
                for (int i = splitIndex[t - 1]; i <= splitIndex[t]; i++)
                {
                    row++;
                    for (int y = 0; y < hasDisList[i].Count; y++)
                    {
                        allData[row, y] = hasDisList[i][y];
                    }
                }


                var rangeTemp = sheet.get_Range(string.Format("A2:M{0}", rows + 1));
                rangeTemp.Value2 = allData;
                CollectBookTemp.Save();
                CollectBookTemp.Close(Type.Missing, Type.Missing, Type.Missing);
                CWB_ExcelHelper.disposeExcel(ref tempApp);
                SingleProject.XlsxToCsv(Destxls);
            }



        }
        private static void WriteDisLQOrSN_Damage2Xls(MSExcel.Worksheet _Worksheet,
        ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType, int cluCount, ref object[,] allDatas, ref bool has, out int outRowCount)
        {
            MSExcel.Range destrange;
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'j', 'k', 'L', 'M', 'N' };
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowCount = 0;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] datas = new object[len, cluCount];
            int typeidx = 0;
            bool res = false;
            // int rowCount = 2;
            int colcnt = 1;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                //沥青
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    colcnt = 0;
                    datas[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    datas[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {


                        if (roadType == 0)
                        {
                            if (dis == 0)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                      + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                  .ToString("f2");

                                dis++; dis++;

                            }
                            else if (dis == 3 || dis == 5 || dis == 7 || dis == 9 || dis == 11 || dis == 13 || dis == 15 || dis == 17 || dis == 20)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                        else
                        {
                            if (dis == 0 || dis == 8 || dis == 14 || dis == 19)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea)
                                  .ToString("f2");
                                dis++;

                            }
                            else if (dis == 2 || dis == 5 || dis == 11)
                            {
                                datas[rowCount, colcnt++] = (RoadDiseaseTypes.roaddis[roadType][dis].totalarea + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea
                                    + RoadDiseaseTypes.roaddis[roadType][dis + 2].totalarea)
                                .ToString("f2");
                                dis++; dis++;
                            }
                            else
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                            }
                        }
                    }
                    rowCount++;
                }
            }
            if (has)
            {
                destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                destrange.Value2 = datas;
                allDatas = datas;

                GlobalExcel.SetBorderLine(destrange, 53);
                if (_Setting.IsExcelSort && prjinfo._Direction > 0) { }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", rowCount + 3));//按桩号排序
                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
            outRowCount = rowCount;
        }
        private static void writeAutoTestXls_RD(MSExcel.Worksheet worksheet_snhz,
    ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LRutVal, double[] RRutVal, double[] SRutVal, int lastCol)
        {
            if (!prjinfo._IsRut)
                return;
            int rowcnt = 0;
            int len = RoadPart10M.Count;
            int lenRut = LRutVal.Length;
            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, lastCol];
            lastCol--;
            if (lenRut < len)
            {
                for (int i = 0; i < lenRut; i++)
                {
                    int _colcnt = 0;
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    disvalsn[i, _colcnt++] = s1;
                    disvalsn[i, _colcnt++] = LRutVal[i].ToString("f1");
                    disvalsn[i, _colcnt++] = RRutVal[i].ToString("f1");
                    disvalsn[i, _colcnt++] = SRutVal[i].ToString("f1");

                    rowcnt++;
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    int _colcnt = 0;
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    disvalsn[i, _colcnt++] = s1;
                    disvalsn[i, _colcnt++] = LRutVal[i];
                    disvalsn[i, _colcnt++] = RRutVal[i];
                    disvalsn[i, _colcnt++] = SRutVal[i];
                    rowcnt++;
                }
            }



            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", len + 1, SingleProject.chars[lastCol]));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        
        private static void writeAutoTestXls_Bump(MSExcel.Worksheet worksheet_snhz,
 ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
        
            int rowcnt = 0;
            int len = RoadPart10M.Count - 1;

            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, lastCol];
            lastCol--;
            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                disvalsn[i, _colcnt++] = PBIVal[i][1];
                disvalsn[i, _colcnt++] = PBIVal[i][2];
                disvalsn[i, _colcnt++] = PBIVal[i][3];
                double lValue = 0;
                double rValue = 0;
                if (i < LDeltaHVal.Length)
                {
                    lValue = LDeltaHVal[i];
                }
                if (prjinfo._IsDIRIMTD)
                {
                    if (i < RDeltaHVal.Length)
                    {
                        rValue = RDeltaHVal[i];
                    }
                }
                double maxVlaue = Math.Max(lValue, rValue);

                disvalsn[i, _colcnt++] = maxVlaue.ToString("f2");

                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", len + 1, SingleProject.chars[lastCol]));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        private static void writeAutoTestXls_Mpd(MSExcel.Worksheet worksheet_snhz,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int lastCol)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
            int rowcnt = 0;
            int len = RoadPart10M.Count - 1;

            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, lastCol];
            lastCol--;
            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                disvalsn[i, _colcnt++] = LMTDVal[i];
                disvalsn[i, _colcnt++] = CMTDVal[i];
                disvalsn[i, _colcnt++] = RMTDVal[i];
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", len + 1, SingleProject.chars[lastCol]));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        private static void writeAutoTestXls_Acc(MSExcel.Worksheet worksheet_snhz,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] SpeedVal, double[] cA, int lastCol)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
            int rowcnt = 0;
            int len = RoadPart10M.Count - 1;

            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, lastCol];
            lastCol--;
            for (int i = 0; i < len; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                disvalsn[i, _colcnt++] = cA[i];
                disvalsn[i, _colcnt++] = cA[i];
                if (SpeedVal != null)
                {
                    disvalsn[i, _colcnt++] = SpeedVal[i];
                }

                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", len + 1, SingleProject.chars[lastCol]));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
            GlobalExcel.SetBorderLine(destrange, 53);
            _Setting.IsExcelSort = true;
            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 30, true);
                // GlobalExcel.Reflection(worksheet_snhz, 2, 1, 2, false);
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:{1}{0}", rowcnt + 1, SingleProject.chars[lastCol]));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        /// <summary>
        /// 沥青破损
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void OutputSNDamage(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\识别结果.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_水泥路面损坏.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //bug
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面破损"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisLQDamage2Xls(_Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, 1, 7, ref has);
            if (has)
            {
                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
            }
            else
            {
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);

                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                File.Delete(Destxls);
            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_Worksheet"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="roadpart"></param>
        /// <param name="arrdis"></param>
        /// <param name="borderType"></param>
        /// <param name="roadType">0沥青 1水泥</param>
        /// rowCount 开始行数
        private static void WriteDisLQDamage2Xls(MSExcel.Worksheet _Worksheet,
          ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType, int cluCount, ref bool has)
        {
            MSExcel.Range destrange;
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I' };
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowCount = 0;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] datas = new object[len, cluCount];
            int typeidx = 0;
            bool res = false;
            // int rowCount = 2;
            int colcnt = 1;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                //沥青
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
                if (roadpart[i].roadtype == roadType)
                {
                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    colcnt = 0;
                    datas[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    datas[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                        // _Worksheet.Cells[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                    }
                    rowCount++;
                }
            }
            if (has)
            {


                destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {

                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount + 1, chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", rowCount + 2));//按桩号排序

                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }

        public static void OutputDis_THSum(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\调绘报表模板.xlsx",
            System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害调绘统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;



            WriteDisLB2Xls_roadpart_THSum(path, _Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart, _GPSInfo);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        private static void WriteDisLB2Xls_roadpart_THSum_HighGps(string outPath, MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
 DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart, ExcelGPS[] gpsInfos)
        {

            List<ValueTuple<double, double, double>> highGPS_Result = new List<ValueTuple<double, double, double>>();
            string prjPath = prjinfo._PrjPath;
            int dir = prjinfo._Direction;
            //List<Disease> diseases = getPorjectAllDiseases(prjPath, dir);

            bool UseHighGPS = true;
            string outHighGpstxtPath = prjinfo._PrjPath + "/HighGps2Mile.txt";
            if (!File.Exists(outHighGpstxtPath))
            {
                UseHighGPS = false;  // 未找到高精度数据文件，回退到默认方法
            }
            else
            {
                List<string> highGpsTxts = File.ReadAllLines(outHighGpstxtPath).ToList();
                List<(double, GPSInfo)> highGpss = new List<(double, GPSInfo)>();
                foreach (var line in highGpsTxts)
                {
                    string[] strings = line.Split(',');
                    GPSInfo gpsInfo = new GPSInfo();
                    gpsInfo._longitude = double.Parse(strings[0]);
                    gpsInfo._latitude = double.Parse(strings[1]);
                    gpsInfo._elevation = double.Parse(strings[2]);
                    highGpss.Add((double.Parse(strings[3]), gpsInfo));
                }
                if (highGpsTxts.Count > 0)
                {
                    HighAccuracyPositioning.UpdateAllImg(prjinfo._PrjPath + "\\RoadImg\\Camera0");
                    foreach (Disease tdis in arrdis)
                    {
                        HighAccuracyDisease dis = new HighAccuracyDisease
                        {
                            DiseaseName = tdis.RoadDisType
                        };
                        if (string.IsNullOrEmpty(dis.DiseaseName))
                        {
                            continue;
                        }
                        int half_x = tdis.rect.X + tdis.rect.Width / 2;
                        int half_y = tdis.rect.Y + tdis.rect.Height / 2;
                        //var point = points[index];
                        double dDiseaseLon = 0, dDiseaseLat = 0, dDiseaseH = 0; //当前像素
                        HighAccuracyPositioning.getHighAccPosition(_Setting.gpsformat, highGpss, _Setting.equipType, prjinfo._PrjPath, tdis.m_mile, half_x, half_y, _RoadConfig.ImageWidth
                            , _RoadConfig.ImageHeight, prjinfo._Direction, _RoadConfig.RealWidth, _RoadConfig.RealHeight,
                                ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH);
                        highGPS_Result.Add((dDiseaseLon, dDiseaseLat, dDiseaseH));

                    }
                }
            }


            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 11];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;
            string thDisPicPath = outPath + "\\调绘病害图片";
            if (Directory.Exists(thDisPicPath))
            {
                Directory.Delete(thDisPicPath, true);
            }
            Directory.CreateDirectory(thDisPicPath);

            

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        if (string.IsNullOrEmpty(prjinfo._RoadNum))
                        {
                            vallist[rowcnt, 1] = "1";
                        }
                        else
                        {
                            vallist[rowcnt, 1] = prjinfo._RoadNum;
                        }

                        vallist[rowcnt, 2] = roadpart[i].roadtype == 0 ? "沥青" : roadpart[i].roadtype == 1 ? "水泥" : "砂石";
                        vallist[rowcnt, 3] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 4] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 4] = "无";
                        }
                        vallist[rowcnt, 5] = arrdis[j].rect.Height * _RoadConfig.HeightScale;

                        vallist[rowcnt, 6] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        if (UseHighGPS)
                        {
                            var (x, y, z) = highGPS_Result[j];
                            vallist[rowcnt, 8] = y.ToString("f7");
                            vallist[rowcnt, 9] = x.ToString("f7");
                            vallist[rowcnt, 10] = z.ToString("f3"); // 即使使用高精度，也保留与此前相同的显示格式
                        }
                        else
                        {
                            vallist[rowcnt, 8] = gpsInfos[i]._latitude;
                            vallist[rowcnt, 9] = gpsInfos[i]._longitude;
                            vallist[rowcnt, 10] = gpsInfos[i]._elevation;
                        }
                        
                        var nowDis = arrdis[j];
                        string picPath = prjinfo._PrjPath + Path.Combine(arrdis[j].imgpath, arrdis[j].imgname);

                        var picRange = _Worksheet.Range[$"L{3 + rowcnt}"];
                        // picRange.RowHeight = 27.682 * 18;
                        float widthC = 27.682f * 1.05f;

                        double ratio;
                        using (System.Drawing.Bitmap map = new System.Drawing.Bitmap(picPath))
                        {
                            ratio = (double)map.Height / (double)map.Width;
                        }

                        picRange.RowHeight = widthC * 3.112f;
                        picRange.ColumnWidth = widthC;
                        Framework.Office.Excel.CWB_ExcelHelper.InsertPicture_Compress(picRange, _Worksheet, picPath, ratio);
                        var hyperRange = _Worksheet.Range[$"M{3 + rowcnt}"];

                        string tempStr = "\\" + arrdis[j].imgpath.Split('\\').Last();
                        string thPicPath = thDisPicPath;
                        string hyperPath = "调绘病害图片\\" + arrdis[j].imgname;


                        Directory.CreateDirectory(thPicPath);
                        thPicPath += "\\" + arrdis[j].imgname;

                        File.Copy(picPath, thPicPath, true);

                        hyperRange.ColumnWidth = widthC * 4;
                        //var o= hyperRange.Select();


                        _Worksheet.Hyperlinks.Add(hyperRange, hyperPath);

                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A3:K{0}", dlen + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Qufen_dis_degree == 1)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
                //if (_Setting.Out_roadimg == 0)
                //{
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
                //}
            }
            else
            {
                //if (_Setting.Out_roadimg == 0)
                //{
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 12]).EntireColumn.Delete();
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                //}
            }
            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, 3, 1, 14, true);
            //}
        }

        private static void WriteDisLB2Xls_roadpart_THSum(string outPath, MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
  DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart, ExcelGPS[] gpsInfos)
        {
            WriteDisLB2Xls_roadpart_THSum_HighGps(outPath, _Worksheet, prjinfo,
prjdir, arrdis, roadpart, gpsInfos);

            /*
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            if (len < 1 || dlen < 1)
                return;

            string errlog = prjdir.FullName + "\\errlog.txt";
            object[,] vallist = new object[dlen, 11];
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;
            string thDisPicPath = outPath + "\\调绘病害图片";
            if (Directory.Exists(thDisPicPath))
            {
                Directory.Delete(thDisPicPath, true);
            }
            Directory.CreateDirectory(thDisPicPath);
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = arrdis[j].RoadDisType.Split('.');
                        vallist[rowcnt, 0] = arrdis[j].m_mile;
                        if (string.IsNullOrEmpty(prjinfo._RoadNum))
                        {
                            vallist[rowcnt, 1] = "1";
                        }
                        else
                        {
                            vallist[rowcnt, 1] = prjinfo._RoadNum;
                        }

                        vallist[rowcnt, 2] = roadpart[i].roadtype == 0 ? "沥青" : roadpart[i].roadtype == 1 ? "水泥" : "砂石";
                        vallist[rowcnt, 3] = s[0];
                        if (s.Length > 1)
                        {
                            vallist[rowcnt, 4] = s[1];
                        }
                        else
                        {
                            vallist[rowcnt, 4] = "无";
                        }
                        vallist[rowcnt, 5] = arrdis[j].rect.Height * _RoadConfig.HeightScale;

                        vallist[rowcnt, 6] = arrdis[j].rect.Width * _RoadConfig.WidthScale;
                        vallist[rowcnt, 7] = arrdis[j].Area;
                        vallist[rowcnt, 8] = gpsInfos[i]._latitude;
                        vallist[rowcnt, 9] = gpsInfos[i]._longitude;
                        vallist[rowcnt, 10] = gpsInfos[i]._elevation;
                        var nowDis = arrdis[j];
                        string picPath = prjinfo._PrjPath + Path.Combine(arrdis[j].imgpath, arrdis[j].imgname);

                        var picRange = _Worksheet.Range[$"L{3 + rowcnt}"];
                        // picRange.RowHeight = 27.682 * 18;
                        float widthC = 27.682f * 1.05f;

                        double ratio;
                        using (System.Drawing.Bitmap map = new System.Drawing.Bitmap(picPath))
                        {
                            ratio = (double)map.Height / (double)map.Width;
                        }

                        picRange.RowHeight = widthC * 3.112f;
                        picRange.ColumnWidth = widthC;
                        Framework.Office.Excel.CWB_ExcelHelper.InsertPicture(picRange, _Worksheet, picPath, ratio);
                        var hyperRange = _Worksheet.Range[$"M{3 + rowcnt}"];

                        string tempStr = "\\" + arrdis[j].imgpath.Split('\\').Last();
                        string thPicPath = thDisPicPath;
                        string hyperPath = "调绘病害图片\\" + arrdis[j].imgname;


                        Directory.CreateDirectory(thPicPath);
                        thPicPath += "\\" + arrdis[j].imgname;

                        File.Copy(picPath, thPicPath, true);

                        hyperRange.ColumnWidth = widthC * 4;
                        //var o= hyperRange.Select();


                        _Worksheet.Hyperlinks.Add(hyperRange, hyperPath);

                        ++rowcnt;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
            }

            MSExcel.Range destrange = _Worksheet.get_Range(String.Format("A3:K{0}", dlen + 2));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Qufen_dis_degree == 1)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
                //if (_Setting.Out_roadimg == 0)
                //{
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
                //}
            }
            else
            {
                //if (_Setting.Out_roadimg == 0)
                //{
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 12]).EntireColumn.Delete();
                //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                //}
            }
            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, 3, 1, 14, true);
            //}
            */
        }
        public static void Convent_Standard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "空间定位数据_LBI.txt";
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);


            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            allDatas.Add(GJTitles.LbiTitle);

            GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref dicGps);
            int len = _RoadPart.Count;
            for (int i = 0; i < len; i++)
            {
                double smile = _RoadPart[i].mile;
                string s1 = (smile * 0.001).ToString("f3");

                if (dicGps != null && dicGps != null)
                {
                    string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"));
                    allDatas.Add(line);
                }
            }
            File.WriteAllLines(Destxls, allDatas);
        }
        #endregion
        #region 报送格式2023农村路
        public static void Convent_Lbi2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "空间定位数据_LBI.txt"; 

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            allDatas.Add(GJTitles.LbiTitle);

            writeAutoTestXls_Lbi2023(Destxls,prjinfo, prjdir, _RoadPart);
        }


        private static void writeAutoTestXls_Lbi2023(string fileName,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            if (_Setting.gjLbiOutHight)
            {
                allDatas.Add("桩号,X,Y,Z,有效性,桩号核对,");
            }
            else
            {
                allDatas.Add(GJTitles.LbiTitle);
            }

            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            for (int i = 0; i < len; i++)
            {
                if (_Setting.gjLbiOutHight)
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = smile.ToString();

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1,double.Parse( dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), double.Parse(dicGps[i]._elevation).ToString("f6"), "A", "0", "0");
                        allDatas.Add(line);
                    }
                }
                else
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = (smile * 0.001).ToString("f3");

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse (dicGps[i]._longitude).ToString("f6"),double.Parse( dicGps[i]._latitude).ToString("f6"), "A");
                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }
        public static void Convent_Damage2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName,Encoding en)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int i = 0;
                string tableNamePart = "";
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "标准沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "标准水泥";
                    i = 1;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                string time = prjinfo._DataDate + prjinfo._DataTime;

                string tempName;
                if (prjinfo._Direction == 1)
                {
                    tempName = string.Format("{0}-DR-{1}-{2}-{3}", fileName, startMile, endMile, time);
                }
                else
                {
                    tempName = string.Format("{0}-DR-{2}-{1}-{3}", fileName, startMile, endMile, time);

                }
                string Destxls = string.Format(@"{0}\{1}.txt", path, tempName);
                WriteDis_Damage2Xls_standard(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }
        public static void Convent_Lbi2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);


            writeAutoTestXls_Lbi2024(Destxls, prjinfo, prjdir, _RoadPart);
        }
        public static void Convent_Lbi2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);


            writeAutoTestXls_Lbi2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart);
        }
        private static void writeAutoTestXls_Lbi2024(string fileName,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            if (_Setting.gjLbiOutHight)
            {
                allDatas.Add("桩号,X,Y,Z,有效性,桩号核对,");
            }
            else
            {
                allDatas.Add(GJTitles.LbiTitle);
            }

            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            for (int i = 0; i < len; i++)
            {
                if (_Setting.gjLbiOutHight)
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = smile.ToString();

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), double.Parse(dicGps[i]._elevation).ToString("f6"), "A", "0", "0");
                        allDatas.Add(line);
                    }
                }
                else
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = (smile * 0.001).ToString("f3");

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"));

                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }

        private static void writeAutoTestXls_Lbi2024_ChongQing(string fileName,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            GJTitles.LbiTitle = "桩号（km）,经度,纬度,海拔高度（m）,有效性";
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            if (_Setting.gjLbiOutHight)
            {
                allDatas.Add("桩号,X,Y,Z,有效性,桩号核对,");
            }
            else
            {
                allDatas.Add(GJTitles.LbiTitle);
            }

            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            for (int i = 0; i < len; i++)
            {
                if (_Setting.gjLbiOutHight)
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = smile.ToString();

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), double.Parse(dicGps[i]._elevation).ToString("f6"), "A", "0", "0");
                        allDatas.Add(line);
                    }
                }
                else
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = (smile * 0.001).ToString("f3");

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), dicGps[i]._elevation,"A");

                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }
        public static void Convent_Iri2023( string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
       
         
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);  
            writeAutoTestXls_iri2023(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
        }

        public static void Convent_Iri2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadName = fileName.Substring(0, fileName.Length - 1);
            string strDirc = fileName.Substring(fileName.Length - 1, 1);
            //文件名
            fileName = string.Format("{0}-{1}-IRI-{2}-{3}", strRoadName, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_iri2024_HuNan(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
        }
        public static void Convent_Iri2024_Standard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            // string strRoadName = fileName.Substring(0, 4);
            string strDirc = fileName.Substring(10, 1);
            //文件名
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);

            writeAutoTestXls_iri2024_Standard(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal,  _SpeedVal);
        }
        private static void writeAutoTestXls_iri2024_Standard(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {//双平整度
                    double value = Math.Max(LIRIMeanVal[i], RIRIMeanVal[i]);
                    line = string.Join(",", (smile * 0.001).ToString("f2"), LIRIMeanVal[i].ToString("f2"), RIRIMeanVal[i].ToString("f2"), value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                else
                {//单平整度
                    line = string.Join(",", (smile * 0.001).ToString("f2"), LIRIMeanVal[i].ToString("f2"), "0.00", LIRIMeanVal[i].ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        private static void writeAutoTestXls_iri2024_HuNan(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {//双平整度
                    double value = Math.Max(LIRIMeanVal[i], RIRIMeanVal[i]);
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), RIRIMeanVal[i].ToString("f2"), value.ToString("f2"));
                }
                else
                {//单平整度
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), "0.00", LIRIMeanVal[i].ToString("f2"));
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        public static void Convent_Lbi2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadNum = fileName.Substring(0, fileName.Length - 1);
            string strDirc = fileName.Substring(fileName.Length - 1, 1);
            fileName = string.Format("{0}{1}-GPS-{2}-标准格式-{3}", strRoadNum, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Lbi2024_HuNan(Destxls, prjinfo, prjdir, _RoadPart);
        }

        private static void writeAutoTestXls_Lbi2024_HuNan(string fileName,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            allDatas.Add(GJTitles.LbiTitle);

            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps);
            int len = RoadPart10M.Count;
            for (int i = 0; i < len; i++)
            {
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 1.0).ToString("f3");

                if (dicGps != null && dicGps != null)
                {
                    string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"),
                        double.Parse(dicGps[i]._latitude).ToString("f6")
                        , double.Parse(dicGps[i]._elevation).ToString("f6"), "A", "0", "0");
                    allDatas.Add(line);
                }
            }
            File.WriteAllLines(fileName, allDatas);
        }

        public static void Convent_Iri2024_LiaoNing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {


            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_iri2024_LiaoNing(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

        }
        private static void writeAutoTestXls_iri2024_LiaoNing(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {

                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {
                    double value = (LIRIMeanVal[i] + RIRIMeanVal[i]) / 2;
                    line = string.Join(",", (smile * 0.001).ToString("f3"),value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                else
                {
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        private static void writeAutoTestXls_iri2023(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {

                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {
                    double value =Math.Max(LIRIMeanVal[i] , RIRIMeanVal[i]) ;
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), RIRIMeanVal[i].ToString("f2"), value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                else
                {
                    line = string.Join(",",( smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), "0.00", LIRIMeanVal[i].ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }


        public static void Convent_LPStandard( string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_LP.xlsx";
          
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
             
            writeAutoTestXls_hightDataStandard(Destxls, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);
            
        }
        private static void writeAutoTestXls_hightDataStandard(string path, List<MilePartD> roadpart, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.RIFileJgTitle);
            bool hasRdlta = rDeltaHVal != null ? true : false;

            for (int i = 0; i < len; i++)
            {

                double smile = roadpart[i].mile;
                double emile = roadpart[i + 1].mile;
                string line = "";
                line += (smile * 0.001).ToString("f4");
                if (lDeltaHVal != null)
                {
                    line += ",";
                    line += lDeltaHVal[i];

                    if (hasRdlta && rDeltaHVal.Length > i)
                    {
                        line += ",";
                        line += rDeltaHVal[i];
                    }
                    else
                    {
                        line += ",";
                        line += "0";
                    }
                    line += ",";
                    line += (speed[i] * 1000 / 3600).ToString("f2");
                }
                else
                {
                    line += ",0,0,0";
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);

        }

        public static void Convent_Rut2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {//
            //拼接文件名称
            //GA04-A-RD-059.000-20230718082358.csv
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadNum = fileName.Substring(0,fileName.Length-1);
            string strDirc = fileName.Substring(fileName.Length - 1,1);
            fileName = string.Format("{0}-{1}-RD-{2}-{3}", strRoadNum, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_RD2024_HuNan(Destxls, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, 4);
        }

        private static void writeAutoTestXls_RD2024_HuNan(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LRutVal, double[] RRutVal, double[] SRutVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;

            int len = RoadPart10M.Count;
            int lenRut = LRutVal.Length;

            datas.Add("桩号(km),左车辙RD1(mm),右车辙RD2(mm),路面车辙RD(mm)");

            if (lenRut < len)
            {
                for (int i = 0; i < lenRut; i++)
                {
                    double smile = RoadPart10M[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, LRutVal[i].ToString("f3"), RRutVal[i].ToString("f3"), SRutVal[i].ToString("f3"));
                    datas.Add(line);
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    
                    double smile = RoadPart10M[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, LRutVal[i].ToString("f3"), RRutVal[i].ToString("f3"), SRutVal[i].ToString("f3"));
                    datas.Add(line);
                }
            }
            File.WriteAllLines(path, datas);
        }

        public static void Convent_Rut2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "车辙_RD.txt";
           
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_RD2023(Destxls, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, 4); 
        }

        public static void Convent_Rut_Standard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "车辙_RD.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);


            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;

            int len = _RoadPart.Count;
            int lenRut = _LRutMeanVal.Length;

            datas.Add("桩号(km),左车辙RD1(mm),右车辙RD2(mm),路面车辙RD(mm)");

            if (lenRut < len)
            {
                for (int i = 0; i < lenRut; i++)
                {
                    double smile = _RoadPart[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, _LRutMeanVal[i].ToString("f1"), _RRutMeanVal[i].ToString("f1"), _SRutMeanVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                  
                    double smile = _RoadPart[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, _LRutMeanVal[i].ToString("f1"), _RRutMeanVal[i].ToString("f1"), _SRutMeanVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            File.WriteAllLines(Destxls, datas);

        }

        public static void Convent_Rut2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "车辙_RD.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_RD2024_chongqing(Destxls, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, 4);
        }

        private static void writeAutoTestXls_RD2024_chongqing(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LRutVal, double[] RRutVal, double[] SRutVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;

            int len = RoadPart10M.Count;
            int lenRut = LRutVal.Length;

            datas.Add("桩号(km),路面车辙RD(mm)");

            if (lenRut < len)
            {
                for (int i = 0; i < lenRut; i++)
                {
                    double smile = RoadPart10M[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, SRutVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    double smile = RoadPart10M[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, SRutVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            File.WriteAllLines(path, datas);
        }
        private static void writeAutoTestXls_RD2023 (string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LRutVal, double[] RRutVal, double[] SRutVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;
          
            int len = RoadPart10M.Count;
            int lenRut = LRutVal.Length;
            
           datas.Add("桩号(km),左车辙RD1(mm),右车辙RD2(mm),路面车辙RD(mm)");
          
            if (lenRut < len)
            {
                for (int i = 0; i < lenRut; i++)
                {
                    double smile = RoadPart10M[i].mile;
                    string s1 = (smile * 0.001).ToString("f3");
                    string line = string.Join(",", s1, LRutVal[i].ToString("f1"), RRutVal[i].ToString("f1"), SRutVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            else
            {
                for (int i = 0; i < len; i++)
                {
                    
                    double smile = RoadPart10M[i].mile; 
                    string s1 = (smile * 0.001).ToString("f3"); 
                    string line = string.Join(",", s1, LRutVal[i].ToString("f1"), RRutVal[i].ToString("f1"), SRutVal[i].ToString("f1"));
                    datas.Add(line);
                }
            }
            File.WriteAllLines(path, datas);
        }

        
        public static void Convent_Mpd2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "磨耗_MPD.xlsx";
          

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
       
            writeAutoTestXls_Mpd2023(Destxls, prjinfo, prjdir, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, 4);

          
           
        }


        public static void Convent_Mpd2024_Standard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "磨耗_MPD.xlsx";


            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            List<string> datas = new List<string>();
            datas.Add("起点桩号(km),MPD_L(mm),MPD_C(mm),MPD_R(mm)");
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            int len = _RoadPart.Count - 1;
            for (int i = 0; i < len; i++)
            {
                string line = "";

                double smile = _RoadPart[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                line = string.Join(",", s1, _LMTDMeanVal[i].ToString("f2"), _CMTDMeanVal[i].ToString("f2"), _RMTDMeanVal[i].ToString("f2"));
                datas.Add(line);
            }
            File.WriteAllLines(Destxls, datas);
        }

        public static void Convent_Mpd2024_chongqing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "磨耗_MPD.xlsx";


            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            writeAutoTestXls_Mpd2024(Destxls, prjinfo, prjdir, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, 4); 
        }

        private static void writeAutoTestXls_Mpd2023(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int lastCol)
        {
            List<string> datas = new List<string>();
            datas.Add("桩号(km),MPD_L(mm),MPD_C(mm),MPD_R(mm)");
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            int len = RoadPart10M.Count - 1;
            for (int i = 0; i < len; i++)
            {
                string line = "";

                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                line = string.Join(",", s1, LMTDVal[i].ToString("f2"), CMTDVal[i].ToString("f2"), RMTDVal[i].ToString("f2"));
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        private static void writeAutoTestXls_Mpd2024(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int lastCol)
        {
            List<string> datas = new List<string>();
            datas.Add("桩号(km),MPD_L(mm),MPD_C(mm),MPD_R(mm)");
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }
          
            int len = RoadPart10M.Count - 1; 
            for (int i = 0; i < len; i++)
            {
                string line = ""; 
          
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                line = string.Join(",", s1, LMTDVal[i].ToString("f2"), CMTDVal[i].ToString("f2"), RMTDVal[i].ToString("f2"));
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        public static void Convent_Bump2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {//G236A-PB-1104.818-20240723104031
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadNum = fileName.Substring(0, fileName.Length - 1);
            string strDirc = fileName.Substring(fileName.Length - 1, 1);

            fileName = string.Format("{0}{1}-PB-{2}-{3}", strRoadNum, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Bump2024(Destxls, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5);
        }

        private static void writeAutoTestXls_Bump2024(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
            
            datas.Add("桩号,	PB_L,PB_M,PB_H,ΔH");
            int len = RoadPart10M.Count - 1;


            for (int i = 0; i < len; i++)
            {
                string line = "";

                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");

                double lValue = 0;
                double rValue = 0;
                if (i < LDeltaHVal.Length)
                {
                    lValue = LDeltaHVal[i];
                }
                if (prjinfo._IsDIRIMTD)
                {
                    if (i < RDeltaHVal.Length)
                    {
                        rValue = RDeltaHVal[i];
                    }
                }

                line = String.Join(",", s1, PBIVal[i][1], PBIVal[i][2], PBIVal[i][3], Math.Max(lValue, rValue));
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }


        public static void Convent_Bump2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "跳车_PB.txt";

            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_Bump2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5);
        }

        public static void Convent_Bump2023(  string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "跳车_PB.txt";
         
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_Bump2023(Destxls, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5); 
        }

        private static void writeAutoTestXls_Bump2024_ChongQing(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
            datas.Add("桩号(km),PB_L,PB_M,PB_H,ΔH(cm)");
            int len = RoadPart10M.Count - 1;
            for (int i = 0; i < len; i++)
            {
                string line = "";

                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");

                double lValue = 0;
                double rValue = 0;
                if (i < LDeltaHVal.Length)
                {
                    lValue = LDeltaHVal[i];
                }
                if (prjinfo._IsDIRIMTD)
                {
                    if (i < RDeltaHVal.Length)
                    {
                        rValue = RDeltaHVal[i];
                    }
                }

                line = String.Join(",", s1, PBIVal[i][1], PBIVal[i][2], PBIVal[i][3], Math.Max(lValue, rValue).ToString("f2"));
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        private static void writeAutoTestXls_Bump2023(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
         
            datas.Add("桩号(km),	PB_L,PB_M,PB_H,ΔH(cm)");
            int len = RoadPart10M.Count - 1;

        
            for (int i = 0; i < len; i++)
            {
                string line = "";
              
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");

                double lValue = 0;
                double rValue = 0;
                if (i < LDeltaHVal.Length)
                {
                    lValue = LDeltaHVal[i];
                }
                if (prjinfo._IsDIRIMTD)
                {
                    if (i < RDeltaHVal.Length)
                    {
                        rValue = RDeltaHVal[i];
                    }
                }
               
                line = String.Join(",", s1, PBIVal[i][1], PBIVal[i][2], PBIVal[i][3], Math.Max(lValue, rValue));
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }
        public static void Convent_TP2023(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_TP.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检数据格式\{1}",
                System.Windows.Forms.Application.StartupPath, excelFileName);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_gc_gj2023(_Worksheet_, _RoadPartF, prjdir, prjinfo, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }
        public static void Convent_TP2024_ChongQing( string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime; 
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, "TP");
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);

            writeAutoTestXls_gc_gj2024_ChongQing( Destxls, _RoadPartF, prjdir, prjinfo, _SpeedVal);

        }
        private static void writeAutoTestXls_gc_gj2024_ChongQing(string filePath, List<MilePartD> roadpart10, DirectoryInfo path, ProjectInfo proj, double[] speedVal)
        {
            List<float[]> allData = new List<float[]>();
            #region 获得高程数据list<float[]>
            int valnum = 0;
            switch (proj._RutMode)
            {
                case 0: valnum = 3; break;
                case 1: valnum = 1; break;
                case 2: valnum = 3; break;
                default: break;
            }
            int cameracnt = valnum == 1 ? 2 : 1;

            try
            {
                string[] dat = { string.Format("{0}\\camera0\\data", path.FullName), string.Format("{0}\\camera1\\data", path.FullName) };
                string[] process = { string.Format("{0}\\RUT\\camera0\\data", path.FullName), string.Format("{0}\\RUT\\camera1\\data", path.FullName) };
                string[] cfg = { string.Format("{0}\\camera0\\rutcfg.ini", path.FullName), string.Format("{0}\\camera1\\rutcfg.ini", path.FullName) };
                short hpix = short.Parse(IniFileOpr.ReadIniData("camera", "hpixel", "2048", cfg[0]));
                for (int i = 0; i < cameracnt; i++)
                {

                    IniFiles rutcfg = new IniFiles(cfg[i]);
                    int m, n, temp, j = 0;
                    float[] objlas = new float[hpix];
                    short[] profile = new short[hpix];
                    string _dtwname = "";
                    float[] tobjlas = new float[hpix];

                    string[] _dats = Directory.GetFiles(dat[i], "*.dat");
                    float _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
                    allData.Clear();
                    for (j = 0; j < _dats.Length; ++j)
                    {
                        _dtwname = _dats[j].Substring(_dats[j].LastIndexOf('\\') + 1);
                        _dtwname = _dtwname.Substring(0, _dtwname.IndexOf('.'));
                        using (FileStream frstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Open))
                        {
                            // fsbar = fsbar / frstream.Length;
                            temp = hpix * 2;
                            byte[] rbarr = new byte[hpix * 2];

                            while (frstream.Read(rbarr, 0, temp) > 0)
                            {

                                Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                for (m = 0, n = 0; m < hpix; ++m)
                                {
                                    // if (profile[m] != 0x7FFF)
                                    // {
                                    objlas[n] = profile[m] / _scaleval;
                                    if (proj._RutMode == 2)
                                    {
                                        objlas[n] = -objlas[n];
                                    }

                                    tobjlas[n] = objlas[m];
                                    ++n;


                                    //  }
                                }

                                allData.Add(tobjlas.Select(t => t).ToArray());
                            }

                        }
                    }
                }
                #endregion
                int len = roadpart10.Count - 1 <= allData.Count - 1 ? roadpart10.Count - 1 : allData.Count - 1;
                int rutDataLength = allData.Count;
                len = len <= rutDataLength ? len : rutDataLength;
              
                List<string> datas = new List<string>() { "桩号(km),高程值1,高程值2,高程值3,高程值4,高程值5,高程值6," +
                    "高程值7,高程值8,高程值9,高程值10,高程值11,高程值12," +
                    "高程值13,速度(m/s)" };

               
                for (int i = 0; i < len; i++)
                {
                   
                    int colcnt = 0;
                    double smile = roadpart10[i].mile;
                    double emile = roadpart10[i + 1].mile;
                    smile = smile * 0.001;
                    string dataLine = smile.ToString("f4")+",";
                    //这个n=100可能有问题
                    try
                    {
                        for (int n = 100; n < allData[i].Length && n < 1350; n += 100)
                        {
                            dataLine +=  (int)(allData[i][n] * 10) + ",";
                            // disvalsn[i, colcnt++] = (int)(allData[i][n] * 10); //原始高程单位 mm

                        }
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                    //disvalsn[i, colcnt++] = speedVal[i]*1000/3600;
                    dataLine += (speedVal[i] * 1000 / 3600).ToString("f1");
                    datas.Add(dataLine);
                }

                File.WriteAllLines(filePath, datas);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("原始高程表导出失败，可能是车辙计算数据缺失！\n" + ex.Message);
            }

        }

        private static void writeAutoTestXls_gc_gj2023(MSExcel.Worksheet _Worksheet, List<MilePartD> roadpart10, DirectoryInfo path, ProjectInfo proj, double[] speedVal)
        {
            List<float[]> allData = new List<float[]>();
            #region 获得高程数据list<float[]>
            int valnum = 0;
            switch (proj._RutMode)
            {
                case 0: valnum = 3; break;
                case 1: valnum = 1; break;
                case 2: valnum = 3; break;
                default: break;
            }
            int cameracnt = valnum == 1 ? 2 : 1;

            try
            {
                string[] dat = { string.Format("{0}\\camera0\\data", path.FullName), string.Format("{0}\\camera1\\data", path.FullName) };
                string[] process = { string.Format("{0}\\RUT\\camera0\\data", path.FullName), string.Format("{0}\\RUT\\camera1\\data", path.FullName) };
                string[] cfg = { string.Format("{0}\\camera0\\rutcfg.ini", path.FullName), string.Format("{0}\\camera1\\rutcfg.ini", path.FullName) };
                short hpix = short.Parse(IniFileOpr.ReadIniData("camera", "hpixel", "2048", cfg[0]));
                for (int i = 0; i < cameracnt; i++)
                {

                    IniFiles rutcfg = new IniFiles(cfg[i]);
                    int m, n, temp, j = 0;
                    float[] objlas = new float[hpix];
                    short[] profile = new short[hpix];
                    string _dtwname = "";
                    float[] tobjlas = new float[hpix];

                    string[] _dats = Directory.GetFiles(dat[i], "*.dat");
                    float _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
                    allData.Clear();
                    for (j = 0; j < _dats.Length; ++j)
                    {
                        _dtwname = _dats[j].Substring(_dats[j].LastIndexOf('\\') + 1);
                        _dtwname = _dtwname.Substring(0, _dtwname.IndexOf('.'));
                        using (FileStream frstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Open))
                        {
                            // fsbar = fsbar / frstream.Length;
                            temp = hpix * 2;
                            byte[] rbarr = new byte[hpix * 2];

                            while (frstream.Read(rbarr, 0, temp) > 0)
                            {

                                Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                for (m = 0, n = 0; m < hpix; ++m)
                                {
                                    // if (profile[m] != 0x7FFF)
                                    // {
                                    objlas[n] = profile[m] / _scaleval;
                                    if (proj._RutMode == 2)
                                    {
                                        objlas[n] = -objlas[n];
                                    }

                                    tobjlas[n] = objlas[m];
                                    ++n;


                                    //  }
                                }

                                allData.Add(tobjlas.Select(t => t).ToArray());
                            }

                        }
                    }
                }

                #endregion
                int len = roadpart10.Count - 1 <= allData.Count - 1 ? roadpart10.Count - 1 : allData.Count - 1;
                int rutDataLength = allData.Count;
                len = len<=rutDataLength? len : rutDataLength;
                object[,] disvalsn = new object[len, 22];
                MSExcel.Range destrange;
                for (int i = 0; i < len; i++)
                {
                    int colcnt = 0;
                    double smile = roadpart10[i].mile;
                    double emile = roadpart10[i + 1].mile;
                    disvalsn[i, colcnt++] = smile * 0.001;
                    //这个n=100可能有问题
                    try
                    {
                        for (int n = 100; n < allData[i].Length && n < 2100; n += 100)
                        {
                            disvalsn[i, colcnt++] = (int)(allData[i][n] * 10); //原始高程单位 mm
                        }
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }
                   
                    disvalsn[i, colcnt++] = speedVal[i];
                }
                destrange = _Worksheet.get_Range(String.Format("A2:V{0}", len + 1));
                destrange.Value2 = disvalsn;
                GlobalExcel.SetBorderLine(destrange, 53);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("原始高程表导出失败，可能是车辙计算数据缺失！\n" + ex.Message);
            }

        }

        public static void Convent_TT2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "——纹理_TT.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            double sMile = prjinfo._StartMile;
            double eMile = prjinfo._EndMile;
            int len = int.Parse(Math.Abs((sMile - eMile) * 1000).ToString());
            double[][] values = new double[3][] { new double[len], new double[len], new double[len] };
            for (int side = 0; side < 3; ++side)
            {
                for (int i = 0; i < len; ++i)
                {
                    values[side][i] = 0;
                }
            }
            for (int side = 0; side < 3; ++side)
            {
                //2中
                string fname = string.Format(@"{0}\IRIMTD\Laser{1}\lasval.txt", prjdir.FullName, side);
                string[] allLines;

                if (File.Exists(fname))
                {
                    int index2 = 0;
                    allLines = File.ReadAllLines(fname);
                    for (int i = 0; i < allLines.Length; ++i)
                    {
                        var splitStr = allLines[i].Split('\t');
                        double value = 0;


                        if (splitStr.Length > 1)
                        {

                            value = double.Parse(splitStr[1]);
                            values[side][index2] = value;
                            values[side][index2 + 1] = value;
                            index2 += 2;
                            if (index2 >= values[side].Length)
                            {
                                index2 = 0;
                            }
                        }
                        else
                        {

                        }
                    }
                }
            }
            object[,] disvalsn = new object[len, 4];
            double smile = prjinfo._StartMile * 0.001;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;

                disvalsn[i, colcnt++] = smile.ToString("f3");
                smile += 0.001;
                disvalsn[i, colcnt++] = values[0][i].ToString("f2");
                disvalsn[i, colcnt++] = values[2][i].ToString("f2");
                disvalsn[i, colcnt++] = values[1][i].ToString("f2");
            }
            //第一行写这个
            File.WriteAllText(Destxls, "桩号(km),左纹理(mm),中纹理(mm),右纹理(mm)\n");
            //后面写disvalsn内容 每列以,间隔
            // 写入 disvalsn 内容
            using (StreamWriter sw = new StreamWriter(Destxls, true))
            {
                for (int i = 0; i < len; i++)
                {
                    sw.WriteLine($"{disvalsn[i, 0]},{disvalsn[i, 1]},{disvalsn[i, 2]},{disvalsn[i, 3]}");
                }
            }

        }
        public static void Convent_TT2024_Standard(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "——纹理_TT.txt";
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            double sMile = prjinfo._StartMile;
            double eMile = prjinfo._EndMile;

            // 计算总点数（1mm间隔）
            int len = (int)Math.Abs((eMile - sMile) * 1000); // 转换为毫米数

            // 创建数组存储三个侧面的数据（如果len很大，可进一步优化为流式计算，但当前假设可接受）
            double[][] values = new double[3][];
            for (int side = 0; side < 3; side++)
            {
                values[side] = new double[len];
            }

            // 处理三个侧面的数据
            for (int side = 0; side < 3; side++)
            {
                string fname = string.Format(@"{0}\IRIMTD\Laser{1}\lasval.txt", prjdir.FullName, side);

                if (File.Exists(fname))
                {
                    string[] allLines = File.ReadAllLines(fname);
                    List<double> sourceValues = new List<double>();

                    // 读取原始数据（2mm间隔）
                    for (int i = 0; i < allLines.Length; i++)
                    {
                        var splitStr = allLines[i].Split('\t');
                        if (splitStr.Length > 1 && double.TryParse(splitStr[1], out double value))
                        {
                            sourceValues.Add(Math.Round(value, 2));
                        }
                    }

                    // 线性插值：从2mm间隔转换为1mm间隔
                    if (sourceValues.Count > 0)
                    {
                        // 计算插值后的数据点数量
                        int interpolatedCount = (sourceValues.Count - 1) * 2 + 1;
                        int pointsToCopy = Math.Min(interpolatedCount, len);

                        for (int i = 0; i < pointsToCopy; i++)
                        {
                            if (i % 2 == 0)
                            {
                                // 偶数索引：使用原始数据点
                                int sourceIndex = i / 2;
                                if (sourceIndex < sourceValues.Count)
                                {
                                    values[side][i] = sourceValues[sourceIndex];
                                }
                            }
                            else
                            {
                                // 奇数索引：线性插值
                                int prevIndex = i / 2;
                                int nextIndex = prevIndex + 1;

                                if (nextIndex < sourceValues.Count)
                                {
                                    // 在两个原始点之间进行线性插值
                                    values[side][i] = (sourceValues[prevIndex] + sourceValues[nextIndex]) / 2.0;
                                }
                                else
                                {
                                    // 如果是最后一个点，使用前一个点的值
                                    values[side][i] = sourceValues[prevIndex];
                                }
                            }
                        }

                        // 如果插值后的点数少于目标长度，用最后一个值填充剩余部分
                        if (pointsToCopy < len)
                        {
                            double lastValue = sourceValues.Count > 0 ? sourceValues[sourceValues.Count - 1] : 0;
                            for (int i = pointsToCopy; i < len; i++)
                            {
                                values[side][i] = lastValue;
                            }
                        }
                    }
                }
                else
                {
                    // 如果文件不存在，用0填充
                    for (int i = 0; i < len; i++)
                    {
                        values[side][i] = 0;
                    }
                }
            }

            // 修改：用 StreamWriter 逐行写入，避免大 List<string> 内存问题
            using (StreamWriter sw = new StreamWriter(Destxls, false, Encoding.UTF8))
            {
                sw.WriteLine("桩号(km),左纹理(mm),中纹理(mm),右纹理(mm)");

                double currentMile = prjinfo._StartMile * 0.001; // 转换为km
                double increment = 0.000001; // 1mm对应的公里数增量

                for (int i = 0; i < len; i++)
                {
                    // 确保索引在有效范围内
                    double leftValue = i < values[0].Length ? values[0][i] : 0;
                    double middleValue = i < values[2].Length ? values[2][i] : 0;
                    double rightValue = i < values[1].Length ? values[1][i] : 0;

                    string line = $"{currentMile.ToString("f6")},{leftValue.ToString("f2")},{middleValue.ToString("f2")},{rightValue.ToString("f2")}";
                    sw.WriteLine(line);
                    if (prjinfo._Direction == 1)
                    {
                        currentMile += increment;
                        if (i % 1000 == 0)
                        {
                            currentMile = (prjinfo._StartMile * 0.001) + (i + 1) * increment;
                        }
                    }
                    else
                    {
                        currentMile -= increment;
                        if (i % 1000 == 0)
                        {
                            currentMile = (prjinfo._StartMile * 0.001) - (i + 1) * increment;
                        }
                    }


                    // 防止浮点数精度累积误差
                   

                    // 可选：每1000行刷新一次（平衡内存与IO性能）
                    if (i % 1000 == 0)
                    {
                        sw.Flush();
                    }
                }
            }
        }


        public static void Convent_TT2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "——纹理_TT.txt"; 

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            double sMile = prjinfo._StartMile;
            double eMile = prjinfo._EndMile;
            int len = int.Parse(Math.Abs((sMile - eMile) * 1000).ToString());
            double[][] values = new double[3][] { new double[len], new double[len], new double[len] };
            for (int side = 0; side < 3; ++side)
            {
                for (int i = 0; i < len; ++i)
                {
                    values[side][i] = 0;
                }
            }
            for (int side = 0; side < 3; ++side)
            {
                //2中
                string fname = string.Format(@"{0}\IRIMTD\Laser{1}\lasval.txt", prjdir.FullName, side);
                string[] allLines;

                if (File.Exists(fname))
                {
                    int index2 = 0;
                    allLines = File.ReadAllLines(fname);
                    for (int i = 0; i < allLines.Length; ++i)
                    {
                        var splitStr = allLines[i].Split('\t');
                        double value = 0;


                        if (splitStr.Length > 1)
                        {

                            value = double.Parse(splitStr[1]);
                            values[side][index2] = value;
                            values[side][index2 + 1] = value;
                            index2 += 2;
                            if (index2 >= values[side].Length)
                            {
                                index2 = 0;
                            }
                        }
                        else
                        {

                        }
                    }
                }
            }
            object[,] disvalsn = new object[len, 4];
            double smile = prjinfo._StartMile * 0.001;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;

                disvalsn[i, colcnt++] = smile.ToString("f3");
                smile += 0.001;
                disvalsn[i, colcnt++] = values[0][i];
                disvalsn[i, colcnt++] = values[2][i];
                disvalsn[i, colcnt++] = values[1][i];
            }
            //第一行写这个
            File.WriteAllText(Destxls, "桩号(km),左纹理(mm),中纹理(mm),右纹理(mm)\n");
            //后面写disvalsn内容 每列以,间隔
            // 写入 disvalsn 内容
            using (StreamWriter sw = new StreamWriter(Destxls, true))
            {
                for (int i = 0; i < len; i++)
                {
                    sw.WriteLine($"{disvalsn[i, 0]},{disvalsn[i, 1]},{disvalsn[i, 2]},{disvalsn[i, 3]}");
                }
            }

        }


        #endregion

        #region 合肥 

        public static void outPutIriCsv_WH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\芜湖csv\iri.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dir = prjinfo._Direction == 1 ? "A" : "B";
            string Destxls = string.Format(@"{0}\{1}-{2}-IRI.csv", path, prjinfo._RoadCode, dir);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            _Workbook.SaveAs(Destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            write_iriCsv_WH(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void write_iriCsv_WH(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal)
        {
            int len = roadpart10.Count - 1;
            object[,] disvalsn = new object[len, 3];
            MSExcel.Range destrange;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                disvalsn[i, colcnt++] = smile * 0.001;
                disvalsn[i, colcnt++] = LIRIMeanVal[i].ToString("f6");
            }
            destrange = _Worksheet.get_Range(String.Format("A2:B{0}", len + 1));
            destrange.Value2 = disvalsn;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction > 0)
            {

            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                //destrange = worksheet_sn.get_Range(String.Format("B2:O{0}", rowcnt_sn - 1));
                //sortrange = worksheet_sn.get_Range(String.Format("C2:C{0}", len + 1));
                //GlobalExcel.ReflectionColnum(worksheet_sn, destrange, sortrange);

                destrange = _Worksheet.get_Range(String.Format("A2:B{0}", len + 1));
                MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
            }
        }

        public static void outPutDrCsv_WH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\芜湖csv\dr.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dir = prjinfo._Direction == 1 ? "A" : "B";
            string Destxls = string.Format(@"{0}\{1}-{2}-DR.csv", path, prjinfo._RoadCode, dir);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            _Workbook.SaveAs(Destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            SmallWritePCI2XlsCsv_WH(_Worksheet_iri, prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void SmallWritePCI2XlsCsv_WH(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
      List<MilePart> roadpart, Disease[] arrdis)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 2];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                vallist[i, 0] = smile / 1000.0;
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 1] = drval;
            }

            destrange = worksheet.get_Range(String.Format("A{0}:B{1}", 2, len + 2 - 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 2, 1, 2, true);
            }
        }

        public static void OutputRoad_Hefei(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\合肥\2023安徽省样表.xlsx",
              System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\合肥路况_{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_jlb = _Workbook.Sheets["记录表"] as MSExcel.Worksheet;
            //WriteDisLB2Xls(_Worksheet_lb, prjinfo, _SmallRoadDisList);
            WriteAll2Xls_Hefei(_Worksheet_jlb, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _DeltaHVal, disval, _GPSInfo);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteAll2Xls_Hefei(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
     Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
     double[] LMTDVal, double[] RMTDVal, double[] MMTDVal, int[][] PBVal, double[] deltahVal, int disval, ExcelGPS[] gpsInfo)
        {
            bool shangxing = prjinfo._Direction == 1 ? true : false;
            //检查区间长度进行处理
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double irival = 0, tpcival = 0;

            object[,] mxlist = new object[len, 25];
            double[] drvals = new double[len];

            string lenstr = "0";
            int tlen = len;
            while ((tlen = tlen / 10) > 0)
            {
                lenstr += "0";
            }

            string errlog = prjdir.FullName + "\\errlog.txt";

            int typeidx = 0;
            bool res = false;

            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile; 
                double drval;

                int milelength = Math.Abs(smile - emile);
                mxlist[i, 0] = prjinfo._DataPerson;
                mxlist[i, 1] = prjinfo._DataDate;
                mxlist[i, 2] = prjinfo._RoadCode;
                mxlist[i, 3] = prjinfo._RoadName;
                mxlist[i, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                double s1 = roadpart[i].mile / 1000.0;
                double s2 = roadpart[i + 1].mile / 1000.0;
                double smile1 = Math.Round(s1, 3);
                double emile1 = Math.Round(s2, 3);
                double milelength1 = Math.Round(Math.Abs(s1 - s2), 3);
                if (prjinfo._Direction > 0)
                {
                    mxlist[i, 8] = smile1;
                    mxlist[i, 9] = emile1;
                }
                else
                {
                    mxlist[i, 8] = emile1;
                   mxlist[i, 9] = smile1;
                }
                mxlist[i, 10] = milelength1;
                mxlist[i, 11] = roadpart[i].width;
                mxlist[i, 12] = prjinfo._City;
                mxlist[i, 13] = prjinfo._District;
                mxlist[i, 14] = roadpart[i].unit;
                if (roadpart[i].isPub)
                {
                    mxlist[i, 15] = "共有路段";
                }
                else
                {
                    mxlist[i, 15] = roadpart[i].degreestr;
                }

                mxlist[i, 16] = roadpart[i].roadtype == 0 ? "沥青" : roadpart[i].roadtype == 1 ? "水泥" : "砂石";
                string sGps = gpsInfo[i]._longitude + "," + gpsInfo[i]._latitude;
                mxlist[i, 22] = sGps;
                if (i + 1 < len)
                {
                    string eGps = gpsInfo[i + 1]._longitude + "," + gpsInfo[i + 1]._latitude;
                    mxlist[i, 23] = eGps;
                }
                else
                {
                    mxlist[i, 23] = sGps;
                }
                mxlist[i, 24] = "市检";
                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                for (int d = 0; d < arrdis.Length; d++)
                {
                    if (d < dlen && ((prjinfo._Direction > 0 && arrdis[d].m_mile >= smile && arrdis[d].m_mile < emile)
                      || (prjinfo._Direction < 0 && arrdis[d].m_mile <= smile && arrdis[d].m_mile > emile)))
                    {
                        res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                           arrdis[d].RoadType, arrdis[d].RoadDisType), out typeidx);
                        if (res)
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[d].Area;
                        }
                        else
                        {
                            string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[d].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                            File.AppendAllText(errlog, errval, Encoding.UTF8);
                        }
                    }
                }
                //PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                mxlist[i, 20] = drval;
                tpcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                drvals[i] = drval;   //dr
                mxlist[i, 6] = string.Format("=100-{0}*POWER(U{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 2, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]); //pci
                mxlist[i, 18] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                   i + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);  //pci评价

                //IRI                
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.IRIExcelSide == 2)
                    {
                        if (_Setting.RQIJudgeType == 0)
                        {
                            irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                        }
                    }
                    else if (_Setting.IRIExcelSide == 0)
                    {
                        irival = Math.Round(LIRIVal[i], 5);
                    }
                    else if (_Setting.IRIExcelSide == 1)
                    {
                        irival = Math.Round(RIRIVal[i], 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                mxlist[i, 21] = irival;
                //rqi
                mxlist[i, 7] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{2})),5)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], irival);
                mxlist[i, 19] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                    i + 2,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);

                mxlist[i, 5] = string.Format("=ROUND(({1}*G{0}+{2}*H{0})/({1}+{2}),5)",
                i + 2,
                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                //PQI 评价
                mxlist[i, 17] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                i + 2,
                _PQIGrade[roadpart[i].roaddegree][0],
                _PQIGrade[roadpart[i].roaddegree][1],
                _PQIGrade[roadpart[i].roaddegree][2],
                _PQIGrade[roadpart[i].roaddegree][3]);
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A2:Y{0}", len + 1));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort)
            {
                MSExcel.Range sortrange = worksheet.get_Range(String.Format("I2:I{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }  //处理过小分段问题
            #region 处理过小分段问题 
            {
                mxlist = destrange.Value2;
                List<object[]> resultValues = new List<object[]>();
                if (_Setting.hefei2MinSplit > 0)
                {
                    int splitValue = _Setting.hefei2MinSplit;
                    for (int i = 1; i <= len; i++)
                    {
                        object[] preValue = new object[26];
                        object[] nowValue = new object[26];
                        if (i == len)
                        {
                            for (int t = 1; t < 26; t++)
                            {
                                preValue[t] = mxlist[i, t];
                            }
                            //遍历到了最后一项 
                            if (resultValues.Count == 0)
                            {

                                resultValues.Add(preValue);
                            }
                            else
                            {
                                nowValue = preValue;
                                //前一个变成了列表最后一项 
                                preValue = resultValues.Last();
                                double preWidth = double.Parse(preValue[12].ToString());
                                string preUnit = "";
                                if (preValue[15] != null)
                                {
                                    preUnit = preValue[15].ToString();

                                }


                                string preGrad = "";
                                if (preValue[16] != null)
                                {
                                    preGrad = preValue[16].ToString();
                                }
                                string preType = "";
                                if (preValue[17] != null)
                                {
                                    preType = preValue[17].ToString();
                                }

                                double preIRI = double.Parse(preValue[22].ToString());
                                double preDr = double.Parse(preValue[21].ToString());

                                double nowLen = double.Parse(nowValue[11].ToString());

                                double nowWidth = double.Parse(nowValue[12].ToString());
                                if (nowWidth == 0)
                                {
                                    nowWidth = preWidth;
                                    nowValue[12] = preWidth;
                                }
                                string nowUnit = "";
                                if (nowValue[15] == null)
                                {
                                    nowUnit = preUnit;
                                    nowValue[15] = preUnit;
                                }
                                else
                                {
                                    nowUnit = nowValue[15].ToString();
                                }
                                


                                string nowGrad = nowValue[16].ToString();
                                string nowType = nowValue[17].ToString();
                                double nowIRI = double.Parse(nowValue[22].ToString());
                                double nowDr = double.Parse(nowValue[21].ToString());

                                double preLen = double.Parse(preValue[11].ToString());
                                //pqi  
                                double prePqi = double.Parse(preValue[6].ToString());
                                //pci
                                double prePci = double.Parse(preValue[7].ToString());
                                //rqi
                                double preRqi = double.Parse(preValue[8].ToString());

                                //pqi  
                                double nowPqi = double.Parse(nowValue[6].ToString());
                                //pci                         
                                double nowPci = double.Parse(nowValue[7].ToString());
                                //rqi                        
                                double nowRqi = double.Parse(nowValue[8].ToString());
                                if ((nowLen * 1000) <= splitValue)
                                {
                                    if (preWidth == nowWidth && preUnit == nowUnit && preGrad == nowGrad && preType == nowType)
                                    {
                                        //第一分段和第二分段合并 
                                        //如果下一个分段小于指定值 
                                        resultValues[resultValues.Count - 1][10] = nowValue[10];
                                        resultValues[resultValues.Count - 1][11] = double.Parse(resultValues[resultValues.Count - 1][10].ToString()) - double.Parse(resultValues[resultValues.Count - 1][9].ToString());
                                        //定位
                                        resultValues[resultValues.Count - 1][24] = nowValue[24];

                                        double newIRI = ((preIRI * preLen) + (nowIRI * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][22] = newIRI;

                                        double newDr = ((preDr * preLen) + (nowDr * nowLen)) / (preLen + nowLen);

                                        double pqi = ((prePqi * preLen) + (nowPqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][6] = pqi;
                                        double pci = ((prePci * preLen) + (nowPci * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][7] = pci;
                                        double rqi = ((preRqi * preLen) + (nowRqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][8] = rqi;

                                        resultValues[resultValues.Count - 1][21] = newDr;


                                    }
                                    else
                                    {
                                        resultValues.Add(nowValue);
                                    }
                                }
                                else if ((preLen * 1000) <= splitValue)
                                {
                                    if (preWidth == nowWidth && preUnit == nowUnit && preGrad == nowGrad && preType == nowType)
                                    {
                                        //第一分段和第二分段合并 
                                        //如果下一个分段小于指定值 
                                        resultValues[resultValues.Count - 1][10] = nowValue[10];
                                        resultValues[resultValues.Count - 1][11] = double.Parse(resultValues[resultValues.Count - 1][10].ToString()) - double.Parse(resultValues[resultValues.Count - 1][9].ToString());
                                        //定位
                                        resultValues[resultValues.Count - 1][24] = nowValue[24];

                                        double newIRI = ((preIRI * preLen) + (nowIRI * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][22] = newIRI;

                                        double newDr = ((preDr * preLen) + (nowDr * nowLen)) / (preLen + nowLen);
                                        double pqi = ((prePqi * preLen) + (nowPqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][6] = pqi;
                                        double pci = ((prePci * preLen) + (nowPci * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][7] = pci;
                                        double rqi = ((preRqi * preLen) + (nowRqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][8] = rqi;

                                        resultValues[resultValues.Count - 1][21] = newDr;


                                    }
                                    else
                                    {
                                        resultValues.Add(nowValue);
                                    }

                                }

                                else
                                {
                                    resultValues.Add(nowValue);
                                }


                            }
                        }
                        else
                        {
                            object[] values = new object[26];
                            for (int t = 1; t < 26; t++)
                            {
                                preValue[t] = mxlist[i, t];
                                nowValue[t] = mxlist[i + 1, t];
                            }

                            double preWidth = double.Parse(preValue[12].ToString());

                            string preUnit = "";
                            if (preValue[15] != null)
                            {
                                preUnit = preValue[15].ToString();

                            }


                            string preGrad = "";
                            if (preValue[16] != null)
                            {
                                preGrad = preValue[16].ToString();
                            }
                            string preType = "";
                            if (preValue[17] != null)
                            {
                                preType = preValue[17].ToString();
                            }

                            double preIRI = double.Parse(preValue[22].ToString());
                            double preDr = double.Parse(preValue[21].ToString());


                            double nowLen = double.Parse(nowValue[11].ToString());
                            double nowWidth = 0;
                            if (double.Parse(nowValue[12].ToString()) == 0)
                            {
                                nowWidth = preWidth;
                                nowValue[12] = preWidth;
                            }
                            else
                            {
                                nowWidth = double.Parse(nowValue[12].ToString());
                            }
                            string nowUnit = "";
                            if (nowValue[15] == null)
                            {
                                nowUnit = preUnit;
                                nowValue[15] = preUnit;
                            }
                            else
                            {
                                nowUnit = nowValue[15].ToString();
                            }
                            string nowGrad = nowValue[16].ToString();
                            string nowType = nowValue[17].ToString();
                            double nowIRI = double.Parse(nowValue[22].ToString());
                            double nowDr = double.Parse(nowValue[21].ToString());

                            //pqi  
                            double prePqi = double.Parse(preValue[6].ToString());
                            //pci
                            double prePci = double.Parse(preValue[7].ToString());
                            //rqi
                            double preRqi = double.Parse(preValue[8].ToString());

                            //pqi  
                            double nowPqi = double.Parse(nowValue[6].ToString());
                            //pci                         
                            double nowPci = double.Parse(nowValue[7].ToString());
                            //rqi                        
                            double nowRqi = double.Parse(nowValue[8].ToString());


                            double preLen = double.Parse(preValue[11].ToString());
                            if ((preLen * 1000) <= splitValue)
                            {
                                if (resultValues.Count == 0)
                                {
                                    //第一个分段就有问题
                                    //看下一个
                                    if (preWidth == nowWidth && preUnit == nowUnit && preGrad == nowGrad && preType == nowType)
                                    {
                                        //第一分段和第二分段合并 
                                        //如果下一个分段小于指定值 
                                        preValue[10] = nowValue[10];
                                        preValue[11] = double.Parse(preValue[10].ToString()) - double.Parse(preValue[9].ToString());
                                        preValue[24] = nowValue[24];

                                        double newIRI = ((preIRI * preLen) + (nowIRI * nowLen)) / (preLen + nowLen);
                                        preValue[22] = newIRI;

                                        double newDr = ((preDr * preLen) + (nowDr * nowLen)) / (preLen + nowLen);
                                        preValue[21] = newDr;

                                        double pqi = ((prePqi * preLen) + (nowPqi * nowLen)) / (preLen + nowLen);
                                        preValue[6] = pqi;
                                        double pci = ((prePci * preLen) + (nowPci * nowLen)) / (preLen + nowLen);
                                        preValue[7] = pci;
                                        double rqi = ((preRqi * preLen) + (nowRqi * nowLen)) / (preLen + nowLen);
                                        preValue[8] = rqi;
                                        resultValues.Add(preValue);
                                        i++;//跳过第二个分段

                                    }
                                    else
                                    {
                                        resultValues.Add(preValue);
                                    }
                                }
                                else
                                {
                                    nowValue = preValue;
                                    //前一个变成了列表最后一项 
                                    preValue = resultValues.Last();
                                    preWidth = double.Parse(preValue[12].ToString());
                                    if (preValue[15] != null)
                                    {
                                        preUnit = preValue[15].ToString();

                                    } 
                                    if (preValue[16] != null)
                                    {
                                        preGrad = preValue[16].ToString();
                                    }   

                                    if (preValue[17] != null)
                                    {
                                        preType = preValue[17].ToString();
                                    }

                                    preIRI = double.Parse(preValue[22].ToString());
                                    preDr = double.Parse(preValue[21].ToString());
                                    nowLen = double.Parse(nowValue[11].ToString());
                                    nowWidth = double.Parse(nowValue[12].ToString());
                                    nowUnit = nowValue[15].ToString();
                                    nowGrad = nowValue[16].ToString();
                                    nowType = nowValue[17].ToString();
                                    nowIRI = double.Parse(nowValue[22].ToString());
                                    nowDr = double.Parse(nowValue[21].ToString());

                                    //pqi  
                                    prePqi = double.Parse(preValue[6].ToString());
                                    //pci
                                    prePci = double.Parse(preValue[7].ToString());
                                    //rqi
                                    preRqi = double.Parse(preValue[8].ToString());

                                    //pqi  
                                    nowPqi = double.Parse(nowValue[6].ToString());
                                    //pci                         
                                    nowPci = double.Parse(nowValue[7].ToString());
                                    //rqi                        
                                    nowRqi = double.Parse(nowValue[8].ToString());

                                    if (preWidth == nowWidth && preUnit == nowUnit && preGrad == nowGrad && preType == nowType)
                                    {
                                        //第一分段和第二分段合并 
                                        //如果下一个分段小于指定值 
                                        resultValues[resultValues.Count - 1][10] = nowValue[10];
                                        resultValues[resultValues.Count - 1][11] = double.Parse(resultValues[resultValues.Count - 1][10].ToString()) - double.Parse(resultValues[resultValues.Count - 1][9].ToString());
                                        resultValues[resultValues.Count - 1][24] = nowValue[24];

                                        double newIRI = ((preIRI * preLen) + (nowIRI * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][22] = newIRI;

                                        double newDr = ((preDr * preLen) + (nowDr * nowLen)) / (preLen + nowLen);

                                        double pqi = ((prePqi * preLen) + (nowPqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][6] = pqi;
                                        double pci = ((prePci * preLen) + (nowPci * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][7] = pci;
                                        double rqi = ((preRqi * preLen) + (nowRqi * nowLen)) / (preLen + nowLen);
                                        resultValues[resultValues.Count - 1][8] = rqi;

                                        resultValues[resultValues.Count - 1][21] = newDr;
                                        //将当前合并到最后一项去
                                        //resultValues.Add(resultValues[resultValues.Count - 1]);

                                    }
                                    else
                                    {
                                        resultValues.Add(nowValue);
                                    }

                                }
                            }
                            else
                            {
                                if (resultValues.Count > 0)
                                {
                                    nowValue = preValue;
                                    //前一个变成了列表最后一项 
                                    preValue = resultValues.Last();
                                    preWidth = double.Parse(preValue[12].ToString());
                                    if (preValue[15] != null)
                                    {
                                        preUnit = preValue[15].ToString();

                                    }


                                     
                                    if (preValue[16] != null)
                                    {
                                        preGrad = preValue[16].ToString();
                                    }
                                  
                                    if (preValue[17] != null)
                                    {
                                        preType = preValue[17].ToString();
                                    }

                                    preIRI = double.Parse(preValue[22].ToString());
                                    preDr = double.Parse(preValue[21].ToString());
                                    nowLen = double.Parse(nowValue[11].ToString());

                                    nowWidth = double.Parse(nowValue[12].ToString());
                                    if (nowWidth == 0)
                                    {
                                        nowWidth = preWidth;
                                        nowValue[12] = preWidth;
                                    }
                                    nowUnit = "";
                                    if (nowValue[15] == null)
                                    {
                                        nowUnit = preUnit;
                                        nowValue[15] = preUnit;
                                    }
                                    else
                                    {
                                        nowUnit = nowValue[15].ToString();
                                    }
                                    nowGrad = nowValue[16].ToString();
                                    nowType = nowValue[17].ToString();
                                    nowIRI = double.Parse(nowValue[22].ToString());
                                    nowDr = double.Parse(nowValue[21].ToString());

                                    //pqi  
                                    prePqi = double.Parse(preValue[6].ToString());
                                    //pci
                                    prePci = double.Parse(preValue[7].ToString());
                                    //rqi
                                    preRqi = double.Parse(preValue[8].ToString());

                                    //pqi  
                                    nowPqi = double.Parse(nowValue[6].ToString());
                                    //pci                         
                                    nowPci = double.Parse(nowValue[7].ToString());
                                    //rqi                        
                                    nowRqi = double.Parse(nowValue[8].ToString());
                                    preLen = double.Parse(preValue[11].ToString());
                                    if ((preLen * 1000) <= splitValue)
                                    {
                                        if (preWidth == nowWidth && preUnit == nowUnit && preGrad == nowGrad && preType == nowType)
                                        {
                                            //第一分段和第二分段合并 
                                            //如果下一个分段小于指定值 
                                            resultValues[resultValues.Count - 1][10] = nowValue[10];
                                            resultValues[resultValues.Count - 1][11] = double.Parse(resultValues[resultValues.Count - 1][10].ToString()) - double.Parse(resultValues[resultValues.Count - 1][9].ToString());
                                            resultValues[resultValues.Count - 1][24] = nowValue[24];

                                            double newIRI = ((preIRI * preLen) + (nowIRI * nowLen)) / (preLen + nowLen);
                                            resultValues[resultValues.Count - 1][22] = newIRI;

                                            double newDr = ((preDr * preLen) + (nowDr * nowLen)) / (preLen + nowLen);

                                            double pqi = ((prePqi * preLen) + (nowPqi * nowLen)) / (preLen + nowLen);
                                            resultValues[resultValues.Count - 1][6] = pqi;
                                            double pci = ((prePci * preLen) + (nowPci * nowLen)) / (preLen + nowLen);
                                            resultValues[resultValues.Count - 1][7] = pci;
                                            double rqi = ((preRqi * preLen) + (nowRqi * nowLen)) / (preLen + nowLen);
                                            resultValues[resultValues.Count - 1][8] = rqi;

                                            resultValues[resultValues.Count - 1][21] = newDr;
                                            //将当前合并到最后一项去
                                            //resultValues.Add(resultValues[resultValues.Count - 1]);

                                        }
                                        else
                                        {
                                            resultValues.Add(nowValue);
                                        }
                                    }
                                    else
                                    {
                                        resultValues.Add(nowValue);
                                    }

                                }
                                else
                                {
                                    resultValues.Add(preValue);

                                }
                            }
                        }

                    }
                }
                object[,] ResultMxlist = new object[resultValues.Count, 25];

                for (int i = 0; i < resultValues.Count; i++)
                {
                    object[] value = resultValues[i];
                    for (int t = 0; t < 25; t++)
                    {
                        ResultMxlist[i, t] = value[t + 1];
                    }
                }
                destrange.Clear();
                destrange = worksheet.get_Range(String.Format("A2:Y{0}", resultValues.Count + 1));
                destrange.Value2 = ResultMxlist;
                GlobalExcel.SetBorderLine(destrange, 53);

            }
            #endregion
        }
        #endregion
        #region 甘肃

        #endregion
    }
}

