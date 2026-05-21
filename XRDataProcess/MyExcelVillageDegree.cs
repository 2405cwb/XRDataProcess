#define 江西车检2024病害顺序
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.XtraCharts;
using DevExpress.XtraCharts.Designer.Native;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors.TextEditController;
using DevExpress.XtraPrinting.Export.Pdf;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraTreeList.Internal;
using Framework.Office.Excel;
using Framework.Other;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using OperateIniFile;
using Spire.Pdf.Exporting.XPS.Schema;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms; 
using System.Windows.Forms.VisualStyles;
using System.Xml;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace XRDataProcess
{
    /// <summary>
    /// 低等级农村公路
    /// </summary>
    class MyExcelVillageDegree
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
        private static double[] _MPDLHVal = null;
        private static double[] _MPDCHVal = null;
        private static double[] _MPDRHVal = null;
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

        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        public static Dictionary<string, int> _RoadGradeDict;

        public static List<MilePart> _RoadPart = null;
        public static List<MilePart> _RoadPart1M = null; //1米桩号分段 
        public static List<MilePart> _RoadPart10 = null; //整10米桩号分段

        
        public static List<MilePartD> _RoadPartF = null;//0.1米桩号分段
        private static double[] _LiriHVal = null;
        private static double[] _RiriHVal = null;


        private static double[] _SpeedVal10 = null;
        private static string[] _MarkVal10 = null;

        
        private static Disease[] _RoadDisList = null;
        private static Disease[] _RoadRepairList = null;

        private static double[] _LIRIMeanVal = null;
        private static double[] _RIRIMeanVal = null;

        private static double[] _LMTDMeanVal = null;
        private static double[] _RMTDMeanVal = null;
        private static double[] _CMTDMeanVal = null;
        //车辙
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
                _RQIGrade[i] = new double[RoadDiseaseTypes.roadtypedict.Count][];
                _RDIGrade[i] = new double[5];
                _PCIGrade[i] = new double[5];
                _PQIGrade[i] = new double[5];
                _PBIGrade[i] = new double[5];
                _PWIGrade[i] = new double[5];

                _RQIa[i] = new double[RoadDiseaseTypes.roadtypedict.Count][];
                _PCIa[i] = new double[RoadDiseaseTypes.roadtypedict.Count][];
                _PQIW[i] = new double[RoadDiseaseTypes.roadtypedict.Count][];
                for (int j = 0; j < RoadDiseaseTypes.roadtypedict.Count; j++)
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

            _RoadSocre = new Dictionary<string, CityRoadDis>[RoadDiseaseTypes.roadtypedict.Count];
            for (int i = 0; i < RoadDiseaseTypes.roadtypedict.Count; i++)
            {
                _RoadSocre[i] = new Dictionary<string, CityRoadDis>();
            }

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

            for (int i = 0; i < RoadDiseaseTypes.roadtypedict.Count; i++)
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
                    if (rootchild.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
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

       

        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
            bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
        {
            _SpeedVal = null;

            bool IRIRes = true, RutRes = true, MTDRes = true, PBIRes = true, GPSRes = true, MaxRutRes = true, MPDRes = true;
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
                MessageBox.Show(string.Format($"{prjinfo._RoadName}【低等级农村公路】不包含【{prjinfo._RoadGrade}】请检查工程数据！"));
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
                if (IsMeanMTD)
                {
                    MTDRes = GlobalExcel.GetMTDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, _Setting.IsWarning);
                  
                }
                if (IsMeanMPD)
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
            }
            else
            {
                RutRes = true;
            }

            if (_Setting.ExcelType == 4||_Setting.ExcelType==8||_Setting.ExcelType==7) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);

            if (_RoadPart[0].roaddegree <= 1)
            {
                return IRIRes && RutRes && MTDRes && GPSRes && MPDRes;
            }
            else
            {
                return IRIRes && MPDRes;
            }
        }
        /// <summary>
        /// 国检转换 TP表 桩号间隔要求为0.1m
        /// </summary>
        /// <param name="prjdir"></param>
        /// <param name="prjinfo"></param>
        /// <param name="disval"></param>
        /// <param name="IsDis"></param>
        /// <param name="IsMeanIRI"></param>
        /// <param name="IsMeanMTD"></param>
        /// <param name="IsMeanRut"></param>
        /// <param name="IsPBI"></param>
        /// <param name="IsSpeed"></param>
        /// <param name="IsMeanMPD"></param>
        /// <returns></returns>
        public static bool InitProDataD(DirectoryInfo prjdir, ProjectInfo prjinfo,double disval,
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
                        GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF,disval, 1, ref _RiriHVal);
                    }
                }

            }
            else
            {
                IRIRes = true;
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

        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面平整度评价等级记录表.xlsx",
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

            
            if (_Setting.ExcelType != 10)
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
       
        
        /// <summary>
          /// 根据湖南农村路计算iri
          /// </summary>
          /// <param name="V">速度</param>
          /// <param name="iri"></param>
        private static double huNanJudgeIri(double V, double iri)
        {
            if (V <= 20)
            {
                return 0.8 * iri;
            }
            else if (V <= 30)
            {
                return 0.9 * iri;
            }
            else if (V <= 40)
            {
                return 0.95 * iri;
            }
            else if (V <= 50)
            {
                return 0.98 * iri;
            }
            else
                return iri;
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
             
                if (SpeedVal != null)
                {
                    vallist[i, 9] = SpeedVal[i];
                }
                //进行iri速度矫正的
                if (_Setting.RQIJudgeType == 2)
                {
                    vallist[i, 3] = huNanJudgeIri(SpeedVal[i], LIRIVal[i]);
                    if (prjinfo._IsDIRIMTD)
                    {
                        vallist[i, 4] = huNanJudgeIri(SpeedVal[i], RIRIVal[i]); 
                        if (_Setting.RQIJudgeType == 0)
                        {
                            vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + DataStartXlsxRow);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            vallist[i, 5] = String.Format("=ROUND(MAX(D{0}, E{0}),5)", i + DataStartXlsxRow);
                        }
                    }
                    else
                    {
                        vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + DataStartXlsxRow);
                    }

                }
                else
                {

                    vallist[i, 3] = LIRIVal[i];
                    if (prjinfo._IsDIRIMTD)
                    {
                        vallist[i, 4] = RIRIVal[i];
                        if (_Setting.RQIJudgeType == 0)
                        {
                            vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + DataStartXlsxRow);
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            vallist[i, 5] = String.Format("=ROUND(MAX(D{0}, E{0}),5)", i + DataStartXlsxRow);
                        }
                    }
                    else
                    {
                        vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + DataStartXlsxRow);
                    }

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



     

        public static void OutputCPMSDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            disval *= 10;
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\CPMS路面病害调查表.xlsx",
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
            MSExcel.Worksheet _Worksheet_ssdc = _Workbook.Sheets["砂石路面损坏调查表"] as MSExcel.Worksheet;

            WritePrj2CPMSXls(_Worksheet_lqdc, prjinfo);
            WritePrj2CPMSXls(_Worksheet_sndc, prjinfo);
            WritePrj2CPMSXls(_Worksheet_ssdc, prjinfo);
            WriteZJGTDisHZTJ2Xls(_Worksheet_sndc, _Worksheet_lqdc, _Worksheet_ssdc, prjinfo, prjdir, _RoadPart, _RoadDisList, disval, _MarkVal);

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
        public static void OutputDis_TH(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\调绘报表模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害调绘统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart_TH(path,_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart, _GPSInfo);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteDisLB2Xls_roadpart_TH(string outPath,MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
     DirectoryInfo prjdir, Disease[] arrdis, List<MilePart> roadpart, ExcelGPS[] gpsInfos)
        {
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

                        string tempStr = "\\" + arrdis[j].imgpath.Split('\\').Last();
                        string thPicPath = thDisPicPath;
                        string hyperPath = "调绘病害图片\\" + arrdis[j].imgname;

                        string picPath = prjinfo._PrjPath + System.IO.Path.Combine(arrdis[j].imgpath, arrdis[j].imgname);
                        Directory.CreateDirectory(thPicPath);
                        thPicPath += "\\" + arrdis[j].imgname;
                        File.Copy(picPath, thPicPath , true);



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

                   //     Framework.Office.Excel.CWB_ExcelHelper.InsertPicture(picRange, _Worksheet,"\\"+ hyperPath, ratio);
                        Framework.Office.Excel.CWB_ExcelHelper.InsertPicture(picRange, _Worksheet, picPath, ratio);

                        var hyperRange = _Worksheet.Range[$"M{3 + rowcnt}"];
                        hyperRange.ColumnWidth = widthC *4;
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


        #endregion
        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            //if (_Setting.ExcelType == 10)
            //{
            //    srcxls = string.Format(@"{0}\报表模板\低等级农村公路\甘肃\路面病害面积统计表.xlsx",
            //    System.Windows.Forms.Application.StartupPath);
            //}
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
           
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //if ( _Setting.ExcelType == 10)
            //{
            //    bool Haslqflag = false;
            //    bool Hassnflag = false;
            //    bool Hasssflag = false;

            //    MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            //    MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            //    MSExcel.Worksheet _Worksheet_sshz = _Workbook.Sheets["砂石病害汇总表"] as MSExcel.Worksheet;
            //    WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, _Worksheet_sshz,
            //        prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, ref Hasssflag, 5, 53);

            //}
            //else
            {
                MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
                WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);

                bool Haslqflag = false;
                bool Hassnflag = false;
                bool Hasssflag = false;

                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sshz = _Workbook.Sheets["砂石病害汇总表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, _Worksheet_sshz,
                    prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, ref Hasssflag, 5, 53);

                
                MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sstj = _Workbook.Sheets["砂石病害统计表"] as MSExcel.Worksheet;
                WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_sstj, prjdir, _RoadPart, Haslqflag, Hassnflag, Hasssflag);

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

        public static void OutputChongQingDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\四川定制\病害明细表.xlsx",
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
                        vallist[rowcnt, 2] = prjinfo._Direction ==1?"上行":"下行";
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

        public static void OutputDis_2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面损坏状况调查表2024.xlsx",
                System.Windows.Forms.Application.StartupPath); 
            string Destxls = string.Format(@"{0}\{1}_农村公路路面损坏状况调查表_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); 

                bool Haslqflag = false;
                bool Hassnflag = false;
                bool Hasssflag = false; 
                MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏状况调查表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥路面损坏状况调查表"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sshz = _Workbook.Sheets["砂石路面损坏状况调查表"] as MSExcel.Worksheet;
                WriteDisHZ2Xls_2024( disval,_Worksheet_snhz, _Worksheet_lqhz, _Worksheet_sshz,
                    prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, ref Hasssflag, 6, 53);

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
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 14, true);
            }
        }

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
        //private static void WriteDisHZ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
        //    ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
        //    ref bool Haslqflag, ref bool Hassnflag,
        //    int DataStartXlsxRow, int borderType)
        //{
        //    MSExcel.Range destrange;
        //    int disnum = 0;
        //    object[,] disval;

        //    Haslqflag = false;//有沥青路段标志
        //    Hassnflag = false;//有水泥路段标志

        //    int rowcnt_sn_s = DataStartXlsxRow;
        //    int rowcnt_sn_e = DataStartXlsxRow;//小计起始的计算范围
        //    int rowcnt_lq_s = DataStartXlsxRow;
        //    int rowcnt_lq_e = DataStartXlsxRow;

        //    int totalsnlen = 0;//水泥路段总长度
        //    int totallqlen = 0;//沥青路段总长度

        //    string errlog = prjdir.FullName + "\\errlog.txt";
        //    int len = roadpart.Count - 1, dlen = arrdis.Length;

        //    int typeidx = 0;
        //    bool res = false;

        //    for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
        //    {
        //        double drval = 0;
        //        int smile = roadpart[i].mile;
        //        int emile = roadpart[i + 1].mile;
        //        int milelength = Math.Abs(smile - emile);

        //        //统计位于这个区域的病害
        //        RoadDiseaseTypes.Clear();
        //        while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
        //            || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
        //        {
        //            res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
        //                    arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
        //            if (res)
        //            {
        //                RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
        //            }
        //            else
        //            {
        //                string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
        //                File.AppendAllText(errlog, errval, Encoding.UTF8);
        //            }
        //            ++j;
        //        }

        //        //病害汇总表
        //        int colcnt = 1;
        //        if (roadpart[i].roadtype == 1)//水泥
        //        {
        //            Hassnflag = true;
        //            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
        //            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
        //            worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = prjinfo._RoadNum;

        //            disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
        //            disval = new object[1, disnum + 3];
        //            for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
        //            {
        //                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
        //            }
        //            drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
        //            disval[0, disnum] = drval;
        //            disval[0, disnum + 1] = string.Format("=100-{0}*POWER(Y{1},{2})",
        //                _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
        //                rowcnt_sn_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

        //            disval[0, disnum + 2] = string.Format("=IF(Z{0}>={1},\"优\",IF(Z{0}>={2},\"良\",IF(Z{0}>={3},\"中\",IF(Z{0}>={4},\"次\",\"差\"))))",
        //                rowcnt_sn_s,
        //                _PCIGrade[roadpart[i].roaddegree][0],
        //                _PCIGrade[roadpart[i].roaddegree][1],
        //                _PCIGrade[roadpart[i].roaddegree][2],
        //                _PCIGrade[roadpart[i].roaddegree][3]);

        //            destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
        //            destrange.Value2 = disval;

        //            totalsnlen += milelength;
        //            rowcnt_sn_s++;
        //        }
        //        else if (roadpart[i].roadtype == 0)//沥青
        //        {
        //            Haslqflag = true;
        //            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
        //            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
        //            worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = prjinfo._RoadNum;

        //            disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
        //            disval = new object[1, disnum + 3];
        //            for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
        //            {
        //                disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
        //            }
        //            drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
        //            disval[0, disnum] = drval;
        //            disval[0, disnum + 1] = string.Format("=100-{0}*POWER(Z{1},{2})",
        //                _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
        //                rowcnt_lq_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

        //            disval[0, disnum + 2] = string.Format("=IF(AA{0}>={1},\"优\",IF(AA{0}>={2},\"良\",IF(AA{0}>={3},\"中\",IF(AA{0}>={4},\"次\",\"差\"))))",
        //                rowcnt_lq_s,
        //                _PCIGrade[roadpart[i].roaddegree][0],
        //                _PCIGrade[roadpart[i].roaddegree][1],
        //                _PCIGrade[roadpart[i].roaddegree][2],
        //                _PCIGrade[roadpart[i].roaddegree][3]);

        //            destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
        //            destrange.Value2 = disval;

        //            totallqlen += milelength;
        //            rowcnt_lq_s++;
        //        }

        //        if (_Setting.IsOutputDisAreaSubtotal)
        //        {
        //            if (emile % 1000 == 0)
        //            {
        //                if (roadpart[i].roadtype == 1)
        //                {
        //                    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
        //                    worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
        //                    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
        //                    disval = new object[1, disnum];
        //                    for (int di = 0; di < disnum; di++)
        //                    {
        //                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
        //                    }
        //                    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                    destrange.Value2 = disval;
        //                    rowcnt_sn_s++;
        //                    rowcnt_sn_e = rowcnt_sn_s;

        //                    if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
        //                    {
        //                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
        //                        worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
        //                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
        //                        disval = new object[1, disnum];
        //                        for (int di = 0; di < disnum; di++)
        //                        {
        //                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
        //                        }
        //                        destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                        destrange.Value2 = disval;
        //                        rowcnt_lq_s++;
        //                        rowcnt_lq_e = rowcnt_lq_s;
        //                    }
        //                }
        //                else if (roadpart[i].roadtype == 0)
        //                {
        //                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
        //                    worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
        //                    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
        //                    disval = new object[1, disnum];
        //                    for (int di = 0; di < disnum; di++)
        //                    {
        //                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
        //                    }
        //                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                    destrange.Value2 = disval;
        //                    rowcnt_lq_s++;
        //                    rowcnt_lq_e = rowcnt_lq_s;

        //                    if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
        //                    {
        //                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
        //                        worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
        //                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
        //                        disval = new object[1, disnum];
        //                        for (int di = 0; di < disnum; di++)
        //                        {
        //                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
        //                        }
        //                        destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                        destrange.Value2 = disval;
        //                        rowcnt_sn_s++;
        //                        rowcnt_sn_e = rowcnt_sn_s;
        //                    }
        //                }
        //            }
        //        }
        //    }

        //    //最后的一个小计
        //    if (_Setting.IsOutputDisAreaSubtotal)
        //    {
        //        if (roadpart[len].mile % 1000 != 0)
        //        {
        //            if (roadpart[len].roadtype == 1)
        //            {
        //                GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
        //                worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
        //                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
        //                disval = new object[1, disnum];
        //                for (int di = 0; di < disnum; di++)
        //                {
        //                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
        //                }
        //                destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                destrange.Value2 = disval;
        //                rowcnt_sn_s++;
        //                rowcnt_sn_e = rowcnt_sn_s;

        //                if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
        //                {
        //                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
        //                    worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
        //                    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
        //                    disval = new object[1, disnum];
        //                    for (int di = 0; di < disnum; di++)
        //                    {
        //                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
        //                    }
        //                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                    destrange.Value2 = disval;
        //                    rowcnt_lq_s++;
        //                    rowcnt_lq_e = rowcnt_lq_s;
        //                }
        //            }
        //            else if (roadpart[len].roadtype == 0)
        //            {
        //                GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
        //                worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
        //                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
        //                disval = new object[1, disnum];
        //                for (int di = 0; di < disnum; di++)
        //                {
        //                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
        //                }
        //                destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                destrange.Value2 = disval;
        //                rowcnt_lq_s++;
        //                rowcnt_lq_e = rowcnt_lq_s;

        //                if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
        //                {
        //                    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
        //                    worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
        //                    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
        //                    disval = new object[1, disnum];
        //                    for (int di = 0; di < disnum; di++)
        //                    {
        //                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
        //                    }
        //                    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //                    destrange.Value2 = disval;
        //                    rowcnt_sn_s++;
        //                    rowcnt_sn_e = rowcnt_sn_s;
        //                }
        //            }
        //        }
        //    }

        //    //总计
        //    //水泥
        //    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0);
        //    worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
        //    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
        //    disval = new object[1, disnum];
        //    for (int di = 0; di < disnum; di++)
        //    {
        //        if (_Setting.IsOutputDisAreaSubtotal)
        //        {
        //            disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
        //        }
        //        else
        //        {
        //            disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
        //        }
                
        //    }
        //    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //    destrange.Value2 = disval;

        //    //沥青
        //    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
        //    worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
        //    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
        //    disval = new object[1, disnum];
        //    for (int di = 0; di < disnum; di++)
        //    {
        //        if (_Setting.IsOutputDisAreaSubtotal)
        //        {
        //            disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
        //        }
        //        else
        //        {
        //            disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
        //        }
        //    }
        //    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
        //    destrange.Value2 = disval;

        //    destrange = worksheet_lqhz.get_Range(String.Format("A1:AB{0}", rowcnt_lq_s));
        //    GlobalExcel.SetBorderLine(destrange, borderType);
        //    destrange = worksheet_snhz.get_Range(String.Format("A1:AA{0}", rowcnt_sn_s));
        //    GlobalExcel.SetBorderLine(destrange, borderType);
        //    RoadDiseaseTypes.Clear();

        //    //if (_Setting.IsExcelSort)
        //    //{
        //    //    MSExcel.Range destrange1 = worksheet_lqhz.get_Range(string.Format("A3:K{0}", dlen + 2));
        //    //    MSExcel.Range sortrange = worksheet_lqhz.get_Range(string.Format("C3:C{0}", dlen + 2));
        //    //    GlobalExcel.ReflectionColnum(worksheet_lqhz, destrange1, sortrange);

        //    //    destrange1 = worksheet_snhz.get_Range(string.Format("A3:K{0}", dlen + 2));
        //    //     sortrange = worksheet_snhz.get_Range(string.Format("C3:C{0}", dlen + 2));
        //    //    GlobalExcel.ReflectionColnum(worksheet_snhz, destrange1, sortrange);
        //    //}



        //    if (!Haslqflag)
        //    {
        //        worksheet_lqhz.Delete();
        //    }

        //    if (!Hassnflag)
        //    {
        //        worksheet_snhz.Delete();
        //    }
        //}
        private static void WriteDisHZ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz, MSExcel.Worksheet worksheet_sshz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            ref bool Haslqflag, ref bool Hassnflag, ref bool Hasssflag,
            int DataStartXlsxRow, int borderType)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志
            Hasssflag = false;

            int rowcnt_sn_s = DataStartXlsxRow;
            int rowcnt_sn_e = DataStartXlsxRow;//小计起始的计算范围
            int rowcnt_lq_s = DataStartXlsxRow;
            int rowcnt_lq_e = DataStartXlsxRow;
            int rowcnt_ss_s = DataStartXlsxRow;
            int rowcnt_ss_e = DataStartXlsxRow;

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
                    try
                    {
                        if (prjinfo._Direction == -1)
                        {

                            if (_Setting.IsExcelSort)
                            {
                                worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] =emile;
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
                    }
                    catch (System.Exception)
                    {

                    }


                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        //if (_Setting.ExcelType == 10)
                        //{
                        //    if (RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname.Contains("裂缝"))
                        //    {
                        //        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totallength * RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].weight;

                        //    }
                        //    else
                        //        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea * RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].weight;

                        //}else

                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER({3}{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_sn_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        GlobalExcel.GetCol((char)('D' + disnum)));

                    disval[0, disnum + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt_sn_s,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3],
                        GlobalExcel.GetCol((char)('D' + disnum + 1)));

                    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
                    destrange.Value2 = disval;

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
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER({3}{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_lq_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                         GlobalExcel.GetCol((char)('D' + disnum)));

                    disval[0, disnum + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt_lq_s,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3],
                        GlobalExcel.GetCol((char)('D' + disnum + 1)));

                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
                    destrange.Value2 = disval;

                    rowcnt_lq_s++;
                }
                else if (roadpart[i].roadtype == 2)//砂石
                {
                    Hasssflag = true;
                    

                    if (prjinfo._Direction == -1)
                    {

                        if (_Setting.IsExcelSort)
                        {
                             worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                        }
                        else
                        {
                             worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                        }
                    }
                    else
                    {
                        worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                        worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                    }


                    worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                          
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER({3}{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_ss_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                         GlobalExcel.GetCol((char)('D' + disnum)));

                    disval[0, disnum + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                        rowcnt_ss_s,
                        _PCIGrade[roadpart[i].roaddegree][0],
                        _PCIGrade[roadpart[i].roaddegree][1],
                        _PCIGrade[roadpart[i].roaddegree][2],
                        _PCIGrade[roadpart[i].roaddegree][3],
                        GlobalExcel.GetCol((char)('D' + disnum + 1)));

                    destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
                    destrange.Value2 = disval;

                    rowcnt_ss_s++;
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
                            else if (Hasssflag && rowcnt_ss_e < rowcnt_ss_s)
                            {
                                GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                                worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                                disval = new object[1, disnum];
                                for (int di = 0; di < disnum; di++)
                                {
                                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                                }
                                destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                                destrange.Value2 = disval;
                                rowcnt_ss_s++;
                                rowcnt_ss_e = rowcnt_ss_s;
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
                            else if (Hasssflag && rowcnt_ss_e < rowcnt_ss_s)
                            {
                                GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                                worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                                disval = new object[1, disnum];
                                for (int di = 0; di < disnum; di++)
                                {
                                    disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                                }
                                destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                                destrange.Value2 = disval;
                                rowcnt_ss_s++;
                                rowcnt_ss_e = rowcnt_ss_s;
                            }
                        }
                        else if (roadpart[i].roadtype == 2)
                        {
                            GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                            worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                            }
                            destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_ss_s++;
                            rowcnt_ss_e = rowcnt_ss_s;

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
                            else if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
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
            if (_Setting.IsOutputDisAreaSubtotal)
            {
                //最后的一个小计
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
                        else if (Hasssflag && rowcnt_ss_e < rowcnt_ss_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                            worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                            }
                            destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_ss_s++;
                            rowcnt_ss_e = rowcnt_ss_s;
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
                        else if (Hasssflag && rowcnt_ss_e < rowcnt_ss_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                            worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                            disval = new object[1, disnum];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                            }
                            destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                            destrange.Value2 = disval;
                            rowcnt_ss_s++;
                            rowcnt_ss_e = rowcnt_ss_s;
                        }
                    }
                    else if (roadpart[len].roadtype == 2)
                    {
                        GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "小计", worksheet_sshz, 0);
                        worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_e, rowcnt_ss_s - 1);
                        }
                        destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                        destrange.Value2 = disval;
                        rowcnt_ss_s++;
                        rowcnt_ss_e = rowcnt_ss_s;

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
                        else if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
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
            if (_Setting.ExcelType != 10)
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
                    {
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
                    }
                    else
                    {
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
                    }

                }
                destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                destrange.Value2 = disval;

                //砂石
                GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "总计", worksheet_sshz, 0);
                worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                    {
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_s - 1);

                    }
                    else
                        disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_s - 1);
                }
                destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
                destrange.Value2 = disval;

            }
                

            destrange = worksheet_lqhz.get_Range(String.Format("A1:M{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_snhz.get_Range(String.Format("A1:M{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_sshz.get_Range(String.Format("A1:J{0}", rowcnt_ss_s));
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
                            MSExcel.Range destrange1 = worksheet_lqhz.get_Range(string.Format("A5:M{0}", rowcnt_lq_s - 1));
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
                            MSExcel.Range destrange1 = worksheet_snhz.get_Range(string.Format("A5:M{0}", rowcnt_sn_s - 1));
                            MSExcel.Range sortrange = worksheet_snhz.get_Range(string.Format("A5:A{0}", rowcnt_sn_s - 1));
                            GlobalExcel.ReflectionColnum(worksheet_snhz, destrange1, sortrange);
                        }
                         
                    }


                }
            }

            if (!Hasssflag)
            {
                worksheet_sshz.Delete();
            }
            else
            {
                if (!_Setting.IsOutputDisAreaSubtotal)
                {
                    if (prjinfo._Direction == -1)
                    {

                        
                        if (_Setting.IsExcelSort)
                        {
                            MSExcel.Range destrange1 = worksheet_sshz.get_Range(string.Format("A5:J{0}", rowcnt_ss_s - 1));
                            MSExcel.Range sortrange = worksheet_sshz.get_Range(string.Format("A5:A{0}", rowcnt_ss_s - 1));
                            GlobalExcel.ReflectionColnum(worksheet_sshz, destrange1, sortrange);
                        }
                    }


                }
            }
        }

        private static void WriteDisHZ2Xls_2024(int xlslen, MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz, MSExcel.Worksheet worksheet_sshz,
          ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
          ref bool Haslqflag, ref bool Hassnflag, ref bool Hasssflag,
          int DataStartXlsxRow, int borderType)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;

            Haslqflag = false;//有沥青路段标志
            Hassnflag = false;//有水泥路段标志
            Hasssflag = false;

            int rowcnt_sn_s = DataStartXlsxRow;
            int rowcnt_sn_e = DataStartXlsxRow;//小计起始的计算范围
            int rowcnt_lq_s = DataStartXlsxRow;
            int rowcnt_lq_e = DataStartXlsxRow;
            int rowcnt_ss_s = DataStartXlsxRow;
            int rowcnt_ss_e = DataStartXlsxRow;

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
                    try
                    {
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
                         
                    }
                    catch (System.Exception)
                    {

                    }


                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {  
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    } 
                    destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum-1))));
                    destrange.Value2 = disval; 
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
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }  
                    destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum-1))));
                    destrange.Value2 = disval; 
                    rowcnt_lq_s++;
                }
                else if (roadpart[i].roadtype == 2)//砂石
                {
                    Hasssflag = true; 
                    if (prjinfo._Direction == -1)
                    {

                        if (_Setting.IsExcelSort)
                        {
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                        }
                        else
                        {
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                            worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                        }
                    }
                    else
                    {
                        worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                        worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
                    }
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    destrange = worksheet_sshz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                    destrange.Value2 = disval;  
                    rowcnt_ss_s++;
                }
         
            }
        
           
                //总计
                //水泥
                GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0); 
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                        disval[0, di] = string.Format("=SUM({0}6:{0}{1})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_s - 1);
                   
                }
                destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                destrange.Value2 = disval;

                //沥青
                GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
               
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                  
                        disval[0, di] = string.Format("=SUM({0}6:{0}{1})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_s - 1);
                    
                }
                destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                destrange.Value2 = disval;

                //砂石
                GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "总计", worksheet_sshz, 0);
              
                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                disval = new object[1, disnum];
                for (int di = 0; di < disnum; di++)
                {
                        disval[0, di] = string.Format("=SUM({0}6:{0}{1})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_ss_s - 1);
                }
                destrange = worksheet_sshz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                destrange.Value2 = disval;


            destrange = worksheet_lqhz.get_Range(String.Format("A6:I{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_snhz.get_Range(String.Format("A6:I{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_sshz.get_Range(String.Format("A6:F{0}", rowcnt_ss_s));
            GlobalExcel.SetBorderLine(destrange, borderType); 

            RoadDiseaseTypes.Clear();

            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }
            else
            {
                worksheet_lqhz.Cells[2, 2] = prjinfo._RoadCode;
                worksheet_lqhz.Cells[2, 3] = "路线名称:" + prjinfo._RoadName;
                worksheet_lqhz.Cells[2, 5] = "调查方向:" + (prjinfo._Direction == 1 ? "上行" : "下行");
                worksheet_lqhz.Cells[2, 8] = "路面宽度:" + _RoadConfig.DetectWidth;

                worksheet_lqhz.Cells[3, 2] = prjinfo._StartMile.ToString("K0+000") + "~" + prjinfo._EndMile.ToString("K0+000");
                worksheet_lqhz.Cells[3, 3] = "单元长度:" + xlslen;

                string input = prjinfo._DataDate;
                string result = "";
                if (input.Length == 8 && int.TryParse(input, out _))
                {
                    int year = int.Parse(input.Substring(0, 4));
                    int month = int.Parse(input.Substring(4, 2));
                    int day = int.Parse(input.Substring(6, 2));

                    DateTime date;
                    if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                    {
                        result = date.ToString("yyyy年MM月dd日");

                    }
                    else
                    {
                        result = "";
                    }
                }
                else
                {
                    result = "";
                }

                worksheet_lqhz.Cells.Cells[3, 5] = "调查时间:" + result;
                worksheet_lqhz.Cells.Cells[3, 8] = "调查人员:" + prjinfo._DataPerson;
            }
           

            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }
            else
            {
               worksheet_snhz.Cells[2, 2] = prjinfo._RoadCode;
               worksheet_snhz.Cells[2, 3] = "路线名称:" + prjinfo._RoadName;
               worksheet_snhz.Cells[2, 5] = "调查方向:" + (prjinfo._Direction == 1 ? "上行" : "下行");
                worksheet_snhz.Cells[2, 8] = "路面宽度:" + _RoadConfig.DetectWidth;

               worksheet_snhz.Cells[3, 2] = prjinfo._StartMile.ToString("K0+000") + "~" + prjinfo._EndMile.ToString("K0+000");
                worksheet_snhz.Cells[3, 3] = "单元长度:" + xlslen;

                string input = prjinfo._DataDate;
                string result = "";
                if (input.Length == 8 && int.TryParse(input, out _))
                {
                    int year = int.Parse(input.Substring(0, 4));
                    int month = int.Parse(input.Substring(4, 2));
                    int day = int.Parse(input.Substring(6, 2));

                    DateTime date;
                    if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                    {
                        result = date.ToString("yyyy年MM月dd日");

                    }
                    else
                    {
                        result = "";
                    }
                }
                else
                {
                    result = "";
                }

                worksheet_snhz.Cells.Cells[3, 5] = "调查时间:" + result;
                worksheet_snhz.Cells.Cells[3, 8] = "调查人员:" + prjinfo._DataPerson;
            }
            if (!Hasssflag)
            {
                worksheet_sshz.Delete();
            }
            else
            {
               worksheet_sshz.Cells[2, 2] = prjinfo._RoadCode;
               worksheet_sshz.Cells[2, 3] = "路线名称:" + prjinfo._RoadName;
               worksheet_sshz.Cells[2, 4] = "调查方向:" + (prjinfo._Direction == 1 ? "上行" : "下行");
                worksheet_sshz.Cells[2, 6] = "路面宽度:" + _RoadConfig.DetectWidth;

               worksheet_sshz.Cells[3, 2] = prjinfo._StartMile.ToString("K0+000") + "~" + prjinfo._EndMile.ToString("K0+000");
                worksheet_sshz.Cells[3, 3] = "单元长度:" + xlslen;

                string input = prjinfo._DataDate;
                string result = "";
                if (input.Length == 8 && int.TryParse(input, out _))
                {
                    int year = int.Parse(input.Substring(0, 4));
                    int month = int.Parse(input.Substring(4, 2));
                    int day = int.Parse(input.Substring(6, 2));

                    DateTime date;
                    if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                    {
                        result = date.ToString("yyyy年MM月dd日");

                    }
                    else
                    {
                        result = "";
                    }
                }
                else
                {
                    result = "";
                }

               worksheet_sshz.Cells.Cells[3, 4] = "调查时间:" + result;
                worksheet_sshz.Cells.Cells[3, 6] = "调查人员:" + prjinfo._DataPerson;
            }
        }
          


        private static void WriteDisTJ2Xls(MSExcel.Worksheet worksheet_sntj, MSExcel.Worksheet worksheet_lqtj, MSExcel.Worksheet worksheet_sstj,
            DirectoryInfo prjdir, List<MilePart> roadpart, bool Haslqflag, bool Hassnflag, bool Hasssflag)
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
                worksheet_lqtj.Cells[2, 5] = Math.Abs(roadpart[0].mile - roadpart[len].mile);

                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    else
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/2", Convert.ToChar('D' + i));
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
                worksheet_sntj.Cells[2, 5] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    else
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/2", Convert.ToChar('D' + i));

                }
                destrange = worksheet_sntj.get_Range("C4:C" + (disnum + 3).ToString());
                destrange.Value2 = disval;
            }
            else
            {
                worksheet_sntj.Delete();
            }

            if (Hasssflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
                worksheet_sstj.Cells[2, 2] = _RoadConfig.DetectWidth;
                worksheet_sstj.Cells[2, 5] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    if (_Setting.IsOutputDisAreaSubtotal)
                        disval[i, 0] = string.Format("=SUMIF(砂石病害汇总表!{0}:{0},\"<>\",砂石病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    else
                        disval[i, 0] = string.Format("=SUMIF(砂石病害汇总表!{0}:{0},\"<>\",砂石病害汇总表!{0}:{0})/2", Convert.ToChar('D' + i));

                }
                destrange = worksheet_sstj.get_Range("C4:C" + (disnum + 3).ToString());
                destrange.Value2 = disval;
            }
            else
            {
                worksheet_sstj.Delete();
            }
        }

        private static double ComputPCI(RoadDiseaseType[][] disarea, int roadtype, double partarea)
        {
            double sumarea = 0;
            int len = _RoadSocre[roadtype].Keys.Count;

            for (int i = 0; i < len; i++)
            {
              /*  if (disarea[roadtype][i].computetype == 1 || disarea[roadtype][i].computetype == 3 || disarea[roadtype][i].computetype == 4)
                {
                    sumarea += disarea[roadtype][i].totalarea * disarea[roadtype][i].totallength;    
                }
                else
                {*/
                    //车辙病害计算结果不一致
                    sumarea += disarea[roadtype][i].totalarea * disarea[roadtype][i].weight;    
               // }

            }
            return 100 * sumarea / partarea;
        }

        public static void OutputPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面破损评价等级记录表.xlsx",
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
                if (drval>150)
                {
                    try
                    {
                        string errorFile = "C:\\Users\\Administrator\\Desktop\\新建文件夹\\errorRoad.txt";
                        using (StreamWriter sw = new StreamWriter(errorFile, true))
                        {
                            sw.WriteLine($"错误工程{prjinfo._RoadCode}_{prjinfo._RoadName},line:{i},errorDrValue{drval}");
                        }
                    }
                    catch (Exception)
                    {
                        
                    } 

                }


                vallist[i, 3] = drval;

                vallist[i, 4] = string.Format("=100-{1}*POWER(D{0},{2})",
                    i + DataStartXlsxRow,
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                double d1 = _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0];
                double d2 = _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1];

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


     

       


        public static void  贵州农村公路检测数据明细表(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            int StartRowInd = 5;
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\贵州2024\农村公路路况检测PQI.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_贵州======.xlsx", path, prjdir.Name); // 文件命名

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode:XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["路况检测PQI"] as MSExcel.Worksheet;


            #region 地区、行政编号、检测日期
            string 市州 = prjinfo._City;
            string 区县 = prjinfo._District;
            string RegionCode = prjinfo._RoadCode.Substring(4, 6) + "000000"; // 12位行政区划编码，取道路编号后6位，再补6位0


            string DetectedDate = ""; // 检测日期
            var timeFormat = "yyyyMMdd";
            if (DateTime.TryParseExact(prjinfo._DataDate, timeFormat, null, System.Globalization.DateTimeStyles.None, out var dt))
            {
                DetectedDate = dt.ToString("yyyy/MM/dd");
            }
            else
            {
                MessageBox.Show("日期格式有误，请检查！");
                // 日期输入错误
            }
            
            #endregion


            #region 年报数据
            string 行政等级 = ""; // 国道、省道、县道、乡道、村道
            string 路线编码 = prjinfo._RoadCode;
            string 路线名称 = prjinfo._RoadName;

            double 养护里程 = 0.000; // ======================================未知
            #endregion


            #region 智能化检测数据
            // 获取roadpart
            List<MilePart> roadpart = _RoadPart;
            double[] LIRIVal = _LIRIMeanVal;
            double[] RIRIVal = _RIRIMeanVal;
            
            string GetPQI(int i, double mileLength)
            {
                // 准备计算PQI所需的参数
                var drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * mileLength);
                var pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                double irival; //IRI 使用平均值 还是 较大值？
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.RQIJudgeType == 0)
                    {
                        irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                    }
                    else
                    {
                        irival = Math.Round(Math.Max(LIRIVal[i], RIRIVal[i]), 5);
                    }
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                var trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                var weight1 = _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0];
                var weight2 = _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1];
                // 根据公式计算PQI
                var PQI = (weight1 * pcival + weight2 * trqival) / (weight1 + weight2);
                return PQI.ToString("f2");
            }
            
            
          
             

            double 路基宽度 = _RoadConfig.DetectWidth;
            double 路面检测宽度 = _RoadConfig.DetectWidth;
            string[] roadType = new string[] { "沥青", "水泥", "砂石" };
            string 路面类型 = roadType[prjinfo._RoadType];
            string 备注 = "";

            #endregion

            object[,] result = new object[roadpart.Count-1, 22];
            for (int ind = 0; ind < roadpart.Count-1 ; ind++)
            {
                double startMile = roadpart[ind].mile;
                double endMile = roadpart[ind+1].mile;
                double mileLength = Math.Abs(endMile - startMile);
                string[] line = new string[22]
                {
                    (ind+1).ToString(),
                    市州,
                    区县,
                    RegionCode,
                    DetectedDate,
                    行政等级,
                    路线编码,
                    路线名称,
                    养护里程.ToString("f3"),
                    startMile.ToString(),
                    endMile.ToString(),
                    mileLength.ToString(),
                    GetPQI(ind, mileLength),
                    "",
                    "",
                    "",
                    "",
                    "",
                    路基宽度.ToString(),
                    路面检测宽度.ToString(),
                    路面类型,
                    备注
                };
                for(int i=0; i<line.Length; i++)
                {
                    result[ind, i] = line[i];
                }
                
            }
            var range = _Worksheet.get_Range($"A{StartRowInd}:V{StartRowInd + result.GetLength(0) - 1}");
            range.Value2 = result;
            GlobalExcel.SetBorderLine(range, 53);



            // 

            // 加框
            //var range = _Worksheet.get_Range(String.Format("A{0}:K{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            //GlobalExcel.SetBorderLine(range, 53);



            // 获取GPS信息
            //GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart10M, ref dicGps); 


            //var Grade_Exp = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
            //        i + DataStartXlsxRow,
            //        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
            //        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
            //        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
            //        _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);


            /*
            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 4, 53);
            WriteStatistics_XMJH(_Worksheet, 4, 3, 22, 'H', "平整度", 1);
            if (_Setting.Out_roadinfo == 0)
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            
            if (_Setting.ExcelType != 10)
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
            */


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void 贵州农村公路检测轨迹(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\贵州2024\农村公路路况检测轨迹.xlsx",
                System.Windows.Forms.Application.StartupPath);
            //string Destxls = string.Format(@"{0}\{1}_IRI_{2}m.xlsx", path, prjdir.Name, disval); // 文件命名

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            //_Workbook.SaveAs(Destxls, AccessMode:XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["路况检测轨迹"] as MSExcel.Worksheet;

            
            var roadpart = _RoadPart;
            int total_row = roadpart.Count;
            object[,] result = new object[total_row, 13];

                        int ind = 0;      // == 读取每一个roadpart， 设法获取检测轨迹文件的名称
                        string filename = "";
                        double 轨迹长度 = 0.000;
                        double 路线长度 = 0.000;

                        double startMile = roadpart[ind].mile;
                        double endMile = roadpart[ind+1].mile;
                        double mileLength = Math.Abs(endMile - startMile);

                        string dateTime = ""; // 检测日期
                        string fileTime = ""; //            ========从文件名中读取检测时间
                        var timeFormat = "yyyyMMddHHmmss";
                        if (DateTime.TryParseExact(fileTime, timeFormat, null, System.Globalization.DateTimeStyles.None, out var dt))
                        {
                            dateTime = dt.ToString("yyyy/MM/dd HH:mm:ss");
                        }
                        else
                        {
                            MessageBox.Show("日期格式有误，请检查！");
                            // 日期输入错误
                        }
                        string 备注 = "";

                        object[] line = new string[13] {
                            (ind+1).ToString(),
                            prjinfo._District,
                            prjinfo._RoadCode.Substring(4, 6) + "000000", // 12位行政区划编码，取道路编号后6位，再补6位0
                            filename,
                            轨迹长度.ToString("f3"),
                            prjinfo._RoadCode,
                            prjinfo._RoadName,
                            路线长度.ToString("f3"),
                            startMile.ToString("f3"),
                            endMile.ToString("f3"),
                            mileLength.ToString("f3"),
                            dateTime,
                            备注
                        };

                        for(int i=0; i<line.Length; i++)
                        {
                            result[ind, i] = line[i];   
                        }



            MSExcel.Range range = _Worksheet.get_Range($"A2:M{total_row}");
            range.Value2 = result;
            GlobalExcel.SetBorderLine(range, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();


        }

        public static void 贵州农村公路检测图片(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\贵州2024\农村公路路况检测照片.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_贵州.xlsx", path, prjdir.Name); // 文件命名==============

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
            _Workbook.SaveAs(Destxls, AccessMode:XlSaveAsAccessMode.xlNoChange);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["路况检测图片"] as MSExcel.Worksheet;
 

            var res_dic = HighAccuracyPositioning.March_GPS(prjinfo._PrjPath);
            int ind = 0;
            int total_row = res_dic.Keys.Count + 1;
            object[,] result = new object[total_row, 13];
            string 行政区划编码 = prjinfo._RoadCode.Substring(4, 6) + "000000"; // 12位行政区划编码，取道路编号后6位，再补6位0

            foreach (var 桩号m in res_dic.Keys)
            {
                string filename = ""; // 图片名称
                double x = 0, y = 0, z = 0; // 经纬度
                double 间隔距离 = 0; // 如何获取？============  50m
                string 备注 = "";
                double 所在桩号km = double.Parse(桩号m) / 1000;
                ((x,y,z), filename) = res_dic[桩号m];

                // 按照贵州要求的图片命名格式
                
                string picTime = prjinfo._DataDate + filename.AsSpan().Slice(16, 6).ToString();
                string page = "01"; // 分文件夹
                char dir = prjinfo._Direction > 0 ? 'A' : 'B';
                string formated_FileName = $"{prjinfo._RoadCode + dir}-{所在桩号km.ToString("f3")}-{page}-{picTime}.jpg";

                string dateTime = ""; // 检测日期
                var timeFormat = "yyyyMMddHHmmss";
                if (DateTime.TryParseExact(picTime, timeFormat,null,System.Globalization.DateTimeStyles.None, out var dt))
                {
                    dateTime = dt.ToString("yyyy/MM/dd HH:mm:ss");
                }
                else
                {
                    MessageBox.Show("日期格式有误，请检查！");
                    // 日期输入错误
                }

                object[] line = new object[13] {
                    (ind+1).ToString(),
                    prjinfo._District,
                    行政区划编码,
                    formated_FileName,
                    所在桩号km.ToString("f3"),
                    间隔距离.ToString("f3"),
                    prjinfo._RoadCode,
                    prjinfo._RoadName,
                    x.ToString("f8"),
                    y.ToString("f8"),
                    y.ToString("f8"),
                    dateTime,
                    备注
                };
                for(int i=0; i<line.Length; i++)
                {
                    result[ind, i] = line[i];
                }
                ind += 1;
            }
            MSExcel.Range range = _Worksheet.get_Range($"A2:M{total_row}");
            range.Value2 = result;
            GlobalExcel.SetBorderLine(range, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

                }

        #region 江西农村路
        public static void 江西农村路沥青病害(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, RoadConfig roadConfig)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\江西2024\低等级沥青路面破损率.xlsx",
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
                if (startMileD > endMileStrD)
                {
                    (startMileD, endMileStrD) = (endMileStrD, startMileD);
                }

                // 文件命名==============
                string Destxls = string.Format(@"{0}\{1}-DR-低等级沥青路面破损-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                (startMileD).ToString("f3"),
                                                                                                (endMileStrD).ToString("f3"));

                startMileD *= 1000; endMileStrD *= 1000; // 转为整形形式
                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, ReadOnly: true);
                _Workbook.SaveAs(Destxls, AccessMode: XlSaveAsAccessMode.xlNoChange);
                MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;

                var roadpart = _RoadPart;
                var arrdis = _RoadDisList;
                int roadType = 0;
                bool has = false;

                string errlog = prjdir.FullName + "\\errlog.txt";
                int startRow = 5;
                int rowCount = 0;
                int len = roadpart.Count - 1, dlen = arrdis.Length;
                object[,] datas = new object[len, 12];
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
                    if (roadpart[i].roadtype == roadType && (startMileD <= roadpart[i].mile && roadpart[i].mile <= endMileStrD))
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
#if 江西车检2024病害顺序
                            if (dis == 1 || dis == 3)
                            {
                                continue;
                            }
#else
                            if (dis == 0 || dis == 2)
                            { // 不统计轻度裂缝
                                //continue;
                            }
#endif



                            datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f3");
                        }

                        datas[rowCount, 10] = roadConfig.DetectWidth;
                        datas[rowCount, 11] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength); // 结果示例表中 未格式化输出
                        rowCount++;
                    }
                }
                if (has)
                {

                    MSExcel.Range destrange = _Worksheet.get_Range($"A{startRow}:L{startRow + rowCount - 1}");
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

        public static void 江西农村路水泥病害(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, RoadConfig roadConfig)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\江西2024\低等级水泥路面破损率.xlsx",
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

                string Destxls = string.Format(@"{0}\{1}-DR-低等级水泥路面破损-{2}-{3}.xlsx", path, prjinfo._RoadCode + ((prjinfo._Direction > 0) ? "A" : "B"),
                                                                                                (startMileD).ToString("f3"),
                                                                                                (endMileStrD).ToString("f3"));

                startMileD *= 1000; endMileStrD *= 1000; // 转为整形形式
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
                object[,] datas = new object[len, 13];
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
                    if (roadpart[i].roadtype == roadType && (startMileD <= roadpart[i].mile && roadpart[i].mile <= endMileStrD))
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
                        double[] area =  new double[colcnt];
                        for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                        {
#if 江西车检2024病害顺序

                            if (dis==1)
                            {
                                continue;
                            }
                            datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f3");
#else
 if (dis == 1)
                            { // 不统计轻度裂缝
                                continue;
                            }
                            else if (dis ==0  || dis == 2)
                            {
                                datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f3");

                            }
                            
                            else
                            {
                                area[dis-3] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea ;
                            }
#endif


                        }
#if 江西车检2024病害顺序
#else
 for (int ttt = 0; ttt < 4; ttt++)
                        {
                            datas[rowCount, colcnt++] = area[ttt]. ToString("f3");
                        }
#endif


                        datas[rowCount, 11] = roadConfig.DetectWidth;
                        datas[rowCount, 12] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength); // 结果示例表中 未格式化输出
                        rowCount++;
                    }
                }
                if (has)
                {
                    int startRow = 5;
                    MSExcel.Range destrange = _Worksheet.get_Range($"A{startRow}:M{startRow + rowCount - 1}");
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

        public static void 江西农村路平整度(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, RoadConfig roadConfig)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\江西2024\低等级路面平整度.xlsx",
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
           
            int colcnt;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
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

            int startRow = 3;
            MSExcel.Range destrange = _Worksheet.get_Range($"A{startRow}:H{startRow + rowCount - 1}");
            destrange.Value2 = datas;
            GlobalExcel.SetBorderLine(destrange, 53);
            

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

#endregion


        public static void OutputPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面综合评价等级记录表.xlsx",
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
        public static void OutputPQI_2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\技术状况评定汇总表_2024.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_农村公路技术状况评定汇总表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePQI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _SpeedVal, _MarkVal);


            if (_StreetDisRecord.Count > 0)
            {
                WriteStreetTCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(),'X',4);
            }
            if (_StreetDisRecord_RoadBed.Count > 0)
            {
                WriteRoadBedSCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(),'V',4);
            }
            //读取数据
            int len = _RoadPart.Count - 1;
            MSExcel.Range range =  _Worksheet.get_Range(String.Format("A5:Z{0}", len+5-1));
            object[,] datas =  range.Value2;
            range.ClearContents();
            // 设置整个 Borders 集合的 LineStyle 为无线条样式
            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            double mqiValue = 0; 
            double pqiValue = 0;
            double sumLen = 0;
            double[] mqiJudge = new double[5];
            double[] pqiJudge = new double[5];
            for (int i = 1; i < len+1; i++)
            {
                double roadLen = Convert.ToDouble( datas[i, 5]);
                double MQISingle = Convert.ToDouble(datas[i, 6]);
                double PQISingle = Convert.ToDouble(datas[i, 14]);
                sumLen += roadLen;
                pqiValue += PQISingle*  roadLen;
                mqiValue += MQISingle * roadLen;
                mqiJudge[getJudgeArea(_MQIGrade, MQISingle)] += roadLen;
                pqiJudge[getJudgeArea(_PQIGrade[_RoadPart[i - 1].roaddegree], PQISingle)] += roadLen;
            }
            pqiValue /= sumLen;
            mqiValue /= sumLen;

            object[,] roadData = new object[1,21];

            int colCnt = 0;
            roadData[0,colCnt++] = prjinfo._RoadCode;
            roadData[0,colCnt++] = prjinfo._RoadName;
            roadData[0,colCnt++] = _RoadPart.First().mile;
            roadData[0, colCnt++] = _RoadPart.Last().mile;
            roadData[0, colCnt++] = "=ABS(C5-D5)/1000";
            roadData[0, colCnt++] = mqiValue;
            for (int i = 0; i < 5; i++)
            {
                roadData[0, colCnt++] = string.Format("={0}/1000",mqiJudge[i]);
            }
            roadData[0, colCnt++] = "=(G5+H5)/E5*100";
            roadData[0, colCnt++] = "=(G5+H5+I5)/E5*100";
            roadData[0, colCnt++] = pqiValue;
            for (int i = 0; i < 5; i++)
            {
                roadData[0, colCnt++] = string.Format("={0}/1000", pqiJudge[i]); 
            }
            roadData[0, colCnt++] = "=(O5+P5)/E5*100";
            roadData[0, colCnt++] = "=(O5+P5+Q5)/E5*100"; 

            MSExcel.Range destrange = _Worksheet.get_Range("A5:U5");
            destrange.Value2 = roadData;
            GlobalExcel.SetBorderLine(destrange, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void OutputPci_2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面破损评定汇总表_2024.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_路面破损评定汇总表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;


            WritePCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _SpeedVal, _MarkVal, 5);

            //读取数据
            int len = _RoadPart.Count - 1;
            MSExcel.Range range = _Worksheet.get_Range(String.Format("A5:F{0}", len + 5 - 1));
            object[,] datas = range.Value2;
            range.ClearContents();
            // 设置整个 Borders 集合的 LineStyle 为无线条样式
            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            double pciValue = 0; 
            double sumLen = 0;
            double[] pciJudge = new double[5]; 
            for (int i = 1; i < len + 1; i++)
            {
                double roadLen = Convert.ToDouble(datas[i, 5]);
                double PCISingle = Convert.ToDouble(datas[i, 6]); 
                sumLen += roadLen;
                pciValue += PCISingle * roadLen;
                pciJudge[getJudgeArea(_PCIGrade[_RoadPart[i - 1].roaddegree], PCISingle)] += roadLen;
                
                
            }
            pciValue /= sumLen; 

            object[,] roadData = new object[1, 14];

            int colCnt = 0;
            roadData[0, colCnt++] = prjinfo._RoadCode;
            roadData[0, colCnt++] = prjinfo._RoadName;
            roadData[0, colCnt++] = _RoadPart.First().mile;
            roadData[0, colCnt++] = _RoadPart.Last().mile;
            roadData[0, colCnt++] = "=ABS(C5-D5)/1000";
            roadData[0, colCnt++] = pciValue;
            for (int i = 0; i < 5; i++)
            {
                roadData[0, colCnt++] = string.Format("={0}/1000", pciJudge[i]);
            }
            roadData[0, colCnt++] = "=(G5+H5)/E5*100";
            roadData[0, colCnt++] = "=(G5+H5+I5)/E5*100";
            roadData[0, colCnt++] = "=(J5+K5)/E5*100";
           

            MSExcel.Range destrange = _Worksheet.get_Range("A5:N5");
            destrange.Value2 = roadData;
            GlobalExcel.SetBorderLine(destrange, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        private static void WritePCI2Xls_2024(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
       List<MilePart> roadpart, Disease[] arrdis, double[] SpeedVal, string[] MarkVal,
       int DataStartXlsxRow)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 6];

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

                vallist[i, 0] = prjinfo._RoadNum;

                vallist[i, 1] = prjinfo._RoadName;
                vallist[i, 2] = smile;
                vallist[i, 3] = emile;
                vallist[i, 4] = milelength;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 5] = string.Format("=100-{1}*POWER({0},{2})",
                    drval,
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
            }

            destrange = worksheet.get_Range(String.Format("A{0}:F{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;
        }


        public static void OutputRqi_2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\路面平整度评定汇总表_2024.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_路面平整度评定汇总表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteIRI2Xls_2024(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal, _MarkVal, 5); 

            //读取数据
            int len = _RoadPart.Count - 1;
            MSExcel.Range range = _Worksheet.get_Range(String.Format("A5:F{0}", len + 5 - 1));
            object[,] datas = range.Value2;
            range.ClearContents();
            // 设置整个 Borders 集合的 LineStyle 为无线条样式
            range.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            double rqiValue = 0;
            double sumLen = 0;
            double[] pqiJudge = new double[5];
            for (int i = 1; i < len + 1; i++)
            {
                double roadLen = Convert.ToDouble(datas[i, 5]);
                double RQISingle = Convert.ToDouble(datas[i, 6]);
                sumLen += roadLen;
                rqiValue += RQISingle * roadLen;
                pqiJudge[getJudgeArea(_RQIGrade[_RoadPart[i - 1].roaddegree][prjinfo._RoadType], RQISingle)] += roadLen;


            }
            rqiValue /= sumLen;

            object[,] roadData = new object[1, 14];

            int colCnt = 0;
            roadData[0, colCnt++] = prjinfo._RoadCode;
            roadData[0, colCnt++] = prjinfo._RoadName;
            roadData[0, colCnt++] = _RoadPart.First().mile;
            roadData[0, colCnt++] = _RoadPart.Last().mile;
            roadData[0, colCnt++] = "=ABS(C5-D5)/1000";
            roadData[0, colCnt++] = rqiValue;
            for (int i = 0; i < 5; i++)
            {
                roadData[0, colCnt++] = string.Format("={0}/1000", pqiJudge[i]);
            }
            roadData[0, colCnt++] = "=(G5+H5)/E5*100";
            roadData[0, colCnt++] = "=(G5+H5+I5)/E5*100";
            roadData[0, colCnt++] = "=(J5+K5)/E5*100";


            MSExcel.Range destrange = _Worksheet.get_Range("A5:N5");
            destrange.Value2 = roadData;
            GlobalExcel.SetBorderLine(destrange, 53);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteIRI2Xls_2024(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
  List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, double[] SpeedVal, string[] MarkVal,
  int DataStartXlsxRow)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 6];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = prjinfo._RoadCode;
                vallist[i, 1] = prjinfo._RoadName;
                vallist[i, 2] = roadpart[i].mile;
                vallist[i, 3] = roadpart[i + 1].mile;
                vallist[i, 4] = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile);


                double iriLeft = 0
                    ;
                double iriRight = 0;
                double iriValue = 0;
                //进行iri速度矫正的
                if (_Setting.RQIJudgeType == 2)
                {
                    iriLeft = huNanJudgeIri(SpeedVal[i], LIRIVal[i]);

                    if (prjinfo._IsDIRIMTD)
                    {

                        iriRight = huNanJudgeIri(SpeedVal[i], RIRIVal[i]);
                        if (_Setting.RQIJudgeType == 0)
                        {
                            iriValue = (iriLeft + iriRight) / 2;

                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {
                            iriValue = Math.Max(iriLeft, iriRight);
                        }
                    }
                    else
                    {
                        iriValue = iriLeft;
                    }

                }
                else
                {

                    iriLeft = LIRIVal[i];
                    if (prjinfo._IsDIRIMTD)
                    {
                        iriRight = RIRIVal[i];
                        if (_Setting.RQIJudgeType == 0)
                        {

                            iriValue = (iriLeft + iriRight) / 2;
                        }
                        else if (_Setting.RQIJudgeType == 1)
                        {

                            iriValue = Math.Max(iriLeft, iriRight);
                        }
                    }
                    else
                    {

                        iriValue = iriLeft;
                    }

                }

                vallist[i, 5] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{2})),5)",
                    _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], iriValue);

            }

            destrange = _Worksheet.get_Range(String.Format("A{0}:F{1}", DataStartXlsxRow, len + DataStartXlsxRow - 1));
            destrange.Value2 = vallist;


        }

        private static int getJudgeArea(double[] judgeArea, double value)
        {
            int index = 4; // 初始化为-1，表示未找到合适的区间

            for (int i = 0; i < judgeArea.Length; i++)
            {
                if (value >= judgeArea[i])
                {
                    index = i;
                    break; // 找到合适的区间后立即退出循环
                }
            }
            return index;
        }

        public static void OutputChongQingSumExcel(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\四川定制\重庆公里指标导入模板.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_公里指标_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["每公里指标 (病害)"] as MSExcel.Worksheet;
            WriteAll2Xls_ChongQing(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
               _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _DeltaHVal, disval, _GPSInfo); 
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
                double pcival = 0, drval = 0;
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

                //IRI
                if (prjinfo._IsDIRIMTD)
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

                colcnt = colcnt + 6;

                vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),5)", rowcnt + 3,
                     _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                     _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

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


        private static void WritePQI2Xls_2024(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
           List<MilePart> roadpart, Disease[] arrdis,
           double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
           double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal, double[] SpeedVal, string[] MarkVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0, wrval = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 26];

            int typeidx = 0;
            bool res = false;

            string input = prjinfo._DataDate;
            worksheet.Cells[2, 2] = "";
            if (input.Length == 8)
            {
                worksheet.Cells[2, 13] = input.Substring(0, 4);
                worksheet.Cells[2, 15] = input.Substring(4, 2);
                worksheet.Cells[2, 17] = input.Substring(6, 2);
            }
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval = 0;
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
                vallist[rowcnt, colcnt++] = prjinfo._RoadCode;
                vallist[rowcnt, colcnt++] = prjinfo._RoadName;
                vallist[rowcnt, colcnt++] = smile;
                vallist[rowcnt, colcnt++] = emile;
                vallist[rowcnt, colcnt++] = milelength;
                vallist[rowcnt, colcnt++] = string.Format("=V{0}*{1}+N{0}*{2}+W{0}*{3}+X{0}*{4}", rowcnt+5, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] = "/";

                vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*Y{0}+{2}*Z{0})/({1}+{2}),5)", rowcnt + 5,
                 _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                 _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] =  "/";
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] = "/";
                vallist[rowcnt, colcnt++] = 100;
                vallist[rowcnt, colcnt++] = 100;
                vallist[rowcnt, colcnt++] = 100;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt++] = Math.Round(pcival, 5); 
                //IRI
                if (prjinfo._IsDIRIMTD)
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
                else
                {

                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, colcnt++] = trqival; 
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:Z{0}", rowcnt+4));
            destrange.Value2 = vallist;
            destrange = worksheet.get_Range(String.Format("A5:U{0}", rowcnt + 4));

            GlobalExcel.SetBorderLine(destrange, 53); 
        }

        //技术状况评定明细表
        public static void OutputPDMX(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\技术状况评定明细表.xlsx",
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
                GlobalExcel.Reflection(_Worksheet, 5, 1, 10, true);
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

        public static void OutputPDMX_2024(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\技术状况评定明细表_2024.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_农村公路技术状况评定明细表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePDMX2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal);
            if (_StreetDisRecord.Count > 0)
            {
                WriteStreetTCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(),'O',3);
            }
            if (_StreetDisRecord_RoadBed.Count > 0)
            {
                WriteRoadBedSCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(),'L',3);
            }

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet,4, 4, 20, true);
                GlobalExcel.Reflection(_Worksheet,4, 4, 5, false);
            } 
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputPDMX_2024Sum(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\农村公路技术状况数据汇总表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_农村公路技术状况数据汇总表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WritePDMX2Xls_2024Sum(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal);
            if (_StreetDisRecord.Count > 0)
            {
                WriteStreetTCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(), 'Q', 2);
            }
            if (_StreetDisRecord_RoadBed.Count > 0)
            {
                WriteRoadBedSCI2Xls_2024(_Worksheet, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), 'O', 2);
            }

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 4, 25, true);
                GlobalExcel.Reflection(_Worksheet, 4, 4, 5, false);
            }
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
            double trqival = 0, irival = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 10];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval = 0;
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
               
                vallist[rowcnt, 2] = string.Format("=G{0}*{1}+D{0}*{2}+H{0}*{3}+I{0}*{4}", rowcnt + 5, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);
                vallist[rowcnt, 6] = "100";
                vallist[rowcnt, 7] = "100";
                vallist[rowcnt, 8] = "100";

                vallist[rowcnt, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] );

                vallist[rowcnt, 4] = Math.Round(pcival, 5);

                //IRI
                if (prjinfo._IsDIRIMTD)
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
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, 5] = Math.Round(trqival, 5);

                vallist[rowcnt, 3] = string.Format("=ROUND(({1}*E{0}+{2}*F{0})/({1}+{2}),5)", rowcnt + 5,
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]); 
                    ++rowcnt;
            } 
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:J{0}", rowcnt + 4));
            destrange.Value2 = vallist;
            destrange = worksheet.get_Range(String.Format("A5:J{0}", rowcnt + 5));
            GlobalExcel.SetBorderLine(destrange, 63);

            worksheet.Cells[2, 2] = prjinfo._RoadCode + prjinfo._RoadName;
            worksheet.Cells[2, 4] = prjinfo._RoadGrade;
            worksheet.Cells[2, 6] = GlobalExcel._RoadTypeStr[prjinfo._RoadType];
            worksheet.Cells[2, 8] = prjinfo._Direction > 0 ? "上行" : "下行";
            worksheet.Cells[2, 9] = prjinfo._DataDate;

            worksheet.Cells[rowcnt + 5, 1] = "合计";
            worksheet.Cells[rowcnt + 5, 2] = String.Format("=SUM(B5:B{0})", rowcnt + 4);
            for (int i = 0; i < 7; ++i)
            {
                //  worksheet.Cells[rowcnt + 5, 3+i] = String.Format("=SUMPRODUCT(B5:B{0},{1}5:{1}{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));

                //
                if (_Setting.JSAverageType)
                {
                    worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=AVERAGE({1}5:{1}{0})", rowcnt + 4, (char)('C' + i));
                }
                else
                {
                    worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=SUMPRODUCT({1}5:{1}{0},B5:B{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));

                }
                //=SUMPRODUCT(C5:C279, B5:B279) / SUM(B5:B279)
            }
        }
        private static void WritePDMX2Xls_2024(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
           List<MilePart> roadpart, Disease[] arrdis,
           double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
           double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 20];

            int typeidx = 0;
            bool res = false;
            string input = prjinfo._DataDate;
            worksheet.Cells[2, 2] = "";
            if (input.Length == 8)
            {
                worksheet.Cells[2, 14] = input.Substring(0, 4);
                worksheet.Cells[2, 16] = input.Substring(4, 2);
                worksheet.Cells[2, 18] = input.Substring(6, 2);
            }
            string result = "";
            if (input.Length == 8 && int.TryParse(input, out _))
            {
                int year = int.Parse(input.Substring(0, 4));
                int month = int.Parse(input.Substring(4, 2));
                int day = int.Parse(input.Substring(6, 2));

                DateTime date;
                if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                {
                    result = date.ToString("yyyy年MM月dd日");

                }
                else
                {
                    result = "";
                }
            }
            else
            {
                result = "";
            }

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval = 0;
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
                vallist[rowcnt, 0] = prjinfo._RoadCode;

                if (prjinfo._RoadCode.Length > 5)
                {
                    vallist[rowcnt, 1] = prjinfo._RoadCode.Substring(prjinfo._RoadCode.Length - 6, 6);

                } 
                vallist[rowcnt, 2] = prjinfo._RoadName;
                //病害汇总表
                if (prjinfo._Direction > 0)
                {
                    vallist[rowcnt, 3] = smile;
                    vallist[rowcnt, 4] = emile;

                }
                else
                {
                    vallist[rowcnt, 3] = emile;
                    vallist[rowcnt, 4] = smile;

                }
                vallist[rowcnt,5] = prjinfo._Direction > 0 ? "上行" : "下行";
                vallist[rowcnt, 6] = prjinfo._RoadGrade.Replace("公路","");
                vallist[rowcnt, 7] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[rowcnt, 8] = milelength;
                vallist[rowcnt, 9] = _RoadConfig.DetectWidth;
                vallist[rowcnt, 10] = string.Format("=L{0}*{1}+M{0}*{2}+N{0}*{3}+O{0}*{4}", rowcnt + 4, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);
                 
                vallist[rowcnt, 11] = "100";
                vallist[rowcnt, 12] = string.Format("=ROUND(({1}*P{0}+{2}*Q{0})/({1}+{2}),5)", rowcnt + 4,
                   _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                   _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, 13] = "100";
                vallist[rowcnt, 14] = "100";


                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                vallist[rowcnt, 15] = Math.Round(pcival, 5);


                //IRI
                if (prjinfo._IsDIRIMTD)
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
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, 16] = Math.Round(trqival, 5);
                vallist[rowcnt, 17] = 100;
                vallist[rowcnt, 18] = "自动";
                vallist[rowcnt, 19] = result.Substring(0,4);
                 
                ++rowcnt;
            }
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:T{0}", rowcnt + 3));
            destrange.Value2 = vallist;
            destrange = worksheet.get_Range(String.Format("A4:T{0}", rowcnt + 4));
            GlobalExcel.SetBorderLine(destrange, 63);
            worksheet.Cells[rowcnt + 4, 1] = "合计";
            worksheet.Cells[rowcnt + 4, 9] = String.Format("=SUM(I4:I{0})", rowcnt + 3);
            for (int i = 0; i < 8; ++i)
            {
               
                if (_Setting.JSAverageType)
                {
                    worksheet.Cells[rowcnt + 4, 11 + i] = String.Format("=AVERAGE({1}4:{1}{0})", rowcnt + 3, (char)('K' + i));
                }
                else
                {
                    worksheet.Cells[rowcnt + 4, 11 + i] = String.Format("=SUMPRODUCT({1}4:{1}{0},I4:I{0})/SUM(I4:I{0})", rowcnt + 3, (char)('K' + i));

                }
            }
        }

        private static void WritePDMX2Xls_2024Sum(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
           List<MilePart> roadpart, Disease[] arrdis,
           double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
           double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, int[][] PBVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 25];

            int typeidx = 0;
            bool res = false; 
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                vallist[rowcnt, 0] = prjinfo._RoadCode;
                if (prjinfo._RoadCode.Length>5)
                {
                    vallist[rowcnt, 1] = prjinfo._RoadCode.Substring(prjinfo._RoadCode.Length - 6, 6);

                }
                vallist[rowcnt, 2] = prjinfo._RoadName;
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
                {
                    vallist[rowcnt, 3] = smile;
                    vallist[rowcnt, 4] = emile;

                }
                else
                {
                    vallist[rowcnt, 3] = emile;
                    vallist[rowcnt, 4] = smile;

                }
                vallist[rowcnt, 5] = prjinfo._Direction > 0 ? "上行" : "下行";
                vallist[rowcnt, 6] = roadpart[i].degreestr;
                vallist[rowcnt, 7] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[rowcnt, 8] = milelength;
                vallist[rowcnt, 9] = _RoadConfig.DetectWidth;
                //mqi
                vallist[rowcnt, 10] = string.Format("=O{0}*{1}+L{0}*{2}+P{0}*{3}+Q{0}*{4}", rowcnt + 3, _MQIW[0], _MQIW[1], _MQIW[2], _MQIW[3]);

                vallist[rowcnt, 11] = string.Format("=ROUND(({1}*R{0}+{2}*S{0})/({1}+{2}),5)", rowcnt + 3,
               _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
               _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, 12] = "自动";
            
                vallist[rowcnt, 13] = "2024";

                vallist[rowcnt, 14] = "100";
                vallist[rowcnt, 15] = "100";
                vallist[rowcnt, 16] = "100";


                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                vallist[rowcnt, 17] = Math.Round(pcival, 5);

                //IRI
                if (prjinfo._IsDIRIMTD)
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
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, 18] = Math.Round(trqival, 5);


                vallist[rowcnt, 19] = "/";
                vallist[rowcnt, 20] = "/";
                vallist[rowcnt, 21] = "/";
                vallist[rowcnt, 22] = "/";
                vallist[rowcnt, 23] = "/"; 
                vallist[rowcnt, 24] = "指南";

                ++rowcnt;
            }
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:Y{0}", rowcnt + 2));
            destrange.Value2 = vallist;
            destrange = worksheet.get_Range(String.Format("A3:Y{0}", rowcnt + 2));
            GlobalExcel.SetBorderLine(destrange, 63);  
        }
        private static void WritePrj2CPMSXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[3, 4] = prjinfo._RoadCode;
            if (prjinfo._Direction > 0)
            {
                _Worksheet.Cells[3, 8] = "上行" + prjinfo._RoadNum;
            }
            else
            {
                _Worksheet.Cells[3, 8] = "下行" + prjinfo._RoadNum;
            }
            _Worksheet.Cells[3, 12] = prjinfo._DataDate;
            _Worksheet.Cells[4, 8] = prjinfo._StartMile;
            _Worksheet.Cells[4, 12] = prjinfo._EndMile;
            _Worksheet.Cells[5, 12] = _RoadConfig.DetectWidth;
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
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\沿线设施损坏汇总表.xlsx",
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

        public static void OutputStreetDis_JSZK(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\沿线设施技术状况调查表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_沿线设施技术状况调查表_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["沿线设施"] as MSExcel.Worksheet;
            WriteStreetDisDC2Xls_2024(_Worksheet_hz, prjinfo, prjdir, _RoadPart, _StreetDisRecord.ToArray(), xlslen);

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
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.streetdislist[k].sumval;
                    ttclval = DiseaseTypes.streetdislist[k].unitscore * DiseaseTypes.streetdislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.streetdislist[k].weight * (100 - ttclval);
                }

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

        private static void WriteStreetDisDC2Xls_2024(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.streetdislist.Count;
            object[,] disval = new object[len, temp + 2];
            worksheet_hz.Cells[2, 2] = prjinfo._RoadCode;
            worksheet_hz.Cells[2, 3] = "路线名称:" + prjinfo._RoadName;
            worksheet_hz.Cells[2, 5] = "调查方向:" + (prjinfo._Direction == 1 ? "上行" : "下行");
            worksheet_hz.Cells[2, 6] = "路面宽度:" + _RoadConfig.DetectWidth;

            worksheet_hz.Cells[3, 2] = prjinfo._StartMile.ToString("K0+000") + "~" + prjinfo._EndMile.ToString("K0+000");
            worksheet_hz.Cells[3, 3] = "单元长度:" + xlslen;

            string input = prjinfo._DataDate;
            string result = "";
            if (input.Length == 8 && int.TryParse(input, out _))
            {
                int year = int.Parse(input.Substring(0, 4));
                int month = int.Parse(input.Substring(4, 2));
                int day = int.Parse(input.Substring(6, 2));

                DateTime date;
                if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                {
                    result = date.ToString("yyyy年MM月dd日");

                }
                else
                {
                    result = "";
                }
            }
            else
            {
                result = "";
            }

            worksheet_hz.Cells[3, 5] = "调查时间:" + result;
            worksheet_hz.Cells[3, 6] = "调查人员:" + prjinfo._DataPerson;
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
                if (prjinfo._Direction == -1)
                {
                    disval[i, 0] = emile;
                    disval[i, 1] = smile;
                }
                else
                {
                    disval[i, 0] = smile;
                    disval[i, 1] = emile;
                }
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.streetdislist[k].sumval;

                }

                smile = emile;
                DiseaseTypes.Clear();
            }
            destrange = worksheet_hz.get_Range(string.Format("A5:{1}{0}", len + 4, GlobalExcel.GetCol((char)('A' + temp + 1))));
            destrange.Value2 = disval;
            GlobalExcel.SetBorderLine(destrange, 53);

            //写入合计 
            worksheet_hz.Cells[len + 5, 1] = "合计";
            destrange = worksheet_hz.get_Range(string.Format("A{0}:B{0}", len + 5));
            destrange.Merge();
            // 设置水平和垂直居中
            destrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            destrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            for (int i = 0; i < temp; i++)
            {
                worksheet_hz.Cells[len + 5, i + 3] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('A' + 2 + i)), 5, len + 4);
            }
            destrange = worksheet_hz.get_Range(string.Format("A{0}:{1}{0}", len + 5, GlobalExcel.GetCol((char)('A' + temp + 1))));
            GlobalExcel.SetBorderLine(destrange, 53);
        }


        private static void WriteRoadBedDisDC2Xls_2024(MSExcel.Worksheet worksheet_hz, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            int dlen = arrdis.Length;

            int temp = DiseaseTypes.roadbeddislist.Count;
            object[,] disval = new object[len, temp + 2];
            worksheet_hz.Cells[2, 2] = prjinfo._RoadCode;
            worksheet_hz.Cells[2, 3] = "路线名称:" + prjinfo._RoadName;
            worksheet_hz.Cells[2, 5] = "调查方向:" + (prjinfo._Direction == 1 ? "上行" : "下行");
            worksheet_hz.Cells[2, 6] = "路面宽度:" + _RoadConfig.DetectWidth;

            worksheet_hz.Cells[3, 2] = prjinfo._StartMile.ToString("K0+000") + "~" + prjinfo._EndMile.ToString("K0+000");
            worksheet_hz.Cells[3, 3] = "单元长度:" + xlslen;

            string input = prjinfo._DataDate;
            string result = "";
            if (input.Length == 8 && int.TryParse(input, out _))
            {
                int year = int.Parse(input.Substring(0, 4));
                int month = int.Parse(input.Substring(4, 2));
                int day = int.Parse(input.Substring(6, 2));

                DateTime date;
                if (DateTime.TryParse($"{year}-{month}-{day}", out date))
                {
                    result = date.ToString("yyyy年MM月dd日");

                }
                else
                {
                    result = "";
                }
            }
            else
            {
                result = "";
            }

            worksheet_hz.Cells[3, 5] = "调查时间:" + result;
            worksheet_hz.Cells[3, 6] = "调查人员:" + prjinfo._DataPerson;
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


                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.roadbeddislist[k].sumval;

                }

                smile = emile;
                DiseaseTypes.Clear();
            }
            destrange = worksheet_hz.get_Range(string.Format("A5:{1}{0}", len + 4, GlobalExcel.GetCol((char)('A' + temp + 1))));
            destrange.Value2 = disval;
            GlobalExcel.SetBorderLine(destrange, 53);

            //写入合计 
            worksheet_hz.Cells[len + 5, 1] = "合计";
            destrange = worksheet_hz.get_Range(string.Format("A{0}:B{0}", len + 5));
            destrange.Merge();
            // 设置水平和垂直居中
            destrange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            destrange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            for (int i = 0; i < temp; i++)
            {
                worksheet_hz.Cells[len + 5, i + 3] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('A' + 2 + i)), 5, len + 4);
            }
            destrange = worksheet_hz.get_Range(string.Format("A{0}:{1}{0}", len + 5, GlobalExcel.GetCol((char)('A' + temp + 1))));
            GlobalExcel.SetBorderLine(destrange, 53);

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
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    ttclval = DiseaseTypes.streetdislist[k].unitscore * DiseaseTypes.streetdislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.streetdislist[k].weight * (100 - ttclval);
                }
                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            }
            destrange = worksheet.get_Range(string.Format("I5:I{0}", len + 4));
            destrange.Value2 = disval;
        }
        private static void WriteStreetTCI2Xls_2024(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis,char tagetCol,int rowCnt)
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
                for (int k = 0; k < DiseaseTypes.streetdislist.Count; ++k)
                {
                    ttclval = DiseaseTypes.streetdislist[k].unitscore * DiseaseTypes.streetdislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.streetdislist[k].weight * (100 - ttclval);
                }
                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            } 
            destrange = worksheet.get_Range(string.Format("{1}{2}:{1}{0}", len + rowCnt, tagetCol,rowCnt+1));
            destrange.Value2 = disval;
        }


        /// <summary>
        /// 图片地址  病害信息
        /// </summary>
        private static Dictionary<(string,int),List< MyStreetMile2DisInfo>> curProjectStreetDic = new Dictionary<(string, int), List<MyStreetMile2DisInfo>>();
        public static void OutputStreetAllDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            curProjectStreetDic.Clear();
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\景观病害明细表.xlsx",
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
            string streetPath = prjinfo._PrjPath  + "\\StreetImg";
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
                dis.StartMile = int.Parse( item.Mile.Replace("K","").Replace("+",""));
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
                    string streetPciturePath =   getStreetPcitureFilePath(prjdir.FullName,int.Parse( item.Mile.Replace("K","").Replace("+","")), item.Side);
                    MyStreetMile2DisInfo info = new MyStreetMile2DisInfo();
                    info.Rect = item.SignRect;
                    info.Mile = int.Parse( item.Mile.Replace("K", "").Replace("+", ""));
                    info.DisInfo = item.DisName + " " + item.Info;
                    if (curProjectStreetDic.ContainsKey((streetPciturePath,item.Side)))
                    {
                        curProjectStreetDic[(streetPciturePath,item.Side)].Add(info);
                    }
                    else
                    {
                        curProjectStreetDic[(streetPciturePath,item.Side)] = new List<MyStreetMile2DisInfo> { info };
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
                    info.Mile = int.Parse( item._mile.Replace("K", "").Replace("+", ""));
                    if (curProjectStreetDic.ContainsKey((streetPciturePath,item.Side)))
                    {
                        curProjectStreetDic[(streetPciturePath,item.Side)].Add(info);
                    }
                    else
                    {
                        curProjectStreetDic[(streetPciturePath,item.Side)] = new List<MyStreetMile2DisInfo> { info };
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

        private static string getStreetPcitureFilePath(string prjdir,int mile,int side)
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
        public static void OutputCPMSStreetDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\CPMS_沿线设施损坏.xlsx",
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
                    GlobalExcel.GetCol((char)('E' + (Math.Min(smile, emile) % 1000) * 10 / 1000)),
                    tablerow * tcnt + 8,
                    tablerow * tcnt + 7 + DiseaseTypes.streetdislist.Count));
                destrange.Value2 = disval;

                if (emile % 1000 == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")) || roadpart[i].roadtype != roadpart[i + 1].roadtype)
                {
                    cemile = emile;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 1] = "路线名称：" + prjinfo._RoadName;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 3] = prjinfo._Direction > 0 ? "上行" : "下行";
                    worksheet_dc.Cells[tablerow * tcnt + 3, 7] = prjinfo._DataDate;
                    worksheet_dc.Cells[tablerow * tcnt + 3, 12] = prjinfo._DataPerson;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 7] = csmile;
                    worksheet_dc.Cells[tablerow * tcnt + 4, 12] = cemile;
                    worksheet_dc.Cells[tablerow * tcnt + 5, 7] = Math.Abs(csmile - cemile);
                    worksheet_dc.Cells[tablerow * tcnt + 5, 12] = _RoadConfig.DetectWidth;
                    if (cemile != prjinfo._EndMile)
                    {
                        srcrange = worksheet_dc.get_Range(String.Format("A{0}:T{1}", tablerow * tcnt + 1, tablerow * (tcnt + 1) - 1));
                        ++tcnt;
                        destrange = worksheet_dc.get_Range(String.Format("A{0}", tablerow * tcnt + 1));
                        srcrange.Copy(destrange);
                        destrange = worksheet_dc.get_Range(String.Format("E{0}:N{1}", tablerow * tcnt + 8, tablerow * tcnt + 7 + DiseaseTypes.streetdislist.Count));
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
                worksheet_dc.Cells[tablerow * tcnt + 3, 3] = prjinfo._Direction > 0 ? "上行" : "下行";
                worksheet_dc.Cells[tablerow * tcnt + 3, 7] = prjinfo._DataDate;
                worksheet_dc.Cells[tablerow * tcnt + 3, 12] = prjinfo._DataPerson;
                worksheet_dc.Cells[tablerow * tcnt + 4, 7] = csmile;
                worksheet_dc.Cells[tablerow * tcnt + 4, 12] = prjinfo._EndMile;
                worksheet_dc.Cells[tablerow * tcnt + 5, 7] = Math.Abs(csmile - prjinfo._EndMile);
                worksheet_dc.Cells[tablerow * tcnt + 5, 12] = _RoadConfig.DetectWidth;
            }
        }
        #endregion

        #region 路基损坏报表
        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\路基损坏汇总表.xlsx",
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

        public static void OutputRoadBedDis_JSZK(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\路基技术状况调查表.xlsx",
                System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_路基技术状况调查表_{2}米.xlsx", path, prjdir.Name, xlslen);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["路基损坏汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["路基"] as MSExcel.Worksheet;
            WriteRoadBedDisDC2Xls_2024(_Worksheet_hz, prjinfo, prjdir, _RoadPart, _StreetDisRecord_RoadBed.ToArray(), xlslen);

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
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    disval[i, k + 2] = DiseaseTypes.roadbeddislist[k].sumval;
                    ttclval = DiseaseTypes.roadbeddislist[k].unitscore * DiseaseTypes.roadbeddislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.roadbeddislist[k].weight * (100 - ttclval);
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
                disval[i, 0] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('C' + i)), len + 4);
            }
            destrange = worksheet_hz.get_Range(string.Format("{0}3:{0}{1}", GlobalExcel.GetCol((char)('A' + temp + 6)), temp + 2));
            destrange.Value2 = disval;

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet_hz, 5, 1, temp + 4, true);
                GlobalExcel.Reflection(worksheet_hz, 5, 1, 2, false);
            }
        }



        private static void WriteRoadBedSCI2Xls_2024(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, StreetDisRecord[] arrdis, char tagetCol, int rowCnt)
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
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    ttclval = DiseaseTypes.roadbeddislist[k].unitscore * DiseaseTypes.roadbeddislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.roadbeddislist[k].weight * (100 - ttclval);
                }

                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            }
           // destrange = worksheet.get_Range(string.Format("L4:L{0}", len + 3));
            destrange = worksheet.get_Range(string.Format("{1}{2}:{1}{0}", len + rowCnt, tagetCol, rowCnt + 1));
            destrange.Value2 = disval;
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
                for (int k = 0; k < DiseaseTypes.roadbeddislist.Count; ++k)
                {
                    ttclval = DiseaseTypes.roadbeddislist[k].unitscore * DiseaseTypes.roadbeddislist[k].sumval;
                    ttclval = ttclval * 1000 / unitlen;
                    ttclval = ttclval > 100 ? 100 : ttclval;
                    tclval += DiseaseTypes.roadbeddislist[k].weight * (100 - ttclval);
                }

                disval[i, 0] = tclval;
                DiseaseTypes.Clear();
            }
            destrange = worksheet.get_Range(string.Format("G5:G{0}", len + 4));
            destrange.Value2 = disval;
        }

        internal static void outMulitExcel_NC(MSExcel.Application excelApp, string outpath, List<string> ExcelJSMXList)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\合并报表模板\南昌农村路\路面技术状况单元评定.xlsx",
                   System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\路面技术状况单元评定.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;

            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitPQI2Xls_NC0(excelApp, destsheet, ExcelJSMXList);
           
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        internal static void outMulitExcel_NC1(MSExcel.Application excelApp, string outpath,List<string> ExcelBHTJList)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\合并报表模板\南昌农村路\水泥混凝土路面单元评定.xlsx",
                   System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\水泥混凝土路面单元评定.xlsx", outpath);


            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;

            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitPQI2Xls_NC1(excelApp, destsheet, ExcelBHTJList);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        internal static void outMulitExcel_NC2(MSExcel.Application excelApp, string outpath, List<string> ExcelBHTJList)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\合并报表模板\南昌农村路\沥青混凝土路面病单元评定.xlsx",
                   System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\沥青混凝土路面病单元评定.xlsx", outpath);


            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;

            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitPQI2Xls_NC2(excelApp, destsheet, ExcelBHTJList);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\CPMS_路基损坏.xlsx",
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
                    tablerow * tcnt + 8,
                    tablerow * tcnt + 7+ DiseaseTypes.roadbeddislist.Count));
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
                        destrange = worksheet_dc.get_Range(String.Format("F{0}:O{1}", tablerow * tcnt + 8, tablerow * tcnt + 7+ DiseaseTypes.roadbeddislist.Count));
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

        private static void WriteZJGTDisHZTJ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz, MSExcel.Worksheet worksheet_sshz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen, string[] MarkVal)
        {
            MSExcel.Range srcrange, destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志
            bool Hasssflag = false;//有砂石路段标志

            int sn_tablerow = _Setting.cmop_rows;
            int lq_tablerow = _Setting.cmop_rows;
            int ss_tablerow = _Setting.cmop_rows;

            int tcnt_sn = 0;
            int tcnt_lq = 0;
            int tcnt_ss = 0;

            int sn_csmile = 0, sn_cemile = 0;
            int lq_csmile = 0, lq_cemile = 0;
            int ss_csmile = 0, ss_cemile = 0;
            bool sn_flag = false, lq_flag = false, ss_flag = false;

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
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;//1
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
                        GlobalExcel.GetCol((char)('E' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
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
                        GlobalExcel.GetCol((char)('E' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
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
                else if (roadpart[i].roadtype == 2)//砂石
                {
                    Hasssflag = true;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
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
                    destrange = worksheet_sshz.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('E' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        ss_tablerow * tcnt_ss + 7,
                        ss_tablerow * tcnt_ss + 6 + disnum));
                    destrange.Value2 = disval;

                    ss_cemile = emile;
                    if (!ss_flag)
                    {
                        ss_flag = true;
                        ss_csmile = smile;
                    }
                }
                //cwb
                //道路材质切换
                if (roadpart[i].roadtype != roadpart[i + 1].roadtype)
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4,8] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 12] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("D{0}:M{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 7 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 12] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("D{0}:M{1}", lq_tablerow * tcnt_lq +7, lq_tablerow * tcnt_lq + 7 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }

                    if (ss_csmile != ss_cemile)
                    {
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 8] = ss_csmile;
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 12] = ss_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_sshz.get_Range(String.Format("A{0}:U{1}", ss_tablerow * tcnt_ss + 1, ss_tablerow * (++tcnt_ss) + 1));
                            destrange = worksheet_sshz.get_Range(String.Format("A{0}", ss_tablerow * tcnt_ss + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_sshz.get_Range(String.Format("D{0}:M{1}", ss_tablerow * tcnt_ss + 7, ss_tablerow * tcnt_ss + 7 + disnum));
                            destrange.ClearContents();
                        }
                        ss_flag = false;
                        ss_csmile = ss_cemile;
                    }
                }
                if (emile % xlslen == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")))
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 12] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("D{0}:M{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 7 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 12] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("D{0}:M{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 7 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }
                    if (ss_csmile != ss_cemile)
                    {
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 8] = ss_csmile;
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 12] = ss_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count - 1;
                            srcrange = worksheet_sshz.get_Range(String.Format("A{0}:U{1}", ss_tablerow * tcnt_ss + 1, ss_tablerow * (++tcnt_ss) + 1));
                            destrange = worksheet_sshz.get_Range(String.Format("A{0}", ss_tablerow * tcnt_ss + 1));
                            srcrange.Copy(destrange);

                            destrange = worksheet_sshz.get_Range(String.Format("D{0}:M{1}", ss_tablerow * tcnt_ss + 7, ss_tablerow * tcnt_ss + 7 + disnum));
                            destrange.ClearContents();
                        }
                        ss_flag = false;
                        ss_csmile = ss_cemile;
                    }
                }
            }
            if (roadpart[len].mile % xlslen != 0 || (MarkVal[len] != null && MarkVal[len].Contains("路面单元")))
            {
                if (sn_csmile != sn_cemile)
                {
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 8] = sn_csmile;
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 12] = roadpart[len].mile;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 8] = lq_csmile;
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 12] = roadpart[len].mile;
                }
                if (ss_csmile != ss_cemile)
                {
                    worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 8] = ss_csmile;
                    worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 12] = roadpart[len].mile;
                }
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (tcnt_sn + 1) + 1));
                destrange = worksheet_snhz.get_Range(String.Format("E{0}:N{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 7 + disnum));
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
                destrange = worksheet_lqhz.get_Range(String.Format("E{0}:N{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq +7 + disnum));
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

            if (Hasssflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count - 1;
                srcrange = worksheet_sshz.get_Range(String.Format("A{0}:U{1}", ss_tablerow * tcnt_ss + 1, ss_tablerow * (tcnt_ss + 1) + 1));
                destrange = worksheet_sshz.get_Range(String.Format("E{0}:N{1}", ss_tablerow * tcnt_ss + 7, ss_tablerow * tcnt_ss + 7 + disnum));
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
            if (!Hasssflag)
            {
                worksheet_sshz.Delete();
            }


            RoadDiseaseTypes.Clear();
        }

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

            _Worksheet.Cells[strrow, strcol] = string.Format("=CONCATENATE(\"路面{6}评价等级“优”率占路段总数\",ROUND({1}{0},4)*100,\"%，“良”率占路段总数\",ROUND({2}{0},4)*100,\"%，“中”率占路段总数\",ROUND({3}{0},4)*100,\"%，“次”率占路段总数\",ROUND({4}{0},4)*100,\"%，“差”率占路段总数\",ROUND({5}{0},4)*100,\"%。\")",
                statisticsrow + 1,
                GlobalExcel.GetCol((char)('A' + statisticscol - 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 2)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 3)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 4)),
                indextype);
        }

        public static void OutputAllRoadStatistics_XN(MSExcel.Application excelApp, string outpath,
        List<string> ExcelJSMXList, List<string> ExcelDRList, List<string> ExcelIRIList)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\咸宁定制\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计-XN.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;

           
            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitPQI2Xls_XN(excelApp, destsheet, ExcelJSMXList, ExcelDRList, ExcelIRIList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        #region 湖南合并出表

        /// <summary>
        /// 湖南定制 合并报表
        /// _ExcelGpss, _DiseaseLQ,_DiseaseSN,_DiseaseAll
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="outpath"></param>
        /// <param name="ExcelJSMXList"></param>
        /// <param name="ExcelDRList"></param>
        /// <param name="ExcelIRIList"></param>
        public static void OutputAllRoadStatistics_hunan(MSExcel.Application excelApp, string outpath,
      List<string> ExcelGpss, List<string> DiseaseLQ, List<string> DiseaseSN,List<string> DiseaseAll)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\湖南定制\表格样式.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        

            MSExcel.Worksheet gpsSheet  = _Workbook.Sheets["空间定位表格样式"] as MSExcel.Worksheet;
            MSExcel.Worksheet disSNsheet = _Workbook.Sheets["水泥病害统计表格样式"] as MSExcel.Worksheet;
            MSExcel.Worksheet disLQsheet = _Workbook.Sheets["沥青路面病害统计表格样式"] as MSExcel.Worksheet;
            MSExcel.Worksheet disAllsheet = _Workbook.Sheets["病害流水表表格样式"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitXls_hunan_0(excelApp, gpsSheet,ExcelGpss,5);
            WriteAllUnitXls_hunan_0(excelApp, disSNsheet, DiseaseSN,8);
            WriteAllUnitXls_hunan_0(excelApp, disLQsheet, DiseaseLQ,7);
            WriteAllUnitXls_hunan_0(excelApp, disAllsheet, DiseaseAll, 8);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
      
        private static void WriteAllUnitXls_hunan_0(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> Excels,int colCount)
        {
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I' };

            int rowidx =3;
            int userow = 0;
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int roadlinenum = Excels.Count;
            for (int ri = 0; ri < roadlinenum; ++ri)
            {
                MSExcel.Workbook tbook = excelApp.Workbooks.Open(Excels[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
                
                MSExcel.Worksheet tsheet = tbook.Sheets["Sheet1"] as MSExcel.Worksheet;
              

                userow = GlobalExcel.judegeusedrow(tsheet, 1, 3);
                object[,] obj_info = null;
                trange = tsheet.get_Range(string.Format("A3:{1}{0}", userow, chars[colCount-1]));
                obj_info = (object[,])trange.Value2;
                userow = userow - 2;
                outobj = new object[userow, colCount];
                


                for (int i = 0; i < userow; ++i)
                {

                    for (int col = 0; col < colCount; col++)
                    {
                        outobj[i, col] = obj_info[i + 1, col + 1];
                    }
                  

                }
                trange = destsheet.get_Range(string.Format("A{0}:{2}{1}", rowidx, rowidx + userow - 1, chars[colCount - 1]));
                trange.Value2 = outobj;
                rowidx += userow;

                tbook.Close();
           
            }
            trange = destsheet.get_Range(string.Format("A1:{1}{0}", rowidx - 1, chars[colCount - 1]));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        #endregion
        private static void WriteAllUnitPQI2Xls_XN(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> ExcelJSMXList,
        List<string> ExcelDRList, List<string> ExcelIRIList)
        {
            if (ExcelDRList.Count != ExcelIRIList.Count || ExcelJSMXList.Count != ExcelDRList.Count)
            {
                MessageBox.Show("技术状况明细表，IRI，PCI的报表数量不一致，请检查！");
            }
            bool isLow = true;
            DialogResult result =  MessageBox.Show("是否按照低等级农村路标准合并?", "提示对话框", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                isLow = true;
            }
            else
            {
                isLow = false;
            }
            int rowidx =4;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int roadlinenum = ExcelDRList.Count;
            for (int ri = 0; ri < roadlinenum; ++ri)
            {
                MSExcel.Workbook tbook_JSMX = excelApp.Workbooks.Open(ExcelJSMXList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
                object[,] tobj_jsmx = null; 
                GetPrjInfo(tbook_iri, ref infoobj);
                if (isLow)
                {
                    MSExcel.Worksheet tsheet_dr = tbook_dr.Sheets["Sheet1"] as MSExcel.Worksheet;
                    MSExcel.Worksheet tsheet_iri = tbook_iri.Sheets["Sheet1"] as MSExcel.Worksheet;
                    MSExcel.Worksheet tsheet_jsmx = tbook_JSMX.Sheets["Sheet1"] as MSExcel.Worksheet;
                    int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                    userow = GlobalExcel.judegeusedrow(tsheet_dr, 1, 3);

                    trange = tsheet_dr.get_Range(string.Format("A3:I{0}", userow));
                    tobj_dr = (object[,])trange.Value2;
                    userow = userow - 2;
                    outobj = new object[userow, 21];

                    trange = tsheet_iri.get_Range(string.Format("A4:I{0}", userow + 3));
                    tobj_iri = (object[,])trange.Value2;

                    trange = tsheet_jsmx.get_Range(string.Format("A5:J{0}", userow + 4));
                    tobj_jsmx = (object[,])trange.Value2;


                    trange = destsheet.get_Range("A1:A1");

                    trange.Value2 = $"附表一：{infoobj[0, 1]}市农村公路技术状况指数评定明细表";


                    for (int i = 0; i < userow; ++i)
                    {
                        for (int ii = 0; ii < 5; ++ii)
                        {
                            //0-4
                            outobj[i, ii] = infoobj[0, ii];
                        }
                        double Smile = double.Parse(tobj_dr[i + 1, 1].ToString());
                        double Emile = double.Parse(tobj_dr[i + 1, 2].ToString());
                        outobj[i, 5] = Smile / 1000;
                        outobj[i, 6] = Emile / 1000;
                        outobj[i, 7] = tobj_dr[i + 1, 7];
                        outobj[i, 8] = infoobj[0, 6];
                        outobj[i, 9] = infoobj[0, 5];
                        outobj[i, 10] = Math.Abs(Smile - Emile);
                        outobj[i, 11] = tobj_jsmx[i + 1, 4];
                        outobj[i, 12] = tobj_dr[i + 1, 4];
                        outobj[i, 13] = tobj_dr[i + 1, 5];
                        if (string.IsNullOrEmpty(tobj_iri[i + 1, 6].ToString()))
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 4];//iri
                        }
                        else
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 6];//iri
                        }

                        outobj[i, 15] = tobj_jsmx[i + 1, 6];   //rqi
                        outobj[i, 16] = " /";
                        outobj[i, 17] = " /";
                        outobj[i, 18] = tobj_jsmx[i + 1, 9];
                        outobj[i, 19] = tobj_jsmx[i + 1, 7];

                    }

                    trange = destsheet.get_Range(string.Format("A{0}:U{1}", rowidx, rowidx + userow - 1));
                    trange.Value2 = outobj;
                    rowidx += userow;

                    tbook_dr.Close();
                    tbook_iri.Close();
                    tbook_JSMX.Close();
                }
                else
                {
                    MSExcel.Worksheet tsheet_dr = tbook_dr.Sheets["Sheet1"] as MSExcel.Worksheet;
                    MSExcel.Worksheet tsheet_iri = tbook_iri.Sheets["Sheet1"] as MSExcel.Worksheet;
                    MSExcel.Worksheet tsheet_jsmx = tbook_JSMX.Sheets["Sheet1"] as MSExcel.Worksheet;
                    int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                    userow = GlobalExcel.judegeusedrow(tsheet_dr, 1, 3);

                    trange = tsheet_dr.get_Range(string.Format("A3:I{0}", userow));
                    tobj_dr = (object[,])trange.Value2;
                    userow = userow - 2;
                    outobj = new object[userow, 21];

                    trange = tsheet_iri.get_Range(string.Format("A4:I{0}", userow + 3));
                    tobj_iri = (object[,])trange.Value2;

                    trange = tsheet_jsmx.get_Range(string.Format("A5:O{0}", userow + 4));
                    tobj_jsmx = (object[,])trange.Value2;


                    trange = destsheet.get_Range("A1:A1");

                    trange.Value2 = $"附表一：{infoobj[0, 1]}市农村公路技术状况指数评定明细表";


                    for (int i = 0; i < userow; ++i)
                    {
                        for (int ii = 0; ii < 5; ++ii)
                        {
                            //0-4
                            outobj[i, ii] = infoobj[0, ii];
                        }
                        double Smile = double.Parse(tobj_dr[i + 1, 1].ToString());
                        double Emile = double.Parse(tobj_dr[i + 1, 2].ToString());
                        outobj[i, 5] = Smile / 1000;
                        outobj[i, 6] = Emile / 1000;
                        outobj[i, 7] = tobj_dr[i + 1, 7];
                        outobj[i, 8] = infoobj[0, 6];
                        outobj[i, 9] = infoobj[0, 5];
                        outobj[i, 10] = Math.Abs(Smile - Emile);
                        outobj[i, 11] = tobj_jsmx[i + 1, 5]; //PQI
                        outobj[i, 12] = tobj_dr[i + 1, 4];  //dr
                        outobj[i, 13] = tobj_dr[i + 1, 5]; //pci
                        if (string.IsNullOrEmpty(tobj_iri[i + 1, 6].ToString()))
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 4];//iri
                        }
                        else
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 6];//iri
                        }

                        outobj[i, 15] = tobj_jsmx[i + 1, 7];   //rqi
                        outobj[i, 16] = " /";
                        outobj[i, 17] = tobj_jsmx[i + 1, 8];
                        outobj[i, 18] = tobj_jsmx[i + 1, 14];
                        outobj[i, 19] = tobj_jsmx[i + 1, 4];

                    }

                    trange = destsheet.get_Range(string.Format("A{0}:U{1}", rowidx, rowidx + userow - 1));
                    trange.Value2 = outobj;
                    rowidx += userow;

                    tbook_dr.Close();
                    tbook_iri.Close();
                    tbook_JSMX.Close();
                }
                
            }
            trange = destsheet.get_Range(string.Format("A1:U{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        public static void OutputAllRoadStatistics(MSExcel.Application excelApp, string outpath, List<string> xlslist)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\农村公路技术状况结果统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;
            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteAllRoadPQI2Xls(excelApp, destsheet, xlslist);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteAllRoadPQI2Xls(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> xlslist)
        {
            double tmp = 0;
            int rowidx = 2;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] idxobj = new object[1, 6];

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
                trange = destsheet.get_Range(string.Format("A{0}:L{0}", rowidx));
                trange.Value2 = infoobj;

                try
                {
                    tsheet = tbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                }
                catch (Exception ) { }

                if (tsheet != null)
                {
                    userow = GlobalExcel.judegeusedrow(tsheet, 1, 2);
                    trange = tsheet.get_Range(string.Format("A2:R{0}", userow));
                    tobj = (object[,])trange.Value2;

                    string[] idxname = { "PQI", "PCI", "RQI", "RDI", "PBI", "PWI" };
                    int[] idxcol = new int[idxname.Length];
                    double[] sumidxval = new double[idxname.Length];
                    for (int i = 0; i < idxname.Length; ++i)
                    {
                        idxcol[i] = 0;
                        sumidxval[i] = 0;
                    }

                    double sumlen = 0;
                    for (int i = 0; i < 18; ++i)
                    {
                        if (tobj[1, i + 1] == null)
                            continue;

                        for (int j = 0; j < idxname.Length; ++j)
                        {
                            if (tobj[1, i + 1].ToString() == idxname[j])
                            {
                                idxcol[j] = i + 1;
                                break;
                            }
                        }
                    }

                    for (int i = 2; i < userow; ++i)
                    {
                        int dmival = Math.Abs(Convert.ToInt32(tobj[i, 1]) - Convert.ToInt32(tobj[i, 2]));
                        for (int j = 0; j < idxname.Length; ++j)
                        {
                            if (idxcol[j] != 0)
                                sumidxval[j] += Convert.ToDouble(tobj[i, idxcol[j]]) * dmival;
                        }

                        sumlen += dmival;
                    }

                    for (int i = 0; i < idxname.Length; ++i)
                    {
                        idxobj[0, i] = sumidxval[i] / sumlen;
                    }

                    trange = destsheet.get_Range(string.Format("M{0}:R{0}", rowidx));
                    trange.Value2 = idxobj;
                    ++rowidx;
                }
                tbook.Save();
                tbook.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:R{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        #region 孝感
        private static void WriteAllUnitPQI2Xls_XG(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> ExcelJSMXList,
  List<string> ExcelDRList, List<string> ExcelIRIList)
        {
            if (ExcelDRList.Count != ExcelIRIList.Count && ExcelJSMXList.Count != ExcelDRList.Count)
            {
                MessageBox.Show("技术状况明细表，IRI，PCI等报表数量不一致，请检查！");
            }


            int rowidx = 4;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int roadlinenum = ExcelDRList.Count;
            for (int ri = 0; ri < roadlinenum; ++ri)
            {
                MSExcel.Workbook tbook_JSMX = excelApp.Workbooks.Open(ExcelJSMXList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
                object[,] tobj_jsmx = null;

                GetPrjInfo(tbook_iri, ref infoobj);

                MSExcel.Worksheet tsheet_dr = tbook_dr.Sheets["Sheet1"] as MSExcel.Worksheet;
                MSExcel.Worksheet tsheet_iri = tbook_iri.Sheets["Sheet1"] as MSExcel.Worksheet;
                MSExcel.Worksheet tsheet_jsmx = tbook_JSMX.Sheets["Sheet1"] as MSExcel.Worksheet;
                int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                userow = GlobalExcel.judegeusedrow(tsheet_dr, 1, 3);

                trange = tsheet_dr.get_Range(string.Format("A3:I{0}", userow));
                tobj_dr = (object[,])trange.Value2;
                userow = userow - 2;
                outobj = new object[userow, 21];

                trange = tsheet_iri.get_Range(string.Format("A4:I{0}", userow + 3));
                tobj_iri = (object[,])trange.Value2;

                trange = tsheet_jsmx.get_Range(string.Format("A5:J{0}", userow + 4));
                tobj_jsmx = (object[,])trange.Value2;


                trange = destsheet.get_Range("A1:A1");

                trange.Value2 = $"附表一：{infoobj[0, 1]}市农村公路技术状况指数评定明细表";


                try
                {
                    for (int i = 0; i < userow; ++i)
                    {
                        for (int ii = 0; ii < 5; ++ii)
                        {
                            //0-4
                            outobj[i, ii] = infoobj[0, ii];
                        }
                        double Smile = double.Parse(tobj_dr[i + 1, 1].ToString());
                        double Emile = double.Parse(tobj_dr[i + 1, 2].ToString());
                        outobj[i, 5] = Smile / 1000;
                        outobj[i, 6] = Emile / 1000;
                        outobj[i, 7] = tobj_dr[i + 1, 7];
                        outobj[i, 8] = infoobj[0, 6];
                        outobj[i, 9] = infoobj[0, 5];
                        outobj[i, 10] = Math.Abs(Smile - Emile);
                        outobj[i, 11] = tobj_jsmx[i + 1, 4];
                        outobj[i, 12] = tobj_dr[i + 1, 4];
                        outobj[i, 13] = tobj_dr[i + 1, 5];
                        if (tobj_iri != null && string.IsNullOrEmpty(tobj_iri[i + 1, 6].ToString()))
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 4];//iri
                        }
                        else
                        {
                            outobj[i, 14] = tobj_iri[i + 1, 6];//iri
                        }

                        outobj[i, 15] = tobj_jsmx[i + 1, 6];   //rqi
                        outobj[i, 16] = " /";
                        outobj[i, 17] = " /";
                        outobj[i, 18] = tobj_jsmx[i + 1, 9];
                        outobj[i, 19] = tobj_jsmx[i + 1, 7];

                    }
                }
                catch (Exception)
                {
                    MessageBox.Show($"{ExcelJSMXList[ri]}表格所在项目下\n技术状况明细表，IRI，PCI某个报表数据条数不一致请检查! ");
                }
                trange = destsheet.get_Range(string.Format("A{0}:U{1}", rowidx, rowidx + userow - 1));
                trange.Value2 = outobj;
                rowidx += userow;

                tbook_dr.Close();
                tbook_iri.Close();
                tbook_JSMX.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:U{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }

        /// <summary>
        /// 孝感合并
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="outpath"></param>
        /// <param name="ExcelJSMXList"></param>
        /// <param name="ExcelDRList"></param>
        /// <param name="ExcelIRIList"></param>
        public static void OutputAllRoadStatistics_XG(MSExcel.Application excelApp, string outpath,
    List<string> ExcelJSMXList, List<string> ExcelDRList, List<string> ExcelIRIList, List<string> ExcelLJSHAndLJList)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\孝感定制\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;
            MSExcel.Worksheet destLJSHsheet = null;
            MSExcel.Worksheet destYXSSsheet = null;

            destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            //DR ->PCI
            WriteAllUnitPQI2Xls_XG(excelApp, destsheet, ExcelJSMXList, ExcelDRList, ExcelIRIList);
            if (ExcelLJSHAndLJList.Count > 0)
            {
                destLJSHsheet = _Workbook.Sheets["路基"] as MSExcel.Worksheet;
                destYXSSsheet = _Workbook.Sheets["沿线设施"] as MSExcel.Worksheet;

                WriteAllUnitStreet_XG(excelApp, destLJSHsheet, destYXSSsheet, ExcelLJSHAndLJList);
            }
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteAllUnitStreet_XG(MSExcel.Application excelApp, MSExcel.Worksheet LJSHsheet, MSExcel.Worksheet YXSSsheet, List<string> ExcelYXSSAndLJList)
        {

            int rowidx = 3;
            int rowidx2 = 3;
            int useLjrow = 0;

            int useYXSSrow = 0;

            object[,] outLJobj = null;
            object[,] outYXSSobj = null;

            MSExcel.Range trange = null;
            int roadlinenum = ExcelYXSSAndLJList.Count;

            for (int ri = 0; ri < roadlinenum; ++ri)
            {
                MSExcel.Workbook tbook_LJSH = excelApp.Workbooks.Open(ExcelYXSSAndLJList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                object[,] tobj_LJ = null;
                object[,] tobj_YXSS = null;

                MSExcel.Worksheet tsheet_LJ = tbook_LJSH.Sheets["sheet1"] as MSExcel.Worksheet;
                useLjrow = GlobalExcel.judegeusedrow(tsheet_LJ, 1, 3);
                DataToExcel_XG(LJSHsheet, ref rowidx, ref useLjrow, out outLJobj, out trange, out tobj_LJ, tsheet_LJ);

                MSExcel.Worksheet tsheet_YXSS = tbook_LJSH.Sheets["sheet2"] as MSExcel.Worksheet;
                useYXSSrow = GlobalExcel.judegeusedrow(tsheet_YXSS, 1, 3);
                DataToExcel_XG(YXSSsheet, ref rowidx2, ref useYXSSrow, out outYXSSobj, out trange, out tobj_YXSS, tsheet_YXSS);
                tbook_LJSH.Close();
            }
            trange = LJSHsheet.get_Range(string.Format("A1:F{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);


            trange = YXSSsheet.get_Range(string.Format("A1:F{0}", rowidx2 - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        private static void DataToExcel_XG(MSExcel.Worksheet LJSHsheet, ref int rowidx, ref int useLjrow, out object[,] outLJobj, out MSExcel.Range trange, out object[,] tobj_LJ, MSExcel.Worksheet tsheet_LJ)
        {
            trange = tsheet_LJ.get_Range(string.Format("A3:F{0}", useLjrow));
            tobj_LJ = (object[,])trange.Value2;
            useLjrow = useLjrow - 2;
            outLJobj = new object[useLjrow, 6];

            for (int i = 0; i < useLjrow; ++i)
            {

                for (int j = 0; j < 6; j++)
                {
                    outLJobj[i, j] = tobj_LJ[i + 1, j + 1];
                }

            }
            trange = LJSHsheet.get_Range(string.Format("A{0}:F{1}", rowidx, rowidx + useLjrow - 1));
            trange.Value2 = outLJobj;
            rowidx += useLjrow;
        }
        #endregion

        #region 南昌
        private static void WriteAllUnitPQI2Xls_NC0(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> ExcelJSMXList)
        {
            int rowidx = 3;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int rowCount = 0;
            for (int ri = 0; ri < ExcelJSMXList.Count; ++ri)
            {
                MSExcel.Workbook tbook_JSMX = excelApp.Workbooks.Open(ExcelJSMXList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);
               
                object[,] tobj_jsmx = null;

                GetPrjInfo(tbook_JSMX, ref infoobj);
                
                MSExcel.Worksheet tsheet_jsmx = tbook_JSMX.Sheets["Sheet1"] as MSExcel.Worksheet;
                int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];

                userow = GlobalExcel.judegeusedrow(tsheet_jsmx, 1,6);
                
               
          
                trange = tsheet_jsmx.get_Range(string.Format("A5:J{0}", userow + 4));
                tobj_jsmx = (object[,])trange.Value2;
                int nodeCont = userow;
                for (int i = 0; i < nodeCont; ++i)
                {
                    try
                    {
                        double dTemp = double.Parse(tobj_jsmx[i + 1,1].ToString());
                    }
                    catch (Exception)
                    {
                        userow--;
                        continue;

                    }
                }
                outobj = new object[userow, 16];
                trange = tsheet_jsmx.get_Range(string.Format("A5:J{0}", userow + 4));
                tobj_jsmx = (object[,])trange.Value2;
                try
                {
                    for (int i = 0; i < userow; ++i)
                    {
                        
                        rowCount++;
                        outobj[i, 0] = rowCount.ToString();
                        outobj[i, 1] = infoobj[0, 1];
                        outobj[i, 2] = infoobj[0, 4];
                        outobj[i, 3] = infoobj[0, 3];
                        string line = infoobj[0, 6].ToString();

                        double Smile = double.Parse(tobj_jsmx[i + 1, 1].ToString());
                        double length = double.Parse(tobj_jsmx[i + 1, 2].ToString());
                        double endMile;
                        endMile = Smile + length;
                        //if (line.Contains("上行"))
                        //{
                           
                        //}
                        //else
                        //{
                        //    endMile = Smile - length;
                        //}
                        outobj[i, 4] = line;  
                        
                        outobj[i, 5] =Smile;
                        outobj[i, 6] = endMile;
                        outobj[i, 7] = length;
                       
                        outobj[i, 8] = infoobj[0, 5];
                        outobj[i, 9] = tobj_jsmx[i + 1, 10].ToString()+"路面";
                        outobj[i, 10] = tobj_jsmx[i + 1, 5];
                        //= IF(E4 >= 90, "优", IF(E4 >= 80, "良", IF(E4 >= 70, "中", IF(E4 >= 60, "次", "差"))))
                        if (double.Parse(tobj_jsmx[i + 1, 5].ToString())>=90)
                        {
                            outobj[i, 11] ="优"; 
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 5].ToString()) >= 80)
                        {
                            outobj[i, 11] = "良"; 
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 5].ToString()) >= 70)
                        {
                            outobj[i, 11] = "中"; 
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 5].ToString()) >= 60)
                        {
                            outobj[i, 11] = "次"; 
                        }
                        else
                        {
                            outobj[i, 11] = "差";  
                        }
                        outobj[i, 12] = tobj_jsmx[i + 1, 6];
                        if (double.Parse(tobj_jsmx[i + 1,6].ToString()) >= 90)
                        {
                            outobj[i, 13] = "优";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1,6].ToString()) >= 80)
                        {
                            outobj[i, 13] = "良";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1,6].ToString()) >= 70)
                        {
                            outobj[i, 13] = "中";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1,6].ToString()) >= 60)
                        {
                            outobj[i, 13] = "次";
                        }
                        else
                        {
                            outobj[i, 13] = "差";
                        }


                        outobj[i, 14] = tobj_jsmx[i + 1, 4];
                        if (double.Parse(tobj_jsmx[i + 1, 4].ToString()) >= 90)
                        {
                            outobj[i, 15] = "优";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 4].ToString()) >= 80)
                        {
                            outobj[i, 15] = "良";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 4].ToString()) >= 70)
                        {
                            outobj[i, 15] = "中";
                        }
                        else if (double.Parse(tobj_jsmx[i + 1, 4].ToString()) >= 60)
                        {
                            outobj[i, 15] = "次";
                        }
                        else
                        {
                            outobj[i, 15] = "差";
                        }
                    }
                }
                catch (Exception)
                {
                    
                }
                trange = destsheet.get_Range(string.Format("A{0}:P{1}", rowidx, rowidx + userow - 1));
                trange.Value2 = outobj;
                rowidx += userow;

                
                tbook_JSMX.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:P{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        private static void WriteAllUnitPQI2Xls_NC1(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> ExcelBHTJList)
        {
            int rowidx = 3;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int rowCount = 0;
            for (int ri = 0; ri < ExcelBHTJList.Count; ++ri)
            {
                MSExcel.Workbook tbook_JSMX = excelApp.Workbooks.Open(ExcelBHTJList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                object[,] tobj_jsmx = null;

                GetPrjInfo(tbook_JSMX, ref infoobj);
                MSExcel.Worksheet tsheet_jsmx;
                try
                {
                    tsheet_jsmx = tbook_JSMX.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                   
                }
                catch (Exception)
                {
                    userow = 0;
                    tbook_JSMX.Close();
                    continue;
                }

                int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];
                userow = GlobalExcel.judegeusedrow(tsheet_jsmx, 1, 6);
               
                trange = tsheet_jsmx.get_Range(string.Format("A5:m{0}", userow + 4));
                tobj_jsmx = (object[,])trange.Value2;
                
                
                outobj = new object[userow, 19];
                object[,] result = new object[userow, 19];
                int yesRowsCount = 0; //有效数据行数
                try
                {
                    for (int i = 0; i < userow; ++i)
                    {
                        double Smile;
                        double eMile;
                        try
                        {
                            Smile = double.Parse(tobj_jsmx[i + 1, 1].ToString());
                            eMile = double.Parse(tobj_jsmx[i + 1, 2].ToString());
                        }
                        catch (Exception)
                        {
                            
                            continue;
                        }
                        yesRowsCount++;
                        double len = Math.Abs(Smile - eMile);
                        rowCount++;
                        outobj[i, 0] = rowCount.ToString();
                        outobj[i, 1] = infoobj[0, 1];
                        outobj[i, 2] = infoobj[0, 4];
                        outobj[i, 3] = infoobj[0, 3];
                        string line = infoobj[0, 6].ToString();
                        outobj[i, 4] = line;
                        outobj[i, 5] = Smile;
                        outobj[i, 6] = eMile;
                        outobj[i, 7] = len;

                        outobj[i, 8] = infoobj[0, 5];
                        outobj[i, 9] =  "水泥路面";
                        outobj[i, 10] = tobj_jsmx[i + 1, 4];
                        outobj[i, 11] = tobj_jsmx[i + 1, 5];
                        outobj[i, 12] = tobj_jsmx[i + 1, 6];
                        outobj[i, 13] = tobj_jsmx[i + 1, 7];
                        outobj[i, 14] = tobj_jsmx[i + 1, 8];
                        outobj[i, 15] = tobj_jsmx[i + 1, 9];
                        outobj[i, 16] = tobj_jsmx[i + 1, 10];
                        outobj[i, 17] = tobj_jsmx[i + 1, 11];
                        outobj[i, 18] = tobj_jsmx[i + 1, 12];

                    }
                    result = new object[yesRowsCount, 19]; int resultRowCount0 = 0;
                    for (int i = 0; i < userow; i++)
                    {
                       
                        if (outobj[i, 1] == null)
                        {
                            continue;
                        }

                        for (int y = 0; y < 19; y++)
                        {
                            result[resultRowCount0, y] = outobj[i, y];
                           
                        }
                        resultRowCount0++;

                    }
                     
                }
                catch (Exception EX)
                {
                   
                }
                trange = destsheet.get_Range(string.Format("A{0}:S{1}", rowidx, rowidx + yesRowsCount - 1));
                trange.Value2 = result;
                rowidx += yesRowsCount;
                tbook_JSMX.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:S{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }

        private static void WriteAllUnitPQI2Xls_NC2(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> ExcelBHTJList)
        {
            int rowidx = 3;
            int userow = 0;
            object[,] infoobj = new object[1, 12];
            object[,] outobj = null;

            MSExcel.Range trange = null;
            int rowCount = 0;
            for (int ri = 0; ri < ExcelBHTJList.Count; ++ri)
            {
                MSExcel.Workbook tbook_JSMX = excelApp.Workbooks.Open(ExcelBHTJList[ri], Type.Missing,
                   false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                   Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                object[,] tobj_jsmx = null;

                GetPrjInfo(tbook_JSMX, ref infoobj);
                MSExcel.Worksheet tsheet_jsmx;
                try
                {
                    tsheet_jsmx = tbook_JSMX.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;

                }
                catch (Exception)
                {
                    userow = 0;
                    tbook_JSMX.Close();
                    continue;
                }

                int roadgrade = _RoadGradeDict[infoobj[0, 5].ToString()];
                userow = GlobalExcel.judegeusedrow(tsheet_jsmx, 1, 6);

                trange = tsheet_jsmx.get_Range(string.Format("A5:M{0}", userow + 4));
                tobj_jsmx = (object[,])trange.Value2;


                outobj = new object[userow, 20];
                object[,] result = new object[userow, 20];
                int yesRowsCount = 0; //有效数据行数
                try
                {
                    for (int i = 0; i < userow; ++i)
                    {
                        double Smile;
                        double eMile;
                        try
                        {
                            Smile = double.Parse(tobj_jsmx[i + 1, 1].ToString());
                            eMile = double.Parse(tobj_jsmx[i + 1, 2].ToString());
                        }
                        catch (Exception)
                        {

                            continue;
                        }
                        yesRowsCount++;
                        double len = Math.Abs(Smile - eMile);
                        rowCount++;
                        outobj[i, 0] = rowCount.ToString();
                        outobj[i, 1] = infoobj[0, 1];
                        outobj[i, 2] = infoobj[0, 4];
                        outobj[i, 3] = infoobj[0, 3];
                        string line = infoobj[0, 6].ToString();
                        outobj[i, 4] = line;
                        outobj[i, 5] = Smile;
                        outobj[i, 6] = eMile;
                        outobj[i, 7] = len;

                        outobj[i, 8] = infoobj[0, 5];
                        outobj[i, 9] = "沥青混凝土";
                        outobj[i, 10] = tobj_jsmx[i + 1, 4];
                        outobj[i, 11] = tobj_jsmx[i + 1, 5];
                        outobj[i, 12] = tobj_jsmx[i + 1, 6];
                        outobj[i, 13] = tobj_jsmx[i + 1, 7];
                        outobj[i, 14] = tobj_jsmx[i + 1, 8];
                        outobj[i, 15] = tobj_jsmx[i + 1, 9];
                        outobj[i, 16] = tobj_jsmx[i + 1, 10];
                        outobj[i, 17] = tobj_jsmx[i + 1, 11];
                        outobj[i, 18] = tobj_jsmx[i + 1, 12];

                    }
                    result = new object[yesRowsCount, 20];
                    int resultRowCount0 = 0;
                    for (int i = 0; i < userow; i++)
                    {

                        if (outobj[i, 1] == null)
                        {
                            continue;
                        }

                        for (int y = 0; y < 20; y++)
                        {
                            result[resultRowCount0, y] = outobj[i, y];

                        }
                        resultRowCount0++;

                    }

                }
                catch (Exception)
                {

                }
                trange = destsheet.get_Range(string.Format("A{0}:S{1}", rowidx, rowidx + yesRowsCount - 1));
                trange.Value2 = result;
                rowidx += yesRowsCount;
                tbook_JSMX.Close();
            }
            trange = destsheet.get_Range(string.Format("A1:S{0}", rowidx - 1));
            GlobalExcel.SetBorderLine(trange, 63);
        }
        #endregion

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

                string roadcode = tobj[4, 1]?.ToString();
                if (roadcode is null)
                {
                    roadcode = "";
                }
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
                    //disval[i, 0] = string.Format("=沥青病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_lq_s);
                    disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
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
                    //disval[i, 0] = string.Format("=水泥病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_sn_s);
                    disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                }
                destrange = worksheet_sntj.get_Range("C4:C" + (disnum + 3).ToString());
                destrange.Value2 = disval;
            }
            else
            {
                worksheet_sntj.Delete();
            }
        }
        #region 孝感定制 
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

        public static void OutputRoadBedDis_XG(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int xlslen)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\路基损坏汇总表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\景观报表模板\孝感定制\景观病害统计表.xlsx",
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
        #region 广西桂兴达
        /////////////////////////////////////////////////////////////////////////////////////////////////
        public static void OutputRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\综合报表模板.xlsx",
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
            bool Hasssflag = false;

            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sshz = _Workbook.Sheets["砂石病害汇总表"] as MSExcel.Worksheet;
            WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, _Worksheet_sshz,
                prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, ref Hasssflag, 5, 53);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
           // WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjdir, _RoadPart, Haslqflag, Hassnflag);
            MSExcel.Worksheet _Worksheet_sstj = _Workbook.Sheets["砂石病害统计表"] as MSExcel.Worksheet;
            WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_sstj, prjdir, _RoadPart, Haslqflag, Hassnflag, Hasssflag);
           // WriteAll2Xls(_Workbook, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _PBIVal, _DeltaHVal, disval);
            WriteAll2Xls(_Workbook, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, disval);

            MSExcel.Worksheet _worksheet_RoadInfo = _Workbook.Sheets["路线信息表"] as MSExcel.Worksheet;
            WriteRoadInfo(_worksheet_RoadInfo, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        #region 低等级农村路专用  
        //养护需求建议表、技术状况明细表、分项指标统计表
        private static void WriteAll2Xls(MSExcel.Workbook workbook, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
            Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, int disval)
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
                prjinfo._IsRut = false;
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
                 
                    mxlist[i, 16] = "";

                    mxlist[i, 10] = "";
                    /*mxlist[i, 27] = string.Format("=IF(K{0}>={1},\"优\",IF(K{0}>={2},\"良\",IF(K{0}>={3},\"中\",IF(K{0}>={4},\"次\",\"差\"))))",
                        i + 4,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);*/
                    mxlist[i, 27] = ""; //rdi  路面车辙
                }

                if (prjinfo._IsIRIMTD)
                {
                    mxlist[i, 17] = "";
                    mxlist[i, 18] = "";
                    mxlist[i, 19] = "";
                   /* mxlist[i, 11] = string.Format("=IF((100-R{0}*{1}-S{0}*{2}-T{0}*{3})>0,(100-R{0}*{1}-S{0}*{2}-T{0}*{3}),0)",
                    i + 4, _PBIScore[1], _PBIScore[2], _PBIScore[3]);*/
                    mxlist[i, 11] = "";

                    /* mxlist[i, 28] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                     i + 4, _PBIGrade[roadpart[i].roaddegree][0],
                     _PBIGrade[roadpart[i].roaddegree][1],
                     _PBIGrade[roadpart[i].roaddegree][2],
                     _PBIGrade[roadpart[i].roaddegree][3],
                     (char)('A' + 11));*/
                    mxlist[i, 28] = "";
                   

                    if (disval == 10)
                    {
                        mxlist[i, 31] =  string.Format("=ROUND(({1}*I{0}+{2}*J{0})/({1}+{2}),5)", i,
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    }
                }

                //构造深度相关 

                if (prjinfo._IsDIRIMTD)
                {
                    mxlist[i, 21] = "";
                    mxlist[i, 22] = "";
                    mxlist[i, 23] = "";
                    //if (prjinfo._IsMMTD)
                    //{
                    //    mxlist[i, 22] = MMTDVal[i];
                    //    //wrval = wrval > 100 ? 100 : (wrval < 0 ? 100 : wrval);
                    //    if (MMTDVal[i] == 0)
                    //    {
                    //        mxlist[i, 23] = 0;
                    //    }
                    //    else  // vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ",i + 4);
                    //    {
                    //        mxlist[i, 23] = string.Format("=IF(W{0}-MIN(U{0},V{0})>0, 100*(W{0}-MIN(U{0},V{0}))/W{0},0)", i + 4);
                    //    }
                    //}
                    //else
                    //{
                    //    mxlist[i, 23] = 0;
                    //}
                }
                   mxlist[i, 12] = "";

                   mxlist[i, 29] = "";

                   
                     mxlist[i, 30] = string.Format("=ROUND(({1}*I{0}+{2}*J{0})/({1}+{2}),5)",
                      i + 4,
                     _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                     _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                   
                mxlist[i, 24] = string.Format("=IF(AE{0}>={1},\"优\",IF(AE{0}>={2},\"良\",IF(AE{0}>={3},\"中\",IF(AE{0}>={4},\"次\",\"差\"))))",
                i + 4,
                _PQIGrade[roadpart[i].roaddegree][0],
                _PQIGrade[roadpart[i].roaddegree][1],
                _PQIGrade[roadpart[i].roaddegree][2],
                _PQIGrade[roadpart[i].roaddegree][3]);
                mxlist[i, 7] = String.Format("=CONCATENATE(TEXT(AE{0},\"0.00\"),\"(\",Y{0},\")\")", i + 4); //pqi
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
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:AE{0}", len + 3));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (disval != 10)
            {
                //((MSExcel.Range)worksheet.Cells[System.Reflection.Missing.Value, 32]).EntireColumn.Delete();
            }
            else
            {
               // GlobalExcel.WriteExcel(2, 32, 2, 1, "跳车值\nH", worksheet, 15);
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
            prjinfo._IsRut = false;
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
        #endregion 低等级农村路专用
       /* private static void WriteAll2Xls(MSExcel.Workbook workbook, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
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
        */
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
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\报表模板5.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
            WriteIRIMTD2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, disval);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteIRIMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
            double[] LIRIVal, double[] RIRIVal/*, double[] LMTDVal, double[] RMTDVal, double[] SpeedVal,*/, int disval)
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
                /*if (SpeedVal != null)
                {
                    vallist[i, 1] = SpeedVal[i];
                }*/
                vallist[i, 1] = "";
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 3] = String.Format("=H{0}*0.6", i + startidx);
                    //vallist[i, 5] = RMTDVal[i];
                    vallist[i, 5] = "";
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
                //vallist[i, 4] = LMTDVal[i];
                vallist[i, 4] = "";
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

       





        public static void outPutAutoTest_GuiZhou(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\定制报表\贵州省农村公路路况检测照片交换模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_贵州省农村公路路况检测照片交换模板_{2}m.xlsx", path, prjdir.Name,disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["路况检测照片"] as MSExcel.Worksheet;
            writeAutoTestXls_GuiZhou(_Worksheet_, prjinfo, prjdir, _RoadPart,disval);



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
            writeAutoTestXls_iri(_Worksheet_iri,prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
           
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// 沥青破损
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void OutputLQDamage(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\沥青路面损坏自动化检测数据.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_沥青路面损坏.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //bug
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            


            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面损坏"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisLQDamage2Xls( _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53,0,9,ref has);
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
        /// 沥青破损
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void OutputSNDamage(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\水泥路面损坏自动化检测数据.xlsx",
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
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["水泥路面损坏"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisLQDamage2Xls(_Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, 1,9,ref has);
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
          ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType,int cluCount,ref bool has)
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


                destrange = _Worksheet.get_Range(String.Format("A2:{1}{0}", rowCount +1, chars[colcnt - 1]));
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
        #region 报送格式
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
            WritePCI2Xls_5(_Worksheet, prjinfo, prjdir, _RoadPart,2, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }
        private static void WritePCI2Xls_5(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
       List<MilePart> roadpart, int DataStartXlsxRow,Disease[] arrdis)
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
                vallist[i, 0] = i+1;
                vallist[i, 1] = smile;
                vallist[i, 2] = emile;
                vallist[i, 3] = LIRIVal[i]; ;
                if (RIRIVal!=null&& RIRIVal.Length>i)
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

            //if (_Setting.IsExcelSort)
            //{
            //    GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 11, true);
            //    GlobalExcel.Reflection(_Worksheet, DataStartXlsxRow, 1, 2, false);
            //}
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

        #region 贵州质安检测
        public static void outPutAccessory04(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件4：交通安全设施排查表.xlsx",
          System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_交通安全设施排查表.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //路况检测图片
            MSExcel.Worksheet worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
           
           
            int length = _RoadPart.Count - 1;
            object[,] data = new object[length,26];

            for (int i = 0; i < length; i++)
            {
                int _colcnt = 0;
                int smile = _RoadPart[i].mile;
                double emile = _RoadPart[i + 1].mile;
                double milelength = Math.Abs(smile - emile);
                data[i, _colcnt++] = i + 1;
                data[i, _colcnt++] = prjinfo._City;
                data[i, _colcnt++] = prjinfo._District;
                data[i, _colcnt++] = "";
                data[i, _colcnt++] = "";
                data[i, _colcnt++] = prjinfo._RoadName;
                data[i, _colcnt++] = "";
                data[i, _colcnt++] = smile/1000.0;
                data[i, _colcnt++] = emile/1000.0;
                data[i, _colcnt++] = Math.Abs(smile / 1000.0- emile / 1000.0) ;
                data[i, _colcnt++] =_RoadConfig.DetectWidth;
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";//四
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] ="";
                data[i, _colcnt++] =""; 
            }
            // 获取数据行数
            int rowCount = data.GetLength(0);
            if (rowCount == 0) return; // 如果数据为空，直接返回

            int startRowIndex = 8;
            // 在指定行插入多行
            MSExcel::Range startRow = worksheet.Rows[startRowIndex];
            for (int i = 0; i < rowCount; i++)
            {
                startRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            // 写入数据
            int columnCount = data.GetLength(1);
            MSExcel::Range writeRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex + rowCount - 1, columnCount]];
            writeRange.Value = data;
            worksheet.Cells[7, 10].Formula = $"=SUBTOTAL(9,J8:J{startRowIndex + rowCount})"; 
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

        private static Dictionary<string, string> CodeOfAdministrativeDivisionDic = new Dictionary<string, string>();

         
       private static void LoadAdministrativeDivisionCodes(MSExcel.Worksheet worksheet)
    {
        try
        {
            

            // 从第2行开始读取（假设第1行是标题）
            int row = 2;

            // 清空字典（如果之前有数据）
            CodeOfAdministrativeDivisionDic.Clear();

            // 循环读取直到遇到空行
            while (true)
            {
                // 获取行政区划名称（A列）和代码（B列）
                MSExcel.Range nameRange = worksheet.Cells[row, 1] as MSExcel.Range;
                MSExcel.Range codeRange = worksheet.Cells[row, 2] as MSExcel.Range;

                // 检查是否为空行
                if (nameRange.Value == null || codeRange.Value == null)
                    break;

                // 获取值并转换为字符串
                string name = nameRange.Value.ToString();
                string code = codeRange.Value.ToString();

                // 添加到字典（名称作为键，代码作为值）
                if (!CodeOfAdministrativeDivisionDic.ContainsKey(name))
                {
                    CodeOfAdministrativeDivisionDic.Add(name, code);
                }

                row++;
            }
        }
        catch (Exception ex)
        {
            // 处理异常
            Console.WriteLine($"加载行政区划代码时出错: {ex.Message}");
            throw; // 根据需求决定是否重新抛出异常
        }
    }
        public static void outPutAccessory05(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件5：村民组数据交换模板2024.xlsx",
          System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_村民组数据交换模板2024.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //读取行政区划代码
            MSExcel.Worksheet worksheet2 = _Workbook.Sheets["填报说明"] as MSExcel.Worksheet;

            LoadAdministrativeDivisionCodes(worksheet2);

            //路况检测图片
            MSExcel.Worksheet worksheet = _Workbook.Sheets["村名组"] as MSExcel.Worksheet;


            int length = _RoadPart.Count - 1; 
            List<object[]> dataLine = new List<object[]>();
            ExcelGPS[] dicGps = null;
             
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref dicGps);
            string sName = prjinfo._Direction > 0 ? "A" : "B";
            string cityCode;
            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._City))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._City];
            }

            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._District))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._District];
            }
            int indexCount = 0; 
            for (int i = 0; i < length; i++)
            {
                if (_MarkVal[i]!= null)
                {
                    if (!_MarkVal[i].Contains("村民"))
                    {
                        continue;
                    }
                    indexCount++;
                    int _colcnt = 0;
                    object[] line = new object[13];
                    line[_colcnt++] = indexCount;
                    line[_colcnt++] =prjinfo._City; 
                    line[_colcnt++] = prjinfo._District; 
                    line[_colcnt++] = prjinfo._CityCode;
                    line[_colcnt++] = "";
                    line[_colcnt++] = _MarkVal[i].Split(' ').Last();
                    line[_colcnt++] = dicGps[i]._longitude;
                    line[_colcnt++] = dicGps[i]._latitude;
                    line[_colcnt++] = dicGps[i]._elevation;

                    string input = prjinfo._DataDate + " " + prjinfo._DataTime;

                    // 尝试解析输入字符串
                    if (DateTime.TryParseExact(input, "yyyyMMdd HHmmss",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                    {
                        // 格式化为目标格式
                        string output = result.ToString("yyyy/M/d HH:mm:ss");
                        Console.WriteLine(output); // 输出: 2025/4/8 17:18:17
                        line[_colcnt++] = output;
                    }
                    else
                    {
                        line[_colcnt++] = input;
                        Console.WriteLine("日期格式不正确");
                    }
                

                  
                    line[_colcnt++] = "";

                    line[_colcnt++] = GenerateGuidFromString(prjinfo._RoadName + sName);//UID
                    line[_colcnt++] = "无"; 
                    dataLine.Add(line);
                }
            }

           
            // 获取数据行数
            int rowCount = dataLine.Count;
            if (rowCount == 0) return; // 如果数据为空，直接返回

            int startRowIndex = 2;
            object[,] datas = ConvertListTo2DArray(dataLine);
            // 在指定行插入多行
            MSExcel::Range startRow = worksheet.Rows[startRowIndex];
            for (int i = 0; i < rowCount; i++)
            {
                startRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            // 写入数据
            int columnCount = datas.GetLength(1);
            MSExcel::Range writeRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex + rowCount - 1, columnCount]];
            writeRange.Value = datas; 
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }
        public static object[,] ConvertListTo2DArray(List<object[]> dataLine)
        {
            if (dataLine == null || dataLine.Count == 0)
                return new object[0, 0];

            // 确定行数和列数
            int rowCount = dataLine.Count;
            int colCount = dataLine[0].Length; // 假设所有内部数组长度相同

            // 创建二维数组
            object[,] result = new object[rowCount, colCount];

            // 填充数据
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    result[i, j] = dataLine[i][j];
                }
            }

            return result;
        }
        /// <summary>
        /// 根据输入字符串生成一个确定的 32 位 GUID。
        /// </summary>
        /// <param name="input">输入字符串</param>
        /// <returns>32 位 GUID 字符串（无连字符，32 个字符）</returns>
        public static string GenerateGuidFromString(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                // 如果输入为空，返回一个随机 GUID 的 32 位格式
                return Guid.NewGuid().ToString("N");
            }

            // 使用 MD5 计算输入字符串的哈希值
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(input);
                byte[] hashBytes = md5.ComputeHash(inputBytes);

                // MD5 哈希值是 16 字节，恰好可以转换为 GUID
                // GUID 需要 16 字节，hashBytes 直接使用
                Guid guid = new Guid(hashBytes);

                // 返回 32 位格式（无连字符）
                return guid.ToString("N");
            }
        }



        public static void outPutAccessory06(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件6：路线轨迹数据交换模板.xlsx",
          System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_路线轨迹数据交换模板.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //读取行政区划代码
            MSExcel.Worksheet worksheet2 = _Workbook.Sheets["填报说明"] as MSExcel.Worksheet;

            LoadAdministrativeDivisionCodes(worksheet2);

            //路况检测图片
            MSExcel.Worksheet worksheet = _Workbook.Sheets["路况检测轨迹"] as MSExcel.Worksheet;

            string dirStr = prjinfo._Direction == 0 ? "A" : "B";
            string cityCode;
            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._City))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._City];
            }

            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._District))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._District];
            }
            string roadCode = prjinfo._RoadCode + ((prjinfo._Direction > 0) ? 'A' : 'B');
            string DestPath = string.Format($"{roadCode}-LOC-{prjinfo._StartMile.ToString("f3")}-{prjinfo._DataDate + prjinfo._DataTime}.txt");
            int length = _RoadPart.Count - 1;
            
            object[,] data = new object[length, 19];

            for (int i = 0; i < length; i++)
            {
                string roadType = _RoadPart[i].roadtype == 0 ? "沥青混凝土" : _RoadPart[i].roadtype == 1 ? "水泥混凝土" : "砂石";
                int _colcnt = 0;
                int smile = _RoadPart[i].mile;
                double emile = _RoadPart[i + 1].mile;
                double milelength = Math.Abs(smile - emile);
                data[i, _colcnt++] = i + 1;
                data[i, _colcnt++] = prjinfo._District;
                data[i, _colcnt++] = prjinfo._CityCode;
                data[i, _colcnt++] = DestPath;
                data[i, _colcnt++] = Math.Abs((prjinfo._StartMile - prjinfo._EndMile)/1000.0);
                data[i, _colcnt++] = prjinfo._RoadCode;
                data[i, _colcnt++] = prjinfo._RoadName;
                data[i, _colcnt++] = roadType;
                data[i, _colcnt++] = _RoadConfig.DetectWidth;
                data[i, _colcnt++] = "";
                data[i, _colcnt++] = "";
                data[i, _colcnt++] = smile / 1000.0;
                data[i, _colcnt++] = emile / 1000.0;
                data[i, _colcnt++] = Math.Abs(smile / 1000.0 - emile / 1000.0);


                string input = prjinfo._DataDate + " " + prjinfo._DataTime;

                // 尝试解析输入字符串
                if (DateTime.TryParseExact(input, "yyyyMMdd HHmmss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                {
                    // 格式化为目标格式
                    string output = result.ToString("yyyy/M/d HH:mm:ss");
                    Console.WriteLine(output); // 输出: 2025/4/8 17:18:17
                    data[i, _colcnt++] = output;
                }
                else
                {
                    data[i, _colcnt++]  = input;
                    Console.WriteLine("日期格式不正确");
                }
                data[i, _colcnt++] = ""; 
                data[i, _colcnt++] = GenerateGuidFromString(prjinfo._RoadName + dirStr);//UID
                data[i, _colcnt++] = ""; 
                data[i, _colcnt++] = ""; 

            }
            // 获取数据行数
            int rowCount = data.GetLength(0);
            if (rowCount == 0) return; // 如果数据为空，直接返回

            int startRowIndex = 2;
            // 在指定行插入多行
            MSExcel::Range startRow = worksheet.Rows[startRowIndex];
            for (int i = 0; i < rowCount; i++)
            {
                startRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            // 写入数据
            int columnCount = data.GetLength(1);
            MSExcel::Range writeRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex + rowCount - 1, columnCount]];
            writeRange.Value = data;
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }


        public static void outPutAccessory07(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件7：安全隐患数据交换模板.xlsx",
          System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_安全隐患数据交换模板.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //读取行政区划代码
            MSExcel.Worksheet worksheet2 = _Workbook.Sheets["填报说明"] as MSExcel.Worksheet;

            LoadAdministrativeDivisionCodes(worksheet2);

            //路况检测图片
            MSExcel.Worksheet worksheet = _Workbook.Sheets["路况检测轨迹"] as MSExcel.Worksheet;

            string dirStr = prjinfo._Direction == 0 ? "A" : "B";
            string cityCode;
            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._City))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._City];
            }

            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._District))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._District];
            }
            string roadCode = prjinfo._RoadCode + ((prjinfo._Direction > 0) ? 'A' : 'B');
            string DestPath = string.Format($"{roadCode}-LOC-{prjinfo._StartMile.ToString("f3")}-{prjinfo._DataDate + prjinfo._DataTime}.txt");
            int length = _RoadPart.Count - 1;

            int rowIdx = 0;

            List<object[]> datas = new List<object[]>();

            for (int i = 0; i < length; i++)
            {
                string markInfo = "";
                if (_MarkVal[i]==null)
                {

                    continue;
                }
                if (_MarkVal[i].Contains("村民"))
                {
                    continue;
                }
                if (!_MarkVal[i].Contains("路面情况"))
                {
                    continue;
                }
                rowIdx++;
                object[] line = new object[ 19];
                markInfo = _MarkVal[i].Split(' ').Last();
                string roadType = _RoadPart[i].roadtype == 0 ? "沥青混凝土" : _RoadPart[i].roadtype == 1 ? "水泥混凝土" : "砂石";
                int _colcnt = 0;
                int smile = _RoadPart[i].mile;
                double emile = _RoadPart[i + 1].mile;
                double milelength = Math.Abs(smile - emile);
                 line[ _colcnt++] = rowIdx;
                 line[ _colcnt++] = prjinfo._District;
                 line[ _colcnt++] = prjinfo._CityCode;
                 line[ _colcnt++] = DestPath;
                 line[ _colcnt++] = Math.Abs((prjinfo._StartMile - prjinfo._EndMile) / 1000.0);
                 line[ _colcnt++] = prjinfo._RoadCode;
                 line[ _colcnt++] = prjinfo._RoadName;
                 line[ _colcnt++] = roadType;
                 line[ _colcnt++] = _RoadConfig.DetectWidth;
                 line[ _colcnt++] = "";
                 line[ _colcnt++] = "";
                 line[ _colcnt++] = smile / 1000.0;
                 line[ _colcnt++] = emile / 1000.0;
                line[ _colcnt++] = Math.Abs(smile / 1000.0 - emile / 1000.0);


                string input = prjinfo._DataDate + " " + prjinfo._DataTime;

                // 尝试解析输入字符串
                if (DateTime.TryParseExact(input, "yyyyMMdd HHmmss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                {
                    // 格式化为目标格式
                    string output = result.ToString("yyyy/M/d HH:mm:ss");
                    Console.WriteLine(output); // 输出: 2025/4/8 17:18:17
                    line[_colcnt++] = output;
                }
                else
                {
                    line[_colcnt++] = input;
                    Console.WriteLine("日期格式不正确");
                }
               line[  _colcnt++] = "";
               line[  _colcnt++] = GenerateGuidFromString(prjinfo._RoadName + dirStr);//UID
               line[  _colcnt++] = "";
                line[  _colcnt++] = markInfo;
                datas.Add(line);
            }
            // 获取数据行数
            object[,] data = ConvertListTo2DArray(datas);
            int rowCount = data.GetLength(0);
            if (rowCount == 0) return; // 如果数据为空，直接返回

            int startRowIndex = 2;
            // 在指定行插入多行
            MSExcel::Range startRow = worksheet.Rows[startRowIndex];
            for (int i = 0; i < rowCount; i++)
            {
                startRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            // 写入数据
            int columnCount = data.GetLength(1);
            MSExcel::Range writeRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex + rowCount - 1, columnCount]];
            writeRange.Value = data;
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }

        public static void outPutAccessory08(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件8：POI交换模板.xlsx",
          System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_POI交换模板.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            //读取行政区划代码
            MSExcel.Worksheet worksheet2 = _Workbook.Sheets["填报说明"] as MSExcel.Worksheet;

            LoadAdministrativeDivisionCodes(worksheet2);

            //路况检测图片
            MSExcel.Worksheet worksheet = _Workbook.Sheets["路况检测POI"] as MSExcel.Worksheet;

            string dirStr = prjinfo._Direction == 0 ? "A" : "B";
            string cityCode;
            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._City))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._City];
            }

            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._District))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._District];
            }
            string roadCode = prjinfo._RoadCode + ((prjinfo._Direction > 0) ? 'A' : 'B');
            string DestPath = string.Format($"{roadCode}-LOC-{prjinfo._StartMile.ToString("f3")}-{prjinfo._DataDate + prjinfo._DataTime}.txt");
            int length = _RoadPart.Count - 1;

            int rowIdx = 0;
            List<object[]> datas = new List<object[]>();
            int dlen = _RoadDisList.Length;
            bool res = false; int typeidx = 0;
            for (int i = 0 ,j = 0 ;  i < length; i++)
            {

                double pcival = 0, drval = 0 , irival = 0 , trqival = 0 ;
                int smile = _RoadPart[i].mile;
                int emile = _RoadPart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && _RoadDisList[j].m_mile >= smile && _RoadDisList[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && _RoadDisList[j].m_mile <= smile && _RoadDisList[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[_RoadPart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                            _RoadDisList[j].RoadType, _RoadDisList[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        RoadDiseaseTypes.roaddis[_RoadPart[i].roadtype][typeidx].totalarea += _RoadDisList[j].Area;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", _RoadDisList[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[_RoadPart[i].roadtype]);
                      
                    }
                    ++j;
                }
                //计算pci
                drval = ComputPCI(RoadDiseaseTypes.roaddis, _RoadPart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][0] * Math.Pow(drval, _PCIa[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][1]);
               
               
                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    if (_Setting.RQIJudgeType == 0)
                    {
                        irival = Math.Round((_LIRIMeanVal[i] + _RIRIMeanVal[i]) * 0.5, 5);
                    }
                    else if (_Setting.RQIJudgeType == 1)
                    {
                        irival = Math.Round(Math.Max(_LIRIMeanVal[i], _RIRIMeanVal[i]), 5);
                    }
                }
                else
                {

                    irival = Math.Round(_LIRIMeanVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][0] * Math.Exp(_RQIa[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][1] * irival));

                double pqiValue = Math.Round((pcival * _PQIW[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][0] + irival * _PQIW[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][1])
                    /(_PQIW[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][0] + _PQIW[_RoadPart[i].roaddegree][_RoadPart[i].roadtype][1]), 2);
                int pqiIdx = 0;
                if (pqiValue >= _PQIGrade[_RoadPart[i].roaddegree][0])
                {
                    pqiIdx = 0;
                }
                else if (pqiValue >= _PQIGrade[_RoadPart[i].roaddegree][1])
                {
                    pqiIdx = 1;
                }
                else if (pqiValue >= _PQIGrade[_RoadPart[i].roaddegree][2])
                {
                    pqiIdx = 2;
                }
                else if (pqiValue >= _PQIGrade[_RoadPart[i].roaddegree][3])
                {
                    pqiIdx = 3;
                }
                else
                {
                    pqiIdx = 4;
                }


                rowIdx++;
                object[] line = new object[24]; 
                string roadType = _RoadPart[i].roadtype == 0 ? "沥青混凝土" : _RoadPart[i].roadtype == 1 ? "水泥混凝土" : "砂石";
                int _colcnt = 0;
               
           
                line[_colcnt++] = rowIdx;
                line[_colcnt++] = prjinfo._City;
                line[_colcnt++] = prjinfo._District;
                line[_colcnt++] = prjinfo._CityCode;

                string input = prjinfo._DataDate;

                // 尝试解析输入字符串
                if (DateTime.TryParseExact(input, "yyyyMMdd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                {
                    // 格式化为目标格式
                    string output = result.ToString("yyyy/M/d");
                    Console.WriteLine(output); // 输出: 2025/4/8
                    line[_colcnt++] = output;
                }
                else
                {
                    line[_colcnt++] = input;
                    Console.WriteLine("日期格式不正确");
                }
                line[_colcnt++] = ""; 
                line[_colcnt++] = prjinfo._RoadCode;
                line[_colcnt++] = prjinfo._RoadName;
                line[_colcnt++] = "";
                line[_colcnt++] = smile / 1000.0;
                line[_colcnt++] = emile / 1000.0;
                line[_colcnt++] = "";
                line[_colcnt++] = pqiValue;//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = "";//PQI
                line[_colcnt++] = _RoadConfig.DetectWidth ; 
                line[_colcnt++] = roadType; 
                line[_colcnt++] = "";
                line[_colcnt++] = GenerateGuidFromString(prjinfo._RoadName + dirStr);//UID
                line[_colcnt++] = "";

                line[pqiIdx + 13] = milelength / 1000.0;
                datas.Add(line);
            }
            // 获取数据行数
            object[,] data = ConvertListTo2DArray(datas);
            int rowCount = data.GetLength(0);
            if (rowCount == 0) return; // 如果数据为空，直接返回

            int startRowIndex = 5;
            // 在指定行插入多行
            MSExcel::Range startRow = worksheet.Rows[startRowIndex];
            for (int i = 0; i < rowCount; i++)
            {
                startRow.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            }
            // 写入数据
            int columnCount = data.GetLength(1);
            MSExcel::Range writeRange = worksheet.Range[worksheet.Cells[startRowIndex, 1], worksheet.Cells[startRowIndex + rowCount - 1, columnCount]];
            writeRange.Value = data;
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

        }
        public static void outPutAccessory09(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            //生成图片
            //找到所有景观图片路径
            string streetIndexFile = prjdir.FullName + "\\StreetImg\\Camera0\\Street2Mile.txt";
            string[] streetInfos = File.ReadAllLines(streetIndexFile);
            if (streetInfos.Length <= 0)
            {
                MessageBox.Show("未找到有效\\StreetImg\\Camera0\\Street2Mile.txt文件");
                return;
            }
            string picResultPath = path + $"\\{prjdir.Name}_景观图片\\";
            if (!Directory.Exists(picResultPath))
            {
                Directory.CreateDirectory(picResultPath);
            }
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\贵州质安检测\附件9：检测照片交换模板.xlsx",
           System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_检测照片交换结果.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //读取行政区划代码
            MSExcel.Worksheet worksheet2 = _Workbook.Sheets["填报说明"] as MSExcel.Worksheet;

            LoadAdministrativeDivisionCodes(worksheet2);
            //路况检测图片
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["路况检测图片"] as MSExcel.Worksheet;
            //填写表格


            ExcelGPS[] dicGps = null;

            int rowcnt = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref dicGps);
            int gpslen = dicGps.Length;

            MSExcel.Range destrange;
            int len = streetInfos.Length;


            List<object[]> objects = new List<object[]>();
            int curDmi = 0;
            int split = 0;
            bool hiddenDangerRoad = false; 
            int lastMile = -100; // 用于跟踪上一次输出的里程


            string cityCode;
            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._City))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._City];
            }

            if (CodeOfAdministrativeDivisionDic.ContainsKey(prjinfo._District))
            {
                prjinfo._CityCode = CodeOfAdministrativeDivisionDic[prjinfo._District];
            }
            string roadCode = prjinfo._RoadCode + ((prjinfo._Direction > 0) ? 'A' : 'B');


            //轨迹路段每过100m输出一次  隐患路段每过100m 5m一张 村名点一个点一张
            for (int i = 0; i < _RoadPart.Count; i++)
            {
                if (i>= streetInfos.Length)
                {
                    continue;
                }
                string curStreetMsg = streetInfos[i];
                if (string.IsNullOrEmpty(curStreetMsg))
                {
                    continue;
                }
                // 更新隐患路段状态
                if (_MarkVal[i] != null && _MarkVal[i].Contains("起点") && _MarkVal[i].Contains("路面情况"))
                    hiddenDangerRoad = true;
                if (i!=0)
                {
                    if (_MarkVal[i - 1] != null && _MarkVal[i - 1].Contains("终点") && _MarkVal[i - 1].Contains("路面情况"))
                        hiddenDangerRoad = false; 
                }
               
                int currentMile = _RoadPart[i].dmi; // 假设第一个元素是里程（米）

                // 判断是否需要输出图片
                bool shouldOutput = false;
                if (hiddenDangerRoad)
                {
                    // 隐患路段：每 5 米输出一次
                    if (currentMile % 5 == 0 && currentMile >= lastMile + 5)
                        shouldOutput = true;
                }
                else if (_MarkVal[i]!=null&& _MarkVal[i].Contains("村民") && _MarkVal[i].Contains("路面情况"))
                {
                    // 村名点：每个点输出一次
                    shouldOutput = true;
                }
                else
                {
                    // 普通路段：每 100 米输出一次
                    if (currentMile % 100 == 0 && currentMile >= lastMile + 100)
                        shouldOutput = true;
                }
                if (!shouldOutput)
                    continue;

                lastMile = currentMile;

                int rowCnt = 0;

                object[] vallist = new object[ 16];

                vallist[  rowCnt++] = i + 1;
                vallist[  rowCnt++] = prjinfo._City;
                vallist[  rowCnt++] = prjinfo._CityCode;
                string imputTime = prjinfo._DataDate + " " + prjinfo._DataTime;
                string imputTime1 = prjinfo._DataDate + prjinfo._DataTime;



                string[] picMsg = curStreetMsg.Split(' ');
                string newPicName = $"{prjinfo._RoadCode}-{(int.Parse(picMsg.First()) / 1000.0).ToString("0.000")}-01-{imputTime1}.jpg";

                string oriPicFilePath = prjdir.FullName + "\\StreetImg\\Camera0" + picMsg.Last();
                string newPicFilePath = picResultPath + newPicName;

                File.Copy(oriPicFilePath, newPicFilePath, true);

                vallist[  rowCnt++] = newPicName;


                vallist[ rowCnt++] = (int.Parse(picMsg.First()) / 1000.0).ToString("0.000");


                vallist[ rowCnt++] = 5;
                string dirStr = "A";
                if (prjinfo._Direction != 1)
                {
                    dirStr = "B";
                }
                vallist[ rowCnt++] = prjinfo._RoadCode + dirStr;
                vallist[ rowCnt++] = prjinfo._RoadName;
                if (gpslen > i)
                {
                    vallist[  rowCnt++] = dicGps[i]._longitude;
                    vallist[  rowCnt++] = dicGps[i]._latitude;
                    vallist[  rowCnt++] = dicGps[i]._elevation;
                }
                else
                {
                    vallist[  rowCnt++] = "经度";
                    vallist[  rowCnt++] = "纬度";
                    vallist[  rowCnt++] = "高程";
                }


                DateTime dateTime;
                if (DateTime.TryParseExact(
                    imputTime,
                    "yyyyMMdd HHmmss",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out dateTime))
                {
                    // 格式化为目标字符串
                    string formattedTime = dateTime.ToString("yyyy/M/d HH:mm:ss");
                    Console.WriteLine(formattedTime); // 输出: 2023/4/24 10:59:46
                    vallist[ rowCnt++] = formattedTime;
                }
                else
                {
                    vallist[ rowCnt++] = imputTime;

                }
                vallist[ rowCnt++] = "";


                vallist[ rowCnt++] = GenerateGuidFromString(prjinfo._RoadCode + dirStr);
                vallist[ rowCnt++] = "";
                if (_MarkVal[i]!=null )
                {

                    vallist[rowCnt++] = _MarkVal[i];
                }
                else
                {
                    vallist[rowCnt++] = "";

                }


                objects.Add(vallist);
            }
            destrange = _Worksheet.get_Range(String.Format("A{0}:P{1}", 2, objects.Count + 2 - 1));
            destrange.Value2 = ConvertListTo2DArray(objects);



            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        #endregion



      
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
               
                disvalsn[i, _colcnt++] = i+1;
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

        public static void outPutAutoTest_9(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string roadCode = prjinfo._RoadCode + ((prjinfo._Direction > 0) ? 'A' : 'B');
            string DestPath = string.Format($"{path}\\{roadCode}-LOC-{prjinfo._StartMile.ToString("f3")}-{prjinfo._DataDate + prjinfo._DataTime}.txt");

            ExcelGPS[] dicGps = null;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref dicGps);
            int len = _RoadPart.Count;
            string[] result = new string[len+1];
            result[0] = "桩号(km),x(m),y(m),z(m),有效性";
            string formatF6(string s)
            {
                if (double.TryParse(s, out var d))
                {
                    return d.ToString("f6");
                }
                else
                {
                    MessageBox.Show("GPS坐标数据无法转化为六位浮点数，请检查数据有效性！");
                    return "";
                }
            }
            for (int i = 0; i < len; i++)
            {
                double smile = _RoadPart[i].mile;
                // 检查是否与上一个点重复
                if (i > 0 && Math.Abs(smile - _RoadPart[i - 1].mile) < 0.001)
                {
                    // 如果是重复点，手动增加微小偏移
                    smile += 0.0001;
                }
                string s1 = (smile * 0.001).ToString("f6");
                
                string x = formatF6(dicGps[i]._longitude),
                       y = formatF6(dicGps[i]._latitude),
                       z = formatF6(dicGps[i]._elevation);
                string line = $"{s1},{x},{y},{z},A";
                result[i + 1] = line;
            }
            File.WriteAllLines(DestPath, result);

        }





        #endregion

        /// <summary>
        /// 断面高程
        /// </summary>
        private static void writeAutoTestXls_hightData(MSExcel.Worksheet workSheet, List<MilePart> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[]speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count - 1;
            object[,] disvalsn = new object[len, 4];
            bool hasRdlta = rDeltaHVal != null ? true : false;
            MSExcel.Range destrange;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;
                int smile = roadpart10[i].mile;
                int emile = roadpart10[i + 1].mile;
                disvalsn[i, colcnt++] = (smile * 0.001).ToString("f3");
                if (lDeltaHVal!=null)
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
                int  _colcnt = 0;
                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] = s1;
                if (dicGps != null&&dicGps != null)
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
            GlobalExcel.SetBorderLine(destrange,53);
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


        private static void writeAutoTestTxt_gpsData(MSExcel.Worksheet worksheet_snhz,
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
                if (dicGps != null && dicGps != null)
                {
                    disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._latitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._elevation;
                }

                disvalsn[i, _colcnt++] = "A";
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


        private static void writeAutoTestXls_GuiZhou(MSExcel.Worksheet worksheet_snhz,
      ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart,int split)
        {
            ExcelGPS[] dicGps = null;

            int rowcnt = 0;
            GlobalExcel.GetGPSInfo(prjinfo, prjdir, RoadPart, ref dicGps);

            string streetPath = prjdir.FullName + "\\StreetImg\\Camera0\\Street2Mile.txt";
             List<string> streetPicList = File.ReadAllLines(streetPath).ToList();

            int len = RoadPart.Count;
            MSExcel.Range destrange;
            object[,] disvalsn = new object[len, 13];
            int divRate = split / prjinfo._StreetImgDis_Left;

            for (int i = 0; i < len-1; i++)
            {
                int _colcnt = 0;
                double smile = RoadPart[i].mile;
                double emile = RoadPart[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (emile * 0.001).ToString("f3");
                disvalsn[i, _colcnt++] =prjinfo._City;
                string streetPic= "";
                string[] strs = null;
                try
                {
                     strs = streetPicList[i * divRate].Split('\\');

                }
                catch (Exception)
                {

                    strs = streetPicList.Last().Split('\\');
                }
                disvalsn[i, _colcnt++] = strs.Last() ;
                disvalsn[i, _colcnt++] =s1;
                disvalsn[i, _colcnt++] =Math.Abs(smile-emile);
                disvalsn[i, _colcnt++] = prjinfo._RoadCode;
                disvalsn[i, _colcnt++] = prjinfo._RoadName;
              
                if (dicGps != null)
                {
                    double hx_longitude = 0;
                    double hx_latitude = 0;
                    CoordinateConverter.Wgs84ToGcj02(double.Parse( dicGps[i]._longitude), double.Parse(dicGps[i]._latitude),out hx_longitude,out hx_latitude);

                    disvalsn[i, _colcnt++] = hx_longitude;
                    disvalsn[i, _colcnt++] = hx_latitude;

                    disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._latitude; 
                    disvalsn[i, _colcnt++] = dicGps[i]._elevation; 
                }
                else
                {
                    _colcnt += 3;
                }
                disvalsn[i, _colcnt++] = prjinfo._DataDate;
                
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A2:M{0}", len + 1));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A2:M{0}", rowcnt + 1));
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

                destrange = worksheet_snhz.get_Range(String.Format("A2:M{0}", rowcnt + 1));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("C2:C{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        #endregion
        #region 湖南定制
        public static void outPutAutoTest_HN0(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\湖南定制\空间定位表格.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_空间定位数据_{2}m.xlsx", path, prjdir.Name, disval);
            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_ = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            writeAutoTestXls_gpsData_hn(_Worksheet_, prjinfo, prjdir, _RoadPart);


            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void writeAutoTestXls_gpsData_hn(MSExcel.Worksheet worksheet_snhz,
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
                disvalsn[i, _colcnt++] = prjinfo._RoadCode;
                disvalsn[i, _colcnt++] = s1;
                if (dicGps[i] != null && dicGps != null)
                {
                    disvalsn[i, _colcnt++] = dicGps[i]._longitude;
                    disvalsn[i, _colcnt++] = dicGps[i]._latitude;
                    //  disvalsn[i, _colcnt++] = dicGps[i]._elevation;
                }

                disvalsn[i, _colcnt++] = "√";
                rowcnt++;
            }
            destrange = worksheet_snhz.get_Range(String.Format("A3:E{0}", len + 2));
            destrange.Value2 = disvalsn;
            destrange = worksheet_snhz.get_Range(String.Format("A3:E{0}", rowcnt + 2));
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

                destrange = worksheet_snhz.get_Range(String.Format("A3:E{0}", rowcnt + 2));
                MSExcel.Range sortrange = worksheet_snhz.get_Range(String.Format("B3:B{0}", len + 2));//按桩号排序

                GlobalExcel.ReflectionColnum(worksheet_snhz, destrange, sortrange);
            }
        }
        public static void OutputSNDamage_hn(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\湖南定制\水泥病害统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_水泥路面损坏_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //bug
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisLQOrSN_Damage2Xls_hn(_Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, 1, 10, ref has);    
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
        private static void WriteDisLQOrSN_Damage2Xls(MSExcel.Worksheet _Worksheet,
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
                  //  datas[rowCount, colcnt++] = prjinfo._RoadCode;
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


                destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {

                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("B3:B{0}", rowCount + 2));//按桩号排序

                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }/// <summary>
       
        private static void WriteDisLQOrSN_Damage2Xls_hn(MSExcel.Worksheet _Worksheet,
       ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int roadType, int cluCount, ref bool has)
        {
            MSExcel.Range destrange;
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I','J','K','L','M' };
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
                     datas[rowCount, colcnt++] = prjinfo._RoadCode;
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


                destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {

                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("B3:B{0}", rowCount + 2));//按桩号排序

                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }/// <summary>
         /// 湖南定制
         /// </summary>
         /// <param name="excelApp"></param>
         /// <param name="path"></param>
         /// <param name="prjinfo"></param>
         /// <param name="prjdir"></param>
         /// <param name="disval"></param>
        public static void OutputLQDamage_hn(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\湖南定制\沥青路面病害统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_沥青路面损坏_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //bug
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);




            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisLQOrSN_Damage2Xls_hn(_Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, 0, 10, ref has);
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

        public static void OutputAllDamage_hn(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\湖南定制\病害流水表表格.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害流水表表格_{2}m.xlsx",
                path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //bug
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_All = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;
            bool has = false;
            WriteDisAll_hn(_Worksheet_All, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, 8, ref has);
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


        private static void WriteDisAll_hn(MSExcel.Worksheet _Worksheet,
   ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int borderType, int cluCount, ref bool has)
        {
            MSExcel.Range destrange;
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I' };
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowCount = 0;


            int typeidx = 0;
            bool res = false;
            // int rowCount = 2;
            int colcnt = 1;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            int rowCounts = 0;


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
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].count++;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                int roadType = roadpart[i].roadtype;
                for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                {
                    if (RoadDiseaseTypes.roaddis[roadType][dis].count > 0)
                    {
                        //有病害
                        rowCounts++;
                    }

                    // _Worksheet.Cells[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                }
            }




            object[,] datas = new object[rowCounts + 1, cluCount];
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
                        RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].count++;
                    }
                    else
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }

                int roadType = roadpart[i].roadtype;
                for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                {
                    if (RoadDiseaseTypes.roaddis[roadType][dis].count > 0)
                    {
                        //有病害
                        has = true;

                        smile = roadpart[i].mile;
                        emile = roadpart[i + 1].mile;
                        int milelength = Math.Abs(smile - emile);

                        //病害汇总表

                        colcnt = 0;
                        datas[rowCount, colcnt++] = prjinfo._RoadCode;
                        datas[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                        datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].disname;
                        datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].count;
                        datas[rowCount, colcnt++] = "㎡";
                        datas[rowCount, colcnt++] = prjinfo._Direction;
                        datas[rowCount, colcnt++] = roadpart[i].roadtype;
                        datas[rowCount, colcnt++] = "";

                        rowCount++;
                    }

                    // _Worksheet.Cells[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                }
            }

            if (has)
            {
                destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                destrange.Value2 = datas;
                GlobalExcel.SetBorderLine(destrange, 53);

                if (_Setting.IsExcelSort && prjinfo._Direction > 0)
                {

                }
                else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {

                    destrange = _Worksheet.get_Range(String.Format("A3:{1}{0}", rowCount + 2, chars[colcnt - 1]));
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("B3:B{0}", rowCount + 2));//按桩号排序

                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }

        #endregion
        #region 国检转换

        public static void Convent_Iri(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\路面平整度自动化检测数据.xlsx",
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

            MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["路面平整度"] as MSExcel.Worksheet;
            writeAutoTestXls_iri(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            SingleProject.XlsxToCsv(Destxls);
        }

        public static void Convent_Iri_JiangXi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\国检转换\江西\平整度_IRI.xlsx",
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
                if (prjinfo._IsDIRIMTD)
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

                disvalsn[i, colcnt++] = _RoadConfig.DetectWidth;



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
        /// <summary>
        /// 空间定位数据
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        public static void Convent_Lbi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\空间定位数据.xlsx",
                System.Windows.Forms.Application.StartupPath);

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-LBI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);


            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);
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
            SingleProject.XlsxToCsv(Destxls);
        }
        /// <summary>
        /// 平整度原始数据  纵断面高程报表
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        public static void Convent_Iri_Original(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\路面平整度自动化检测原始数据.xlsx",
                System.Windows.Forms.Application.StartupPath);
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-LP-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, fileName);
            // string Destxls = string.Format(@"{0}\{1}_路面平整度.xlsx", path, prjdir.Name);
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

        private static void writeAutoTestXls_Bump(MSExcel.Worksheet worksheet_snhz,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            if (!prjinfo._IsRut)
                return;
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
                disvalsn[i, _colcnt++] = Math.Max(lValue, rValue);

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
        /// 破损
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        /// <param name="prjinfo"></param>
        /// <param name="prjdir"></param>
        /// <param name="disval"></param>
        public static void Convent_Damage(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            string[] srcxlsArr = new string[] {
                string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\沥青路面损坏识别结果.xlsx",
                System.Windows.Forms.Application.StartupPath),
                string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\水泥路面损坏识别结果.xlsx",
                System.Windows.Forms.Application.StartupPath),
                string.Format(@"{0}\报表模板\低等级农村公路\自动化报表\砂石路面损坏识别结果.xlsx",
                System.Windows.Forms.Application.StartupPath)

            };
            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD =double.Parse( roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());
                int col ;
                string srcxls ;
                int[] tableCol = new[] { 6, 7, 6 };
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
                string tableNamePart = System.IO.Path.GetFileNameWithoutExtension(srcxls).Substring(0, 6);
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, endMile, tableNamePart);
                string Destxls = string.Format(@"{0}\{1}.xlsx", path, tempName);

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["sheet1"] as MSExcel.Worksheet;


                WriteDis_Damage2Xls_gj(_Worksheet_snhz, prjinfo, prjdir, _RoadPart, _RoadDisList, 53, i, col, startMileD,endMileStrD);


                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                SingleProject.XlsxToCsv(Destxls);
            }
         
            



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
            List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I' };
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
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        if (dis==1||dis==2)
                        {
                            datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totallength.ToString("f2");

                        }
                        else
                        {
                            datas[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");

                        }
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
                    MSExcel.Range sortrange = _Worksheet.get_Range(String.Format("A2:A{0}", rowCount + 1));//按桩号排序
                    GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                }
            }
        }



        public static void OutConvent(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {

        }
        #endregion
     
  

        #region 车道统计 湖南
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
            WriteAllRoadPQI2Xls_2(excelApp, destsheet, ExcelJSMXList);

            destsheet = _Workbook.Sheets["公里单元技术状况评定明细表"] as MSExcel.Worksheet;
            WriteAllUnitPQI2Xls(excelApp, destsheet, ExcelDRList, ExcelIRIList);

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
                        rowidx + i, _PQIGrade[roadgrade][0], _PQIGrade[roadgrade][1], _PQIGrade[roadgrade][2], _PQIGrade[roadgrade][3]);
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

        private static void WriteAllRoadPQI2Xls_2(MSExcel.Application excelApp, MSExcel.Worksheet destsheet, List<string> xlslist)
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
                    trange = tsheet.get_Range(string.Format("A5:J{0}", userow));
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
                            tidxval[j] = Convert.ToDouble(tobj[i, 4 + j]);
                            sumidxval[j] += tidxval[j] * dmival;
                        }
                        sumlen += dmival;

                        roadtypestr = tobj[i, 10].ToString();
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

        #region 国检

        public static void Convent_Lbi2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Lbi2024(Destxls, prjinfo, prjdir, _RoadPart);
        }

        public static void Convent_Lbi2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff); 
            writeAutoTestXls_Lbi2023(Destxls, prjinfo, prjdir, _RoadPart);
        }

        public static void Convent_Lbi2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Lbi2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart);
        }
        public static void Convent_Lbi2024_1(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Lbi2024_1(Destxls, prjinfo, prjdir, _RoadPart);
        }
        public static void Convent_Lbi2024_HeNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}_{3}_{1}_{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);


            writeAutoTestXls_Lbi2024_HeNan(Destxls, prjinfo, prjdir, _RoadPart);
        }
        public static void Convent_Lbi2024_GanSu(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string excelFileName = "空间定位数据_LBI.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);


            writeAutoTestXls_Lbi2024_GanSu(Destxls, prjinfo, prjdir, _RoadPart);
        }
        public static void Convent_Lbi2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadNum = fileName.Substring(0, 4);
            string strDirc = fileName.Substring(10, 1);
            fileName = string.Format("{0}{1}-GPS-{2}-标准格式-{3}", strRoadNum, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_Lbi2024_HuNan(Destxls, prjinfo, prjdir, _RoadPart);
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
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), "A");
                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }

        private static void writeAutoTestXls_Lbi2024_ChongQing(string fileName,
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
                    string s1 = (smile * 0.001).ToString("f3");

                    if (dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6") );
                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }

        private static void writeAutoTestXls_Lbi2024_1(string fileName,
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

        private static void writeAutoTestXls_Lbi2024_GanSu(string fileName,
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
            int len = RoadPart10M.Count - 1;
            for (int i = 0; i < len; i++)
            {
                if (_Setting.gjLbiOutHight)
                {
                    
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = smile.ToString();

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"),
                                                           double.Parse(dicGps[i]._latitude).ToString("f6"),
                                                           "A", "0", "0");  
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
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"),
                                                            double.Parse(dicGps[i]._latitude).ToString("f6")
                                                        );
                        allDatas.Add(line);
                    }
                }

            }

            double smileLast = prjinfo._EndMile;
            //double emile = RoadPart10M[i + 1].mile;
            string s1Last = (smileLast * 0.001).ToString("f3");

            if (dicGps != null && dicGps != null)
            {
                string line = string.Join(",", s1Last, double.Parse(dicGps[dicGps.Length - 1]._longitude).ToString("f6"),
                                                        double.Parse(dicGps[dicGps.Length - 1]._latitude).ToString("f6")
                                                        , "1");
                allDatas.Add(line);
            }

            File.WriteAllLines(fileName, allDatas);
        }

        private static void writeAutoTestXls_Lbi2024_HeNan(string fileName,
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
            int len = RoadPart10M.Count-1;
            for (int i = 0; i < len; i++)
            {
                if (_Setting.gjLbiOutHight)
                {
                    double smile = RoadPart10M[i].mile;
                    //double emile = RoadPart10M[i + 1].mile;
                    string s1 = smile.ToString();

                    if (dicGps != null && dicGps != null)
                    {
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), double.Parse(dicGps[i]._elevation).ToString("f2"), "A", "0", "0");
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
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), 
                            double.Parse(dicGps[i]._latitude).ToString("f6")
                            , double.Parse(dicGps[i]._elevation).ToString("f2"), "1");
                        allDatas.Add(line);
                    }
                }

            }

            double smileLast = prjinfo._EndMile;
            //double emile = RoadPart10M[i + 1].mile;
            string s1Last = (smileLast * 0.001).ToString("f3");

            if (dicGps != null && dicGps != null)
            {
                string line = string.Join(",", s1Last, double.Parse(dicGps[dicGps.Length-1]._longitude).ToString("f6"),
                    double.Parse(dicGps[dicGps.Length-1]._latitude).ToString("f6")
                    , double.Parse(dicGps[dicGps.Length-1]._elevation).ToString("f2"), "1");
                allDatas.Add(line);
            }

            File.WriteAllLines(fileName, allDatas);
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

        private static void writeAutoTestXls_Lbi2024(string fileName,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M)
        {
            ExcelGPS[] dicGps = null;
            List<string> allDatas = new List<string>();
            allDatas.Add("桩号（Km）,经度,纬度,有效性"); 
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
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), "A");
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
                        string line = string.Join(",", s1, double.Parse(dicGps[i]._longitude).ToString("f6"), double.Parse(dicGps[i]._latitude).ToString("f6"), "A");
                        allDatas.Add(line);
                    }
                }

            }
            File.WriteAllLines(fileName, allDatas);
        }

        public static void Convent_Damage2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, string suff,Encoding encoding)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName;
                tempName = string.Format("{0}-DR-{1}-{2}-{3}", fileName, startMile, endMile, time);
                string Destxls = string.Format(@"{0}\{1}{2}", path, tempName, suff);
                WriteDis_Damage2Xls_gj2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD,encoding);
            }
        }
         
        public static void Convent_Damage2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName,string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
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
                string Destxls = string.Format(@"{0}\{1}{2}", path, tempName,suff);
                WriteDis_Damage2Xls_gj2024(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }


        public static void Convent_Damage2024_AnHui2(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
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
                string Destxls = string.Format(@"{0}\{1}{2}", path, tempName, suff);
                WriteDis_Damage2Xls_gj2024_AnHui(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }

        public static void Convent_Damage2024_AnHui(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

               
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
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName;
                tempName = string.Format("{0}-DR-{1}-{2}-{3}", fileName, startMile, endMile, time);
                
                string Destxls = string.Format(@"{0}\{1}{2}", path, tempName, suff);
                WriteDis_Damage2Xls_gj2024_AnHui(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }



        public static void Convent_Damage2024_GuangDong(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
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
                string Destxls = string.Format(@"{0}\{1}{2}", path, tempName, suff);
                WriteDis_Damage2Xls_gj2024_GuangDong(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }
        public static void Convent_Damage2024_GanSu(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName, string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName;
                if (prjinfo._Direction == 1)
                {

                    tempName = string.Join("-", fileName, "DR", startMile, endMile, time);
                }
                else
                {
                    tempName = string.Join("-", fileName, "DR", endMile, startMile, time);

                }
                string Destxls = string.Format(@"{0}\{1}" + suff, path, tempName);
                WriteDis_Damage2Xls_gj2024(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }

        public static void Convent_Damage2024_HeNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName,string suff)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName;
                if (prjinfo._Direction == 1)
                {
                  
                    tempName = string.Join("_", fileName, "识别结果", startMile, endMile, time, tableNamePart);
                }
                else
                {
                    tempName = string.Format("{0}-DR-{2}-{1}-{3}", fileName, startMile, endMile, time);
                    tempName = string.Join("_", fileName, "识别结果", endMile, startMile , time, tableNamePart);

                }
                string Destxls = string.Format(@"{0}\{1}"+ suff, path, tempName);
                WriteDis_Damage2Xls_gj2024(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);
            }
        }
        public static void Convent_Iri2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);
            writeAutoTestXls_iri2023(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

        }
      
        public static void Convent_Iri2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;

            //文件名
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_iri2024_ChongQing(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
        }

        public static void Convent_Iri2024_HuNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string strRoadName = fileName.Substring(0, 4);
            string strDirc = fileName.Substring(10, 1);
            //文件名
            fileName = string.Format("{0}-{1}-IRI-{2}-{3}", strRoadName, strDirc, startMile, time);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_iri2024_HuNan(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
        }

        public static void Convent_Iri2024_HeNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}_IRI_{1}_{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);
            writeAutoTestXls_iri2024_HeNan(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);

        }

        public static void Convent_Iri2024_GanSu(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string endMile = (prjinfo._EndMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-IRI-{1}-{2}", fileName, startMile,  time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_iri2024_GanSu(Destxls, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _SpeedVal);
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
                    line = string.Join(",", (smile * 0.001).ToString("f3"), value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
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
            int len = roadpart10.Count-1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {

                int smile = roadpart10[i].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {
                    double value = Math.Max(LIRIMeanVal[i], RIRIMeanVal[i]);
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

        private static void writeAutoTestXls_iri2024_ChongQing(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
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
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), RIRIMeanVal[i].ToString("f2"), value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                else
                {//单平整度
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"), "0.00", LIRIMeanVal[i].ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        private static void writeAutoTestXls_iri2024_HeNan(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count-1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {

                int smile = roadpart10[i].mile; 
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {
                    double value = Math.Max(LIRIMeanVal[i], RIRIMeanVal[i]);
                    line = string.Join(",", (smile * 0.001).ToString("f3"),  value.ToString("f2"), (SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                else
                {
                    line = string.Join(",", (smile * 0.001).ToString("f3"), LIRIMeanVal[i].ToString("f2"),(SpeedVal[i] * 1000 / 3600).ToString("f2"));
                }
                datas.Add(line);
            }
            string lineLast = "";
            if (prjinfo._IsDIRIMTD)
            {
                double value = Math.Max(LIRIMeanVal[LIRIMeanVal.Length-1], RIRIMeanVal[RIRIMeanVal.Length-1]);
                lineLast = string.Join(",", (prjinfo._EndMile * 0.001).ToString("f3"), value.ToString("f2"), (SpeedVal[SpeedVal.Length-1] * 1000 / 3600).ToString("f2"));
            }
            else
            {
                lineLast = string.Join(",", (prjinfo._EndMile * 0.001).ToString("f3"), LIRIMeanVal[LIRIMeanVal.Length-1].ToString("f2"), (SpeedVal[SpeedVal.Length-1] * 1000 / 3600).ToString("f2"));
            }
            datas.Add(lineLast);
            File.WriteAllLines(path, datas);
        }
        private static void writeAutoTestXls_iri2024_GanSu(string path, ProjectInfo prjinfo, List<MilePart> roadpart10, double[] LIRIMeanVal, double[] RIRIMeanVal, double[] SpeedVal)
        {
            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.IriTitle);
            for (int i = 0; i < len; i++)
            {

                int smile = roadpart10[i].mile;
                string line = "";
                if (prjinfo._IsDIRIMTD)
                {
                    double value = Math.Max(LIRIMeanVal[i], RIRIMeanVal[i]);
                    line = string.Join(",", (smile * 0.001).ToString("f3"),
                                             Math.Round(LIRIMeanVal[i], 2),
                                             Math.Round(RIRIMeanVal[i],2),
                                             Math.Round(value, 2),
                                             Math.Round((SpeedVal[i] * 1000 / 3600), 1)
                                            );
                }
                else
                {
                    line = string.Join(",", (smile * 0.001).ToString("f3"),
                                            Math.Round(LIRIMeanVal[i], 2),
                                             "0",                                     // 写死了IRI_右
                                             Math.Round(LIRIMeanVal[i], 2),
                                             Math.Round((SpeedVal[i] * 1000 / 3600), 1)
                                             );
                }
                datas.Add(line);
            }
            string lineLast = "";
            if (prjinfo._IsDIRIMTD)
            {
                double value = Math.Max(LIRIMeanVal[LIRIMeanVal.Length - 1], RIRIMeanVal[RIRIMeanVal.Length - 1]);
                lineLast = string.Join(",", (prjinfo._EndMile * 0.001).ToString("f3"),
                                              Math.Round(LIRIMeanVal.Last(), 2),
                                             Math.Round(RIRIMeanVal.Last(), 2),
                                             Math.Round(value, 2),
                                             Math.Round((SpeedVal.Last() * 1000 / 3600), 1)
                                            );
            }
            else
            {
                lineLast = string.Join(",", (prjinfo._EndMile * 0.001).ToString("f3"),
                                             Math.Round(LIRIMeanVal.Last(), 2),
                                             "0",
                                             Math.Round(LIRIMeanVal.Last(), 2),
                                             Math.Round((SpeedVal.Last() * 1000 / 3600), 1)
                                            );
            }
            datas.Add(lineLast);
            File.WriteAllLines(path, datas);
        }




        public static void Convent_LP2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string excelFileName = "高程_LP.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);

            writeAutoTestXls_hightData2023(Destxls, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);

        }

        public static void Convent_LP2024_GanSu(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string excelFileName = "高程_LP.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);

            writeAutoTestXls_hightData2024_GanSu(Destxls, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);
        }

            public static void Convent_LP2024_HeNan(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string excelFileName = "高程_LP.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}_{3}_{1}_{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);

            writeAutoTestXls_hightData2024_HeNan(Destxls, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);

        }

        public static void Convent_LP2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "高程_LP.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            writeAutoTestXls_hightData2024(Destxls, _RoadPartF, _LiriHVal, _RiriHVal, _SpeedVal, prjinfo);

        }
        private static void writeAutoTestXls_hightData2024_HeNan(string path, List<MilePartD> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count-1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.RIFileJgTitle);
            bool hasRdlta = rDeltaHVal != null ? true : false;

            for (int i = 0; i < len; i++)
            {

                double smile = roadpart10[i].mile;
                string line = "";
                line += (smile * 0.001).ToString("f3");
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
            string lastLine = (prjinfo._EndMile * 0.001).ToString("f3") ;
            if (lDeltaHVal != null)
            {
                lastLine += ",";
                lastLine += lDeltaHVal[lDeltaHVal.Length-1];

                if (hasRdlta)
                {  
                    lastLine += ",";
                    lastLine += rDeltaHVal[rDeltaHVal.Length-1];
                }
                else
                {
                   lastLine += ",";
                    lastLine += "0";
                }
                lastLine+= ",";
                lastLine += (speed[speed.Length-1] * 1000 / 3600).ToString("f2");
            }
            else
            {
                lastLine += ",0,0,0";
            }
            datas.Add(lastLine);

            File.WriteAllLines(path, datas);

        }


        private static void writeAutoTestXls_hightData2023(string path, List<MilePartD> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count-1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.RIFileJgTitle);
            bool hasRdlta = rDeltaHVal != null ? true : false;

            for (int i = 0; i < len; i++)
            {

                double smile = roadpart10[i].mile;
                double emile = roadpart10[i + 1].mile;
                string line = "";
                line += (smile * 0.001).ToString("f6");
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


        private static void writeAutoTestXls_hightData2024_GanSu(string path, List<MilePartD> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add(GJTitles.RIFileJgTitle);
            bool hasRdlta = rDeltaHVal != null ? true : false;

            for (int i = 0; i < len; i++)
            {

                double smile = roadpart10[i].mile;
                double emile = roadpart10[i + 1].mile;
                string line = "";
                line += (smile * 0.001).ToString("f6");
                if (lDeltaHVal != null)
                {
                    line += ",";
                    line += lDeltaHVal[i].ToString("f2");  // 甘肃要求数据精确到小数点后两位

                    if (hasRdlta && rDeltaHVal.Length > i)
                    {
                        line += ",";
                        line += rDeltaHVal[i].ToString("f2");  // 甘肃要求数据精确到小数点后两位
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


        private static void writeAutoTestXls_hightData2024(string path, List<MilePartD> roadpart10, double[] lDeltaHVal, double[] rDeltaHVal, double[] speed, ProjectInfo prjinfo)
        {

            int len = roadpart10.Count - 1;
            List<string> datas = new List<string>();
            datas.Add("桩号(km),左高程(mm),右高程(mm),速度(m/s)");
            bool hasRdlta = rDeltaHVal != null ? true : false;

            for (int i = 0; i < len; i++)
            {

                double smile = roadpart10[i].mile;
                double emile = roadpart10[i + 1].mile;
                string line = "";
                line += (smile * 0.001).ToString("f6");
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
        public static void Convent_Rut2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "车辙_RD.txt";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            if (_LRutMaxVal!=null)
            {
                writeAutoTestXls_RD2023(Destxls, prjinfo, prjdir, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, 4);

            }
        }
       
        private static void writeAutoTestXls_RD2023(string path,
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

            if (_LMTDMeanVal != null)
            {
                writeAutoTestXls_Mpd2023(Destxls, prjinfo, prjdir, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, 4);

            }
        }

        private static void writeAutoTestXls_Mpd2024_ChongQing(string path,
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
                string s1 = (smile * 0.001).ToString("f2");
                string s2 = LMTDVal[i].ToString("f2");
                string s3 = CMTDVal[i].ToString("f2");
                string s4 = RMTDVal[i].ToString("f2");
                line = string.Join(",", s1, s2, s3, s4);
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        public static void Convent_Mpd2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;

            fileName = string.Format("{0}-MPD-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);

            if (_LMTDMeanVal!=null)
            {
                writeAutoTestXls_Mpd2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, 4);
            }
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
                line = string.Join(",", s1, LMTDVal[i], CMTDVal[i], RMTDVal[i]);
                datas.Add(line);
            }
            File.WriteAllLines(path, datas);
        }

        public static void Convent_Bump2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName, string suff)
        {
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;

            fileName = string.Format("{0}-PB-{1}-{2}", fileName, startMile, time);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName, suff);
            writeAutoTestXls_Bump2024_ChongQing(Destxls, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5);
        }
        private static void writeAutoTestXls_Bump2024_ChongQing(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;
            datas.Add("桩号(km),	PB_L,PB_M,PB_H,ΔH(cm)");
            int len = RoadPart10M.Count - 1;


            for (int i = 0; i < len; i++)
            {
                string line = "";

                double smile = RoadPart10M[i].mile;
                //double emile = RoadPart10M[i + 1].mile;
                string s1 = (smile * 0.001).ToString("f2");

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


        public static void Convent_Bump2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "跳车_PB.txt";

            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);
            writeAutoTestXls_Bump2023(Destxls, prjinfo, prjdir, _RoadPart, _PBIVal, _LDeltaHVal, _RDeltaHVal, 5);
        }
        private static void writeAutoTestXls_Bump2023(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;
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

        private static void writeAutoTestXls_Bump2024(string path,
ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> RoadPart10M, int[][] PBIVal, double[] LDeltaHVal, double[] RDeltaHVal, int lastCol)
        {
            List<string> datas = new List<string>();
            if (!prjinfo._IsRut)
                return;
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
        public static void Convent_TP2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, "TP");
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            //string Destxls = string.Format(@"{0}\{1}_IRIMTD_{2}m.xlsx", path, prjdir.Name, disval);

            writeAutoTestXls_gc_gj2024_ChongQing(Destxls, _RoadPartF, prjdir, prjinfo, _SpeedVal);

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
                object[,] disvalsn = new object[len, 22];
                List<string> datas = new List<string>() { "桩号(km),高程值1(0.1mm),高程值2(0.1mm),高程值3(0.1mm),高程值4(0.1mm),高程值5(0.1mm),高程值6(0.1mm)," +
                    "高程值7(0.1mm),高程值8(0.1mm),高程值9(0.1mm),高程值10(0.1mm),高程值11(0.1mm),高程值12(0.1mm)," +
                    "高程值13(0.1mm),高程值14(0.1mm),高程值15(0.1mm),高程值16(0.1mm),高程值17(0.1mm)," +
                    "高程值18(0.1mm),高程值19(0.1mm),高程值20(0.1mm),速度(m/s）\r\n" };

                MSExcel.Range destrange;
                for (int i = 0; i < len; i++)
                {
                    string dataLine = "";
                    int colcnt = 0;
                    double smile = roadpart10[i].mile;
                    double emile = roadpart10[i + 1].mile;
                    disvalsn[i, colcnt++] = smile * 0.001;
                    //这个n=100可能有问题
                    try
                    {
                        for (int n = 100; n < allData[i].Length && n < 2100; n += 100)
                        {
                            dataLine += (int)(allData[i][n] * 10) + ",";
                            // disvalsn[i, colcnt++] = (int)(allData[i][n] * 10); //原始高程单位 mm

                        }
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                    //disvalsn[i, colcnt++] = speedVal[i]*1000/3600;
                    dataLine += speedVal[i] * 1000 / 3600;
                    datas.Add(dataLine);
                }

                File.WriteAllLines(filePath, datas);

            }
            catch (System.Exception ex)
            {
                MessageBox.Show("原始高程表导出失败，可能是车辙计算数据缺失！\n" + ex.Message);
            }

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

        public static void Convent_TT2023(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "——纹理_TT.xlsx";
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\国检转换格式2023\{1}",
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
            writeTTData2023(_Worksheet_, prjdir, prjinfo);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// 断面高程  平整度
        /// </summary>

        private static void writeTTData2023(MSExcel.Worksheet workSheet, DirectoryInfo prjdir, ProjectInfo prjinfo)
        {

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
            MSExcel.Range destrange;
            double smile = prjinfo._StartMile * 0.001;
            for (int i = 0; i < len; i++)
            {
                int colcnt = 0;

                disvalsn[i, colcnt++] = smile;
                smile += 0.001;
                disvalsn[i, colcnt++] = values[0][i];
                disvalsn[i, colcnt++] = values[2][i];
                disvalsn[i, colcnt++] = values[1][i];
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
                destrange = workSheet.get_Range(String.Format("A2:C{0}", len + 1));
                MSExcel.Range sortrange = workSheet.get_Range(String.Format("A2:A{0}", len + 1));//按桩号排序

                GlobalExcel.ReflectionColnum(workSheet, destrange, sortrange);
            }
        }


        public static void Convent_Damage2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, string fileName)
        {
            string[] roadTypeSplit = File.ReadAllLines(prjdir.FullName + "\\RoadTypeInfo.txt");

            foreach (string roadTypeInfo in roadTypeSplit)
            {
                double startMileD = double.Parse(roadTypeInfo.Split('-').FirstOrDefault());
                double endMileStrD = double.Parse(roadTypeInfo.Split('k').FirstOrDefault().Split('-').LastOrDefault());

                string tableNamePart = "";
                int i = 0;
                if (roadTypeInfo.Contains("沥青"))
                {
                    tableNamePart = "指南沥青";
                    i = 0;
                }
                else if (roadTypeInfo.Contains("水泥"))
                {
                    tableNamePart = "指南水泥";
                    i = 1;
                }
                else
                {
                    tableNamePart = "指南砂石";
                    i = 2;
                }
                //拼接文件名称
                string startMile = startMileD.ToString("f3");
                string endMile = endMileStrD.ToString("f3");
                //string time = prjinfo._DataDate + prjinfo._DataTime;
                string tempName;
                if (prjinfo._Direction == 1)
                {
                    tempName = string.Format("{0}-DR-{1}-{2}-{3}-{4}", fileName, startMile, endMile, prjinfo._DataDate, tableNamePart);
                }
                else
                {
                    tempName = string.Format("{0}-DR-{2}-{1}-{3}-{4}", fileName, startMile, endMile, prjinfo._DataDate, tableNamePart);
                }

                string Destxls = string.Format(@"{0}\{1}.txt", path, tempName);


                WriteDis_Damage2Xls_gj2023(Destxls, prjinfo, prjdir, _RoadPart, _RoadDisList, i, startMileD, endMileStrD);


            }
        }
        public static void Convent_GdLP2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName,string suff)
        {
            string excelFileName = "a.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}{2}", path, fileName,suff);

            //找到惯导文件 
            string filePath = prjinfo._PrjPath + "\\IRIMTD\\DAQ0\\acc1mmDatas.txt";


            if (File.Exists(filePath))
            {
                var allTxt = File.ReadAllLines(filePath).ToList();
                List<string> newTexts = allTxt
.Select(line =>
{
    string[] columns = line.Split(',');
    if (columns.Length >= 1 && double.TryParse(columns[0], out double value0))
    {
        columns[0] = value0.ToString("f6");
    }
    if (columns.Length >= 2 && double.TryParse(columns[1], out double value1)) 
    {
        columns[1] = (value1 * 0.001).ToString("f6"); // 第二列单位转换为m
    }
    if (columns.Length >= 3 && double.TryParse(columns[2], out double value2))
    {
        columns[2] = value2.ToString("f5");
    }
    if (columns.Length >= 4 && double.TryParse(columns[3], out double value3))
    {
        columns[3] = value3.ToString("f5");
    }
    if (columns.Length >= 5 && double.TryParse(columns[4], out double value4))
    {
        columns[4] = value4.ToString("f5");
    }
    return string.Join(",", columns);
})
.ToList();
                newTexts.Insert(0, GJTitles.RIFileGdTitle);
                File.WriteAllLines(Destxls, newTexts);
            }
            else
            {
                MessageBox.Show("工程" + prjinfo._RoadName + "请进行惯导平整度计算!");
            }


        }

        public static void Convent_GdLP2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, string fileName)
        {
            string excelFileName = "LP.xlsx";

            //拼接文件名称
            string startMile = (prjinfo._StartMile * 0.001).ToString("f3");
            string time = prjinfo._DataDate + prjinfo._DataTime;
            string excelName = excelFileName.Split('.')[0].Split('_').Last();
            fileName = string.Format("{0}-{3}-{1}-{2}", fileName, startMile, time, excelName);
            string Destxls = string.Format(@"{0}\{1}.txt", path, fileName);

            //找到惯导文件 
            string filePath = prjinfo._PrjPath + "\\IRIMTD\\DAQ0\\acc1mmDatas.txt";


            if (File.Exists(filePath))
            {
                var allTxt = File.ReadAllLines(filePath).ToList();
                allTxt.Insert(0, "时长,桩号(km),左加速度(m/s2),右加速度(m/s2),速度(m/s)");
                File.WriteAllLines(Destxls, allTxt);
            }
            else
            {
                MessageBox.Show("工程" + prjinfo._RoadName + "请进行惯导平整度计算!");
            }


        }
        private static void WriteDis_Damage2Xls_gj2023(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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
                if (roadpart[i].roadtype == roadType)
                {

                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        line += ",";
                        double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                        line += area == 0 ? area.ToString() : area.ToString("f2");
                         
                        // _Worksheet.Cells[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                    }

                }
                allDatas.Add(line);
            }
            File.WriteAllLines(path, allDatas);
        }

        private static void WriteDis_Damage2Xls_gj2024_ChongQing(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile, Encoding encoding)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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
                if (roadpart[i].roadtype == roadType)
                {

                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表
                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        switch (roadpart[i].roadtype)
                        {
                            case 0:
                                if (dis == 1 )
                                {
                                    line += ",";
                                    double area =  RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                    dis++;
                                }
                              else  if (dis == 3)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                    dis++;
                                }
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                }
                                break;
                            case 1:
                                if (dis == 1)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                    dis++;
                                }
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                }
                                break;
                            case 2:
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString("f3") : area.ToString("f3");
                                }
                                break;
                            default:
                                break;
                        } 
                    }

                }
                //桩号(km),识别宽度(m),破损率DR(%),网裂(m2),纵向裂缝(m),横向裂缝(m),坑槽(m2),松散(m2)

                //桩号(km),识别宽度(m),破损率DR(%),纵向裂缝(m),横向裂缝(m),网裂(m2),松散(m2),坑槽(m2)
                if (roadType == 0)
                {
                    string[] strSplit =line.Split(',');

                    // 按照新的列顺序重新组合数据
                      line= $"{strSplit[0]},{strSplit[1]},{strSplit[2]},{strSplit[4]},{strSplit[5]},{strSplit[3]},{strSplit[7]},{strSplit[6]}";

                }
                //桩号(km),识别宽度(m),破损率DR(%),破碎板(m2),裂缝(m),错台(m2),拱起(m2),坑洞(m2),露骨(m2)
                //桩号(km),识别宽度(m),破损率DR(%),破碎板(m2),裂缝(m),坑洞(m2),露骨(m2),错台(m2),拱起(m2)
                else if (roadType == 1)
                {

                    string[] strSplit = line.Split(',');

                    // 按照新的列顺序重新组合数据
                    line = $"{strSplit[0]},{strSplit[1]},{strSplit[2]},{strSplit[3]},{strSplit[4]},{strSplit[7]},{strSplit[8]},{strSplit[5]},{strSplit[6]}";
                }
               // 桩号(km),识别宽度(m),破损率DR(%),沉陷(m2),波浪搓扳(m2),车辙(m2),坑槽(m2)
                //桩号(km),识别宽度(m),破损率DR(%),坑槽(m2),沉陷(m2),车辙(m2),波浪搓扳(m2)
                else
                {
                    string[] strSplit = line.Split(',');

                    // 按照新的列顺序重新组合数据
                    line = $"{strSplit[0]},{strSplit[1]},{strSplit[2]},{strSplit[6]},{strSplit[3]},{strSplit[5]},{strSplit[4]}";

                }


                allDatas.Add(line);
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

        private static void WriteDis_Damage2Xls_gj2024(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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

                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        switch (roadpart[i].roadtype)
                        {
                            case 0:
                                if (dis == 0 )
                                {
                                    line += ",";
                                    double area =  RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength;
                                    line += area.ToString("f2");
                                    dis++;
                                }
                                else if ( dis == 2)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength;
                                    line += area.ToString("f2");
                                    dis++;
                                }
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area.ToString("f2");
                                }
                                break;
                            case 1:
                                if (dis == 1)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis + 1].totallength;
                                    line += area.ToString("f2");
                                    dis++;
                                }
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area.ToString("f2");
                                }
                                break;
                            case 2:
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area.ToString("f2");
                                }

                                break;
                            default:
                                break;
                        }


                    }

                    if (roadpart[i].roadtype == 0 )
                    {
                        line =  GlobalExcel.MoveElement(line, 5, 3); 
                    }

                    if (roadpart[i].roadtype == 1)
                    {
                        line = GlobalExcel.MoveElement(line, 7, 5);
                        line = GlobalExcel.MoveElement(line, 8, 6);
                    }

                }

               

                allDatas.Add(line);
            }
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, GlobalExcel.getConvert_GJ_TempAreaStr(6)));

            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, GlobalExcel.getConvert_GJ_TempAreaStr(7)));

            }
            else
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, GlobalExcel.getConvert_GJ_TempAreaStr(5)));

            }
            File.WriteAllLines(path, allDatas);
        }


        private static void WriteDis_Damage2Xls_gj2024_AnHui(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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
                if (roadpart[i].roadtype == roadType)
                {

                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    string F2_Except0(double area)
                    {   // 保留0， 而不会被f2解析成 0.00
                        return area == 0 ? area.ToString() : area.ToString("f2");
                    }
                    var sn_temp_arr = new double[5];
                    int ind = 0;
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        switch (roadpart[i].roadtype)
                        {
                            case 0:
                                if (dis == 0 || dis == 2)
                                {
                                    
                                    double area =  RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    sn_temp_arr[ind] += area;
                                    dis++;
                                }
                                else
                                {
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    sn_temp_arr[ind] += area;
                                }
                                break;
                            case 1:
                                if (dis == 1)
                                {
                                    line += ",";
                                    double area =  RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                    dis++;
                                }  
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }
                                break;
                            case 2:
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }

                                break;
                            default:
                                break;
                        }
                        ind += 1;

                    }
                    if(roadType == 0)
                    {
                        
                        var sn_string_arr = new string[5];
                        for(int t=0; t<sn_temp_arr.Length; t++)
                        {
                            sn_string_arr[t] = F2_Except0(sn_temp_arr[t]);
                        }
                        line += "," + string.Join(",", sn_string_arr);
                    }

                }
                allDatas.Add(line);
            }
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0"));

            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0", "0"));

            }
            else
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0"));

            }
            File.WriteAllLines(path, allDatas);
        }
        private static void WriteDis_Damage2Xls_gj2024_GuangDong(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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
                if (roadpart[i].roadtype == roadType)
                {

                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    string ct = "", gq = "", lg = "", kd = "";
                    string F2_Except0(double area)
                    {   // 保留0， 而不会被f2解析成 0.00
                        return area == 0 ? area.ToString() : area.ToString("f2");
                    }
                    var sn_temp_arr = new double[5];
                    int ind = 0;
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        switch (roadpart[i].roadtype)
                        {
                            case 0:
                                if (dis == 1 || dis == 3)
                                {

                                    double area =   RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    sn_temp_arr[ind] += area;
                                    dis++;
                                }
                                else
                                {
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    sn_temp_arr[ind] += area;
                                }
                                break;
                            case 1:
                                if (dis == 1)
                                {
                                    line += ",";
                                    double area =  RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                    dis++;
                                }
                                else if (dis == 3)
                                {
                                    ct = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                }
                                else if (dis == 4)
                                    gq = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                else if (dis == 5)
                                    kd = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                else if (dis == 6)
                                {
                                    lg = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                    string values = string.Join(",",
                                                kd, lg, ct, gq
                                            );
                                    line += "," + values;
                                }


                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }
                                break;
                            case 2:
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }

                                break;
                            default:
                                break;
                        }
                        ind += 1;

                    }
                    if (roadType == 0)
                    {
                        // 需求： 纵裂、横裂、网裂、松散、坑槽
                        // 原有：网裂、纵裂、横裂、坑槽、松散
                        void swap(int a, int b)
                        {
                            (sn_temp_arr[a], sn_temp_arr[b]) = (sn_temp_arr[b], sn_temp_arr[a]);
                        }
                        swap(1, 0); swap(2, 1); swap(3, 4);
                        var sn_string_arr = new string[5];
                        for (int t = 0; t < sn_temp_arr.Length; t++)
                        {
                            sn_string_arr[t] = F2_Except0(sn_temp_arr[t]);
                        }
                        line += "," + string.Join(",", sn_string_arr);
                    }

                }
                allDatas.Add(line);
            }
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0"));

            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0", "0"));

            }
            else
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0"));

            }
            File.WriteAllLines(path, allDatas);
        }

        private static void WriteDis_Damage2Xls_gj2024_AnHui2(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int roadType, double sMile, double eMile)
        {
            List<string> allDatas = new List<string>();
            if (roadType == 0)
            {
                allDatas.Add(GJTitles.LQ_NC_BIG);
            }
            else if (roadType == 1)
            {
                allDatas.Add(GJTitles.SN_NC_BIG);
            }
            else
            {
                allDatas.Add(GJTitles.SS_NC_BIG);
            }
            string errlog = prjdir.FullName + "\\errlog.txt";

            int len = roadpart.Count - 1,
                dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
            bool has = false;
            // int rowCount = 2;
            int colcnt = 1;
            sMile *= 1000;
            eMile *= 1000;
            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                string line = "";
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
                //统计位于这个区域的病害  cwb
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
                if (roadpart[i].roadtype == roadType)
                {

                    //有病害
                    has = true;
                    smile = roadpart[i].mile;
                    emile = roadpart[i + 1].mile;
                    int milelength = Math.Abs(smile - emile);

                    //病害汇总表

                    line = string.Join(",", (smile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2"));
                    //_Worksheet.Cells[rowCount, colcnt++] = (smile * 0.001).ToString("f3");
                    // _Worksheet.Cells[rowCount, colcnt++] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength).ToString("f2");
                    string ct = "", gq = "", lg = "", kd = "";
                    string F2_Except0(double area)
                    {   // 保留0， 而不会被f2解析成 0.00
                        return area == 0 ? area.ToString() : area.ToString("f2");
                    }
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        switch (roadpart[i].roadtype)
                        {
                            case 0:
                                if (dis == 1 || dis == 3)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea
                                        + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                    dis++;
                                }
                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }
                                break;
                            case 1:
                                if (dis == 1)
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea
                                        + RoadDiseaseTypes.roaddis[roadType][dis + 1].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                    dis++;
                                }
                                else if (dis == 3)
                                {
                                    ct = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                }
                                else if (dis == 4)
                                    gq = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                else if (dis == 5)
                                    kd = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                else if (dis == 6)
                                {
                                    lg = F2_Except0(RoadDiseaseTypes.roaddis[roadType][dis].totalarea);
                                    string values = string.Join(",",
                                                kd, lg, ct, gq
                                            );
                                    line += "," + values;
                                }


                                else
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }
                                break;
                            case 2:
                                {
                                    line += ",";
                                    double area = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                                    line += area == 0 ? area.ToString() : area.ToString("f2");
                                }

                                break;
                            default:
                                break;
                        }


                    }

                }
                allDatas.Add(line);
            }
            if (roadType == 0)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0"));

            }
            else if (roadType == 1)
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0", "0", "0"));

            }
            else
            {
                allDatas.Add(string.Join(",", (eMile * 0.001).ToString("f3"), _RoadConfig.DetectWidth, "0.00", "0", "0", "0", "0"));

            }
            File.WriteAllLines(path, allDatas);
        }



        #endregion
        #region 合肥

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
            }
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
                                if (preValue[15] !=null )
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
                              else  if ((preLen * 1000) <= splitValue)
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
                                   
                                    if (nowValue[15] != null)
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
                                if (resultValues.Count>0)
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
                                    if (nowValue[15]==null)
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

        #region 四川
        private static void WriteAll2Xls_ChongQing(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart,
 Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
 double[] LMTDVal, double[] RMTDVal, double[] MMTDVal, int[][] PBVal, double[] deltahVal, int disval, ExcelGPS[] gpsInfo)
        {
            bool shangxing = prjinfo._Direction == 1 ? true : false;
            //检查区间长度进行处理
            int len = roadpart.Count-1, dlen = arrdis.Length;
            double irival = 0, tpcival = 0;

            object[,] mxlist = new object[len, 36];
            double[] drvals = new double[len];

          
            string errlog = prjdir.FullName + "\\errlog.txt";

            int typeidx = 0;
            bool res = false;

            for (int i = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile; 

                double drval;
                int rowCount = 0;
                int milelength = Math.Abs(smile - emile);
                mxlist[i, rowCount++] = prjinfo._AreaCode;
                mxlist[i, rowCount++] = prjinfo._RoadCodePart;
                mxlist[i, rowCount++] = prjinfo._RoadName;
                mxlist[i, rowCount++] = prjinfo._City;
                mxlist[i, rowCount++] = prjinfo._District;

                double s1 = roadpart[i].mile / 1000.0;
                double s2 = roadpart[i + 1].mile / 1000.0;
                double smile1 = Math.Round(s1, 3);
                double emile1 = Math.Round(s2, 3);
                double milelength1 = Math.Round(Math.Abs(s1 - s2), 3);
                if (prjinfo._Direction > 0)
                {
                    mxlist[i, rowCount++] = smile1;
                    mxlist[i, rowCount++] = emile1;
                }
                else
                {
                    mxlist[i, rowCount++] = emile1;
                    mxlist[i, rowCount++] = smile1; 
                }
                mxlist[i, rowCount++] = milelength1;
                mxlist[i, rowCount++] = roadpart[i].degreestr;
                mxlist[i, rowCount++] = roadpart[i].roadtype == 0 ? "沥青" : roadpart[i].roadtype == 1 ? "水泥" : "砂石";
                mxlist[i, rowCount++] = prjinfo._Direction == 1 ? "上行" : "下行";

                //破损面积（平方米）
                mxlist[i, rowCount++] = String.Format("=sum(V{0}:AD{0})", i + 3);

                //检测面积（平方米）
                mxlist[i, rowCount++] = _RoadConfig.DetectWidth*milelength;

                 
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
                mxlist[i, rowCount++] = drval; 
                mxlist[i, rowCount++] = string.Format("=100-{0}*POWER(N{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]); //pci
                mxlist[i, rowCount++] = string.Format("=IF(O{0}>={1},\"优\",IF(O{0}>={2},\"良\",IF(O{0}>={3},\"中\",IF(O{0}>={4},\"次\",\"差\"))))",
                   i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);  //pci评价

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

                //rqi
                mxlist[i, rowCount++] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{2})),5)", _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], irival);
                mxlist[i, rowCount++] = string.Format("=IF(Q{0}>={1},\"优\",IF(Q{0}>={2},\"良\",IF(Q{0}>={3},\"中\",IF(Q{0}>={4},\"次\",\"差\"))))",
                    i + 3,
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][2],
                    _RQIGrade[roadpart[i].roaddegree][roadpart[i].roadtype][3]);

                mxlist[i, rowCount++] = string.Format("=ROUND(({1}*O{0}+{2}*Q{0})/({1}+{2}),5)",
                i + 3,
                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                //PQI 评价
                mxlist[i, rowCount++] = string.Format("=IF(S{0}>={1},\"优\",IF(S{0}>={2},\"良\",IF(S{0}>={3},\"中\",IF(S{0}>={4},\"次\",\"差\"))))",
                i + 3,
                _PQIGrade[roadpart[i].roaddegree][0],
                _PQIGrade[roadpart[i].roaddegree][1],
                _PQIGrade[roadpart[i].roaddegree][2],
                _PQIGrade[roadpart[i].roaddegree][3]);
                mxlist[i, rowCount++] = irival;

                int roadType = roadpart[i].roadtype;
                if (roadType == 0 )
                {
                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        
                        {
                            mxlist[i, rowCount++] = RoadDiseaseTypes.roaddis[roadType][dis] .totalarea;
                        }

                        // _Worksheet.Cells[rowCount, colcnt++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea.ToString("f2");
                    }
                    for (int ttt = 0; ttt < 7; ttt++)
                    {
                        mxlist[i, rowCount++] = "0";
                    }
                }
                else
                {
                    for (int ttt = 0; ttt < 7; ttt++)
                    {
                        mxlist[i, rowCount++] = "0";
                    }

                    for (int dis = 0; dis < RoadDiseaseTypes.roaddis[roadType].Length; dis++)
                    {
                        mxlist[i, rowCount++] = RoadDiseaseTypes.roaddis[roadType][dis].totalarea;
                    } 
                }
            } 
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:AJ{0}", len + 2));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);
             if (_Setting.IsExcelSort)
            {
                MSExcel.Range sortrange = worksheet.get_Range(String.Format("F3:F{0}", len + 2));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
        }
        #endregion


        #region 上海惠普
        public static void OutputDis_HPcsv_0(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
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
            string srcxls = string.Format(@"{0}\报表模板\低等级农村公路\上海惠浦\X636310115X2黑色.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-黑色-{6}.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime, disval);
            // string destxls = string.Format(@"{0}\{1}_水泥混凝土路面破损{2}汇总.csv", path, prjdir.Name, subdname);
            _WorkbookSN = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            _Worksheet_snhz = _WorkbookSN.Sheets[string.Format("Sheet1")] as MSExcel.Worksheet;

            _WorkbookSN.SaveAs(destxls, MSExcel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);



            bool Haslqflag = false;
            WriteDisHZ2Xls_modle10(_Worksheet_snhz, prjinfo, prjdir, _RoadPart1M, _RoadDisList, ref Haslqflag, 53, prjinfo._RoadImgDis, "Road");

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
       ref bool Haslqflag, int borderType, int ImgDis, string ImgType)
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
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;

                double smile = RoadPart1M[i].mile;
                double emile = RoadPart1M[i + 1].mile;
                //int emile = RoadPart1M[i + 1].mile;
                //int milelength = Math.Abs(smile - emile);
                disvalsn[i, colcnt++] = prjinfo._RoadCode;//路线代码
                disvalsn[i, colcnt++] = prjinfo._DataDate.Substring(0, 4) + "/" + prjinfo._DataDate.Substring(4, 2) + "/" + prjinfo._DataDate.Substring(6, 2) + " " + prjinfo._DataTime.Substring(0, 2) + ":" + prjinfo._DataTime.Substring(2, 2);
                disvalsn[i, colcnt++] = prjinfo._RoadName;
                disvalsn[i, colcnt++] = prjinfo._DataPerson;
                disvalsn[i, colcnt++] = "dis";
                disvalsn[i, colcnt++] = "XR-M";
                disvalsn[i, colcnt++] = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
                disvalsn[i, colcnt++] = "Start";
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
                            string name = RoadDiseaseTypes.roaddis[RoadPart1M[i].roadtype][typeidx].disname;
                            int roadType = RoadPart1M[i].roadtype;
                            switch (roadType)
                            {
                                case 0:
                                    if (name.Contains("网裂"))
                                    {
                                        name = "01" + name;
                                        break;
                                    }
                                    if (name.Contains("纵向裂缝"))
                                    {
                                        name = "02" + name;
                                        break;
                                    }
                                    if (name.Contains("横向裂缝"))
                                    {
                                        name = "03" + name;
                                        break;
                                    }
                                    if (name.Contains("坑槽"))
                                    {
                                        name = "04" + name;
                                        break;
                                    }
                                    if (name.Contains("松散"))
                                    {
                                        name = "05" + name;
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
                                    if (name.Contains("错台"))
                                    {
                                        name = "23" + name;
                                        break;
                                    }
                                    if (name.Contains("拱起"))
                                    {
                                        name = "24" + name;
                                        break;
                                    }
                                    if (name.Contains("拱起"))
                                    {
                                        name = "25" + name;
                                        break;
                                    }
                                    if (name.Contains("露骨"))
                                    {
                                        name = "26" + name;
                                        break;
                                    }
                                     
                                    break;
                                case 2:
                                    if (name.Contains("沉陷"))
                                    {
                                        name = "31" + name;
                                        break;
                                    }
                                    if (name.Contains("波浪搓板"))
                                    {
                                        name = "32" + name;
                                        break;
                                    }
                                    if (name.Contains("车辙"))
                                    {
                                        name = "33" + name;
                                        break;
                                    }
                                    if (name.Contains("坑槽"))
                                    {
                                        name = "34" + name;
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
        #endregion
    }
}
