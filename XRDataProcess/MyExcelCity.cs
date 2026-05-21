using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms; 
using System.Threading;
using Org.BouncyCastle.Asn1.Sec;
using DevExpress.XtraEditors.TextEditController;
using DevExpress.XtraPrinting.Export.Pdf;
using Microsoft.Office.Interop.Excel;

namespace XRDataProcess
{
    /// <summary>
    /// 城镇道路报表，CJJ 36-2016 城镇道路养护技术规范
    /// </summary>
    class MyExcelCity
    {
        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public static double[][] _RQIGrade;//道路等级 等级区间
        private static double[][] _RDIGrade;
        public static double[][] _MTDGrade;
        private static double[][] _IRIGrade;
        public static double[][] _PCIGrade;
        public static double[][] _PQIGrade;
        private static double[][] _PWIGrade;
        private static double[][] _RDIRD;
        private static double[] _RDIa;
        private static double[] _PWIa;

        /// <summary>
        /// 0-PCI系数，1-RQI系数
        /// </summary>
        private static double[][] _PQIW;
        private static double _PQIT;

        private static double[][] _WeightParm;//0-沥青，1-水泥
        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "快速路", "主干路", "次干路", "支路" };

        private static Dictionary<string, int> _RoadGradeDict;
        private static Dictionary<string, int>[] _DisType;

        private static ExcelGPS[] _GPSInfo = null;
        private static double[] _RQIa;
        private static List<MilePart> _RoadPart = null;
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
        private static double[] _SRutDisValL = null;
        private static double[] _SRutDisValR = null;
        private static int[] _SRutDisMile = null;
        private static double[][] _SRutCountMeanVal = null;
        private static double[] _rutThresh = new double[1];
        private static string[] _MarkVal = null;
        private static double[] _SpeedVal = null;

        private static void InitXlsParm()
        {
            int len = _RoadGradeStr.Length;
            _RQIGrade = new double[len][];
            _RDIGrade = new double[len][];
            _MTDGrade = new double[len][];
            _IRIGrade = new double[len][];
            _PCIGrade = new double[len][];
            _PQIGrade = new double[len][];
            _PQIW = new double[len][];
            _PWIGrade = new double[len][];

            for (int i = 0; i < 4; i++)
            {
                _RQIGrade[i] = new double[4];
                _RDIGrade[i] = new double[5];
                _PWIGrade[i] = new double[5];
                _MTDGrade[i] = new double[4];
                _IRIGrade[i] = new double[5];
                _PCIGrade[i] = new double[4];
                _PQIGrade[i] = new double[4];
                _PQIW[i] = new double[2];
            }

            _RoadSocre = new Dictionary<string, CityRoadDis>[2];
            for (int i = 0; i < 2; i++)
            {
                _RoadSocre[i] = new Dictionary<string, CityRoadDis>();
            }

            _WeightParm = new double[2][];

            _DisType = new Dictionary<string, int>[2];
            for (int i = 0; i < 2; i++)
            {
                _DisType[i] = new Dictionary<string, int>();
            }
            _DisType[0].Add("裂缝类", 0);
            _DisType[0].Add("变形类", 1);
            _DisType[0].Add("松散类", 2);
            _DisType[0].Add("其他类", 3);
            _DisType[1].Add("裂缝类", 0);
            _DisType[1].Add("接缝料破坏类", 1);
            _DisType[1].Add("表面破坏类", 2);
            _DisType[1].Add("其他类", 3);

            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }

            _PWIa = new double[2];
            _RQIa = new double[2];
            _RDIa = new double[2];
            _RDIRD = new double[2][];
            for (int i = 0; i < 2; i++)
            {
                _RDIRD[i] = new double[2];
            }
        }

        public static void LoadXlsParm()
        {
            InitXlsParm();

            XmlDocument Doc = new XmlDocument();
            Doc = new XmlDocument();
            XmlElement Elem;
            XmlNodeList xmlNodes;
            Doc.Load(System.Windows.Forms.Application.StartupPath + "\\ParaVal.xml");    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点  
            xmlNodes = Elem.ChildNodes;

            //读取病害类型
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
                                    roaddis._DisType = Convert.ToInt32(((XmlElement)node).GetAttribute("损坏类型"));
                                    string dismidu = ((XmlElement)node).GetAttribute("损坏密度");
                                    string disscore = ((XmlElement)node).GetAttribute("单项扣分值");
                                    string[] smidu = dismidu.Split(' ');
                                    string[] sscore = disscore.Split(' ');
                                    int len = smidu.Length > sscore.Length ? smidu.Length : sscore.Length;
                                    roaddis._MiduScore = new double[2][];//损坏密度，扣分值
                                    for (int j = 0; j < 2; j++)
                                    {
                                        roaddis._MiduScore[j] = new double[len];
                                    }
                                    for (int k = 0; k < len; k++)
                                    {
                                        if (k < smidu.Length)
                                        {
                                            roaddis._MiduScore[0][k] = Convert.ToDouble(smidu[k]) * 0.01;
                                        }
                                        else
                                        {
                                            roaddis._MiduScore[0][k] = Convert.ToDouble(smidu[smidu.Length - 1]) * 0.01;
                                        }
                                        roaddis._MiduScore[1][k] = Convert.ToDouble(sscore[k]);
                                    }

                                    _RoadSocre[i].Add(roaddis._DisName, roaddis);
                                }
                            }
                            if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面权函数曲线")
                            {
                                string str = ((XmlElement)subnode).GetAttribute("Wi");
                                string[] s = str.Split(' ');
                                _WeightParm[i] = new double[s.Length + 1];
                                for (int j = 0; j < s.Length; j++)
                                {
                                    _WeightParm[i][j] = Convert.ToDouble(s[j]);
                                }
                                _WeightParm[i][s.Length] = 0;
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
                            //读取计算RQI的系数
                            else if (i == 0 && subnode.Name == "RQI系数")
                            {
                                _RQIa[0] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a0"));
                                _RQIa[1] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a1"));
                            }
                            //读取计算PWI的系数
                            else if (i == 0 && subnode.Name == "PWI系数")
                            {
                                _PWIa[0] = double.Parse(((XmlElement)subnode).GetAttribute("a0"));
                                _PWIa[1] = double.Parse(((XmlElement)subnode).GetAttribute("a1"));
                            }
                        }
                    }
                }
            }

            //读取等级区间
            for (int i = 0; i < 4; i++)
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
                                    string strval = ((XmlElement)node).GetAttribute("等级区间");
                                    string[] s = strval.Split(' ');
                                    double[] val = new double[s.Length];
                                    for (int j = 0; j < s.Length; j++)
                                    {
                                        val[j] = Convert.ToDouble(s[j]);
                                    }
                                    if (node.Name == "RQI")
                                    {
                                        val.CopyTo(_RQIGrade[i], 0);
                                    }
                                    else if (node.Name == "RDI")
                                    {
                                        val.CopyTo(_RDIGrade[i], 0);
                                    }
                                    else if (node.Name == "PWI")
                                    {
                                        val.CopyTo(_PWIGrade[i], 0);
                                    }
                                    else if (node.Name == "MTD")
                                    {
                                        val.CopyTo(_MTDGrade[i], 0);
                                    }
                                    else if (node.Name == "IRI")
                                    {
                                        val.CopyTo(_IRIGrade[i], 0);
                                    }
                                    else if (node.Name == "PCI")
                                    {
                                        val.CopyTo(_PCIGrade[i], 0);
                                    }
                                    else if (node.Name == "PQI")
                                    {
                                        val.CopyTo(_PQIGrade[i], 0);
                                        _PQIW[i][0] = Convert.ToDouble(((XmlElement)node).GetAttribute("w2"));
                                        _PQIW[i][1] = Convert.ToDouble(((XmlElement)node).GetAttribute("w1"));
                                        _PQIT = Convert.ToDouble(((XmlElement)node).GetAttribute("T"));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private class CityRoadDis
        {
            public int _DisType = -1;
            public string _DisName = null;
            public double _UseWidth = 0.0;
            public double[][] _MiduScore = null;
        }
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
                        GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF, disval, 1, ref _RiriHVal);
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

        public static List<MilePartD> _RoadPartF = null;//0.1米桩号分段
        private static double[] _LiriHVal = null;
        private static double[] _RiriHVal = null;

        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
             bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsMaxRut, int PartType, bool IsSpeed, bool IsGPS)
        {
            bool IRIRes = true, RutRes = true, MTDRes = true, GPSRes = true;
            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }
            _RoadPart = new List<MilePart>();
            MilePart spart = null;
            MilePart tpart = null;
            try
            {
                spart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade.Replace("主干路次干路", "主干路")], degreestr = prjinfo._RoadGrade };
            }
            catch
            {
                MessageBox.Show(string.Format("【市政道路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
                System.Environment.Exit(0);
            }
            _RoadPart.Add(spart);
            if (PartType == 0)
            {
                GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart, ref _MarkVal);
            }
            else if (PartType == 1)
            {
                GlobalExcel.GetAllMilePart_Dmi(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                GlobalExcel.GetMarkInfo_Dmi(prjinfo, prjdir, _RoadPart, ref _MarkVal);
            }

            if (IsDis)
            {
                if (_Setting.OutRut == 1)
                {
                    if (_RoadPart1M != null)
                    {
                        _RoadPart1M.Clear();
                        _RoadPart1M = null;
                    }
                    _RoadPart1M = new List<MilePart>();

                    tpart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade.Replace("主干路次干路", "主干路")], degreestr = prjinfo._RoadGrade };
                    _RoadPart1M.Add(tpart);

                    GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 1, prjinfo._Direction, _RoadGradeStr, ref _RoadPart1M, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
                    GlobalExcel.GetRutDisVal(prjinfo, prjdir, _RoadPart1M, ref _SRutDisValL, ref _SRutDisValR, ref _SRutDisMile);

                }
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, _RoadGradeDict, _SRutDisValL, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh, _RoadPart, _SRutDisValR, true);
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
                if (IsMeanMTD)
                {
                    MTDRes = GlobalExcel.GetMTDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, _Setting.IsWarning);
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
                }
            }
            else
            {
                RutRes = true;
            }

            if (IsGPS) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);
            
            return IRIRes && RutRes && MTDRes && GPSRes;
        }

        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面平整度评价等级记录表.xlsx",
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

            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputIRI_WithSpeed(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面平整度评价等级记录表(带车速).xlsx",
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

            WriteIRI2Xls_WithSpeed(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _MarkVal,_SpeedVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, string[] MarkVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LIRIVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RIRIVal[i];
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + 4);
                }
                vallist[i, 6] = String.Format("=IF({0}+{1}*F{2}>=0,{0}+{1}*F{2},0)", _RQIa[0], _RQIa[1], i + 4);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"A\",IF(G{0}>={2},\"B\",IF(G{0}>={3},\"C\",\"D\")))",
                    i + 4, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 9] = roadpart[i].degreestr;
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:K{0}", len + 3));
            destrange.Value2 = vallist;

            WriteIRIStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }

        private static void WriteIRI2Xls_WithSpeed(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
         List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal, string[] MarkVal, double[] SpeedVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 12];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LIRIVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RIRIVal[i];
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + 4);
                }
                vallist[i, 6] = String.Format("=IF({0}+{1}*F{2}>=0,{0}+{1}*F{2},0)", _RQIa[0], _RQIa[1], i + 4);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"A\",IF(G{0}>={2},\"B\",IF(G{0}>={3},\"C\",\"D\")))",
                    i + 4, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 9] = roadpart[i].degreestr;
                vallist[i, 10] = SpeedVal[i];
                vallist[i, 11] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:L{0}", len + 3));
            destrange.Value2 = vallist;

            WriteIRIStatistics_Speed(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:L{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteIRIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "A", "B", "C", "D" };
            MSExcel.Range destrange = _Worksheet.get_Range("S3:V5");
            object[,] val = new object[3, 4];
            for (int i = 0; i < degstr.Length; i++)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('S' + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 12] = "=CONCATENATE(\"路面平整度评价等级“A”率占路段总数\",ROUND(S4,4)*100,\"%，“B”率占路段总数\",ROUND(T4,4)*100,\"%，“C”率占路段总数\",ROUND(U4,4)*100,\"%，“D”率占路段总数\",ROUND(V4,4)*100,\"%。\")";
        }

        private static void WriteIRIStatistics_Speed(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "A", "B", "C", "D" };
            MSExcel.Range destrange = _Worksheet.get_Range("S3:V5");
            object[,] val = new object[3, 4];
            for (int i = 0; i < degstr.Length; i++)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('S' + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 13] = "=CONCATENATE(\"路面平整度评价等级“A”率占路段总数\",ROUND(S4,4)*100,\"%，“B”率占路段总数\",ROUND(T4,4)*100,\"%，“C”率占路段总数\",ROUND(U4,4)*100,\"%，“D”率占路段总数\",ROUND(V4,4)*100,\"%。\")";
        }
        public static void OutputDis_HPcsv_0(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            // MSExcel.Workbook _WorkbookLQ = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            // MSExcel.Worksheet _Worksheet_lqhz = null;

            string subdname = null;
            subdname = "两米";
            
            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
            string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            // string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\CICS\{1}\水泥混凝土路面破损{1}汇总.xlsx", System.Windows.Forms.Application.StartupPath, subdname);
            // string destxls = string.Format(@"{0}\{1}{2}{3}({4}{5}{3}(已识别)-路面破损-水泥路面).csv", path, prjinfo._RoadCode, prjinfo._RoadName, strDirection, prjinfo._DataDate, prjinfo._DataTime);
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\上海惠普\病害模板.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-病害-{6}.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime, disval);
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
                File.Delete(string.Format(@"{0}\{2} {3} {1} {4}{5}-病害.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime));
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

            List<(int,string)> markDic = new List<(int, string)>();

            for (int i = 0; i < RoadPart1M.Count; i++)
            {
                if (markInfo[i]!=null)
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
            object[,] disvalsn = new object[len, 23];
            int typeidx = 0;
            bool res = false; 
            string roadSplitStr = "Start";
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int colcnt = 0;

                double smile = RoadPart1M[i].mile;
                double emile = RoadPart1M[i + 1].mile;
                int curDmi = RoadPart1M[i].dmi; 
                disvalsn[i, colcnt++] = prjinfo._RoadCode;//路线代码
                disvalsn[i, colcnt++] = prjinfo._DataDate.Substring(0, 4) + "/" + prjinfo._DataDate.Substring(4, 2) + "/" + prjinfo._DataDate.Substring(6, 2) + " " + prjinfo._DataTime.Substring(0, 2) + ":" + prjinfo._DataTime.Substring(2, 2);
                disvalsn[i, colcnt++] = prjinfo._RoadName;

                disvalsn[i, colcnt++] = leftimgsinfo[i].Split('\\').Last();//

                for (int t = 0; t < markDic.Count; t++)
                {
                    if (markDic[markDic.Count-1].Item1 <= curDmi )
                    {
                        roadSplitStr = "End";
                        break;
                    }

                    if (markDic[t].Item1 > curDmi)
                    {
                        if (t == 0 )
                        {
                            roadSplitStr = "Start";
                        } 
                        else
                        {
                            roadSplitStr = "Reference" + " " + t;
                        }

                            break;
                    } 
                }

                disvalsn[i, colcnt++] = roadSplitStr;


                string s1 = (smile * 0.001).ToString("f3");
                string s2 = (smile * 0.001).ToString("f3");
                disvalsn[i, colcnt++] = s1;
                disvalsn[i, colcnt++] = s2;
                 
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
                                disvalsn[i, colcnt++] = Math.Round((double)arrdis[j].Area,2);
                            }
                            else if (name.Contains("裂缝"))
                            {
                                disvalsn[i, colcnt++] = Math.Round(arrdis[j].calcheight,2);
                                disvalsn[i, colcnt++] = "";
                            }

                            else
                            {
                                disvalsn[i, colcnt++] = Math.Round(arrdis[j].calcheight,2);
                                disvalsn[i, colcnt++] = Math.Round((double)arrdis[j].Area,2);
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
            }
            #endregion
            if (Haslqflag)
            {
                destrange = worksheet_snhz.get_Range(String.Format("A2:W{0}", len + 1));
                destrange.Value2 = disvalsn;
                destrange = worksheet_snhz.get_Range(String.Format("A2:W{0}", rowcnt_sn + 1));
                GlobalExcel.SetBorderLine(destrange, borderType);
                
            }
        }
        public static void OutputDis_HPcsv_IRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            MSExcel.Workbook _WorkbookSN = null;
            MSExcel.Worksheet _Worksheet_snhz = null;
            string strDirection = prjinfo._Direction > 0 ? "上行" : "下行";
            string[] _RoadTypeStr = { "沥青", "水泥", "砂石" };
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\上海惠浦\x640310115s2-iri-1.xlsx", System.Windows.Forms.Application.StartupPath/*, subdname*/);
            string destxls = string.Format(@"{0}\{2}-{3} {1} {4}{5}-iri-{6}m.csv", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime, disval);
            string destxlsTxt = string.Format(@"{0}\{2}-{3} {1} {4}{5}-iri-{6}m.txt", path, prjinfo._RoadCode, prjdir.Name, strDirection, prjinfo._DataDate, prjinfo._DataTime, disval);

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
                disvalsn[i, colcnt++] = marks[i];//事件
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

        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\车辙深度评价等级记录表.xlsx",
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

            WriteRut2Xls(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }


        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, double disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\车辙深度评价等级记录表.xlsx",
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

            WriteRut2Xls(_Worksheet, prjinfo, _RoadPartF, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteRut2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
           List<MilePartD> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal, string[] MarkVal)
        {
            MSExcel.Range destrange;

            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 11];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = Math.Round( roadpart[i].mile,1);
                vallist[i, 1] = Math.Round(roadpart[i + 1].mile,1);
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];
                vallist[i, 5] = SRutVal[i];
                //vallist[i, 5] = string.Format("=ROUND(MAX(D{0},E{0}),2)", i+4);
                vallist[i, 6] = string.Format("=IF(F{0}<={1},{2}-{3}*F{0},IF(F{0}<={4},{5}-{6}*(F{0}-{1}),0))",
                    i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 9] = roadpart[i].degreestr;
               // vallist[i, 10] = MarkVal[i];
            }
            destrange = _Worksheet.get_Range(String.Format("A4:B{0}", len + 3));
            destrange.ClearFormats();
            destrange = _Worksheet.get_Range(String.Format("A4:K{0}", len + 3));
            destrange.Value2 = vallist;

            WriteRutStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
       

        private static void WriteRut2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal, string[] MarkVal)
        {
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
                //if (i == 1)
                //{
                //    using (StreamWriter sw = new StreamWriter("C:\\Users\\Administrator\\Desktop\\新建文件夹 (3)\\data.txt", true))
                //    {
                //        sw.WriteLine(SRutVal[i]);
                //    } 
                //}
                //vallist[i, 5] = string.Format("=ROUND(MAX(D{0},E{0}),2)", i+4);
                vallist[i, 6] = string.Format("=IF(F{0}<={1},{2}-{3}*F{0},IF(F{0}<={4},{5}-{6}*(F{0}-{1}),0))",
                    i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 9] = roadpart[i].degreestr;
                vallist[i, 10] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:K{0}", len + 3));
            destrange.Value2 = vallist;

            WriteRutStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 11, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
            
        }
        private static void WriteRutStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("S3:W5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('S' + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 12] = "=CONCATENATE(\"沥青路面车辙深度评价等级“优”率占路段总数\",ROUND(S4,4)*100,\"%，“良”率占路段总数\",ROUND(T4,4)*100,\"%，“中”率占路段总数\",ROUND(U4,4)*100,\"%，“次”率占路段总数\",ROUND(V4,4)*100,\"%，“差”率占路段总数\",ROUND(W4,4)*100,\"%。\")";
        }

        public static void OutputMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面构造深度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_MTD_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal, string[] MarkVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 10];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = LMTDVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RMTDVal[i];
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND(D{0},5)", i + 4);
                }
                vallist[i, 6] = string.Format("=IF(F{0}>={1},\"A\",IF(F{0}>={2},\"B\",IF(F{0}>={3},\"C\",\"D\")))",
                    i + 4, _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);
                vallist[i, 7] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 8] = roadpart[i].degreestr;
                vallist[i, 9] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:J{0}", len + 3));
            destrange.Value2 = vallist;

            WriteMTDStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:J{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 10, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteMTDStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "A", "B", "C", "D" };
            object[,] val = new object[3, 4];
            MSExcel.Range destrange = _Worksheet.get_Range("R3:U5");
            for (int i = 0; i < degstr.Length; i++)
            {
                val[0, i] = string.Format("=ABS(SUMIF(G:G,\"{0}\",A:A)-SUMIF(G:G,\"{0}\",B:B))", degstr[i]);
                val[1, i] = string.Format("={0}3/{0}5", (char)('R' + i));
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 11] = "=CONCATENATE(\"沥青路面构造深度评价等级“A”率占路段总数\",ROUND(R4,4)*100,\"%，“B”率占路段总数\",ROUND(S4,4)*100,\"%，“C”率占路段总数\",ROUND(T4,4)*100,\"%，“D”率占路段总数\",ROUND(U4,4)*100,\"%。\")";
        }

        public static void OutputPWI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面磨耗评价等级记录表.xlsx",
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

            WritePWI2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _CMTDMeanVal, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePWI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal, double[] CMTDVal, string[] MarkVal)
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
                    vallist[i, 6] = string.Format("=IF(F{0}-MIN(D{0},E{0})>0, 100*(F{0}-MIN(D{0},E{0}))/F{0},0) ", i + 4);
                }
                vallist[i, 7] = string.Format("=100-{0}*POWER(G{1},{2})", _PWIa[0], i + 4, _PWIa[1]);
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _PWIGrade[roadpart[i].roaddegree][0],
                    _PWIGrade[roadpart[i].roaddegree][1],
                    _PWIGrade[roadpart[i].roaddegree][2],
                    _PWIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 10] = roadpart[i].degreestr;
                vallist[i, 11] = MarkVal[i];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:L{0}", len + 3));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            WritePWIStatistics(_Worksheet);
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 12, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WritePWIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("V3:Z5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(I:I,\"{0}\",A:A)-SUMIF(I:I,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('V' + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 15] = "=CONCATENATE(\"路面磨耗评价等级“优”率占路段总数\",ROUND(V4,4)*100,\"%，“良”率占路段总数\",ROUND(W4,4)*100,\"%，“中”率占路段总数\",ROUND(X4,4)*100,\"%，“次”率占路段总数\",ROUND(Y4,4)*100,\"%，“差”率占路段总数\",ROUND(Z4,4)*100,\"%。\")";
        }

        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面病害面积统计表.xlsx",
                 System.Windows.Forms.Application.StartupPath);

            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, _RoadDisList, _RoadPart);

            MSExcel.Worksheet _Worksheet_xb = _Workbook.Sheets["修补列表"] as MSExcel.Worksheet;
            WriteRepairLB2Xls(_Worksheet_xb, prjinfo, _RoadRepairList);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青路面病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥路面病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青路面病害区间汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥路面病害区间汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lqmx = _Workbook.Sheets["沥青路面病害区间明细表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snmx = _Workbook.Sheets["水泥路面病害区间明细表"] as MSExcel.Worksheet;
            WriteDisHZTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_snhz, _Worksheet_lqhz, _Worksheet_snmx, _Worksheet_lqmx,
               prjinfo, prjdir, _RoadPart, _RoadDisList);

            MSExcel.Worksheet _Worksheet_CADlb = _Workbook.Sheets["CAD病害列表"] as MSExcel.Worksheet;
            WriteCADDisLB2xls(_Worksheet_CADlb, prjinfo, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteDisLB2Xls_roadpart(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist, List<MilePart> roadpart)
        {
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0, troadtype = -1;
            object[,] val = new object[len, 13];
            foreach (Disease tdis in dislist)
            {
                for (int k = 0; k < roadpart.Count - 1; ++k)
                {
                    if ((prjinfo._Direction > 0 && roadpart[k].mile <= tdis.m_mile && tdis.m_mile < roadpart[k + 1].mile)
                      || (prjinfo._Direction < 0 && roadpart[k].mile >= tdis.m_mile && tdis.m_mile > roadpart[k + 1].mile))
                    {
                        troadtype = RoadDiseaseTypes.roadtypedict[tdis.RoadType];
                        if (troadtype == roadpart[k].roadtype)
                        {
                            val[i, 0] = tdis.m_mile;
                            val[i, 1] = prjinfo._RoadNum;
                            val[i, 2] = tdis.RoadDisType;
                            val[i, 3] = tdis.rect.Height * _RoadConfig.HeightScale;
                            val[i, 4] = tdis.rect.Width * _RoadConfig.WidthScale;
                            val[i, 5] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                            val[i, 6] = tdis.Area;
                            val[i, 7] = tdis.calcheight;
                            if (tdis.depth > 0)
                            {
                                val[i, 8] = tdis.depth;
                            }
                            else
                            {
                                val[i, 8] = "/";
                            }
                            val[i, 9] = tdis.calcwidth;
                            val[i, 10] = tdis.imgname;
                            val[i, 11] = tdis.imgpath;
                            val[i, 12] = tdis.RoadType;
                            troadtype = -1;
                            ++i;
                            break;
                        }
                    }
                }
            }
            int tlen = 0;
            for (int k = 0; k < len; ++k)
            {
                if (val[k, 0] == null)
                {
                    break;
                }
                tlen++;
            }
            destrange = _Worksheet.get_Range(String.Format("A3:M{0}", tlen + 2));
            destrange.Value2 = val;

            if (_Setting.showGpsInfoToPicture)
            {
                string outHighGpstxtPath = prjinfo._PrjPath + "/HighGps2Mile.txt";
                if (File.Exists(outHighGpstxtPath))
                {
                    List<string> highGpsTxts = File.ReadAllLines(outHighGpstxtPath).ToList();
                    List<(double, GPSInfo)> highGpss = new List<(double, GPSInfo)>();
                    foreach (var line in highGpsTxts)
                    {
                        string[] strings = line.Split(',');
                        GPSInfo gpsInfo = new GPSInfo();
                        gpsInfo._longitude =double.Parse( strings[0]);
                        gpsInfo._latitude = double.Parse(strings[1]);
                        gpsInfo._elevation = double.Parse(strings[2]);
                        highGpss.Add((double.Parse(strings[3]),gpsInfo));
                    }
                    if (highGpsTxts.Count > 0)
                    {
                        HighAccuracyPositioning.UpdateAllImg(prjinfo._PrjPath+"\\RoadImg\\Camera0");
                        _Worksheet.Cells[2,14].Value = "经度";
                        _Worksheet.Cells[2,15].Value = "纬度";
                        _Worksheet.Cells[2,16].Value = "高程";

                        for (int dd = 0; dd < dislist.Length; dd++)
                        { 
                            Disease tdis = dislist[dd];
                            int xPoint = tdis.rect.X + tdis.rect.Width / 2;
                            int yPoint = tdis.rect.Y + tdis.rect.Height / 2;
                            double dDiseaseLon = 0, dDiseaseLat = 0, dDiseaseH = 0; //当前像素
                            HighAccuracyPositioning.getHighAccPosition(_Setting.gpsformat, highGpss,_Setting.equipType, prjinfo._PrjPath, tdis.m_mile, xPoint, yPoint, _RoadConfig.ImageWidth
                                , _RoadConfig.ImageHeight, prjinfo._Direction, _RoadConfig.RealWidth, _RoadConfig.RealHeight,
                                  ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH); 
                            _Worksheet.Cells[2 + 1 + dd, 14].Value = dDiseaseLon;
                            _Worksheet.Cells[2 + 1 + dd, 15].Value = dDiseaseLat;
                            _Worksheet.Cells[2 + 1 + dd, 16].Value = dDiseaseH;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请选中软件上方【显示定位】选项后,重新点击GPS桩号匹配后重试");
                }
                destrange = _Worksheet.get_Range(String.Format("A1:P{0}", tlen + 2));
                GlobalExcel.SetBorderLine(destrange, 53);
            }
            else
            {
                destrange = _Worksheet.get_Range(String.Format("A1:M{0}", tlen + 2));
                GlobalExcel.SetBorderLine(destrange, 53);
            }
            

            //if (_Setting.Out_roadimg == 0) //不导出路面图像
            //{
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
            //    ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            //}

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 13, true);
            }
        }
        private static void WriteRepairLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0;
            object[,] val = new object[len, 7];
            foreach (Disease tdis in dislist)
            {
                val[i, 0] = tdis.m_mile;
                val[i, 1] = prjinfo._RoadNum;
                val[i, 2] = tdis.RoadDisType;
                val[i, 3] = tdis.rect.Height * _RoadConfig.HeightScale;
                val[i, 4] = tdis.rect.Width * _RoadConfig.WidthScale;
                val[i, 5] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                val[i, 6] = tdis.Area;
                ++i;
            }
            destrange = _Worksheet.get_Range(String.Format("A3:G{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:G{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 7, true);
            }
        }
        public static void WriteDisHZTJ2Xls(MSExcel.Worksheet worksheet_sntj, MSExcel.Worksheet worksheet_lqtj,
            MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            MSExcel.Worksheet worksheet_snmx, MSExcel.Worksheet worksheet_lqmx,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;
            object[,] disvalmx;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int rowcnt_sn_s = 5;
            int rowcnt_sn_e = 5;//小计起始的计算范围
            int rowcnt_lq_s = 5;
            int rowcnt_lq_e = 5;

            int rowcnt_lqmx = 4;
            int rowcnt_snmx = 4;

            int totalsnlen = 0;//水泥路段总长度
            int totallqlen = 0;//沥青路段总长度

            double partarea = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double md = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                partarea = _RoadConfig.DetectWidth * milelength;

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
                if (roadpart[i].roadtype == 1)
                {
                    Hassnflag = true;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 2];
                    disvalmx = new object[1, disnum * 3 + 5];
                    disvalmx[0, 0] = prjinfo._RoadName;
                    disvalmx[0, 1] = i;
                    disvalmx[0, 2] = milelength;
                    disvalmx[0, 3] = smile;
                    disvalmx[0, 4] = emile;

                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        md = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea / partarea;
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        disvalmx[0, kk * 3 + 5] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        disvalmx[0, kk * 3 + 6] = md;
                        disvalmx[0, kk * 3 + 7] = ChaZhi(_RoadSocre[roadpart[i].roadtype][RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname]._MiduScore, md);
                    }

                    pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, partarea);
                    disval[0, disnum] = Math.Round(pcival, 5);
                    disval[0, disnum + 1] = string.Format("=IF(R{0}>={1},\"A\",IF(R{0}>={2},\"B\",IF(R{0}>={3},\"C\",\"D\")))",
                        rowcnt_sn_s, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);
                    destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum + 1))));
                    destrange.Value2 = disval;

                    destrange = worksheet_snmx.get_Range(string.Format("A{0}:{1}{0}", rowcnt_snmx, GlobalExcel.GetCol((char)('A' + disnum * 3 + 4))));
                    destrange.Value2 = disvalmx;

                    totalsnlen += milelength;
                    rowcnt_sn_s++;
                    rowcnt_snmx++;
                }
                else if (roadpart[i].roadtype == 0)
                {
                    Haslqflag = true;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 2];
                    disvalmx = new object[1, disnum * 3 + 5];
                    disvalmx[0, 0] = prjinfo._RoadName;
                    disvalmx[0, 1] = i;
                    disvalmx[0, 2] = milelength;
                    disvalmx[0, 3] = smile;
                    disvalmx[0, 4] = emile;

                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        md = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea / partarea;
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        disvalmx[0, kk * 3 + 5] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        disvalmx[0, kk * 3 + 6] = md;
                        disvalmx[0, kk * 3 + 7] = ChaZhi(_RoadSocre[roadpart[i].roadtype][RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].disname]._MiduScore, md);
                    }
                    pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, partarea);
                    disval[0, disnum] = Math.Round(pcival, 5);
                    disval[0, disnum + 1] = string.Format("=IF(Q{0}>={1},\"A\",IF(Q{0}>={2},\"B\",IF(Q{0}>={3},\"C\",\"D\")))",
                        rowcnt_lq_s, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);
                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 1))));
                    destrange.Value2 = disval;

                    destrange = worksheet_lqmx.get_Range(string.Format("A{0}:{1}{0}", rowcnt_lqmx, GlobalExcel.GetCol((char)('A' + disnum * 3 + 4))));
                    destrange.Value2 = disvalmx;

                    totallqlen += milelength;
                    rowcnt_lq_s++;
                    rowcnt_lqmx++;
                }

                if (emile % 1000 == 0 && _Setting.IsOutputDisAreaSubtotal)
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
                    else if (roadpart[i].roadtype == 0 )
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

                        if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s && _Setting.IsOutputDisAreaSubtotal)
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

            //最后的一个小计
            if (roadpart[len].mile % 1000 != 0 && _Setting.IsOutputDisAreaSubtotal)
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
            destrange = worksheet_snhz.get_Range(String.Format("A1:S{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet_snmx.get_Range(String.Format("A1:AU{0}", rowcnt_snmx - 1));
            GlobalExcel.SetBorderLine(destrange, 53);

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
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
            }
            destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
            destrange.Value2 = disval;
            destrange = worksheet_lqhz.get_Range(String.Format("A1:R{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet_lqmx.get_Range(String.Format("A1:AR{0}", rowcnt_lqmx - 1));
            GlobalExcel.SetBorderLine(destrange, 53);

            int rowval = 0;
            RoadDiseaseTypes.Clear();

            if (Haslqflag)
            {
                worksheet_lqtj.Cells[2, 3] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len].mile);
                worksheet_lqtj.Cells[2, 5] = _RoadConfig.DetectWidth;
                worksheet_lqtj.Cells[4, 6] = string.Format("={0}*E2", Math.Abs(roadpart[0].mile - roadpart[len].mile));
                disnum = RoadDiseaseTypes.roaddis[0].Length;
                for (int i = 0; i < disnum; i++)
                {
                    rowval = i + 4;
                    if (_Setting.IsOutputDisAreaSubtotal)
                    {
                        worksheet_lqtj.Cells[rowval, 3] = string.Format("=SUMIF(沥青路面病害区间汇总表!{0}:{0},\"<>\",沥青路面病害区间汇总表!{0}:{0})/3", Convert.ToChar('D' + i));

                    }
                    else
                    {
                        worksheet_lqtj.Cells[rowval, 3] = string.Format("=SUMIF(沥青路面病害区间汇总表!{0}:{0},\"<>\",沥青路面病害区间汇总表!{0}:{0})/2", Convert.ToChar('D' + i));

                    }
                    worksheet_lqtj.Cells[rowval, 4] = string.Format("=C{0}/F4", rowval);
                    string disname = ((MSExcel.Range)worksheet_lqtj.Cells[rowval, 2]).Value.ToString();

                    double tmval = Convert.ToDouble(((MSExcel.Range)worksheet_lqtj.Cells[rowval, 4]).Value.ToString());
                    worksheet_lqtj.Cells[rowval, 5] = ChaZhi(_RoadSocre[0][disname]._MiduScore, tmval);
                    RoadDiseaseTypes.roaddis[0][i].totalarea = Convert.ToDouble(((MSExcel.Range)worksheet_lqtj.Cells[rowval, 3]).Value);
                }
                worksheet_lqtj.Cells[disnum + 4, 3] = String.Format("=SUM(C4:C{0})", disnum + 3);
                worksheet_lqtj.Cells[disnum + 4, 4] = String.Format("=C{0}/F4", disnum + 4);
                worksheet_lqtj.Cells[disnum + 4, 5] = String.Format("=SUM(E4:E{0}", disnum + 3);
            }
            else
            {
                worksheet_lqhz.Delete();
                worksheet_lqtj.Delete();
                worksheet_lqmx.Delete();
            }

            if (Hassnflag)
            {
                worksheet_sntj.Cells[2, 3] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len].mile);
                worksheet_sntj.Cells[2, 5] = _RoadConfig.DetectWidth;
                worksheet_sntj.Cells[4, 6] = string.Format("={0}*E2", Math.Abs(roadpart[0].mile - roadpart[len].mile));
                disnum = RoadDiseaseTypes.roaddis[1].Length;
                for (int i = 0; i < disnum; i++)
                {
                    rowval = i + 4;
                    if (_Setting.IsOutputDisAreaSubtotal)
                    {
                        worksheet_sntj.Cells[i + 4, 3] = string.Format("=SUMIF(水泥路面病害区间汇总表!{0}:{0},\"<>\",水泥路面病害区间汇总表!{0}:{0})/3", Convert.ToChar('D' + i));

                    }
                    else
                    {
                        worksheet_sntj.Cells[i + 4, 3] = string.Format("=SUMIF(水泥路面病害区间汇总表!{0}:{0},\"<>\",水泥路面病害区间汇总表!{0}:{0})/2", Convert.ToChar('D' + i));

                    }
                    worksheet_sntj.Cells[rowval, 4] = string.Format("=C{0}/F4", rowval);
                    string disname = ((MSExcel.Range)worksheet_sntj.Cells[rowval, 2]).Value.ToString();
                    double tmval = Convert.ToDouble(((MSExcel.Range)worksheet_sntj.Cells[rowval, 4]).Value.ToString());
                    worksheet_sntj.Cells[rowval, 5] = ChaZhi(_RoadSocre[1][disname]._MiduScore, tmval);
                    RoadDiseaseTypes.roaddis[1][i].totalarea = Convert.ToDouble(((MSExcel.Range)worksheet_sntj.Cells[i + 4, 3]).Value);
                }
                worksheet_sntj.Cells[disnum + 4, 3] = String.Format("=SUM(C4:C{0}", disnum + 3);
                worksheet_sntj.Cells[disnum + 4, 4] = String.Format("=C{0}/F4", disnum + 4);
                worksheet_sntj.Cells[disnum + 4, 5] = String.Format("=SUM(E4:E{0}", disnum + 3);
            }
            else
            {
                worksheet_snhz.Delete();
                worksheet_sntj.Delete();
                worksheet_snmx.Delete();
            }

        }

        public static double ComputPCI(RoadDiseaseType[][] disarea, int roadtype, double partarea)
        {
            double[] totalareatmp = new double[disarea[roadtype].Length];
            double uij = 0, wij = 0;
            double[] DPa = { 0, 0, 0, 0, 0 };
            int len = _RoadSocre[roadtype].Keys.Count;

            for (int i = 0; i < len; i++)
            {
                totalareatmp[i] = disarea[roadtype][i].totalarea / partarea;
                totalareatmp[i] = ChaZhi(_RoadSocre[roadtype][disarea[roadtype][i].disname]._MiduScore, totalareatmp[i]);
            }

            // 类别内的扣分和
            for (int i = 0; i < len; i++)
            {
                DPa[_RoadSocre[roadtype][disarea[roadtype][i].disname]._DisType] += totalareatmp[i];
            }

            // 每种病害的uij，得到每种的权重曲线的扣分
            for (int i = 0; i < len; i++)
            {
                if (DPa[_RoadSocre[roadtype][disarea[roadtype][i].disname]._DisType] > 0)
                {
                    uij = totalareatmp[i] / DPa[_RoadSocre[roadtype][disarea[roadtype][i].disname]._DisType];
                }
                else
                {
                    uij = 0;
                }

                wij = _WeightParm[roadtype][0];
                for (int k = 1; k < _WeightParm[roadtype].Length; k++)
                {
                    wij = wij * uij + _WeightParm[roadtype][k];
                }

                totalareatmp[i] = totalareatmp[i] * wij;
            }

            for (int i = 0; i < DPa.Length; ++i)
            {
                DPa[i] = 0;
            }

            // 每类病害的扣分
            double DP = 0;
            for (int i = 0; i < len; i++)
            {
                DPa[_RoadSocre[roadtype][disarea[roadtype][i].disname]._DisType] += totalareatmp[i];
                DP += totalareatmp[i];
            }

            double DP2 = 0;
            for (int i = 0; i < DPa.Length; ++i)
            {
                if (DP > 0)
                {
                    uij = DPa[i] / DP;
                }
                else
                {
                    uij = 0;
                }

                wij = _WeightParm[roadtype][0];
                for (int k = 1; k < _WeightParm[roadtype].Length; k++)
                {
                    wij = wij * uij + _WeightParm[roadtype][k];
                }

                DPa[i] = DPa[i] * wij;

                DP2 = DP2 + DPa[i];
            }

            return 100 - DP2;
        }

        public static double ChaZhi(double[][] MiduScore, double mval)
        {
            double sval = 0;
            int len = MiduScore[0].Length;
            for (int i = 1; i < len; i++)
            {
                if (mval < MiduScore[0][i] && mval >= MiduScore[0][i - 1])
                {
                    if (i == 1)
                    {
                        sval = mval * (MiduScore[1][i] - MiduScore[1][i - 1]) / (MiduScore[0][i] - MiduScore[0][i - 1]);
                    }
                    else
                    {
                        if (i < len - 1)
                        {
                            sval = (mval - MiduScore[0][i - 1])
                                * (MiduScore[1][i] - MiduScore[1][i - 1])
                                / (MiduScore[0][i] - MiduScore[0][i - 1])
                                + MiduScore[1][i - 1];
                        }
                        else
                        {
                            if (MiduScore[0][len - 1] == MiduScore[0][len - 2])
                            {
                                sval = MiduScore[1][len - 1];
                            }
                            else
                            {
                                sval = (mval - MiduScore[0][i - 1])
                                    * (MiduScore[1][i] - MiduScore[1][i - 1])
                                    / (MiduScore[0][i] - MiduScore[0][i - 1])
                                    + MiduScore[1][i - 1];
                            }
                        }
                    }
                }
            }
            return sval;
        }

        public static void OutputPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面破损评价等级记录表.xlsx",
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

            if (_Setting.PartType == 0)
            {
                WritePCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _MarkVal);
            }
            else if (_Setting.PartType == 1)
            {
                WritePCI2Xls_Dmi(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _MarkVal);
            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, string[] MarkVal)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 8];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (pcival < 0) pcival = 0;
                vallist[i, 3] = Math.Round(pcival, 5);
                vallist[i, 4] = string.Format("=IF(D{0}>={1},\"A\",IF(D{0}>={2},\"B\",IF(D{0}>={3},\"C\",\"D\")))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 6] = roadpart[i].degreestr;
                vallist[i, 7] = MarkVal[i];
            }
            destrange = worksheet.get_Range(String.Format("A3:H{0}", len + 2));
            destrange.Value2 = vallist;
            WritePCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:H{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 8, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePCIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "A", "B", "C", "D" };
            MSExcel.Range destrange = _Worksheet.get_Range("P3:S5");
            object[,] val = new object[3, 4];
            for (int i = 0; i < 4; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(E:E,\"{0}\",A:A)-SUMIF(E:E,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('P' + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 9] = "=CONCATENATE(\"沥青路面PCI评价等级“A”率占路段总数\",ROUND(P4,4)*100,\"%，“B”率占路段总数\",ROUND(Q4,4)*100,\"%，“C”率占路段总数\",ROUND(R4,4)*100,\"%，“D”率占路段总数\",ROUND(S4,4)*100,\"%。\")";
        }
        private static void WritePCI2Xls_Dmi(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, string[] MarkVal)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 8];

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (pcival < 0) pcival = 0;
                vallist[i, 3] = Math.Round(pcival, 5);
                vallist[i, 4] = string.Format("=IF(D{0}>={1},\"A\",IF(D{0}>={2},\"B\",IF(D{0}>={3},\"C\",\"D\")))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 6] = roadpart[i].degreestr;
                vallist[i, 7] = MarkVal[i];
            }
            destrange = worksheet.get_Range(String.Format("A3:H{0}", len + 2));
            destrange.Value2 = vallist;
            WritePCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:H{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 8, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }

        public static void OutputPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\路面综合评价等级记录表.xlsx",
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

            if (_Setting.PartType == 0)
            {
                WritePQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _MarkVal);
            }
            else if (_Setting.PartType == 1)
            {
                WritePQI2Xls_Dmi(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _MarkVal);
            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, string[] MarkVal)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 12];
            double trqival = 0, irival = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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
                vallist[i, 0] = smile;
                vallist[i, 1] = emile;
                vallist[i, 2] = prjinfo._RoadNum;

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (pcival < 0) pcival = 0;
                vallist[i, 3] = Math.Round(pcival, 5);
                vallist[i, 4] = string.Format("=IF(D{0}>={1},\"A\",IF(D{0}>={2},\"B\",IF(D{0}>={3},\"C\",\"D\")))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = _RQIa[0] + _RQIa[1] * irival;
                vallist[i, 5] = trqival > 0 ? trqival : 0;

                vallist[i, 6] = string.Format("=IF(F{0}>={1},\"A\",IF(F{0}>={2},\"B\",IF(F{0}>={3},\"C\",\"D\")))",
                    i + 3, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                //PQI计算公式与路面材质没关系，只和道路等级有关系，快速路和主干路，支路和次干路
                vallist[i, 7] = string.Format("=ROUND({1}*D{0}+{2}*{3}*F{0},5)", i + 3, _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"A\",IF(H{0}>={2},\"B\",IF(H{0}>={3},\"C\",\"D\")))",
                    i + 3, _PQIGrade[roadpart[i].roaddegree][0], _PQIGrade[roadpart[i].roaddegree][1], _PQIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 10] = roadpart[i].degreestr;
                vallist[i, 11] = MarkVal[i];
            }

            destrange = worksheet.get_Range(String.Format("A3:L{0}", len + 2));
            destrange.Value2 = vallist;
            WritePQIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:L{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 12, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePQIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "A", "B", "C", "D" };
            MSExcel.Range destrange = _Worksheet.get_Range("T3:W5");
            object[,] val = new object[3, 4];
            for (int i = 0; i < 4; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(I:I,\"{0}\",A:A)-SUMIF(I:I,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('T' + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 13] = "=CONCATENATE(\"沥青路面PQI评价等级“A”率占路段总数\",ROUND(T4,4)*100,\"%，“B”率占路段总数\",ROUND(U4,4)*100,\"%，“C”率占路段总数\",ROUND(V4,4)*100,\"%，“D”率占路段总数\",ROUND(W4,4)*100,\"%。\")";
        }
        private static void WritePQI2Xls_Dmi(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal, string[] MarkVal)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 12];
            double trqival = 0, irival = 0;

            int typeidx = 0;
            bool res = false;

            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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
                vallist[i, 0] = smile;
                vallist[i, 1] = emile;
                vallist[i, 2] = prjinfo._RoadNum;

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (pcival < 0) pcival = 0;
                vallist[i, 3] = Math.Round(pcival, 5);
                vallist[i, 4] = string.Format("=IF(D{0}>={1},\"A\",IF(D{0}>={2},\"B\",IF(D{0}>={3},\"C\",\"D\")))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = _RQIa[0] + _RQIa[1] * irival;
                vallist[i, 5] = trqival > 0 ? trqival : 0;

                vallist[i, 6] = string.Format("=IF(F{0}>={1},\"A\",IF(F{0}>={2},\"B\",IF(F{0}>={3},\"C\",\"D\")))",
                    i + 3, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                //PQI计算公式与路面材质没关系，只和道路等级有关系，快速路和主干路，支路和次干路
                vallist[i, 7] = string.Format("=ROUND({1}*D{0}+{2}*{3}*F{0},5)", i + 3, _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"A\",IF(H{0}>={2},\"B\",IF(H{0}>={3},\"C\",\"D\")))",
                    i + 3, _PQIGrade[roadpart[i].roaddegree][0], _PQIGrade[roadpart[i].roaddegree][1], _PQIGrade[roadpart[i].roaddegree][2]);
                vallist[i, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 10] = roadpart[i].degreestr;
                vallist[i, 11] = MarkVal[i];
            }

            destrange = worksheet.get_Range(String.Format("A3:L{0}", len + 2));
            destrange.Value2 = vallist;
            WritePQIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:L{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 12, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }

        //合计病害长度为区间内病害长度之和，桩号为最长的病害桩号
        //宽度为最宽的病害宽度，中心位置为最宽病害的中心位置
        private static void WriteCADDisLB2xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] disbak)
        {
            MSExcel.Range destrange;
            int len = disbak.Length;
            object[,] val = new object[len, 7];
            int smile = prjinfo._StartMile;
            int dislen = disbak.Length;

            int cadrownum = 0;
            for (int i = 0; i < dislen; )
            {
                int maxlen = 0, maxlenmile = 0;
                List<Disease> caddislist = new List<Disease>();
                RoadDiseaseTypes.Clear();
                while (i < dislen && ((prjinfo._Direction > 0 && disbak[i].m_mile >= smile && disbak[i].m_mile < smile + prjinfo._Direction * _Setting.CADLength)
                    || (prjinfo._Direction < 0 && disbak[i].m_mile <= smile && disbak[i].m_mile > smile + prjinfo._Direction * _Setting.CADLength)))
                {
                    bool pushflag = false;
                    int pushidx = -1;
                    if (disbak[i].RoadDisType == "线裂")
                    {
                        if (disbak[i].rect.Width > disbak[i].rect.Height)
                        {
                            disbak[i].RoadDisType = "横向" + disbak[i].RoadDisType;
                        }
                        else
                        {
                            disbak[i].RoadDisType = "纵向" + disbak[i].RoadDisType;
                        }
                    }

                    foreach (Disease cdis in caddislist)
                    {
                        pushidx++;
                        if (cdis.RoadDisType == disbak[i].RoadDisType)
                        {
                            pushflag = true;
                            break;
                        }
                    }

                    if (pushflag)
                    {
                        caddislist[pushidx].rect.Height += disbak[i].rect.Height;
                        if (caddislist[pushidx].rect.Width < disbak[i].rect.Width)
                        {
                            caddislist[pushidx].rect.Width = disbak[i].rect.Width;
                            caddislist[pushidx].rect.X = disbak[i].rect.X;
                        }
                        if (maxlen < disbak[i].rect.Height)
                        {
                            maxlen = disbak[i].rect.Height;
                            maxlenmile = disbak[i].m_mile;
                        }
                        caddislist[pushidx].m_mile = maxlenmile;
                    }
                    else
                    {
                        caddislist.Add(disbak[i]);
                    }
                    ++i;
                }
                smile = smile + prjinfo._Direction * _Setting.CADLength;

                foreach (Disease cdis in caddislist)
                {
                    val[cadrownum, 0] = cdis.m_mile;
                    val[cadrownum, 1] = prjinfo._RoadNum;
                    val[cadrownum, 2] = cdis.RoadDisType;
                    val[cadrownum, 3] = cdis.rect.Height * _RoadConfig.HeightScale;
                    val[cadrownum, 4] = cdis.rect.Width * _RoadConfig.WidthScale;
                    val[cadrownum, 5] = (cdis.rect.Width / 2 + cdis.rect.X) * _RoadConfig.WidthScale;
                    ++cadrownum;
                }
            }

            destrange = _Worksheet.get_Range(String.Format("A3:F{0}", cadrownum + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:F{0}", cadrownum + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 6, true);
            }
        }

        //带GPS
        public static void OutputGPSRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_PQI = _Workbook.Sheets["PQI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_PCI = _Workbook.Sheets["PCI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_RQI = _Workbook.Sheets["RQI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_IRI = _Workbook.Sheets["IRI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sn = _Workbook.Sheets["水泥病害"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lq = _Workbook.Sheets["沥青病害"] as MSExcel.Worksheet;

            WriteGPSAll2Xls(_Worksheet_PQI, _Worksheet_PCI, _Worksheet_RQI, _Worksheet_IRI, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _GPSInfo);
            WriteGPSDis2Xls(_Worksheet_lq, _Worksheet_sn, prjinfo, prjdir, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteGPSAll2Xls(
            MSExcel.Worksheet worksheetPQI, MSExcel.Worksheet worksheetPCI,
            MSExcel.Worksheet worksheetRQI, MSExcel.Worksheet worksheetIRI,
            ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] IDObj = new object[len, 1];

            object[,] IRIObj = new object[len, 3];


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
                IDObj[rowcnt, 0] = i + 1;
                GPSStartObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;
                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                //病害相关
                PCIObj[rowcnt, 0] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);

                if (Convert.ToDouble(PCIObj[rowcnt, 0]) < 0)
                    PCIObj[rowcnt, 0] = 0;

                //平整度相关
                if (prjinfo._IsDIRIMTD)
                {
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                    IRIObj[rowcnt, 1] = RIRIVal[i];
                    IRIObj[rowcnt, 2] = string.Format("=(J{0}+K{0})/2", rowcnt + 2);
                }
                else
                {
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                    IRIObj[rowcnt, 2] = string.Format("=J{0}", rowcnt + 2);
                }

                RQIObj[rowcnt, 0] = string.Format("=IF({0}+{1}*IRI!L{2}>=0,{0}+{1}*IRI!L{2},0)",
                    _RQIa[0], _RQIa[1], rowcnt + 2);

                PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})*{3}),5)",
                         rowcnt + 2,
                        _PQIW[roadpart[i].roaddegree][0],
                        _PQIW[roadpart[i].roaddegree][1],
                        _PQIT
                         );
                ++rowcnt;

            }
            //将结果复制进Excel
            MSExcel.Range destrange, sortrange;
            MSExcel.Worksheet[] tsheet = { worksheetRQI, worksheetPQI, worksheetIRI, worksheetPCI };
            object[] tobj = { RQIObj, PQIObj, IRIObj, PCIObj };
            char[] valnum = { 'J', 'J', 'L', 'J' };
            for (int i = 0; i < tsheet.Length; ++i)
            {
                destrange = tsheet[i].get_Range(String.Format("A2:A{0}", len + 1));
                destrange.Value2 = IDObj;
                // 非公式数据
                if (i > 1)
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
                    //destrange = tsheet[i].get_Range(String.Format("B2:{0}{1}", valnum[i], len + 1));
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
                if (i <= 1)//公式数据
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

                colcnt = 0;
                disinfo[0, colcnt++] = i + 1;
                disinfo[0, colcnt++] = tempgpsinfo._utctime;
                //disinfo[0, colcnt++] = tempgpsinfo._mile;
                disinfo[0, colcnt++] = dislist[i].m_mile;
                disinfo[0, colcnt++] = tempgpsinfo._longitude;
                disinfo[0, colcnt++] = tempgpsinfo._latitude;

                int A = tempgpsinfo._mile;
                int B = dislist[i].m_mile;

                disinfo[0, colcnt++] = tempgpsinfo._utctime;
                //disinfo[0, colcnt++] = tempgpsinfo._mile;
                disinfo[0, colcnt++] = dislist[i].m_mile;
                disinfo[0, colcnt++] = tempgpsinfo._longitude;
                disinfo[0, colcnt++] = tempgpsinfo._latitude;

                disinfo[0, colcnt++] = dislist[i].RoadDisType;
                disinfo[0, colcnt++] = dislist[i].calcheight;
                disinfo[0, colcnt++] = dislist[i].calcwidth;

                if (dislist[i].depth > 0)
                {
                    disinfo[0, colcnt++] = dislist[i].depth;
                }
                else
                {
                    disinfo[0, colcnt++] = "/";
                }

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

            if (rowcnt_lq < 3)
            {
                worksheet_lq.Delete();
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_lq.get_Range(String.Format("B2:Q{0}", rowcnt_lq - 1));
                sortrange = worksheet_lq.get_Range(String.Format("C2:C{0}", len + 1));
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
            if (!((File.Exists(fnamemile0) || File.Exists(fnamemile1))))
            {
                MessageBox.Show("工程文件缺少景观图像数据");
                return;
            }
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS _景观图像.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS _全景图像.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS _路面图像.xlsx",
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
                        // tdmi = leftidx[i] * ImgDis;
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

                            //tmpmile = tempinfos[gi + 1]._mile - prjinfo._Direction * ImgDis;
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
                                    if (tempinfos[gi]._mile > tmile)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    if (tempinfos[gi]._mile < tmile)
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

        // 所有带GPS的指标，10米、100米、200米
        public static void OutputGPSAll2Xls_2(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = null;
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = null;
            if (disval == 10)
            {
                srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS_2.xlsx", System.Windows.Forms.Application.StartupPath);
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

            string sheetname = string.Format("桩号按{0}m输出", disval);
            MSExcel.Worksheet _Worksheet = _Workbook.Sheets[sheetname] as MSExcel.Worksheet;
            if (disval == 10)
            {
                MSExcel.Worksheet _Worksheet2 = _Workbook.Sheets["景观图像"] as MSExcel.Worksheet;
                if (prjinfo._IsStreet)
                {
                    WriteGPSImg2Xls_2(_Worksheet2, prjinfo, prjdir, "Street", prjinfo._StreetImgDis_Left);
                }
                else
                {
                    _Worksheet2.Delete();
                }

                MSExcel.Worksheet _Worksheet3 = _Workbook.Sheets["病害图像"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet4 = _Workbook.Sheets["路面图像"] as MSExcel.Worksheet;
                if (prjinfo._IsRoad)
                {
                    WriteGPSDis2Xls_2(_Worksheet3, prjinfo, prjdir, _RoadPart, _RoadDisList);
                    WriteGPSImg2Xls_2(_Worksheet4, prjinfo, prjdir, "Road", prjinfo._RoadImgDis);
                }
                else
                {
                    _Worksheet3.Delete();
                    _Worksheet4.Delete();
                }

                MSExcel.Worksheet _Worksheet5 = _Workbook.Sheets["全景图像"] as MSExcel.Worksheet;
                if (prjinfo._IsPano)
                {
                    WriteGPSImg2Xls_2(_Worksheet5, prjinfo, prjdir, "Pano", prjinfo._PanoImgDis);
                }
                else
                {
                    _Worksheet5.Delete();
                }
            }

            WriteGPSAll2Xls_2(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _GPSInfo, disval);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteGPSAll2Xls_2(MSExcel.Worksheet worksheet,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo, int disval)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] ObjVal = new object[len, 23];

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

                //GPS、桩号、工程信息等
                ObjVal[rowcnt, 0] = i + 1;
                ObjVal[rowcnt, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                ObjVal[rowcnt, 2] = prjinfo._RoadNum;

                //ObjVal[rowcnt, 3] = GPSInfo[i]._utctime;
                //ObjVal[rowcnt, 4] = smile;
                //ObjVal[rowcnt, 5] = GPSInfo[i]._longitude;
                //ObjVal[rowcnt, 6] = GPSInfo[i]._latitude;
                //ObjVal[rowcnt, 7] = GPSInfo[i + 1]._utctime;
                //ObjVal[rowcnt, 8] = emile;
                //ObjVal[rowcnt, 9] = GPSInfo[i + 1]._longitude;
                //ObjVal[rowcnt, 10] = GPSInfo[i + 1]._latitude;

                GPSStartObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;

                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                ObjVal[rowcnt, 11] = string.Format("=ABS(E{0}-I{0})", rowcnt + 2);

                //PCI
                ObjVal[rowcnt, 14] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);

                if (Convert.ToDouble(ObjVal[rowcnt, 14]) < 0) ObjVal[rowcnt, 14] = 0;
                ObjVal[rowcnt, 15] = string.Format("=IF(O{0}>={1},\"A\",IF(O{0}>={2},\"B\",IF(O{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                //RQI、MTD
                if (prjinfo._IsDIRIMTD)
                {
                    ObjVal[rowcnt, 18] = (LIRIVal[i] + RIRIVal[i]) / 2;
                    ObjVal[rowcnt, 21] = (LMTDVal[i] + RMTDVal[i]) / 2;
                }
                else
                {
                    ObjVal[rowcnt, 18] = LIRIVal[i];
                    ObjVal[rowcnt, 21] = LMTDVal[i];
                }
                ObjVal[rowcnt, 19] = string.Format("=IF(S{0}<={1},\"A\",IF(S{0}<={2},\"B\",IF(S{0}<={3},\"C\",\"D\")))",
                    rowcnt + 2, _IRIGrade[roadpart[i].roaddegree][1], _IRIGrade[roadpart[i].roaddegree][2], _IRIGrade[roadpart[i].roaddegree][3]);

                ObjVal[rowcnt, 16] = string.Format("=IF({0}+{1}*S{2}>=0,{0}+{1}*S{2},0)", _RQIa[0], _RQIa[1], rowcnt + 2);
                ObjVal[rowcnt, 17] = string.Format("=IF(Q{0}>={1},\"A\",IF(Q{0}>={2},\"B\",IF(Q{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                ObjVal[rowcnt, 22] = string.Format("=IF(V{0}>={1},\"A\",IF(V{0}>={2},\"B\",IF(V{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);

                //Rut
                ObjVal[rowcnt, 20] = SRutVal[i];

                // PQI
                ObjVal[rowcnt, 12] = string.Format("=ROUND(({1}*(O{0})+{2}*(Q{0})*{3}),5)",
                         rowcnt + 2,
                        _PQIW[roadpart[i].roaddegree][0],
                        _PQIW[roadpart[i].roaddegree][1],
                        _PQIT);
                ObjVal[rowcnt, 13] = string.Format("=IF(M{0}>={1},\"A\",IF(M{0}>={2},\"B\",IF(M{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A2:W{0}", len + 1));
            destrange.Value2 = ObjVal;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (disval == 10)
            {
                destrange = worksheet.get_Range("W:W");
                destrange.Delete();
                destrange = worksheet.get_Range("T:T");
                destrange.Delete();
                destrange = worksheet.get_Range("M:R");
                destrange.Delete();
            }

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSEndObj;
                destrange = worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSStartObj;

                destrange = worksheet.get_Range(string.Format("B2:W{0}", len + 1));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("E2:E{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
            else
            {
                destrange = worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSStartObj;
                destrange = worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSEndObj;
            }
        }

        // 所有带GPS的指标，10米、100米、200米
        public static void OutputGPSAll2Xls_2_Dmi(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = null;
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = null;

            srcxls = string.Format(@"{0}\报表模板\城镇道路\综合报表模板GPS_2_Dmi.xlsx", System.Windows.Forms.Application.StartupPath);
            _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing, false,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet2 = _Workbook.Sheets["景观图像"] as MSExcel.Worksheet;
            if (prjinfo._IsStreet)
            {
                WriteGPSImg2Xls_2(_Worksheet2, prjinfo, prjdir, "Street", prjinfo._StreetImgDis_Left) ;
            }
            else
            {
                _Worksheet2.Delete();
            }

            MSExcel.Worksheet _Worksheet3 = _Workbook.Sheets["病害图像"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet4 = _Workbook.Sheets["路面图像"] as MSExcel.Worksheet;
            if (prjinfo._IsStreet)
            {
                WriteGPSDis2Xls_2(_Worksheet3, prjinfo, prjdir, _RoadPart, _RoadDisList);
                WriteGPSImg2Xls_2(_Worksheet4, prjinfo, prjdir, "Road", prjinfo._RoadImgDis);
            }
            else
            {
                _Worksheet3.Delete();
                _Worksheet4.Delete();
            }

            MSExcel.Worksheet _Worksheet5 = _Workbook.Sheets["全景图像"] as MSExcel.Worksheet;
            if (prjinfo._IsStreet)
            {
                WriteGPSImg2Xls_2(_Worksheet5, prjinfo, prjdir, "Pano", prjinfo._PanoImgDis);
            }
            else
            {
                _Worksheet5.Delete();
            }

            MSExcel.Worksheet _Worksheet = _Workbook.Sheets["按单元输出"] as MSExcel.Worksheet;
            WriteGPSAll2Xls_2_Dmi(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _GPSInfo, disval, _MarkVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteGPSImg2Xls_2(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, string ImgType, int ImgDis)
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
            object[,] dataobj = new object[len, 11];
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

                            //tmpmile = tempinfos[gi + 1]._mile - prjinfo._Direction * ImgDis;
                            //if (tmpmile < 0)
                            //    tmpmile = 0;
                            //dataobj[leftidx[i], colcnt++] = tmpmile;
                            dataobj[leftidx[i], colcnt++] = tmile;

                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._longitude;
                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._latitude;
                            dataobj[leftidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                            dataobj[leftidx[i], colcnt++] = prjinfo._RoadNum;
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
                            dataobj[leftidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                            dataobj[leftidx[i], colcnt++] = prjinfo._RoadNum;
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
                                dataobj[rightidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                                dataobj[rightidx[i], colcnt++] = prjinfo._RoadNum;

                                colcnt = 9;
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
                                dataobj[rightidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                                dataobj[rightidx[i], colcnt++] = prjinfo._RoadNum;

                                colcnt = 9;
                                temp = rightimgsinfo[i].LastIndexOf('\\');
                                dataobj[rightidx[i], colcnt++] = rightimgsinfo[i].Substring(temp + 1);
                                int temp2 = rightimgsinfo[i].IndexOf(' ') + 2;
                                dataobj[rightidx[i], colcnt++] = string.Format("\\{0}Img\\Camera1\\{1}", ImgType, rightimgsinfo[i].Substring(temp2, temp - temp2));
                            }
                        }
                    }
                }
            }

            MSExcel.Range destrange = worksheet.get_Range(string.Format("A2:K{0}", len + 1));
            destrange.Value2 = dataobj;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(string.Format("B2:K{0}", len + 1));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("C2:C{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
        }
        private static void WriteGPSDis2Xls_2(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

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
            else
            {
                MessageBox.Show("请先进行GPS桩号匹配！");
                return;
            }

            int rowcnt = 0;
            int gi = 0;
            ExcelGPS tempgpsinfo = null;
            object[,] vallist = new object[dlen, 15];

            int typeidx = 0;
            bool res = false;
			
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;

                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (!res)
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                        ++j;
                        continue;
                    }

                    for (; gi < tempinfos.Length; ++gi)
                    {
                        if (prjinfo._Direction > 0)
                        {
                            if (tempinfos[gi]._mile >= arrdis[j].m_mile)
                            {
                                break;
                            }
                        }
                        else
                        {
                            if (tempinfos[gi]._mile <= arrdis[j].m_mile)
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
                    vallist[rowcnt, 0] = rowcnt + 1;
                    vallist[rowcnt, 1] = tempgpsinfo._utctime;
                    vallist[rowcnt, 2] = arrdis[j].m_mile;
                    vallist[rowcnt, 3] = tempgpsinfo._longitude;
                    vallist[rowcnt, 4] = tempgpsinfo._latitude;
                    vallist[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                    vallist[rowcnt, 6] = arrdis[j].RoadDisType;
                    vallist[rowcnt, 7] = arrdis[j].calcheight;
                    vallist[rowcnt, 8] = arrdis[j].calcwidth;
                    if (arrdis[j].depth > 0)
                    {
                        vallist[rowcnt, 9] = arrdis[j].depth;
                    }
                    else
                    {
                        vallist[rowcnt, 9] = "/";
                    }
                    vallist[rowcnt, 10] = arrdis[j].Area;
                    vallist[rowcnt, 11] = prjinfo._Direction > 0 ? "上行" : "下行";
                    vallist[rowcnt, 12] = prjinfo._RoadNum;
                    vallist[rowcnt, 13] = arrdis[j].imgname;
                    vallist[rowcnt, 14] = arrdis[j].imgpath;
                    ++rowcnt;
                    ++j;
                }
            }

            MSExcel.Range destrange = worksheet.get_Range(string.Format("A2:O{0}", dlen + 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(string.Format("B2:O{0}", len + 1));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("C2:C{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
        }

        private static void WriteGPSAll2Xls_2_Dmi(MSExcel.Worksheet worksheet,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo, int disval, string[] MarkVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] ObjVal = new object[len, 26];

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

                //GPS、桩号、工程信息等
                ObjVal[rowcnt, 0] = i + 1;
                ObjVal[rowcnt, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                ObjVal[rowcnt, 2] = prjinfo._RoadNum;

                ////ObjVal[rowcnt, 3] = GPSInfo[i]._utctime;
                ////ObjVal[rowcnt, 4] = smile;
                ////ObjVal[rowcnt, 5] = GPSInfo[i]._longitude;
                ////ObjVal[rowcnt, 6] = GPSInfo[i]._latitude;
                ////ObjVal[rowcnt, 7] = GPSInfo[i + 1]._utctime;
                ////ObjVal[rowcnt, 8] = emile;
                ////ObjVal[rowcnt, 9] = GPSInfo[i + 1]._longitude;
                ////ObjVal[rowcnt, 10] = GPSInfo[i + 1]._latitude;

                GPSStartObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;

                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                ObjVal[rowcnt, 11] = string.Format("=ABS(E{0}-I{0})", rowcnt + 2);

                //PCI
                ObjVal[rowcnt, 14] = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (Convert.ToDouble(ObjVal[rowcnt, 14]) < 0) ObjVal[rowcnt, 14] = 0;
                ObjVal[rowcnt, 15] = string.Format("=IF(O{0}>={1},\"A\",IF(O{0}>={2},\"B\",IF(O{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                //RQI、MTD
                if (prjinfo._IsDIRIMTD)
                {
                    ObjVal[rowcnt, 18] = (LIRIVal[i] + RIRIVal[i]) / 2;
                    ObjVal[rowcnt, 21] = (LMTDVal[i] + RMTDVal[i]) / 2;
                }
                else
                {
                    ObjVal[rowcnt, 18] = LIRIVal[i];
                    ObjVal[rowcnt, 21] = LMTDVal[i];
                }
                ObjVal[rowcnt, 19] = string.Format("=IF(S{0}<={1},\"A\",IF(S{0}<={2},\"B\",IF(S{0}<={3},\"C\",\"D\")))",
                    rowcnt + 2, _IRIGrade[roadpart[i].roaddegree][1], _IRIGrade[roadpart[i].roaddegree][2], _IRIGrade[roadpart[i].roaddegree][3]);

                ObjVal[rowcnt, 16] = string.Format("=IF({0}+{1}*S{2}>=0,{0}+{1}*S{2},0)", _RQIa[0], _RQIa[1], rowcnt + 2);
                ObjVal[rowcnt, 17] = string.Format("=IF(Q{0}>={1},\"A\",IF(Q{0}>={2},\"B\",IF(Q{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
                ObjVal[rowcnt, 22] = string.Format("=IF(V{0}>={1},\"A\",IF(V{0}>={2},\"B\",IF(V{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);

                //Rut
                ObjVal[rowcnt, 20] = SRutVal[i];

                // PQI
                ObjVal[rowcnt, 12] = string.Format("=ROUND(({1}*(O{0})+{2}*(Q{0})*{3}),5)",
                         rowcnt + 2,
                        _PQIW[roadpart[i].roaddegree][0],
                        _PQIW[roadpart[i].roaddegree][1],
                        _PQIT);
                ObjVal[rowcnt, 13] = string.Format("=IF(M{0}>={1},\"A\",IF(M{0}>={2},\"B\",IF(M{0}>={3},\"C\",\"D\")))",
                    rowcnt + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                ObjVal[rowcnt, 23] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                ObjVal[rowcnt, 24] = roadpart[i].degreestr;
                ObjVal[rowcnt, 25] = _MarkVal[rowcnt];

                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A2:Z{0}", len + 1));
            destrange.Value2 = ObjVal;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (disval == 10)
            {
                destrange = worksheet.get_Range("W:W");
                destrange.Delete();
                destrange = worksheet.get_Range("T:T");
                destrange.Delete();
                destrange = worksheet.get_Range("M:R");
                destrange.Delete();
            }

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSEndObj;
                destrange = worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSStartObj;

                destrange = worksheet.get_Range(string.Format("B2:Z{0}", len + 1));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("E2:E{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
            else
            {
                destrange = worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSStartObj;
                destrange = worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSEndObj;
            }
        }
        #region 奥路通
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
                srcxls = string.Format(@"{0}\报表模板\城镇道路\奥路通\001-路面病害模板.xlsx",
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
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
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

            int rowcnt_sn_s = 4;
            int rowcnt_sn_e = 4;//小计起始的计算范围
            int rowcnt_lq_s = 4;
            int rowcnt_lq_e = 4;

            //int totalsnlen = 0;//水泥路段总长度
            //int totallqlen = 0;//沥青路段总长度

            double partarea = 0;
            if (prjinfo._Direction > 0)
            {
                worksheet_snhz.Cells[2, 15] = "上行";
                worksheet_lqhz.Cells[2, 15] = "上行";
            }
            else
            {
                worksheet_snhz.Cells[2, 15] = "下行";
                worksheet_lqhz.Cells[2, 15] = "下行";
            }

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                partarea = _RoadConfig.DetectWidth * milelength;

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
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt] = smile.ToString() + emile.ToString();

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];

                    for (int di = 0, kk = 0; di < RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count; ++di, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }

                    pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, partarea);
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt + 1] = Math.Round(pcival, 5);

                    destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                    destrange.Value2 = disval;

                    rowcnt_sn_s++;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt] = smile.ToString() + emile.ToString();

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];

                    for (int di = 0, kk = 0; di < RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count; ++di, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }

                    pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype,  partarea);
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt + 1] = Math.Round(pcival, 5);

                    destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
                    destrange.Value2 = disval;

                    rowcnt_lq_s++;

                }

                if (emile % 1000 == 0)
                {
                    if (roadpart[i].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                        disval = new object[1, disnum];
                        for (int di = 0; di < disnum; di++)
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
                            for (int di = 0; di < disnum; di++)
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
                        for (int di = 0; di < disnum; di++)
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
                            for (int di = 0; di < disnum; di++)
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
                    for (int di = 0; di < disnum; di++)
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
                        for (int di = 0; di < disnum; di++)
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
                    for (int di = 0; di < disnum; di++)
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
                        for (int di = 0; di < disnum; di++)
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
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_s - 1);
            }
            destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
            destrange.Value2 = disval;
            destrange = worksheet_snhz.get_Range(String.Format("A1:P{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, 53);

            //沥青
            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_s - 1);
            }
            destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum - 1))));
            destrange.Value2 = disval;
            destrange = worksheet_lqhz.get_Range(String.Format("A1:O{0}", rowcnt_lq_s)); 
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

        private static void WriteALTPQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal
            )
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double irival = 0;
            worksheet.Cells[2, 6] = prjinfo._RoadNum;
            worksheet.Cells[2, 2] = prjinfo._RoadName;
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 16];

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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

                vallist[rowcnt, colcnt] = string.Format("K{00:0+000}-K{1:00+000}", roadpart[i].mile, roadpart[i + 1].mile);
                vallist[rowcnt, colcnt + 1] = Math.Abs(emile - smile);

                //  PCI
                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[rowcnt, colcnt + 3] = Math.Round(pcival, 5);

                //IRI RQI
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }

                //trqival = _RQIa[0] + _RQIa[1] * irival;
                vallist[rowcnt, colcnt + 2] = Math.Round(irival, 5);
                vallist[rowcnt, colcnt + 4] = string.Format("=IF（ROUND({0}+{1}*C{2},2)>0,ROUND({0}+{1}*C{2},5),0)",
                    _RQIa[0], _RQIa[1], i + 5);

                // PQI
                vallist[rowcnt, colcnt + 5] = string.Format("=ROUND({1}*D{0}+{2}*{3}*E{0},5)", i + 5, _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);

                vallist[rowcnt, colcnt + 6] = roadpart[i].degreestr;
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:G{0}", rowcnt + 4));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(string.Format("A5:G{0}", rowcnt + 4));
                MSExcel.Range sortrange = worksheet.get_Range(string.Format("A5:A{0}", rowcnt + 4));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
            int chartlen = len + 4;
            MSExcel.ChartObject chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            MSExcel.Chart chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0},D4:D{0}, E4:E{0}, F4:F{0}", chartlen));

            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "PQI各项指标", Type.Missing, Type.Missing, Type.Missing);

            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(2);
            chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0}, C4:C{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "IRI", Type.Missing, Type.Missing, Type.Missing);


        }
        #endregion

        #region 模板1
        //pci
        public static void OutputZYPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            //武汉+南湖大道-上行-右2车道-20170925073424-按路线--车行道技术状况检测结果明细表.xlsx
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\PCI\按路线--车行道技术状况检测结果明细表.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\PCI";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "PCI-按路线--车行道技术状况检测结果明细表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteZYPCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYPCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
              List<MilePart> roadpart, Disease[] arrdis)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 7];
            worksheet.Cells[1, 1] = prjinfo._District + "-车行道技术状况检测结果明细表";
            int UnitNum = int.Parse(设置单元编号.unitnum);

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0;
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

                vallist[i, 0] = string.Format("{0:000000}-{1}", UnitNum++, prjinfo._RoadNum);
                vallist[i, 1] = smile;
                vallist[i, 2] = emile;
                vallist[i, 3] = Math.Abs(emile - smile);
                vallist[i, 4] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 5] = Math.Round(pcival, 5);
                vallist[i, 6] = string.Format("=IF(F{0}>={1},\"A\",IF(F{0}>={2},\"B\",IF(F{0}>={3},\"C\",\"D\")))",
                    i + 4, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

            }
            destrange = worksheet.get_Range(String.Format("A4:G{0}", len + 3));
            destrange.Value2 = vallist;
            // WriteZYPCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:G{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 4, 2, 7, true);
                // GlobalExcel.Reflection(worksheet, 5, 1, 2, false);
            }
        }
        //rut
        public static void OutputZYRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\车辙\按单元--车辙状况与评定表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\车辙";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "Rut-按单元--车辙状况与评定表");

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteZYRut2Xls(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYRut2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 9];
            _Worksheet.Cells[1, 1] = prjinfo._District + " - " + "车辙状况与判定表";
            _Worksheet.Cells[3, 2] = Math.Abs(prjinfo._StartMile - prjinfo._EndMile);
            _Worksheet.Cells[2, 1] = prjinfo._StartMile;
            _Worksheet.Cells[2, 3] = prjinfo._EndMile;
            _Worksheet.Cells[3, 4] = _RoadConfig.DetectWidth;
            _Worksheet.Cells[3, 7] = 设置单元编号.roadnum;
            int UnitNum = int.Parse(设置单元编号.unitnum);
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = string.Format("{0:000000}-{1}", UnitNum++, prjinfo._RoadNum);
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;
                vallist[i, 3] = Math.Abs(roadpart[i + 1].mile - roadpart[i].mile);
                vallist[i, 4] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 5] = LRutVal[i];
                vallist[i, 6] = RRutVal[i];
                vallist[i, 7] = SRutVal[i];
                vallist[i, 8] = string.Format("=IF(H{0}<={1},{2}-{3}*H{0},IF(H{0}<={4},{5}-{6}*(H{0}-{1}),0))",
                   i + 6, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);

            }

            destrange = _Worksheet.get_Range(String.Format("A6:I{0}", len + 5));
            destrange.Value2 = vallist;

            // WriteRutStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A6:I{0}", len + 5));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 6, 2, 9, true);
                // GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        //dis
        public static void OutputZYDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\路面病害明细\按路线--路面病害明细表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\路面病害明细";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "按路线--路面病害统计明细表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["路面病害明细表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青路面病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥路面病害汇总表"] as MSExcel.Worksheet;
            WriteZYDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, _RoadDisList, _RoadPart);
            WriteZYDisHZ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjinfo, prjdir, _RoadPart, _RoadDisList);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYDisLB2Xls_roadpart(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist, List<MilePart> roadpart)
        {
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0, troadtype = -1;
            object[,] val = new object[len, 11];
            _Worksheet.Cells[1, 1] = prjinfo._District + " - " + "路面病害明细表";
            foreach (Disease tdis in dislist)
            {
                for (int k = 0; k < roadpart.Count - 1; ++k)
                {
                    if ((prjinfo._Direction > 0 && roadpart[k].mile <= tdis.m_mile && tdis.m_mile < roadpart[k + 1].mile)
                      || (prjinfo._Direction < 0 && roadpart[k].mile >= tdis.m_mile && tdis.m_mile > roadpart[k + 1].mile))
                    {
                        string[] s;
                        s = tdis.RoadDisType.Split('.');
                        troadtype = RoadDiseaseTypes.roadtypedict[tdis.RoadType];
                        if (troadtype == roadpart[k].roadtype)
                        {
                            val[i, 0] = prjinfo._Direction > 0 ? "上行" : "下行";
                            val[i, 1] = prjinfo._RoadNum;
                            val[i, 2] = tdis.m_mile;
                            val[i, 3] = tdis.RoadType;
                            if (s.Length > 1)
                            {
                                val[i, 4] = s[0];
                                val[i, 5] = s[1];
                            }
                            else
                            {
                                val[i, 4] = tdis.RoadDisType;
                                val[i, 5] = "无";
                            }

                            val[i, 6] = tdis.calcheight;
                            val[i, 7] = tdis.calcwidth;
                            val[i, 8] = tdis.depth;
                            val[i, 9] = tdis.Area;

                            troadtype = -1;
                            ++i;
                            break;
                        }
                    }
                }
            }
            int tlen = 0;
            for (int k = 0; k < len; ++k)
            {
                if (val[k, 0] == null)
                {
                    break;
                }
                tlen++;
            }
            destrange = _Worksheet.get_Range(String.Format("A3:K{0}", tlen + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", tlen + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 11, true);
            }
        }
        public static void WriteZYDisHZ2Xls(MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis)
        {
            MSExcel.Range destrange;
            int disnum = 0;
            object[,] disval;

            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int rowcnt_sn_s = 4;
            int rowcnt_sn_e = 4;
            int rowcnt_lq_s = 4;
            int rowcnt_lq_e = 4;

            double partarea = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[1, 2];

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);
                partarea = _RoadConfig.DetectWidth * milelength;

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

                if (roadpart[i].roadtype == 0) //0--沥青
                {
                    Haslqflag = true;
                    vallist[0, 0] = i + 1;
                    vallist[0, 1] = roadpart[i].mile.ToString("K0+000") + "-" + roadpart[i + 1].mile.ToString("K0+000");
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];
                    for (int di = 0; di < disnum; di++)
                    {
                        disval[0, di] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    destrange = worksheet_lqhz.get_Range(string.Format("C{0}:O{0}", rowcnt_lq_s));
                    destrange.Value2 = disval;
                    destrange = worksheet_lqhz.get_Range(string.Format("A{0}:B{0}", rowcnt_lq_s));
                    destrange.Value2 = vallist;

                    ++rowcnt_lq_e;
                    rowcnt_lq_s = rowcnt_lq_e;
                }
                else if (roadpart[i].roadtype == 1)
                {
                    Hassnflag = true;
                    vallist[0, 0] = i + 1;
                    vallist[0, 1] = roadpart[i].mile.ToString("K0+000") + "-" + roadpart[i + 1].mile.ToString("K0+000");
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum];
                    for (int di = 0; di < disnum; di++)
                    {
                        disval[0, di] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    destrange = worksheet_snhz.get_Range(string.Format("C{0}:P{0}", rowcnt_sn_s));
                    destrange.Value2 = disval;
                    destrange = worksheet_snhz.get_Range(string.Format("A{0}:B{0}", rowcnt_sn_s));
                    destrange.Value2 = vallist;

                    ++rowcnt_sn_e;
                    rowcnt_sn_s = rowcnt_sn_e;
                }
            }

            //总计
            //水泥
            destrange = worksheet_snhz.get_Range(String.Format("A4:P{0}", rowcnt_sn_e - 1));
            GlobalExcel.SetBorderLine(destrange, 53);

            //沥青

            destrange = worksheet_lqhz.get_Range(String.Format("A4:O{0}", rowcnt_lq_e - 1));
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
        // iri  按单元--平整度状况与评定表
        public static void OutputZYIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\平整度\按单元--平整度状况与评定表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\平整度";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "按单元-IRI-平整度状况与评定表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteZYIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 10];
            // 南湖大道 - 平整度状况与评定表
            _Worksheet.Cells[1, 1] = prjinfo._District + " - " + "平整度状况与评定表";
            _Worksheet.Cells[3, 2] = Math.Abs(prjinfo._StartMile - prjinfo._EndMile);
            _Worksheet.Cells[2, 1] = prjinfo._StartMile;
            _Worksheet.Cells[2, 3] = prjinfo._EndMile;
            _Worksheet.Cells[3, 4] = _RoadConfig.DetectWidth;
            _Worksheet.Cells[3, 7] = 设置单元编号.roadnum;
            int UnitNum = int.Parse(设置单元编号.unitnum);
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = string.Format("{0:000000}-{1}", UnitNum++, prjinfo._RoadNum);
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;
                vallist[i, 3] = Math.Abs(roadpart[i + 1].mile - roadpart[i].mile);
                vallist[i, 4] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype] + "路面";
                vallist[i, 5] = LIRIVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 6] = RIRIVal[i];
                    vallist[i, 7] = String.Format("=ROUND((F{0}+G{0})/2,5)", i + 6);
                }
                else
                {
                    vallist[i, 7] = String.Format("=ROUND((F{0}),5)", i + 6);
                }
                vallist[i, 8] = String.Format("=IF({0}+{1}*H{2}>=0,{0}+{1}*H{2},0)", _RQIa[0], _RQIa[1], i + 6);
                vallist[i, 9] = string.Format("=IF(I{0}>={1},\"A\",IF(I{0}>={2},\"B\",IF(I{0}>={3},\"C\",\"D\")))",
                    i + 6, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);

            }

            destrange = _Worksheet.get_Range(String.Format("A6:J{0}", len + 5));
            destrange.Value2 = vallist;

            destrange = _Worksheet.get_Range(String.Format("A1:J{0}", len + 5));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 6, 2, 10, true);
                // GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }

        //mtd
        public static void OutputZYMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\磨耗\按单元--磨耗状况与评定表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\磨耗";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "按单元-MTD-磨耗状况与评定表");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteZYMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal);
            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 10];
            // 南湖大道 - 平整度状况与评定表
            _Worksheet.Cells[1, 1] = prjinfo._District + " - " + "磨耗状况与评定表";
            _Worksheet.Cells[3, 2] = Math.Abs(prjinfo._StartMile - prjinfo._EndMile);
            _Worksheet.Cells[2, 1] = prjinfo._StartMile;
            _Worksheet.Cells[2, 3] = prjinfo._EndMile;
            _Worksheet.Cells[3, 4] = _RoadConfig.DetectWidth;
            _Worksheet.Cells[3, 7] = 设置单元编号.roadnum;
            int UnitNum = int.Parse(设置单元编号.unitnum);
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = string.Format("{0:000000}-{1}", UnitNum++, prjinfo._RoadNum);
                vallist[i, 1] = roadpart[i].mile;
                vallist[i, 2] = roadpart[i + 1].mile;
                vallist[i, 3] = Math.Abs(roadpart[i + 1].mile - roadpart[i].mile);
                vallist[i, 4] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype] + "路面";
                vallist[i, 5] = LMTDVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 6] = RMTDVal[i];
                    vallist[i, 7] = String.Format("=ROUND((F{0}+G{0})/2,5)", i + 6);
                }
                else
                {
                    vallist[i, 7] = String.Format("=ROUND(F{0},5)", i + 6);
                }
                vallist[i, 8] = string.Format("=IF(H{0}>={1},\"A\",IF(H{0}>={2},\"B\",IF(H{0}>={3},\"C\",\"D\")))",
                    i + 6, _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);

            }

            destrange = _Worksheet.get_Range(String.Format("A6:I{0}", len + 5));
            destrange.Value2 = vallist;

            destrange = _Worksheet.get_Range(String.Format("A1:I{0}", len + 5));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 6, 2, 10, true);
            }
        }

        //pqi
        public static void OutputZYPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            //武汉+南湖大道-上行-右2车道-20170925073424-按路线--车行道技术状况检测结果明细表.xlsx
            string direction = prjinfo._Direction > 0 ? "上行" : "下行";
            string srcxls = string.Format(@"{0}\报表模板\城镇道路\模板1\综合评价\路面技术状况评定结果.xlsx",
               System.Windows.Forms.Application.StartupPath);
            string dirpath = path + "\\综合评价";
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }
            string Destxls = string.Format(@"{0}\{1}+{2}-{3}-{4}车道-{5}-{6}.xlsx", dirpath, prjinfo._City, prjinfo._District, direction, prjinfo._RoadNum, prjinfo._DataDate + prjinfo._DataTime, "按路线-PQI-路面技术状况评定结果");
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet _Worksheet = null;
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Worksheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

            WriteZYPQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZYPQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
              List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 15];
            double pcival = 0, irival = 0;

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {

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

                vallist[i, 0] = prjinfo._RoadCode;
                vallist[i, 1] = prjinfo._RoadName;
                vallist[i, 2] = smile;   //起点桩号
                vallist[i, 3] = emile;   //终点桩号
                vallist[i, 4] = Math.Abs(emile - smile); //检测长度
                vallist[i, 5] = _RoadConfig.DetectWidth; //路面宽度
                vallist[i, 6] = 设置单元编号.roadnum;
                vallist[i, 7] = Math.Abs(emile - smile) * _RoadConfig.DetectWidth; //检测面积
                //PCI
                vallist[i, 14] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                pcival = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 8] = Math.Round(pcival, 5);
                vallist[i, 9] = string.Format("=IF(I{0}>={1},\"A\",IF(I{0}>={2},\"B\",IF(I{0}>={3},\"C\",\"D\")))",
                    i + 4, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                //RQI
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) / 2, 5);
                }
                else
                {
                    irival = LIRIVal[i];
                }
                vallist[i, 10] = String.Format("=IF({0}+{1}*{2}>=0,{0}+{1}*{2},0)", _RQIa[0], _RQIa[1], irival);
                vallist[i, 11] = string.Format("=IF(K{0}>={1},\"A\",IF(K{0}>={2},\"B\",IF(K{0}>={3},\"C\",\"D\")))",
                    i + 4, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);

                //PQI
                vallist[i, 12] = string.Format("=ROUND({1}*I{0}+{2}*{3}*K{0},5)", i + 4, _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);
                vallist[i, 13] = string.Format("=IF(M{0}>={1},\"A\",IF(M{0}>={2},\"B\",IF(M{0}>={3},\"C\",\"D\")))",
                    i + 4, _PQIGrade[roadpart[i].roaddegree][0], _PQIGrade[roadpart[i].roaddegree][1], _PQIGrade[roadpart[i].roaddegree][2]);


            }
            destrange = worksheet.get_Range(String.Format("A4:O{0}", len + 3));
            destrange.Value2 = vallist;
            // WriteZYPCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A3:O{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 5, 2, 7, true);
                // GlobalExcel.Reflection(worksheet, 5, 1, 2, false);
            }
        }
        #endregion

        #region 上海浦公报表模板，模板5 
        public static void OutputSHPG2Xls(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval, LaneProjectClass laneinfo = null)
        {
            bool Hassnflag = false;
            bool Haslqflag = false;

            string srcxls = null;
            string Destxls = string.Format(@"{0}\{1}.xlsx", path, prjdir.Name);
            MSExcel.Workbook _Workbook = null;

            if (disval == 10)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_iri = _Workbook.Sheets["IRI-10m"] as MSExcel.Worksheet;
                WriteIRI2Xls_SHPG(_Worksheet_iri, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal, _MarkVal, _GPSInfo, _SpeedVal);

                MSExcel.Worksheet _Worksheet_mtd = _Workbook.Sheets["TD-10m"] as MSExcel.Worksheet;
                WriteMTD2Xls_SHPG(_Worksheet_mtd, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal, _MarkVal, _GPSInfo, _SpeedVal);
            }
            else if (disval == 1)
            {
                _Workbook = excelApp.Workbooks.Open(Destxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_rd = _Workbook.Sheets["RD-1m"] as MSExcel.Worksheet;
                WriteRut2Xls_SHPG(_Worksheet_rd, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);
            }
            else
            {
                if (laneinfo != null&&laneinfo.m_wcDataClasses.Count > 0)
                {
                    srcxls = string.Format(@"{0}\报表模板\城镇道路\模板5_弯沉.xlsx", System.Windows.Forms.Application.StartupPath);

                }
                else
                {
                    srcxls = string.Format(@"{0}\报表模板\城镇道路\模板5.xlsx", System.Windows.Forms.Application.StartupPath);

                }
                _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                MSExcel.Worksheet _Worksheet_prj = _Workbook.Sheets["控制信息"] as MSExcel.Worksheet;

                if (laneinfo != null)
                {
                    WritePrjInfo2Xls(_Worksheet_prj, prjinfo, laneinfo, path);
                   
                   
                }
                else
                {
                    WritePrjInfo2Xls(_Worksheet_prj, prjinfo);
                }
               
                MSExcel.Worksheet _Worksheet_unit = _Workbook.Sheets["单元划分"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_lq = _Workbook.Sheets["病害面积计算（沥青）"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_sn = _Workbook.Sheets["病害面积计算（水泥）"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_hz = _Workbook.Sheets["指标汇总"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_rqi = _Workbook.Sheets["RQI"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_pci = _Workbook.Sheets["PCI"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_pqi = _Workbook.Sheets["PQI"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_rd = _Workbook.Sheets["RD"] as MSExcel.Worksheet;
                MSExcel.Worksheet _Worksheet_td = _Workbook.Sheets["TD"] as MSExcel.Worksheet;
                if (laneinfo != null)
                {
                    WriteSHPG2Xls(_Worksheet_unit, _Worksheet_lq, _Worksheet_sn, _Worksheet_hz,
                                      _Worksheet_rqi, _Worksheet_pci, _Worksheet_pqi, _Worksheet_rd, _Worksheet_td,
                                      prjinfo, prjdir, _RoadPart, _RoadDisList,
                                      _LIRIMeanVal, _RIRIMeanVal,
                                      _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                                      _LMTDMeanVal, _RMTDMeanVal, _GPSInfo, _MarkVal,
                                      ref Hassnflag, ref Haslqflag,laneinfo);
                }
                else
                {
                    WriteSHPG2Xls(_Worksheet_unit, _Worksheet_lq, _Worksheet_sn, _Worksheet_hz,
                  _Worksheet_rqi, _Worksheet_pci, _Worksheet_pqi, _Worksheet_rd, _Worksheet_td,
                  prjinfo, prjdir, _RoadPart, _RoadDisList,
                  _LIRIMeanVal, _RIRIMeanVal,
                  _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                  _LMTDMeanVal, _RMTDMeanVal, _GPSInfo, _MarkVal,
                  ref Hassnflag, ref Haslqflag);
                }
              
               

                MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害明细"] as MSExcel.Worksheet;
                WriteGPSDis2Xls_SHPG(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);

                MSExcel.Worksheet _Worksheet2 = _Workbook.Sheets["景观图像"] as MSExcel.Worksheet;
                WriteGPSImg2Xls_SHPG(_Worksheet2, prjinfo, prjdir, "Street", prjinfo._StreetImgDis_Left);
                if (laneinfo != null)
                {
                    WriteTj2Xls_SHPG(excelApp, _Workbook,laneinfo);

                }
                else
                {
                    WriteTj2Xls_SHPG(excelApp, _Workbook);

                }
                if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)//弯沉为false 需要删除部分列
                {
                    MSExcel.Worksheet _WorksheetWc = _Workbook.Sheets["弯沉"] as MSExcel.Worksheet;
                    WriteWc2Xls_SHPG(_WorksheetWc,prjinfo, laneinfo, _RoadPart,_MarkVal);
                    int value = 0;
                    for (int i = 0; i < laneinfo.m_wcDataClasses[0].wcDatas.Rows.Count; i++)
                    {
                      int smile= int.Parse( laneinfo.m_wcDataClasses[0].wcDatas.Rows[0]["Mile"].ToString());
                      int eMile = int.Parse(laneinfo.m_wcDataClasses[0].wcDatas.Rows[1]["Mile"].ToString());
                      value = Math.Abs(smile - eMile); 
                    }
                    MSExcel.Worksheet newSheet = (MSExcel.Worksheet)_Workbook.Sheets.Add(Type.Missing, _Workbook.Sheets[_Workbook.Sheets.Count-1], 1, Type.Missing);
                    newSheet.Name = $"弯沉-{value}m";
                    WriteWcData2Xls_SHPG(newSheet, prjinfo, laneinfo, _RoadPart, _MarkVal);
                }

                if (laneinfo!=null && laneinfo.m_wcDataClasses.Count==0)//弯沉为false 需要删除部分列
                {
                    //删除指标汇总 Y,Z列
                    // 获取Y列的Range对象
                    MSExcel.Range rangeY = _Worksheet_hz.Columns["Y"] as MSExcel.Range;
                    // 删除Y列的数据
                    rangeY.Delete();
                   rangeY = _Worksheet_hz.Columns["Y"] as MSExcel.Range;
                    // 删除Y列的数据
                    rangeY.Delete();

                    rangeY = _Worksheet_unit.Columns["H"] as MSExcel.Range;
                    rangeY.Delete();
                    rangeY = _Worksheet_unit.Columns["H"] as MSExcel.Range;
                    rangeY.Delete();

                   
                     
                    //删除弯沉sheet页面
                    // 遍历所有工作表
                    foreach (MSExcel.Worksheet sheet in _Workbook.Sheets)
                    {
                        // 检查工作表名称
                        if (sheet.Name == "弯沉")
                        {
                            // 删除工作表
                            sheet.Delete(); 
                            break;
                        }
                    } 

                }

                if (laneinfo != null)
                {
                    string tstr = ((MSExcel.Range)_Worksheet_prj.Cells[7, 8]).Text.ToString();
                    if (tstr != laneinfo.m_lane.m_pavementtype
                        || tstr != laneinfo.m_lane.m_roadpartinfo.m_type
                        || tstr != laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadtype)
                    {
                        string info = "车道ID=" + laneinfo.m_lane.m_id
                            + "，实际工程记录路面材质为" + tstr
                            + "，基础信息导入的车道路面材质为" + laneinfo.m_lane.m_pavementtype
                            + "，路段路面材质为" + laneinfo.m_lane.m_roadpartinfo.m_type
                            + "，道路路面材质为" + laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadtype + "\n";
                        File.AppendAllText(path.Replace("\\车道报表数据", "\\路面材质不一致记录.txt"), info);
                    }
                   
                }

            }

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WritePrjInfo2Xls(MSExcel.Worksheet _Worksheet_prj, ProjectInfo prjinfo)
        {
            _Worksheet_prj.Cells[3, 2] = prjinfo._Province;
            _Worksheet_prj.Cells[4, 2] = prjinfo._City;
            _Worksheet_prj.Cells[5, 2] = prjinfo._District;

            char[] splitchr = { '(', '（', '-', '_', ')', '）' };  
            string[] strs1 = prjinfo._RoadName.Split(splitchr, System.StringSplitOptions.RemoveEmptyEntries);
            if (strs1.Length > 0)
            {
                _Worksheet_prj.Cells[8, 2] = strs1[0];
                _Worksheet_prj.Cells[9, 2] = prjinfo._RoadCode;
                _Worksheet_prj.Cells[11, 2] = prjinfo._RoadGrade;

                if (strs1.Length > 2)
                {
                    if (prjinfo._Direction > 0)
                    {
                        _Worksheet_prj.Cells[17, 2] = strs1[1];
                        _Worksheet_prj.Cells[18, 2] = strs1[2];
                        _Worksheet_prj.Cells[19, 2] = prjinfo._StartMile;
                        _Worksheet_prj.Cells[20, 2] = prjinfo._EndMile;
                    }
                    else
                    {
                        _Worksheet_prj.Cells[17, 2] = strs1[2];
                        _Worksheet_prj.Cells[18, 2] = strs1[1];
                        _Worksheet_prj.Cells[19, 2] = prjinfo._EndMile;
                        _Worksheet_prj.Cells[20, 2] = prjinfo._StartMile;
                    }
                }
            }

            _Worksheet_prj.Cells[4, 8] = prjinfo._Direction > 0 ? "上行" : "下行";
            _Worksheet_prj.Cells[5, 8] = prjinfo._RoadNum;
            _Worksheet_prj.Cells[8, 8] = prjinfo._DataDate;
            _Worksheet_prj.Cells[5, 8] = _RoadConfig.DetectWidth;
        }

        private static void WritePrjInfo2Xls(MSExcel.Worksheet _Worksheet_prj, ProjectInfo prjinfo, 
            LaneProjectClass laneinfo = null, string path = null)
        {
            object[,] obj_roadinfo = new object[32, 1];
            object[,] obj_roadpartinfo = new object[12, 1];
            object[,] obj_laneinfo = new object[11, 1];
            object[,] obj_projectinfo = new object[13, 1];
            object[,] obj_reportinfo = new object[5, 1];
            object[,] obj_indexinfo = new object[laneinfo.m_report.m_projectinfo.m_indexlist.Count, 5];

            int rowidx = 0;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_id;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_province;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_city;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_district;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_town;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_village;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_name;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_code;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_properity;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_grade;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_length;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_width;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadway_area;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_sidewalk_area;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadtype;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadstartlocation;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadendlocation;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadsartmile;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_roadendmile;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_buildyear;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_buildunit;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_designunit;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_constructionunit;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_controlunit;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_managementunit_province;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_managementunit_city;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_managementunit_district;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_managementunit_department;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_maintenance_center;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_maintenance_section;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_maintenance_unit;
            obj_roadinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_roadinfo.m_project_department;

            rowidx = 0;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_id;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_startlocation;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_endlocation;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_startmile;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_endmile;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_part_grade;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_length;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_width;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_area;
            obj_roadpartinfo[rowidx++, 0] = laneinfo.m_lane.m_roadpartinfo.m_type;
            obj_roadpartinfo[rowidx++, 0] = "";//路段gps坐标
            if (laneinfo.m_wcDataClasses.Count>0)
            {
                obj_roadpartinfo[rowidx++, 0] = laneinfo.m_wcDataClasses[0].traffic;//交通量等级 
            }

            rowidx = 0;
            obj_laneinfo[rowidx++, 0] = laneinfo.m_lane.m_id;
            obj_laneinfo[rowidx++, 0] = laneinfo.m_lane.m_roadfunctiontype;
            obj_laneinfo[rowidx++, 0] = prjinfo._Direction > 0 ? "上行" : "下行";
            obj_laneinfo[rowidx++, 0] = prjinfo._RoadNum;
            obj_laneinfo[rowidx++, 0] = laneinfo.m_lane.m_width;
            obj_laneinfo[rowidx++, 0] = "=IF(AND('病害面积计算（沥青）'!D3>0,'病害面积计算（水泥）'!D3>0),\"混合\",IF('病害面积计算（沥青）'!D3>0,\"沥青\",IF('病害面积计算（水泥）'!D3>0,\"水泥\",\"\")))";
            obj_laneinfo[rowidx++, 0] = prjinfo._DataDate;
            if (prjinfo._Direction > 0)
            {
                obj_laneinfo[rowidx++, 0] = prjinfo._StartMile;
                obj_laneinfo[rowidx++, 0] = prjinfo._EndMile;
            }
            else
            {
                obj_laneinfo[rowidx++, 0] = prjinfo._EndMile;
                obj_laneinfo[rowidx++, 0] = prjinfo._StartMile;
            }
            obj_laneinfo[rowidx++, 0] = laneinfo.m_lane.m_carwaytype;
            if (laneinfo.m_wcDataClasses.Count > 0)
            {

                string wcTime = laneinfo.m_wcDataClasses.First().time;
                string[] timeSplit = wcTime.Split('/');
                if (timeSplit.Length > 1)
                {
                    string tempStr = timeSplit[0] + "年" + timeSplit[1] + "月" + timeSplit[2] + "日";

                    obj_laneinfo[rowidx++, 0] = tempStr;
                }
                else
                {
                    timeSplit = wcTime.Split('-');
                    if (timeSplit.Length>1)
                    {

                        string tempStr = timeSplit[0] + "年" + timeSplit[1] + "月" + timeSplit[2] + "日";

                        obj_laneinfo[rowidx++, 0] = tempStr;
                    }
                    else
                    {
                        MessageBox.Show("弯沉时间解析错误,请确定时间格式正确");
                        return;
                    }
                }
               

            }

            if (Convert.ToDouble(laneinfo.m_lane.m_width) > 0)
            {
                _RoadConfig.DetectWidth = Convert.ToDouble(laneinfo.m_lane.m_width);
            }

            rowidx = 0;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_id;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_project_name;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_entrust_client;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_entrust_serial;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_contract_num;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_entrust_date;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_testing_unit;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_project_dutyperson;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_testing_start_date;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_testing_end_date;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_testing_standard.m_name;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_testing_standard.m_code;
            obj_projectinfo[rowidx++, 0] = laneinfo.m_report.m_projectinfo.m_date;

            rowidx = 0;
            obj_reportinfo[rowidx++, 0] = laneinfo.m_report.m_report_num;
            obj_reportinfo[rowidx++, 0] = laneinfo.m_report.m_report_name;
            obj_reportinfo[rowidx++, 0] = laneinfo.m_report.m_project_name;
            obj_reportinfo[rowidx++, 0] = laneinfo.m_report.m_report_start_date;
            obj_reportinfo[rowidx++, 0] = laneinfo.m_report.m_report_end_date;

            rowidx = 0;
            foreach (IndexInfoClass tindex in laneinfo.m_report.m_projectinfo.m_indexlist)
            {
                if (tindex.m_index == "弯沉" )
                {
                    if (laneinfo.m_wcDataClasses.Count  == 0 )
                    {
                        tindex.m_tesing = "否";
                    }
                }
                obj_indexinfo[rowidx, 0] = tindex.m_id;
                obj_indexinfo[rowidx, 1] = tindex.m_name;
                obj_indexinfo[rowidx, 2] = tindex.m_index;
                obj_indexinfo[rowidx, 3] = tindex.m_pavementtype;
                obj_indexinfo[rowidx, 4] = tindex.m_tesing;
                ++rowidx;
            }

            MSExcel.Range destrange = null;
            destrange = _Worksheet_prj.get_Range("B2:B33");
            destrange.Value2 = obj_roadinfo;

            destrange = _Worksheet_prj.get_Range("E2:E13");
            destrange.Value2 = obj_roadpartinfo;
            if (laneinfo.m_wcDataClasses.Count > 0)
                destrange = _Worksheet_prj.get_Range("H2:H12");
            else
            {
                destrange = _Worksheet_prj.get_Range("H2:H11");

            }
            destrange.Value2 = obj_laneinfo;

            destrange = _Worksheet_prj.get_Range("J3:N" + (laneinfo.m_report.m_projectinfo.m_indexlist.Count + 2).ToString());
            destrange.Value2 = obj_indexinfo;
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = _Worksheet_prj.get_Range("Q2:Q14");
            destrange.Value2 = obj_projectinfo;

            destrange = _Worksheet_prj.get_Range("T2:T6");
            destrange.Value2 = obj_reportinfo;
        }
       
        private static void WriteSHPG2Xls(MSExcel.Worksheet _Worksheet_unit,
            MSExcel.Worksheet _Worksheet_lq, MSExcel.Worksheet _Worksheet_sn,
            MSExcel.Worksheet _Worksheet_hz, MSExcel.Worksheet _Worksheet_rqi,
            MSExcel.Worksheet _Worksheet_pci, MSExcel.Worksheet _Worksheet_pqi,
            MSExcel.Worksheet _Worksheet_rd, MSExcel.Worksheet _Worksheet_td,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal,
            double[] LRutVal, double[] RRutVal, double[] SRutVal,
            double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo, string[] MarkVal,
            ref bool Hassnflag, ref bool Haslqflag, LaneProjectClass laneinfo = null)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;

            int rowcnt = 0;
            int rowcnt_sn_s = 0, rowcnt_lq_s = 0;
            int rownum_sn = 0, rownum_lq = 0;

            for (int i = 0; i < len; ++i)
            {
                if (roadpart[i].roadtype == 0)
                {
                    ++rownum_lq;
                }
                else if (roadpart[i].roadtype == 1)
                {
                    ++rownum_sn;
                }
            }
            if (rownum_lq > 0)
            {
                Haslqflag = true;
            }

            if (rownum_sn > 0)
            {
                Hassnflag = true;
            }

            object[,] obj_unit = new object[len, 9];
            object[,] obj_hz = new object[len, 27];
            object[,] obj_rqi = new object[len, 11];
            object[,] obj_pci = new object[len, 9];
            object[,] obj_pqi = new object[len, 8];
            object[,] obj_rd = new object[len, 9];
            object[,] obj_td = new object[len, 10];
            object[,] obj_lq = new object[rownum_lq, 5 + RoadDiseaseTypes.DiseaseTypeDict[0].Count];
            object[,] obj_sn = new object[rownum_sn, 5 + RoadDiseaseTypes.DiseaseTypeDict[1].Count];

            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];

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

                //单元划分表
                obj_unit[rowcnt, 0] = i + 1;
                obj_unit[rowcnt, 1] = smile;
                obj_unit[rowcnt, 2] = emile;
                obj_unit[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_unit[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_unit[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                obj_unit[rowcnt, 6] = _RoadConfig.DetectWidth;
                if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
                {
                    obj_unit[rowcnt, 7] = laneinfo.m_wcDataClasses[0].traffic;
                    try
                    {
                        if (prjinfo._Direction < 0)
                        {
                            //obj_unit[rowcnt, 8] = laneinfo.m_wcDataClasses[0].wcResultDatas.Rows[len - 1 - i]["基层类型"];
                            obj_unit[rowcnt, 8] = laneinfo.m_wcDataClasses[0].wcResultDatas.Rows[i]["基层类型"]; 


                        }
                        else
                        {
                            obj_unit[rowcnt, 8] = laneinfo.m_wcDataClasses[0].wcResultDatas.Rows[i]["基层类型"];
                             

                        }
                    }
                    catch (Exception ex)
                    {
                        string msg = prjdir.Name +"\n文件解析错误: " + ex.Message+ "\n请检查工程数据【单元划分】与对应弯沉数据【单元信息】是否一致！";

                        throw new Exception(msg);
                    }
                  
                }

                //RQI
                obj_rqi[rowcnt, 0] = i + 1;
                obj_rqi[rowcnt, 1] = smile;
                obj_rqi[rowcnt, 2] = emile;
                obj_rqi[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_rqi[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_rqi[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                obj_rqi[rowcnt, 6] = LIRIVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    obj_rqi[rowcnt, 7] = RIRIVal[i];
                    //obj_rqi[rowcnt, 8] = String.Format("=ROUND((G{0}+H{0})/2,5)", rowcnt + 2);
                }
                else
                {
                    obj_rqi[rowcnt, 7] = "";
                    //obj_rqi[rowcnt, 8] = String.Format("=ROUND(G{0},5)", rowcnt + 2);
                }
                obj_rqi[rowcnt, 8] = String.Format("=IF(OR(ISBLANK(G{0}),G{0}=0),IF(OR(ISBLANK(H{0}),H{0}=0),0,ROUND(VALUE(H{0}),5)),ROUND(AVERAGE(G{0},H{0}),5))", rowcnt + 2);
                obj_rqi[rowcnt, 9] = String.Format("=IF({0}+{1}*I{2}>=0,{0}+{1}*I{2},0)", _RQIa[0], _RQIa[1], rowcnt + 2);
                obj_rqi[rowcnt, 10] = string.Format("=IF(J{0}>={1},\"A\",IF(J{0}>={2},\"B\",IF(J{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);

                //RD
                obj_rd[rowcnt, 0] = i + 1;
                obj_rd[rowcnt, 1] = smile;
                obj_rd[rowcnt, 2] = emile;
                obj_rd[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_rd[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_rd[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                if (prjinfo._IsRut)
                {
                    obj_rd[rowcnt, 6] = LRutVal[i];
                    obj_rd[rowcnt, 7] = RRutVal[i];
                    obj_rd[rowcnt, 8] = SRutVal[i];
                }

                //TD
                obj_td[rowcnt, 0] = i + 1;
                obj_td[rowcnt, 1] = smile;
                obj_td[rowcnt, 2] = emile;
                obj_td[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_td[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_td[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                obj_td[rowcnt, 6] = LMTDVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    obj_td[rowcnt, 7] = RMTDVal[i];
                    //obj_td[rowcnt, 8] = String.Format("=ROUND((G{0}+H{0})/2,5)", rowcnt + 2);
                }
                else
                {
                    obj_td[rowcnt, 7] = "";
                    //obj_td[rowcnt, 8] = String.Format("=ROUND(G{0},5)", rowcnt + 2);
                }
                obj_td[rowcnt, 8] = String.Format("=IF(OR(ISBLANK(G{0}),G{0}=0),IF(OR(ISBLANK(H{0}),H{0}=0),0,ROUND(VALUE(H{0}),5)),ROUND(AVERAGE(G{0},H{0}),5))", rowcnt + 2);
                //obj_td[rowcnt, 9] = string.Format("=IF(I{0}>={1},\"A\",IF(I{0}>={2},\"B\",IF(I{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                //    _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);
                obj_td[rowcnt, 9] = string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",IF(F{0}=\"水泥\",\"/\", IF(I{0}>={1},\"A\",IF(I{0}>={2},\"B\",IF(I{0}>={3},\"C\",\"D\")))))", rowcnt + 2,
                    _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);


                //PCI
                obj_pci[rowcnt, 0] = i + 1;
                obj_pci[rowcnt, 1] = smile;
                obj_pci[rowcnt, 2] = emile;
                obj_pci[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_pci[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_pci[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                obj_pci[rowcnt, 6] = 100 - ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                obj_pci[rowcnt, 7] = string.Format("=100 - G{0}", rowcnt + 2);
                obj_pci[rowcnt, 8] = string.Format("=IF(H{0}>={1},\"A\",IF(H{0}>={2},\"B\",IF(H{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                //PQI
                obj_pqi[rowcnt, 0] = i + 1;
                obj_pqi[rowcnt, 1] = smile;
                obj_pqi[rowcnt, 2] = emile;
                obj_pqi[rowcnt, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt + 2);
                obj_pqi[rowcnt, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                obj_pqi[rowcnt, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                //PQI计算公式与路面材质没关系，只和道路等级有关系，快速路和主干路，支路和次干路
                obj_pqi[rowcnt, 6] = string.Format("=ROUND({1}*PCI!H{0}+{2}*{3}*RQI!J{0},5)", rowcnt + 2,
                    _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);
                obj_pqi[rowcnt, 7] = string.Format("=IF(G{0}>={1},\"A\",IF(G{0}>={2},\"B\",IF(G{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _PQIGrade[roadpart[i].roaddegree][0], _PQIGrade[roadpart[i].roaddegree][1], _PQIGrade[roadpart[i].roaddegree][2]);

                //指标汇总
                obj_hz[rowcnt, 0] = i + 1;
                obj_hz[rowcnt, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                obj_hz[rowcnt, 2] = prjinfo._RoadNum;
                obj_hz[rowcnt, 3] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                //obj_hz[rowcnt, 4] = GPSInfo[i]._utctime;
                //obj_hz[rowcnt, 5] = smile;
                //obj_hz[rowcnt, 6] = GPSInfo[i]._longitude;
                //obj_hz[rowcnt, 7] = GPSInfo[i]._latitude;
                //obj_hz[rowcnt, 8] = GPSInfo[i + 1]._utctime;
                //obj_hz[rowcnt, 9] = emile;
                //obj_hz[rowcnt, 10] = GPSInfo[i + 1]._longitude;
                //obj_hz[rowcnt, 11] = GPSInfo[i + 1]._latitude;

                GPSStartObj[rowcnt, 0] = GPSInfo[i]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;

                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                obj_hz[rowcnt, 12] = string.Format("=ABS(F{0}-J{0})", rowcnt + 2);

                obj_hz[rowcnt, 13] = string.Format("=ROUND({1}*P{0}+{2}*{3}*R{0},5)", rowcnt + 2,
                    _PQIW[roadpart[i].roaddegree][0], _PQIW[roadpart[i].roaddegree][1], _PQIT);
                obj_hz[rowcnt, 14] = string.Format("=IF(N{0}>={1},\"A\",IF(N{0}>={2},\"B\",IF(N{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _PQIGrade[roadpart[i].roaddegree][0], _PQIGrade[roadpart[i].roaddegree][1], _PQIGrade[roadpart[i].roaddegree][2]);

                obj_hz[rowcnt, 15] = string.Format("=PCI!H{0}", rowcnt + 2);
                obj_hz[rowcnt, 16] = string.Format("=IF(P{0}>={1},\"A\",IF(P{0}>={2},\"B\",IF(P{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);

                obj_hz[rowcnt, 17] = String.Format("=IF({0}+{1}*T{2}>=0,{0}+{1}*T{2},0)", _RQIa[0], _RQIa[1], rowcnt + 2);
                obj_hz[rowcnt, 18] = string.Format("=IF(R{0}>={1},\"A\",IF(R{0}>={2},\"B\",IF(R{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                    _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);

                obj_hz[rowcnt, 19] = string.Format("=RQI!I{0}", rowcnt + 2);
                obj_hz[rowcnt, 20] = string.Format("=IF(T{0}<={1},\"A\",IF(T{0}<={2},\"B\",IF(T{0}<={3},\"C\",\"D\")))", rowcnt + 2,
                    _IRIGrade[roadpart[i].roaddegree][1], _IRIGrade[roadpart[i].roaddegree][2], _IRIGrade[roadpart[i].roaddegree][3]);

                obj_hz[rowcnt, 21] = string.Format("=RD!I{0}", rowcnt + 2);

                obj_hz[rowcnt, 22] = string.Format("=TD!I{0}", rowcnt + 2);
                //obj_hz[rowcnt, 23] = string.Format("=IF(W{0}>={1},\"A\",IF(W{0}>={2},\"B\",IF(W{0}>={3},\"C\",\"D\")))", rowcnt + 2,
                //    _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);
                obj_hz[rowcnt, 23] = string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",IF(D{0}=\"水泥\",\"/\", IF(W{0}>={1},\"A\",IF(W{0}>={2},\"B\",IF(W{0}>={3},\"C\",\"D\")))))", rowcnt + 2,
                    _MTDGrade[roadpart[i].roaddegree][0], _MTDGrade[roadpart[i].roaddegree][1], _MTDGrade[roadpart[i].roaddegree][2]);
                if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
                {
                    double wcValue = double.Parse(laneinfo.m_wcDataClasses.First().wcResultDatas.Rows[i]["弯沉值"].ToString());
                      
                        if (double.IsNaN(wcValue))
                        {
                            obj_hz[rowcnt, 24] = "/";
                            obj_hz[rowcnt, 25] = "/";

                        }
                        else
                        {
                            obj_hz[rowcnt, 24] = laneinfo.m_wcDataClasses.First().wcResultDatas.Rows[i]["弯沉值"];
                            obj_hz[rowcnt, 25] = laneinfo.m_wcDataClasses.First().wcResultDatas.Rows[i]["评价等级"];

                        }
                    
                     

                    
                }
                
                //病害面积
                if (roadpart[i].roadtype == 1)
                {
                    if (prjinfo._Direction > 0)
                        obj_sn[rowcnt_sn_s, 0] = i + 1;
                    else
                        obj_sn[rowcnt_sn_s, 0] = len - i;
                    obj_sn[rowcnt_sn_s, 1] = smile;
                    obj_sn[rowcnt_sn_s, 2] = emile;
                    obj_sn[rowcnt_sn_s, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt_sn_s + 3);
                    obj_sn[rowcnt_sn_s, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                    for (int di = 0, kk = 0; di < RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count; ++di, ++kk)
                    {
                        obj_sn[rowcnt_sn_s, kk + 5] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    ++rowcnt_sn_s;
                }
                else if (roadpart[i].roadtype == 0)
                {
                    if (prjinfo._Direction > 0)
                        obj_lq[rowcnt_lq_s, 0] = i + 1;
                    else
                        obj_lq[rowcnt_lq_s, 0] = len - i;
                    obj_lq[rowcnt_lq_s, 1] = smile;
                    obj_lq[rowcnt_lq_s, 2] = emile;
                    obj_lq[rowcnt_lq_s, 3] = string.Format("=ABS(C{0}-B{0})", rowcnt_lq_s + 3);
                    obj_lq[rowcnt_lq_s, 4] = MarkVal[i] == "路口单元" ? "是" : "否";
                    for (int di = 0, kk = 0; di < RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count; ++di, ++kk)
                    {
                        obj_lq[rowcnt_lq_s, kk + 5] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    ++rowcnt_lq_s;
                }

                ++rowcnt;
            }

            //写入xls
            MSExcel.Range sortrange = null;
            MSExcel.Range destrange = null;

            //单元划分
            destrange = _Worksheet_unit.get_Range(string.Format("A2:I{0}", rowcnt + 1));
            destrange.Value2 = obj_unit;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_unit.get_Range(string.Format("B2:G{0}", rowcnt + 1));
                sortrange = _Worksheet_unit.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_unit, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_unit, 2, 2, 2, false);
            }

            //RQI
            destrange = _Worksheet_rqi.get_Range(string.Format("A2:K{0}", rowcnt + 1));
            destrange.Value2 = obj_rqi;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_rqi.get_Range(string.Format("B2:K{0}", rowcnt + 1));
                sortrange = _Worksheet_rqi.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_rqi, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_rqi, 2, 2, 2, false);
            }

            //PCI
            destrange = _Worksheet_pci.get_Range(string.Format("A2:I{0}", rowcnt + 1));
            destrange.Value2 = obj_pci;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_pci.get_Range(string.Format("B2:I{0}", rowcnt + 1));
                sortrange = _Worksheet_pci.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_pci, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_pci, 2, 2, 2, false);
            }

            //PQI
            destrange = _Worksheet_pqi.get_Range(string.Format("A2:H{0}", rowcnt + 1));
            destrange.Value2 = obj_pqi;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_pqi.get_Range(string.Format("B2:F{0}", rowcnt + 1));
                sortrange = _Worksheet_pqi.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_pqi, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_pqi, 2, 2, 2, false);
            }

            //RD
            destrange = _Worksheet_rd.get_Range(string.Format("A2:I{0}", rowcnt + 1));
            destrange.Value2 = obj_rd;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_rd.get_Range(string.Format("B2:I{0}", rowcnt + 1));
                sortrange = _Worksheet_rd.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_rd, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_rd, 2, 2, 2, false);
            }

            //TD
            destrange = _Worksheet_td.get_Range(string.Format("A2:J{0}", rowcnt + 1));
            destrange.Value2 = obj_td;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_td.get_Range(string.Format("B2:J{0}", rowcnt + 1));
                sortrange = _Worksheet_td.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_td, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet_td, 2, 2, 2, false);
            }

            //沥青病害
            if (Haslqflag)
            {
                destrange = _Worksheet_lq.get_Range(string.Format("A3:R{0}", rowcnt_lq_s + 2));
                destrange.Value2 = obj_lq;
                GlobalExcel.SetBorderLine(destrange, 53);
                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = _Worksheet_lq.get_Range(string.Format("A3:R{0}", rowcnt_lq_s + 2));
                    sortrange = _Worksheet_lq.get_Range(string.Format("B3:B{0}", rowcnt_lq_s + 2));
                    GlobalExcel.ReflectionColnum(_Worksheet_lq, destrange, sortrange);
                    GlobalExcel.Reflection(_Worksheet_lq, 3, 2, 2, false);
                }
            }
            //else
            //{
            //    _Worksheet_lq.Delete();
            //}

            //水泥病害
            if (Hassnflag)
            {
                destrange = _Worksheet_sn.get_Range(string.Format("A3:S{0}", rowcnt_sn_s + 2));
                destrange.Value2 = obj_sn;
                GlobalExcel.SetBorderLine(destrange, 53);
                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = _Worksheet_sn.get_Range(string.Format("A3:S{0}", rowcnt_sn_s + 2));
                    sortrange = _Worksheet_sn.get_Range(string.Format("B3:B{0}", rowcnt_sn_s + 2));
                    GlobalExcel.ReflectionColnum(_Worksheet_sn, destrange, sortrange);
                    GlobalExcel.Reflection(_Worksheet_sn, 3, 2, 2, false);
                }
            }
            //else
            //{
            //    _Worksheet_sn.Delete();
            //}

            //指标汇总
            destrange = _Worksheet_hz.get_Range(string.Format("A2:AA{0}", rowcnt + 1));
            destrange.Value2 = obj_hz;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_hz.get_Range(String.Format("E2:H{0}", len + 1));
                destrange.Value2 = GPSEndObj;
                destrange = _Worksheet_hz.get_Range(String.Format("I2:L{0}", len + 1));
                destrange.Value2 = GPSStartObj;

                destrange = _Worksheet_hz.get_Range(string.Format("B2:L{0}", rowcnt + 1));
                sortrange = _Worksheet_hz.get_Range(string.Format("F2:F{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_hz, destrange, sortrange);
            }
            else
            {
                destrange = _Worksheet_hz.get_Range(String.Format("E2:H{0}", len + 1));
                destrange.Value2 = GPSStartObj;
                destrange = _Worksheet_hz.get_Range(String.Format("I2:L{0}", len + 1));
                destrange.Value2 = GPSEndObj;
            }
        }

        private static void WriteGPSDis2Xls_SHPG(MSExcel.Worksheet _Worksheet_lb, ProjectInfo prjinfo, DirectoryInfo prjdir, Disease[] dislist, List<MilePart> roadpart)
        {
            int i = 0;
            int len = dislist.Length;
            object[,] disinfo = new object[len, 18];

            string[] gpsinfostrs = null;
            ExcelGPS[] tempinfos = null;
            if (File.Exists(prjdir.FullName + "\\GPS2Mile.txt"))
            {
                gpsinfostrs = File.ReadAllLines(prjdir.FullName + "\\GPS2Mile.txt");
                tempinfos = new ExcelGPS[gpsinfostrs.Length];
                for (i = 0; i < gpsinfostrs.Length; ++i)
                {
                    tempinfos[i] = new ExcelGPS(gpsinfostrs[i]);
                }
            }

            i = 0;
            int troadtype = -1;
            int gi = 0;
            ExcelGPS tempgpsinfo;
            foreach (Disease tdis in dislist)
            {
                for (int k = 0; k < roadpart.Count - 1; ++k)
                {
                    if ((prjinfo._Direction > 0 && roadpart[k].mile <= tdis.m_mile && tdis.m_mile < roadpart[k + 1].mile)
                      || (prjinfo._Direction < 0 && roadpart[k].mile >= tdis.m_mile && tdis.m_mile > roadpart[k + 1].mile))
                    {
                        troadtype = RoadDiseaseTypes.roadtypedict[tdis.RoadType];
                        if (troadtype == roadpart[k].roadtype)
                        {
                            // GPS信息
                            for (; gi < tempinfos.Length; ++gi)
                            {
                                if (prjinfo._Direction > 0)
                                {
                                    if (tempinfos[gi]._mile >= tdis.m_mile)
                                    {
                                        break;
                                    }
                                }
                                else
                                {
                                    if (tempinfos[gi]._mile <= tdis.m_mile)
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

                            disinfo[i, 0] = i + 1;
                            disinfo[i, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                            disinfo[i, 2] = prjinfo._RoadNum;
                            disinfo[i, 3] = tempgpsinfo._utctime;
                            //disinfo[i, 4] = tempgpsinfo._mile;
                            disinfo[i, 4] = tdis.m_mile;
                            disinfo[i, 5] = tempgpsinfo._longitude;
                            disinfo[i, 6] = tempgpsinfo._latitude;
                            disinfo[i, 7] = tdis.RoadType;
                            disinfo[i, 8] = tdis.RoadDisType;
                            disinfo[i, 9] = tdis.rect.Height * _RoadConfig.HeightScale;
                            disinfo[i, 10] = tdis.rect.Width * _RoadConfig.WidthScale;
                            disinfo[i, 11] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                            disinfo[i, 12] = tdis.calcheight;
                            disinfo[i, 13] = tdis.calcwidth;
                            if (tdis.depth > 0)
                            {
                                disinfo[i, 14] = tdis.depth;
                            }
                            else
                            {
                                disinfo[i, 14] = "/";
                            }
                            disinfo[i, 15] = tdis.Area;
                            disinfo[i, 16] = tdis.imgname;
                            disinfo[i, 17] = tdis.imgpath;
                            troadtype = -1;
                            ++i;
                            break;
                        }
                    }
                }
            }
            int tlen = 0;
            for (int k = 0; k < len; ++k)
            {
                if (disinfo[k, 0] == null)
                {
                    break;
                }
                tlen++;
            }

            //写入xls
            MSExcel.Range sortrange = null;
            MSExcel.Range destrange = null;

            //单元划分
            destrange = _Worksheet_lb.get_Range(string.Format("A2:R{0}", tlen + 1));
            destrange.Value2 = disinfo;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_lb.get_Range(string.Format("B2:R{0}", tlen + 1));
                sortrange = _Worksheet_lb.get_Range(string.Format("E2:E{0}", tlen + 1));
                GlobalExcel.ReflectionColnum(_Worksheet_lb, destrange, sortrange);
            }
        }

        private static void WriteGPSImg2Xls_SHPG(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, string ImgType, int ImgDis)
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

            int tmpmile = 0;
            int len = prjinfo._EndDmi / ImgDis + 1;
            int temp = 0;
            object[,] dataobj = new object[len, 11];
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
                            dataobj[leftidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                            dataobj[leftidx[i], colcnt++] = prjinfo._RoadNum;

                            dataobj[leftidx[i], colcnt++] = tempgpsinfo._utctime;

                            //tmpmile = tempinfos[gi + 1]._mile - prjinfo._Direction * ImgDis;
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
                            dataobj[leftidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                            dataobj[leftidx[i], colcnt++] = prjinfo._RoadNum;

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
                                dataobj[rightidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                                dataobj[rightidx[i], colcnt++] = prjinfo._RoadNum;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._utctime;

                                //tmpmile = tempinfos[gi + 1]._mile - prjinfo._Direction * ImgDis;
                                //if (tmpmile < 0)
                                //    tmpmile = 0;
                                //dataobj[rightidx[i], colcnt++] = tmpmile;
                                dataobj[rightidx[i], colcnt++] = tmile;

                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._longitude;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._latitude;

                                colcnt = 9;
                                rightidx[i] = i;
                                dataobj[rightidx[i], colcnt++] = "图像丢帧";
                                dataobj[rightidx[i], colcnt++] = "图像丢帧";
                            }
                            else
                            {
                                dataobj[rightidx[i], colcnt++] = i + 1;
                                dataobj[rightidx[i], colcnt++] = prjinfo._Direction > 0 ? "上行" : "下行";
                                dataobj[rightidx[i], colcnt++] = prjinfo._RoadNum;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._utctime;

                                //tmpmile = tempgpsinfo._mile;
                                //if (tmpmile < 0)
                                //    tmpmile = 0;
                                //dataobj[rightidx[i], colcnt++] = tmpmile;
                                dataobj[rightidx[i], colcnt++] = tmile;

                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._longitude;
                                dataobj[rightidx[i], colcnt++] = tempgpsinfo._latitude;

                                colcnt = 9;
                                temp = rightimgsinfo[i].LastIndexOf('\\');
                                dataobj[rightidx[i], colcnt++] = rightimgsinfo[i].Substring(temp + 1);
                                int temp2 = rightimgsinfo[i].IndexOf(' ') + 2;
                                dataobj[rightidx[i], colcnt++] = string.Format("\\{0}Img\\Camera1\\{1}", ImgType, rightimgsinfo[i].Substring(temp2, temp - temp2));
                            }
                        }
                    }
                }
            }

            MSExcel.Range destrange = null;
            MSExcel.Range sortrange = null;

            // 把景观照片粘贴进去
            float picleft, pictop;
            float height = (float)(5.5 * 0.3937008 * 72); // 单位从厘米转换为磅            
            float width = (float)(7.33 * 0.3937008 * 72); // 单位从厘米转换为磅

            string picname0 = "";
            string picname1 = "";
            try
            {
                if (prjinfo._Direction > 0)
                {
                    picname0 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[1].Split(' ')[1]);
                    if (leftimgsinfo.Length==2)
                    {
                        picname0 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[0].Split(' ')[1]);
                    }
                    picname1 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[leftimgsinfo.Length - 1].Split(' ')[1]);
                }
                else
                {
                    picname0 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[leftimgsinfo.Length - 1].Split(' ')[1]);
                    picname1 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[1].Split(' ')[1]);
                    if (leftimgsinfo.Length == 2)
                    {
                        picname1 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[0].Split(' ')[1]);
                    }
                }
            }
            catch (Exception)
            {

                 
            }

#if DEBUG
            if (!File.Exists(picname0))
            {
                picname0 = "E:\\job\\测试数据\\a.png";
                
            }
            if (!File.Exists(picname1))
            {
                picname1 = "E:\\job\\测试数据\\a.png";
            }
#endif

            destrange = worksheet.get_Range("L2");
            picleft = Convert.ToSingle(destrange.Left);
            pictop = Convert.ToSingle(destrange.Top);
            worksheet.Shapes.AddPicture(picname0, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, picleft, pictop, width, height);

            destrange = worksheet.get_Range("L20");
            picleft = Convert.ToSingle(destrange.Left);
            pictop = Convert.ToSingle(destrange.Top);
            worksheet.Shapes.AddPicture(picname1, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, picleft, pictop, width, height);


            destrange = worksheet.get_Range(string.Format("A2:K{0}", len + 1));
            destrange.Value2 = dataobj;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(string.Format("B2:K{0}", len + 1));
                sortrange = worksheet.get_Range(string.Format("E2:E{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            }
        }

        private static void WriteMTD2Xls_SHPG(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal,
            string[] MarkVal, ExcelGPS[] GPSInfo, double[] SpeedVal)
        {
            int len = roadpart.Count - 1;

            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] vallist = new object[len, 17];

            for (int i = 0; i < len; i++)
            {

                vallist[i, 0] = i + 1;
                vallist[i, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                vallist[i, 2] = prjinfo._RoadNum;

                //vallist[i, 3] = GPSInfo[i]._utctime;
                //vallist[i, 4] = roadpart[i].mile;
                //vallist[i, 5] = GPSInfo[i]._longitude;
                //vallist[i, 6] = GPSInfo[i]._latitude;
                //vallist[i, 7] = GPSInfo[i + 1]._utctime;
                //vallist[i, 8] = roadpart[i + 1].mile;
                //vallist[i, 9] = GPSInfo[i + 1]._longitude;
                //vallist[i, 10] = GPSInfo[i + 1]._latitude;

                GPSStartObj[i, 0] = GPSInfo[i]._utctime;
                GPSStartObj[i, 1] = roadpart[i].mile;
                GPSStartObj[i, 2] = GPSInfo[i]._longitude;
                GPSStartObj[i, 3] = GPSInfo[i]._latitude;

                GPSEndObj[i, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[i, 1] = roadpart[i + 1].mile;
                GPSEndObj[i, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[i, 3] = GPSInfo[i + 1]._latitude;

                vallist[i, 11] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 12] = SpeedVal[i];
                vallist[i, 13] = LMTDVal[i];
                //if (prjinfo._IsDIRIMTD)
                //{
                //    vallist[i, 14] = RMTDVal[i];
                //    vallist[i, 15] = String.Format("=ROUND((N{0}+O{0})/2,5)", i + 2);
                //}
                //else
                //{
                //    vallist[i, 15] = String.Format("=ROUND((N{0}),5)", i + 2);
                //}
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 14] = RMTDVal[i];
                }
                vallist[i, 15] = String.Format("=IF(OR(ISBLANK(N{0}),N{0}=0),IF(OR(ISBLANK(O{0}),O{0}=0),0,ROUND(VALUE(O{0}),5)),ROUND(AVERAGE(N{0},O{0}),5))", i + 2);
                vallist[i, 16] = _MarkVal[i];
            }

            MSExcel.Range destrange = null;
            MSExcel.Range sortrange = null;

            destrange = _Worksheet.get_Range(String.Format("A2:Q{0}", len + 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSEndObj;
                destrange = _Worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSStartObj;

                destrange = _Worksheet.get_Range(String.Format("B2:Q{0}", len + 1));
                sortrange = _Worksheet.get_Range(string.Format("E2:E{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
            }
            else
            {
                destrange = _Worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSStartObj;
                destrange = _Worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSEndObj;
            }
        }

        private static void WriteIRI2Xls_SHPG(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal,
            string[] MarkVal, ExcelGPS[] GPSInfo, double[] SpeedVal)
        {
            int len = roadpart.Count - 1;

            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] vallist = new object[len, 17];
            for (int i = 0; i < len; i++)
            {

                vallist[i, 0] = i + 1;
                vallist[i, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                vallist[i, 2] = prjinfo._RoadNum;

                //vallist[i, 3] = GPSInfo[i]._utctime;
                //vallist[i, 4] = roadpart[i].mile;
                //vallist[i, 5] = GPSInfo[i]._longitude;
                //vallist[i, 6] = GPSInfo[i]._latitude;
                //vallist[i, 7] = GPSInfo[i + 1]._utctime;
                //vallist[i, 8] = roadpart[i + 1].mile;
                //vallist[i, 9] = GPSInfo[i + 1]._longitude;
                //vallist[i, 10] = GPSInfo[i + 1]._latitude;

                GPSStartObj[i, 0] = GPSInfo[i]._utctime;
                GPSStartObj[i, 1] = roadpart[i].mile;
                GPSStartObj[i, 2] = GPSInfo[i]._longitude;
                GPSStartObj[i, 3] = GPSInfo[i]._latitude;

                GPSEndObj[i, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[i, 1] = roadpart[i + 1].mile;
                GPSEndObj[i, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[i, 3] = GPSInfo[i + 1]._latitude;

                vallist[i, 11] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 12] = SpeedVal[i];
                vallist[i, 13] = LIRIVal[i];
                //if (prjinfo._IsDIRIMTD)
                //{
                //    vallist[i, 14] = RIRIVal[i];
                //    vallist[i, 15] = String.Format("=ROUND((N{0}+O{0})/2,5)", i + 2);
                //}
                //else
                //{
                //    vallist[i, 15] = String.Format("=ROUND((N{0}),5)", i + 2);
                //}                
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 14] = RIRIVal[i];
                }
                vallist[i, 15] = String.Format("=IF(OR(ISBLANK(N{0}),N{0}=0),IF(OR(ISBLANK(O{0}),O{0}=0),0,ROUND(VALUE(O{0}),5)),ROUND(AVERAGE(N{0},O{0}),5))", i + 2);
                vallist[i, 16] = _MarkVal[i];
            }

            MSExcel.Range destrange = null;
            MSExcel.Range sortrange = null;

            destrange = _Worksheet.get_Range(String.Format("A2:Q{0}", len + 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSEndObj;
                destrange = _Worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSStartObj;

                destrange = _Worksheet.get_Range(String.Format("B2:Q{0}", len + 1));
                sortrange = _Worksheet.get_Range(string.Format("E2:E{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
            }
            else
            {
                destrange = _Worksheet.get_Range(String.Format("D2:G{0}", len + 1));
                destrange.Value2 = GPSStartObj;
                destrange = _Worksheet.get_Range(String.Format("H2:K{0}", len + 1));
                destrange.Value2 = GPSEndObj;
            }
        }

        private static void WriteRut2Xls_SHPG(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 10];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = i + 1;
                vallist[i, 1] = prjinfo._Direction > 0 ? "上行" : "下行";
                vallist[i, 2] = prjinfo._RoadNum;
                vallist[i, 3] = roadpart[i].mile;
                vallist[i, 4] = roadpart[i + 1].mile;
                vallist[i, 5] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                vallist[i, 6] = LRutVal[i];
                vallist[i, 7] = RRutVal[i];
                vallist[i, 8] = SRutVal[i];
            }

            MSExcel.Range destrange = null;
            MSExcel.Range sortrange = null;

            destrange = _Worksheet.get_Range(String.Format("A2:J{0}", len + 1));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet.get_Range(String.Format("B2:J{0}", len + 1));
                sortrange = _Worksheet.get_Range(string.Format("D2:D{0}", len + 1));
                GlobalExcel.ReflectionColnum(_Worksheet, destrange, sortrange);
                GlobalExcel.Reflection(_Worksheet, 2, 4, 2, false);
            }
        }
        public static  void WriteWc2Xls_SHPG(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, LaneProjectClass laneinfo, List<MilePart> roadpart,string[] markValue)
        {
            int len = roadpart.Count-1;

            object[,] datas = new object[len, 10];
            var wcData = laneinfo.m_wcDataClasses.First().wcResultDatas;

            
            for (int i = 0; i < len; i++)
            {
                datas[i, 0] = i + 1;
              
                    datas[i, 1] = roadpart[i].mile;
                    datas[i, 2] = roadpart[i + 1].mile;
                
                
                datas[i, 3] = Math.Abs( roadpart[i+1].mile- roadpart[i].mile);
                datas[i, 4] = markValue[i] == "路口单元" ? "是" : "否";
                datas[i, 5] = roadpart[i].roadtype==0?"沥青": roadpart[i].roadtype == 1?"水泥":"砂石";
                datas[i, 6] = wcData.Rows[i]["交通量等级"];
                datas[i, 7] = wcData.Rows[i]["基层类型"];

                if (prjinfo._Direction < 0)
                {
                    double wcValue = double.Parse(wcData.Rows[len - i - 1]["弯沉值"].ToString());

                    if (double.IsNaN(wcValue))
                    {
                        datas[i, 8] ="/";
                        datas[i, 9] = "/";
                    }
                    else
                    {
                        datas[i, 8] = double.Parse(wcData.Rows[len - i - 1]["弯沉值"].ToString()).ToString("F1");
                        datas[i, 9] = wcData.Rows[len - i - 1]["评价等级"];
                    } 

                }
                else
                {
                    double wcValue = double.Parse(wcData.Rows[i]["弯沉值"].ToString());

                    if (double.IsNaN(wcValue))
                    {
                        datas[i, 8] = "/";
                        datas[i, 9] = "/";
                    }
                    else
                    {
                        datas[i, 8] = double.Parse(wcData.Rows[i]["弯沉值"].ToString()).ToString("F1");
                        datas[i, 9] = wcData.Rows[i]["评价等级"];
                    }
                   

                }

            }
            Microsoft.Office.Interop.Excel.Range destrange = null;
            Microsoft.Office.Interop.Excel.Range sortrange = null;

            destrange = worksheet.get_Range(String.Format("A2:J{0}", len + 1));
            destrange.Value2 = datas;

            // 设置字体为宋体，大小为10号，水平居中
            MSExcel.Font font = destrange.Font;
            font.Name = "宋体"; // 设置字体名称
            font.Size = 10; // 设置字体大小

            MSExcel.Borders borders = destrange.Borders;
            borders.LineStyle = MSExcel.XlLineStyle.xlContinuous;
            borders.Weight = 2d;

            destrange.HorizontalAlignment = MSExcel.XlHAlign.xlHAlignCenter; // 设置水平居中
            destrange.VerticalAlignment = MSExcel.XlVAlign.xlVAlignCenter; // 设置垂直居中（可选）

            GlobalExcel.SetBorderLine(destrange, 53);
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet.get_Range(String.Format("B2:J{0}", len + 1));
                sortrange = worksheet.get_Range(string.Format("B2:B{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
                GlobalExcel.Reflection(worksheet, 2, 2, 2, false);
            }
        }

        public static void WriteWcData2Xls_SHPG(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, LaneProjectClass laneinfo, List<MilePart> roadpart, string[] markValue)
        {
            var dt = laneinfo.m_wcDataClasses.First().wcDatas;
            int len =dt.Rows.Count;

            object[,] datas = new object[len+1, 5];
            datas[0, 0] = "序号";
            datas[0, 1] ="行车方向";
            datas[0, 2] = "车道编号";
            datas[0, 3] = "桩号";
            datas[0, 4] = "弯沉修正值\r\n(0.01mm)";
            for (int i = 1; i < len+1; i++)
            {
                datas[i, 0] = i ;
                datas[i, 1] = dt.Rows[i-1]["方向"];
                datas[i, 2] = dt.Rows[i-1]["车道编号"];
                string mile = int.Parse( dt.Rows[i-1]["Mile"].ToString()).ToString("K0+000");
                datas[i, 3] =mile;
                datas[i, 4] = dt.Rows[i - 1]["弯沉值"];


            }
            MSExcel.Range destrange = null;
            MSExcel.Range sortrange = null;

            destrange = worksheet.get_Range(String.Format("A1:E{0}", len + 1));
            destrange.Value2 = datas;


            // 设置字体为宋体，大小为10号，水平居中
            MSExcel.Font font = destrange.Font;
            font.Name = "宋体"; // 设置字体名称
            font.Size = 10; // 设置字体大小

            MSExcel.Borders borders = destrange.Borders;
            borders.LineStyle = MSExcel.XlLineStyle.xlContinuous;
            borders.Weight = 2d;

            destrange.HorizontalAlignment = MSExcel.XlHAlign.xlHAlignCenter; // 设置水平居中
            destrange.VerticalAlignment = MSExcel.XlVAlign.xlVAlignCenter; // 设置垂直居中（可选）

            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet.get_Range(String.Format("E2:E{0}", len + 1));
            destrange.NumberFormat = "#0.00";
        }
        public static void WriteTj2Xls_SHPG(MSExcel.Application excelApp, MSExcel.Workbook srcbook, LaneProjectClass laneinfo = null)
        {
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Worksheet srcsheet2 = null;
            MSExcel.Range srcrange = null;
            object[,] obj = null;
            object[,] obj_lq = null;
            object[,] obj_sn = null;
            int userownum = 0;

            int lq_dis_num = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
            int sn_dis_num = RoadDiseaseTypes.DiseaseTypeDict[1].Count;

            srcsheet = srcbook.Sheets["统计"] as MSExcel.Worksheet;

            // 病害面积统计
            try
            {
                srcsheet2 = srcbook.Sheets["病害面积计算（沥青）"] as MSExcel.Worksheet;

                obj = new object[lq_dis_num + 1, 1];
                for (int i = 0; i < lq_dis_num; ++i)
                {
                    obj[i, 0] = string.Format("=SUM('病害面积计算（沥青）'!{0}:{0})", GlobalExcel.GetCol((char)('F' + i)));
                }
                obj[lq_dis_num, 0] = string.Format("=SUM(C4:C{0})", lq_dis_num + 3);
                srcrange = srcsheet.get_Range(string.Format("C4:C{0}", lq_dis_num + 4));
                srcrange.Value2 = obj;

                userownum = GlobalExcel.judegeusedrow(srcsheet2, 2, 2);
                srcsheet.Cells[4, 4] = string.Format("=ABS(SUM('病害面积计算（沥青）'!B{0}:B{1})-SUM('病害面积计算（沥青）'!C{0}:C{1}))*控制信息!H6", 3, userownum);

                srcrange = srcsheet.get_Range(string.Format("A4:D{0}", lq_dis_num + 3));
                obj_lq = (object[,])srcrange.Value2;
            }
            catch (System.Exception ex)
            {
            }

            try
            {
                srcsheet2 = srcbook.Sheets["病害面积计算（水泥）"] as MSExcel.Worksheet;
                obj = new object[sn_dis_num + 1, 1];
                for (int i = 0; i < sn_dis_num; ++i)
                {
                    obj[i, 0] = string.Format("=SUM('病害面积计算（水泥）'!{0}:{0})", GlobalExcel.GetCol((char)('F' + i)));
                }
                obj[sn_dis_num, 0] = string.Format("=SUM(C24:C{0})", sn_dis_num + 23);
                srcrange = srcsheet.get_Range(string.Format("C24:C{0}", sn_dis_num + 24));
                srcrange.Value2 = obj;

                userownum = GlobalExcel.judegeusedrow(srcsheet2, 2, 2);
                srcsheet.Cells[24, 4] = string.Format("=ABS(SUM('病害面积计算（水泥）'!B{0}:B{1})-SUM('病害面积计算（水泥）'!C{0}:C{1}))*控制信息!H6", 3, userownum);

                srcrange = srcsheet.get_Range(string.Format("A24:D{0}", sn_dis_num + 23));
                obj_sn = (object[,])srcrange.Value2;
            }
            catch (System.Exception ex)
            {
            }

            // 评价等级统计
            obj = new object[6, 9];
            obj[0, 0] = "评定等级";
            string[] headstr = { "PQI", "PCI", "RQI", "TD" };
            for (int i = 0; i < headstr.Length; ++i)
            {
                obj[0, 2 * i + 1] = headstr[i];
                obj[1, 2 * i + 1] = "长度(m)";
                obj[1, 2 * i + 2] = "百分比(%)";
                obj[i + 2, 0] = GlobalExcel.GetCol((char)('A' + i));
            }

            char[] rowidx = { 'A', 'B', 'C', 'D' };
            char[] colidx = { 'O', 'Q', 'S', 'X' };
            for (int i = 0; i < colidx.Length; ++i)
            {
                for (int k = 0; k < rowidx.Length; ++k)
                {
                    if (i < colidx.Length - 1)
                    {
                        obj[2 + k, 2 * i + 1] =
                            string.Format("=ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J))", colidx[i], rowidx[k]);
                        obj[2 + k, 2 * i + 2] =
                            string.Format("=ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J))*100/ABS(SUM(指标汇总!F:F)-SUM(指标汇总!J:J))", colidx[i], rowidx[k]);
                    }
                    else
                    {
                        obj[2 + k, 2 * i + 1] =
                            string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J)))", colidx[i], rowidx[k]);
                        obj[2 + k, 2 * i + 2] =
                            string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J))*100/SUM($O$23:$O$26))", colidx[i], rowidx[k]);
                    }
                }
            }


            srcrange = srcsheet.get_Range("H21:P26");
            srcrange.Delete();
            srcrange = srcsheet.get_Range("H21:P26");
            GlobalExcel.SetBorderLine(srcrange, 63);
            srcrange.Value2 = obj;
            srcrange = srcsheet.get_Range("H21:H22");
            srcrange.MergeCells = true; //合并单元格
            for (int i = 0; i < 4; ++i)
            {
                srcrange = srcsheet.get_Range(string.Format("{0}21:{1}21", GlobalExcel.GetCol((char)('I' + i * 2)), GlobalExcel.GetCol((char)('J' + i * 2))));
                srcrange.MergeCells = true; //合并单元格
                srcrange = srcsheet.get_Range(string.Format("{0}23:{0}26", GlobalExcel.GetCol((char)('J' + i * 2))));
                srcrange.NumberFormat = "#0.00";
            }

            //弯沉统计

            if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
            {
                obj = new object[5, 3];
                obj[0, 0] = "评定等级";
                string[] headWcstr= { "结构强度"};
                for (int i = 0; i < headWcstr.Length; ++i)
                {
                    obj[0, 2 * i + 1] = headWcstr[i];
                    obj[1, 2 * i + 1] = "长度(m)";
                    obj[1, 2 * i + 2] = "百分比(%)"; 
                }
                obj[2, 0] = "足够";
                obj[3, 0] = "临界";
                obj[4, 0] = "不足";
                string[]  rowidxWc = { "足够", "临界", "不足" };
                for (int k = 0; k < rowidxWc.Length; ++k)
                {
                    obj[2 + k,1] =
                            string.Format("=ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J))", 'Z', rowidxWc[k]);
                    obj[2 + k,2] =
                        string.Format("=ABS(SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!F:F)-SUMIF(指标汇总!{0}:{0},\"{1}\",指标汇总!J:J))*100/ABS(SUM(指标汇总!F:F)-SUM(指标汇总!J:J))", 'Z', rowidxWc[k]);
                } 
                srcrange = srcsheet.get_Range("H30:J34");
                srcrange.Delete();
                srcrange = srcsheet.get_Range("H30:J34");
                GlobalExcel.SetBorderLine(srcrange, 63);
                srcrange.Value2 = obj;
                srcrange = srcsheet.get_Range("H30:H31");

                srcrange.MergeCells = true; //合并单元格
                 
                    srcrange = srcsheet.get_Range("I30:J30");
                    srcrange.MergeCells = true; //合并单元格
                    srcrange = srcsheet.get_Range("I32:J34");
                    srcrange.NumberFormat = "#0.00";  
            }
            // 车道统计
            if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
                obj = new object[1, 14];
            else
                obj = new object[1, 12];
            srcsheet2 = srcbook.Sheets["控制信息"] as MSExcel.Worksheet;
            string roaddegreestr = ((MSExcel.Range)(srcsheet2.Cells[7, 5])).Value.ToString();
            int roaddegree = 0;
            if (roaddegreestr == "快速路")
                roaddegree = 0;
            else if (roaddegreestr == "主干路")
                roaddegree = 1;
            else if (roaddegreestr == "次干路")
                roaddegree = 2;
            else if (roaddegreestr == "支路")
                roaddegree = 3;

            obj[0, 0] = "=CONCATENATE(控制信息!H4,控制信息!H5,\"车道\")";
            obj[0, 1] = "=控制信息!H9";
            obj[0, 2] = "=控制信息!H10";
            obj[0, 3] = "=ABS(I14-J14)";
            obj[0, 4] = "=SUMPRODUCT(指标汇总!M:M,指标汇总!P:P)/SUM(指标汇总!M:M)";
            obj[0, 6] = "=SUMPRODUCT(指标汇总!M:M,指标汇总!R:R)/SUM(指标汇总!M:M)";
            //obj[0, 8] = "=SUMPRODUCT(指标汇总!M:M,指标汇总!N:N)/SUM(指标汇总!M:M)";
            obj[0, 8] = string.Format("=ROUND({0}*L14+{1}*{2}*N14,5)", _PQIW[roaddegree][0], _PQIW[roaddegree][1], _PQIT);
            obj[0, 10] = "=IF(控制信息!$E$7=\"支路\",\"/\",IF(SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M)>0, SUMPRODUCT(N(指标汇总!D:D=\"沥青\"),指标汇总!M:M,指标汇总!W:W)/SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M), \"/\"))";

            obj[0, 5] = string.Format("=IF(L14>={0},\"A\",IF(L14>={1},\"B\",IF(L14>={2},\"C\",\"D\")))",
            _PCIGrade[roaddegree][0], _PCIGrade[roaddegree][1], _PCIGrade[roaddegree][2]);
            obj[0, 7] = string.Format("=IF(N14>={0},\"A\",IF(N14>={1},\"B\",IF(N14>={2},\"C\",\"D\")))",
            _RQIGrade[roaddegree][0], _RQIGrade[roaddegree][1], _RQIGrade[roaddegree][2]);
            obj[0, 9] = string.Format("=IF(P14>={0},\"A\",IF(P14>={1},\"B\",IF(P14>={2},\"C\",\"D\")))",
            _PQIGrade[roaddegree][0], _PQIGrade[roaddegree][1], _PQIGrade[roaddegree][2]);
            obj[0, 11] = string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",IF(R14>={0},\"A\",IF(R14>={1},\"B\",IF(R14>={2},\"C\",\"D\"))))",
            _MTDGrade[roaddegree][0], _MTDGrade[roaddegree][1], _MTDGrade[roaddegree][2]);

            if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
            {
                  obj[0, 12] = laneinfo.m_wcDataClasses.First().WcValue;
                  obj[0, 13] = laneinfo.m_wcDataClasses.First().WcJudge;
                  srcrange = srcsheet.get_Range("H14:U14");
            }
            else
            {
                srcrange = srcsheet.get_Range("H14:S14");
                //删除弯沉栏目

                MSExcel.Range srcrangeTemp = srcsheet.get_Range("T12:U16");
                srcrangeTemp.Delete();
            }


            srcrange.Value2 = obj;
            srcrange.NumberFormat = "#0.00";
            if (laneinfo != null && laneinfo.m_wcDataClasses.Count > 0)
            {
                srcrange = srcsheet.get_Range("T14:T14");
                srcrange.NumberFormat = "#0.0"; 
            }
                

            srcrange = srcsheet.get_Range("I14:J14");
            srcrange.NumberFormat = "#K0+000";

            srcrange = srcsheet.get_Range("K14:K14");
            srcrange.NumberFormat = "#0";

            srcrange = srcsheet.get_Range("K15:K16");
            obj = (object[,])srcrange.Value2;

            double pcivallq = 0.0;
            double pcivalsn = 0.0;
            RoadDiseaseTypes.Clear();
            for (int i = 1; i <= lq_dis_num; ++i)
            {
                string ttype = "沥青." + obj_lq[i, 2].ToString();
                int typeidx = RoadDiseaseTypes.DiseaseTypeDict[0][ttype];
                RoadDiseaseTypes.roaddis[0][typeidx].totalarea = Convert.ToDouble(obj_lq[i, 3]);
            }
            for (int i = 1; i <= sn_dis_num; ++i)
            {
                string ttype = "水泥." + obj_sn[i, 2].ToString();
                int typeidx = RoadDiseaseTypes.DiseaseTypeDict[1][ttype];
                RoadDiseaseTypes.roaddis[1][typeidx].totalarea = Convert.ToDouble(obj_sn[i, 3]);
            } 
            double sumarealq = Convert.ToDouble(obj_lq[1, 4]);
            double sumareasn = Convert.ToDouble(obj_sn[1, 4]);
            pcivallq = ComputPCI(RoadDiseaseTypes.roaddis, 0, sumarealq);
            pcivalsn = ComputPCI(RoadDiseaseTypes.roaddis, 1, sumareasn);
            if (sumarealq > 0)
            {
                if (pcivallq < 0) pcivallq = 0;
            }
            else
            {
                pcivallq = 0;
            }

            if (sumareasn > 0)
            {
                if (pcivalsn < 0) pcivalsn = 0;
            }
            else
            {
                pcivalsn = 0;
            }

            obj = new object[2, 9];
            obj[0, 0] = "=SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M)";
            obj[1, 0] = "=SUMIF(指标汇总!D:D,\"水泥\",指标汇总!M:M)";

            obj[0, 1] = pcivallq;
            obj[1, 1] = pcivalsn;

            obj[0, 2] = string.Format("=IF(L15>={0},\"A\",IF(L15>={1},\"B\",IF(L15>={2},\"C\",\"D\")))",
            _PCIGrade[roaddegree][0], _PCIGrade[roaddegree][1], _PCIGrade[roaddegree][2]);
            obj[1, 2] = string.Format("=IF(L16>={0},\"A\",IF(L16>={1},\"B\",IF(L16>={2},\"C\",\"D\")))",
            _PCIGrade[roaddegree][0], _PCIGrade[roaddegree][1], _PCIGrade[roaddegree][2]);

            obj[0, 3] = "=IF(4.98+-0.34*SUMPRODUCT((指标汇总!D2:D1048576=\"沥青\")*(指标汇总!M2:M1048576)*(指标汇总!T2:T1048576))/SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M)>=0,4.98+-0.34*SUMPRODUCT((指标汇总!D2:D1048576=\"沥青\")*(指标汇总!M2:M1048576)*(指标汇总!T2:T1048576))/SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M))";
            obj[1, 3] = "=IF(4.98+-0.34*SUMPRODUCT((指标汇总!D2:D1048576=\"水泥\")*(指标汇总!M2:M1048576)*(指标汇总!T2:T1048576))/SUMIF(指标汇总!D:D,\"水泥\",指标汇总!M:M)>=0,4.98+-0.34*SUMPRODUCT((指标汇总!D2:D1048576=\"水泥\")*(指标汇总!M2:M1048576)*(指标汇总!T2:T1048576))/SUMIF(指标汇总!D:D,\"水泥\",指标汇总!M:M))";

            obj[0, 4] = string.Format("=IF(N15>={0},\"A\",IF(N15>={1},\"B\",IF(N15>={2},\"C\",\"D\")))",
            _RQIGrade[roaddegree][0], _RQIGrade[roaddegree][1], _RQIGrade[roaddegree][2]);
            obj[1, 4] = string.Format("=IF(N16>={0},\"A\",IF(N16>={1},\"B\",IF(N16>={2},\"C\",\"D\")))",
            _RQIGrade[roaddegree][0], _RQIGrade[roaddegree][1], _RQIGrade[roaddegree][2]);

            obj[0, 5] = string.Format("=ROUND({0}*L15+{1}*{2}*N15,5)", _PQIW[roaddegree][0], _PQIW[roaddegree][1], _PQIT);
            obj[1, 5] = string.Format("=ROUND({0}*L16+{1}*{2}*N16,5)", _PQIW[roaddegree][0], _PQIW[roaddegree][1], _PQIT);

            obj[0, 6] = string.Format("=IF(P15>={0},\"A\",IF(P15>={1},\"B\",IF(P15>={2},\"C\",\"D\")))",
            _PQIGrade[roaddegree][0], _PQIGrade[roaddegree][1], _PQIGrade[roaddegree][2]);
            obj[1, 6] = string.Format("=IF(P16>={0},\"A\",IF(P16>={1},\"B\",IF(P16>={2},\"C\",\"D\")))",
            _PQIGrade[roaddegree][0], _PQIGrade[roaddegree][1], _PQIGrade[roaddegree][2]);

            obj[0, 7] = "=SUMPRODUCT((指标汇总!D2:D1048576=\"沥青\")*(指标汇总!M2:M1048576)*(指标汇总!W2:W1048576))/SUMIF(指标汇总!D:D,\"沥青\",指标汇总!M:M)";
            obj[1, 7] = "=SUMPRODUCT((指标汇总!D2:D1048576=\"水泥\")*(指标汇总!M2:M1048576)*(指标汇总!W2:W1048576))/SUMIF(指标汇总!D:D,\"水泥\",指标汇总!M:M)";

            obj[0, 8] = string.Format("=IF(控制信息!$E$7=\"支路\",\"/\",IF(R15>={0},\"A\",IF(R15>={1},\"B\",IF(R15>={2},\"C\",\"D\"))))",
            _MTDGrade[roaddegree][0], _MTDGrade[roaddegree][1], _MTDGrade[roaddegree][2]);
            obj[1, 8] = "/";

            srcrange = srcsheet.get_Range("K15:S16");
            srcrange.Value2 = obj;

            // 养护对策信息
            try
            {
                srcsheet = srcbook.Sheets["对策"] as MSExcel.Worksheet;
            }
            catch (System.Exception ex)
            {
                MSExcel.Workbook templatebook = null;
                MSExcel.Worksheet templatesheet = null;
                MSExcel.Range templaterange = null;

                templatebook = excelApp.Workbooks.Open(string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\模板5.xlsx",
                    System.Windows.Forms.Application.StartupPath), Type.Missing, true,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                templatesheet = templatebook.Sheets["对策"] as MSExcel.Worksheet;
                templaterange = templatesheet.get_Range("A1:AI50");

                srcsheet = srcbook.Sheets["TD-10m"] as MSExcel.Worksheet;
                srcsheet2 = srcbook.Sheets.Add(Type.Missing, srcsheet, 1, MSExcel.XlSheetType.xlWorksheet) as MSExcel.Worksheet;
                srcsheet2.Name = "对策";
                srcrange = srcsheet2.get_Range("A1:AI50");

                templaterange.Copy(srcrange);
                templatebook.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            // 单元对策
            srcsheet = srcbook.Sheets["指标汇总"] as MSExcel.Worksheet;
            srcsheet2 = srcbook.Sheets["单元对策"] as MSExcel.Worksheet;
            userownum = GlobalExcel.judegeusedrow(srcsheet, 5, 2);
            obj = new object[userownum, 14];

            string[] colstr2 = { "A", "B", "C", "D", "F", "J", "M", "Q", "S", "X" };
            int[] colidx2 = { 0, 1, 2, 3, 4, 5, 6, 8, 9, 10 };
            int ii = 0;
            for (int i = 0; i < userownum; ++i)
            {
                ii = i + 1;
                for (int k = 0; k < colidx2.Length; ++k)
                {
                    obj[i, colidx2[k]] = "=指标汇总!" + colstr2[k] + ii.ToString();
                }
                if (i == 0)
                {
                    obj[i, 7] = "强度等级";
                    obj[i, 8] = "PCI等级";
                    obj[i, 9] = "RQI等级";
                    obj[i, 10] = "TD等级";
                    obj[i, 11] = "辅助列";
                    obj[i, 12] = "松";
                    obj[i, 13] = "严";
                }
                else
                {
                    if (laneinfo!=null && laneinfo.m_wcDataClasses.Count>0)
                    {
                        double wcValue = double.Parse(laneinfo.m_wcDataClasses.First().wcResultDatas.Rows[ii-2]["弯沉值"].ToString());
                        if (double.IsNaN(wcValue))
                        {
                            obj[i, 7] = $"足够";

                        }
                        else
                        {
                            obj[i, 7] = $"=IF(控制信息!$N$6=\"否\",\"足够\",弯沉!J{ii})";

                        }
                    }
                    else
                    {
                        obj[i, 7] = "=IF(控制信息!$N$6=\"否\",\"足够\",\"按实际输入\")";
                    }
                  
                    obj[i, 11] = string.Format("=IF(K{0}=\"/\",H{0}&IF(I{0}>J{0},I{0},J{0}),H{0}&IF(I{0}>J{0},I{0},J{0})&K{0})", ii);
                    obj[i, 12] = string.Format("=IF($D{0}=\"沥青\",IF(OR(控制信息!$N$8=\"否\",控制信息!$E$7=\"支路\"),INDEX(对策!I:I,MATCH(L{0},对策!G:G,0)),INDEX(对策!H:H,MATCH(L{0},对策!F:F,0))),INDEX(对策!AC:AC,MATCH(L{0},对策!AA:AA,0)))", ii);
                    obj[i, 13] = string.Format("=IF($D{0}=\"沥青\",IF(OR(控制信息!$N$8=\"否\",控制信息!$E$7=\"支路\"),INDEX(对策!K:K,MATCH(L{0},对策!G:G,0)),INDEX(对策!J:J,MATCH(L{0},对策!F:F,0))),INDEX(对策!AC:AC,MATCH(L{0},对策!AA:AA,0)))", ii);
                }
            }

            srcrange = srcsheet2.get_Range("A1:N" + userownum.ToString());
            srcrange.Value2 = obj;

            srcrange = srcsheet2.get_Range("E:F");
            srcrange.NumberFormat = "#K0+000";

            srcbook.Save();
        }
        #endregion
    }
}
