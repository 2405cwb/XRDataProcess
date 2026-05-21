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

namespace XRDataProcess
{
    /// <summary>
    /// 上海城镇道路，地标，DGTJ 08-92-2013上海城市道路养护技术规程
    /// </summary>
    class MyExcelCitySH2013
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
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, _RoadGradeDict, _SRutDisValL, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh, _RoadPart,_SRutDisValR, true);
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面平整度评价等级记录表.xlsx",
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

        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\车辙深度评价等级记录表.xlsx",
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
                //vallist[i, 5] = string.Format("=ROUND(MAX(D{0},E{0}),2)", i+4);
                vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面构造深度评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面磨耗评价等级记录表.xlsx",
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
            string srcxls = null;
            srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面病害面积统计表.xlsx",
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

            destrange = _Worksheet.get_Range(String.Format("A1:M{0}", tlen + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

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
                    //disval[0, disnum + 1] = string.Format("=IF(R{0}>={1},\"A\",IF(R{0}>={2},\"B\",IF(R{0}>={3},\"C\",\"D\")))",
                    //    rowcnt_sn_s, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2]);
                    disval[0, disnum + 1] = string.Format("=IF(Q{0}>={1},\"A\",IF(Q{0}>={2},\"B\",IF(Q{0}>={3},\"C\",\"D\")))",
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
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
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
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
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
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count;
                for (int i = 0; i < disnum; i++)
                {
                    rowval = i + 4;
                    worksheet_lqtj.Cells[rowval, 3] = string.Format("=SUMIF(沥青路面病害区间汇总表!{0}:{0},\"<>\",沥青路面病害区间汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    worksheet_lqtj.Cells[rowval, 4] = string.Format("=C{0}/F4", rowval);
                    string disname = ((MSExcel.Range)worksheet_lqtj.Cells[rowval, 2]).Value.ToString();

                    double tmval = Convert.ToDouble(((MSExcel.Range)worksheet_lqtj.Cells[rowval, 4]).Value.ToString());
                    worksheet_lqtj.Cells[rowval, 5] = ChaZhi(_RoadSocre[0][disname]._MiduScore, tmval);
                    RoadDiseaseTypes.roaddis[0][i].totalarea = Convert.ToDouble(((MSExcel.Range)worksheet_lqtj.Cells[rowval, 3]).Value);
                }
                worksheet_lqtj.Cells[disnum + 4, 3] = String.Format("=SUM(C4:C{0}", disnum + 3);
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
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count;
                for (int i = 0; i < disnum; i++)
                {
                    rowval = i + 4;
                    worksheet_sntj.Cells[i + 4, 3] = string.Format("=SUMIF(水泥路面病害区间汇总表!{0}:{0},\"<>\",水泥路面病害区间汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面破损评价等级记录表.xlsx",
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
                    j++;
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
                    j++;
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
            string srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\路面综合评价等级记录表.xlsx",
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
                srcxls = string.Format(@"{0}\报表模板\城镇道路 上海 DGTJ 08-92-2013\模板5.xlsx", System.Windows.Forms.Application.StartupPath);
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
                WriteSHPG2Xls(_Worksheet_unit, _Worksheet_lq, _Worksheet_sn, _Worksheet_hz,
                    _Worksheet_rqi, _Worksheet_pci, _Worksheet_pqi, _Worksheet_rd, _Worksheet_td,
                    prjinfo, prjdir, _RoadPart, _RoadDisList,
                    _LIRIMeanVal, _RIRIMeanVal,
                    _LRutMeanVal, _RRutMeanVal, _SRutMeanVal,
                    _LMTDMeanVal, _RMTDMeanVal, _GPSInfo, _MarkVal,
                    ref Hassnflag, ref Haslqflag);

                MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害明细"] as MSExcel.Worksheet;
                WriteGPSDis2Xls_SHPG(_Worksheet_lb, prjinfo, prjdir, _RoadDisList, _RoadPart);

                MSExcel.Worksheet _Worksheet2 = _Workbook.Sheets["景观图像"] as MSExcel.Worksheet;
                WriteGPSImg2Xls_SHPG(_Worksheet2, prjinfo, prjdir, "Street", prjinfo._StreetImgDis_Left);

                WriteTj2Xls_SHPG(excelApp, _Workbook);

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

            char[] splitchr = { '（', '（', '-', '_', ')', '）' };
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
            _Worksheet_prj.Cells[8, 8] = prjinfo._DataDate.Insert(6, "/").Insert(4, "/");
            _Worksheet_prj.Cells[5, 8] = _RoadConfig.DetectWidth;
        }

        private static void WritePrjInfo2Xls(MSExcel.Worksheet _Worksheet_prj, ProjectInfo prjinfo,
            LaneProjectClass laneinfo = null, string path = null)
        {
            object[,] obj_roadinfo = new object[32, 1];
            object[,] obj_roadpartinfo = new object[10, 1];
            object[,] obj_laneinfo = new object[10, 1];
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

            destrange = _Worksheet_prj.get_Range("E2:E11");
            destrange.Value2 = obj_roadpartinfo;

            destrange = _Worksheet_prj.get_Range("H2:H11");
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
            ref bool Hassnflag, ref bool Haslqflag)
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

            object[,] obj_unit = new object[len, 7];
            object[,] obj_hz = new object[len, 25];
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
            destrange = _Worksheet_unit.get_Range(string.Format("A2:G{0}", rowcnt + 1));
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
            destrange = _Worksheet_hz.get_Range(string.Format("A2:Y{0}", rowcnt + 1));
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
                ++tlen;
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
            if (prjinfo._Direction > 0)
            {
                picname0 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[0].Split(' ')[1]);
                picname1 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[leftimgsinfo.Length - 1].Split(' ')[1]);
            }
            else
            {
                picname0 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[leftimgsinfo.Length - 1].Split(' ')[1]);
                picname1 = string.Format("{0}\\{1}Img\\Camera0\\{2}", prjdir.FullName, ImgType, leftimgsinfo[0].Split(' ')[1]);
            }
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

        public static void WriteTj2Xls_SHPG(MSExcel.Application excelApp, MSExcel.Workbook srcbook)
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
            catch (System.Exception  )
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

            // 车道统计
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

            srcrange = srcsheet.get_Range("H14:S14");
            srcrange.Value2 = obj;
            srcrange.NumberFormat = "#0.00";

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
                    obj[i, 7] = "强度";
                    obj[i, 8] = "PCI等级";
                    obj[i, 9] = "RQI等级";
                    obj[i, 10] = "TD等级";
                    obj[i, 11] = "辅助列";
                    obj[i, 12] = "松";
                    obj[i, 13] = "严";
                }
                else
                {
                    obj[i, 7] = "=IF(控制信息!$N$6=\"否\",\"足够\",\"按实际输入\")";
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