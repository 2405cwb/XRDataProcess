using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms;
using Framework.Other.MyGlobal;

namespace XRDataProcess
{
    /// <summary>
    /// 重庆农村路
    /// </summary>
    class MyExcelCQDegree
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
                    if (rootchild.Name == Global.g_ParmStyles[(int)_Setting.ParmStyle])
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
                    if (rootchild.Name == Global.g_ParmStyles[(int)_Setting.ParmStyle])
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
                    if (prjinfo._IsDIRIMTD)
                    {
                        GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart10, 1, ref _RDeltaHVal);
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

            if (_Setting.ExcelType == 4) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);
            
            if (_RoadPart[0].roaddegree <= 1)
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
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\路面平整度评价等级记录表.xlsx",
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

            MSExcel.Worksheet _WorksheetPrjInfo = null;
            _WorksheetPrjInfo = _Workbook.Sheets["工程信息"] as MSExcel.Worksheet;
            WriteProjectInfo2Xls(_WorksheetPrjInfo, prjinfo, prjdir);

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
                vallist[i, 3] = LIRIVal[i];

                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RIRIVal[i];
                    if (_Setting.RQIJudgeType == 0)
                    {
                        vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + DataStartXlsxRow);
                    }
                    else if (_Setting. RQIJudgeType == 1)
                    {
                        vallist[i, 5] = String.Format("=ROUND(MAX(D{0}, E{0}),5)", i + DataStartXlsxRow);
                    }
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + DataStartXlsxRow);
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
        
        public static void OutputCPMSDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            disval *= 10;
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\CPMS路面病害调查表.xlsx",
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

        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

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
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
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
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
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
                    worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = smile;
                    worksheet_sshz.Cells[rowcnt_ss_s, colcnt++] = emile;
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

            //砂石
            GlobalExcel.WriteExcel(rowcnt_ss_s, 1, 1, 2, "总计", worksheet_sshz, 0);
            worksheet_sshz.Cells[rowcnt_ss_s, 3] = prjinfo._RoadNum;
            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_ss_s - 1);
            }
            destrange = worksheet_sshz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_ss_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
            destrange.Value2 = disval;
            
            destrange = worksheet_lqhz.get_Range(String.Format("A1:L{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_snhz.get_Range(String.Format("A1:N{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, borderType);
            destrange = worksheet_sshz.get_Range(String.Format("A1:K{0}", rowcnt_ss_s));
            GlobalExcel.SetBorderLine(destrange, borderType);

            RoadDiseaseTypes.Clear();

            if (!Haslqflag)
            {
                worksheet_lqhz.Delete();
            }

            if (!Hassnflag)
            {
                worksheet_snhz.Delete();
            }

            if (!Hasssflag)
            {
                worksheet_sshz.Delete();
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
                worksheet_lqtj.Cells[2, 4] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                }
                destrange = worksheet_lqtj.get_Range("B4:B" + (disnum + 3).ToString());
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
                worksheet_sntj.Cells[2, 4] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                }
                destrange = worksheet_sntj.get_Range("B4:B" + (disnum + 3).ToString());
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
                worksheet_sstj.Cells[2, 4] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=SUMIF(砂石病害汇总表!{0}:{0},\"<>\",砂石病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                }
                destrange = worksheet_sstj.get_Range("B4:B" + (disnum + 3).ToString());
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
                sumarea += disarea[roadtype][i].totalarea * disarea[roadtype][i].weight;
            }
            return 100 * sumarea / partarea;
        }

        public static void OutputPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\路面破损评价等级记录表.xlsx",
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

                vallist[i, 4] = string.Format("=100-{1}*POWER(D{0},{2})",
                    i + DataStartXlsxRow,
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

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
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\路面综合评价等级记录表.xlsx",
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

        //技术状况评定明细表
        public static void OutputPDMX(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\技术状况评定明细表.xlsx",
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
                vallist[rowcnt, 6] = 100;
                vallist[rowcnt, 7] = 100;
                vallist[rowcnt, 8] = 100;

                vallist[rowcnt, 9] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
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
                worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=SUMPRODUCT(B5:B{0},{1}5:{1}{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));
            }
        }

        private static void WritePrj2CPMSXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[3, 3] = prjinfo._RoadCode;
            if (prjinfo._Direction > 0)
            {
                _Worksheet.Cells[3, 7] = "上行" + prjinfo._RoadNum;
            }
            else
            {
                _Worksheet.Cells[3, 7] = "下行" + prjinfo._RoadNum;
            }
            _Worksheet.Cells[3, 11] = prjinfo._DataDate;
            _Worksheet.Cells[4, 7] = prjinfo._StartMile;
            _Worksheet.Cells[4, 11] = prjinfo._EndMile;
            _Worksheet.Cells[5, 11] = _RoadConfig.DetectWidth;
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
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\景观报表模板\沿线设施损坏汇总表.xlsx",
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

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSStreetDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\景观报表模板\CPMS_沿线设施损坏.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\景观报表模板\路基损坏汇总表.xlsx",
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

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\景观报表模板\CPMS_路基损坏.xlsx",
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
                        GlobalExcel.GetCol((char)('D' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
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
                        GlobalExcel.GetCol((char)('D' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
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
                else if (roadpart[i].roadtype == 2)
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
                        GlobalExcel.GetCol((char)('D' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
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

                if (emile % xlslen == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")))
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 7] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 11] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("D{0}:M{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 7] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 11] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("D{0}:M{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
                            destrange.ClearContents();
                        }
                        lq_flag = false;
                        lq_csmile = lq_cemile;
                    }
                    if (ss_csmile != ss_cemile)
                    {
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 7] = ss_csmile;
                        worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 11] = ss_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[2].Count - 1;
                            srcrange = worksheet_sshz.get_Range(String.Format("A{0}:U{1}", ss_tablerow * tcnt_ss + 1, ss_tablerow * (++tcnt_ss) + 1));
                            destrange = worksheet_sshz.get_Range(String.Format("A{0}", ss_tablerow * tcnt_ss + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_sshz.get_Range(String.Format("D{0}:M{1}", ss_tablerow * tcnt_ss + 7, ss_tablerow * tcnt_ss + 6 + disnum));
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
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 7] = sn_csmile;
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 11] = roadpart[len].mile;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 7] = lq_csmile;
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 11] = roadpart[len].mile;
                }
                if (ss_csmile != ss_cemile)
                {
                    worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 7] = ss_csmile;
                    worksheet_sshz.Cells[ss_tablerow * tcnt_ss + 4, 11] = roadpart[len].mile;
                }
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (tcnt_sn + 1) + 1));
                destrange = worksheet_snhz.get_Range(String.Format("D{0}:M{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
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
                destrange = worksheet_lqhz.get_Range(String.Format("D{0}:M{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
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
                destrange = worksheet_sshz.get_Range(String.Format("D{0}:M{1}", ss_tablerow * tcnt_ss + 7, ss_tablerow * tcnt_ss + 6 + disnum));
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

            _Worksheet.Cells[strrow, strcol] = string.Format("=CONCATENATE(\"沥青路面{6}评价等级“优”率占路段总数\",ROUND({1}{0},4)*100,\"%，“良”率占路段总数\",ROUND({2}{0},4)*100,\"%，“中”率占路段总数\",ROUND({3}{0},4)*100,\"%，“次”率占路段总数\",ROUND({4}{0},4)*100,\"%，“差”率占路段总数\",ROUND({5}{0},4)*100,\"%。\")",
                statisticsrow + 1,
                GlobalExcel.GetCol((char)('A' + statisticscol - 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 1)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 2)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 3)),
                GlobalExcel.GetCol((char)('A' + statisticscol - 1 + 4)),
                indextype);
        }

        public static void OutputAllRoadStatistics(MSExcel.Application excelApp, string outpath, List<string> xlslist)
        {
            string srcxls = string.Format(@"{0}\报表模板\重庆农村公路\多车道统计.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\多车道统计.xlsx", outpath);

            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet destsheet = null;
            destsheet = _Workbook.Sheets["技术指标"] as MSExcel.Worksheet;
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
                catch (Exception ex) { }

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
            catch (Exception  ) { }
        }
    }
}
