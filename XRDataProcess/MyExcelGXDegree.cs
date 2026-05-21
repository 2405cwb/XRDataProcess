using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms; 

namespace XRDataProcess
{
    /// <summary>
    /// 广西农村公路，地标
    /// </summary>
    class MyExcelGXDegree
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面平整度评价等级记录表.xlsx",
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
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + DataStartXlsxRow);
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

        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\车辙深度评价等级记录表.xlsx",
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

                vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
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

                vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面磨耗评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面构造深度评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面构造深度MPD评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\CPMS路面病害调查表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面病害面积统计表.xlsx",
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

            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            WriteDisHZ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, ref Haslqflag, ref Hassnflag, 5, 53);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
            WriteDisTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, prjdir, _RoadPart, Haslqflag, Hassnflag);

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

                    totalsnlen += milelength;
                    ++rowcnt_sn_s;
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

                    totallqlen += milelength;
                    ++rowcnt_lq_s;
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
                        ++rowcnt_sn_s;
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
                            ++rowcnt_lq_s;
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
                        ++rowcnt_lq_s;
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
                            ++rowcnt_sn_s;
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
                    ++rowcnt_sn_s;
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
                        ++rowcnt_lq_s;
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
                    ++rowcnt_lq_s;
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
                        ++rowcnt_sn_s;
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
            destrange = worksheet_snhz.get_Range(String.Format("A1:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum +  2))));
            GlobalExcel.SetBorderLine(destrange, borderType);

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
            destrange = worksheet_lqhz.get_Range(String.Format("A1:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 2))));
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
                worksheet_lqtj.Cells[2, 4] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    //disval[i, 0] = string.Format("=沥青病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_lq_s);
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
                    //disval[i, 0] = string.Format("=水泥病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_sn_s);
                    disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                }
                destrange = worksheet_sntj.get_Range("B4:B" + (disnum + 3).ToString());
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面破损评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面综合评价等级记录表.xlsx",
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

                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\技术状况评定明细表.xlsx",
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

                //IRI
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, 6] = Math.Round(trqival, 5);

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
                worksheet.Cells[rowcnt + 5, 3 + i] = String.Format("=SUMPRODUCT(B5:B{0},{1}5:{1}{0})/SUM(B5:B{0})", rowcnt + 4, (char)('C' + i));
            }
        }

        //输出空的技术状况评定明细表--用于提供给别人填充设备无法检测的技术指标数值
        public static void OutputPDMX_Empty(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\技术状况评定明细表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\路面跳车评价等级记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\景观报表模板\沿线设施损坏汇总表.xlsx",
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
            destrange = worksheet.get_Range(string.Format("N5:N{0}", len + 4));
            destrange.Value2 = disval;
        }

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSStreetDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\景观报表模板\CPMS_沿线设施损坏.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\景观报表模板\路基损坏汇总表.xlsx",
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
            destrange = worksheet.get_Range(string.Format("D5:D{0}", len + 4));
            destrange.Value2 = disval;
        }

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSRoadBedDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\景观报表模板\CPMS_路基损坏.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\板块病害列表.xlsx",
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
            #region
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

            #endregion
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

        #region 中南安环
        public static void OutputZNRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\中南安环\路况信息综合表2.xlsx",
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
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
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
            double   irival = 0;

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
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\中南安环\原始数据记录表.xlsx",
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\中南安环\病害统计表.xlsx",
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
                    ++n;
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\综合报表模板.xlsx",
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
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
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
            string srcxls = string.Format(@"{0}\报表模板\广西农村公路\报表模板5.xlsx",
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
                    vallist[i, 7] = RIRIVal[i];
                    vallist[i, 8] = String.Format("=(G{0}+H{0})/2", i + startidx);
                    vallist[i, 10] = String.Format("=ROUND(100/(1+{0}*EXP({1}*H{2})),5)",
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + startidx);
                }
                else
                {
                    vallist[i, 8] = String.Format("=G{0}", i + startidx);
                }

                vallist[i, 2] = String.Format("=G{0}*0.6", i + startidx);
                vallist[i, 4] = LMTDVal[i];
                vallist[i, 6] = LIRIVal[i];
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

                if (emile % xlslen == 0 || (MarkVal[i + 1] != null && MarkVal[i + 1].Contains("路面单元")))
                {
                    if (sn_csmile != sn_cemile)
                    {
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 7] = sn_csmile;
                        worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 12] = sn_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_snhz.get_Range(String.Format("A{0}", sn_tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_snhz.get_Range(String.Format("E{0}:N{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 7] = lq_csmile;
                        worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 12] = lq_cemile;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqhz.get_Range(String.Format("A{0}:U{1}", lq_tablerow * tcnt_lq + 1, lq_tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqhz.get_Range(String.Format("A{0}", lq_tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqhz.get_Range(String.Format("E{0}:N{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
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
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 7] = sn_csmile;
                    worksheet_snhz.Cells[sn_tablerow * tcnt_sn + 4, 12] = roadpart[len].mile;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 7] = lq_csmile;
                    worksheet_lqhz.Cells[lq_tablerow * tcnt_lq + 4, 12] = roadpart[len].mile;
                }
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                srcrange = worksheet_snhz.get_Range(String.Format("A{0}:U{1}", sn_tablerow * tcnt_sn + 1, sn_tablerow * (tcnt_sn + 1) + 1));
                destrange = worksheet_snhz.get_Range(String.Format("E{0}:N{1}", sn_tablerow * tcnt_sn + 7, sn_tablerow * tcnt_sn + 6 + disnum));
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
                destrange = worksheet_lqhz.get_Range(String.Format("E{0}:N{1}", lq_tablerow * tcnt_lq + 7, lq_tablerow * tcnt_lq + 6 + disnum));
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

    }
}
