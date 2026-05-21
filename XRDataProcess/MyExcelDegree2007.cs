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
using Framework.Other.MyGlobal;

namespace XRDataProcess
{
    /// <summary>
    /// 等级公路规范2007，JTG H20-2007《公路技术状况评定标准》
    /// </summary>
    class MyExcelDegree2007
    {
        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

        private static double[][] _RQIGrade;//道路等级 等级区间
        private static double[][] _RDIGrade;
        private static double[][] _MTDGrade;
        private static double[][] _PCIGrade;
        private static double[][] _PQIGrade;
        private static double[][] _RDIRD;
        private static double[] _RDIa;

        private static double[][] _RQIa;//公路等级 参数序号
        private static double[][][] _PCIa;//公路等级 路面材质 参数序号
        private static double[][][] _PQIW;//公路等级 路面材质 参数序号
        private static double[][] _WeightParm;//0-沥青，1-水泥
        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        public static Dictionary<string, int> _RoadGradeDict;

        public static List<MilePart> _RoadPart = null;
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

        private static ExcelGPS[] _GPSInfo = null;

        private static void InitXlsParm()
        {
            int len = _RoadGradeStr.Length;

            _RQIGrade = new double[len][];
            _RDIGrade = new double[len][];
            _MTDGrade = new double[len][];
            _PCIGrade = new double[len][];
            _PQIGrade = new double[len][];

            _RQIa = new double[len][];
            _PCIa = new double[len][][];
            _PQIW = new double[len][][];

            for (int i = 0; i < len; i++)
            {
                _RQIGrade[i] = new double[5];
                _RDIGrade[i] = new double[5];
                _MTDGrade[i] = new double[5];
                _PCIGrade[i] = new double[5];
                _PQIGrade[i] = new double[5];

                _RQIa[i] = new double[2];
                _PCIa[i] = new double[2][];
                _PQIW[i] = new double[2][];
                for (int j = 0; j < 2; j++)
                {
                    _PCIa[i][j] = new double[2];
                    _PQIW[i][j] = new double[4];
                }
            }
            _RDIa = new double[2];
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
            Doc.Load(System.Windows.Forms.Application.StartupPath + "\\ParaVal.xml");    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点  
            xmlNodes = Elem.ChildNodes;

            //读取病害类型
            for (int i = 0; i < 2; i++)
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
                                        _RQIa[i][0] = Convert.ToDouble(((XmlElement)node).GetAttribute("w1"));
                                        _RQIa[i][1] = Convert.ToDouble(((XmlElement)node).GetAttribute("w2"));
                                    }
                                    else if (node.Name == "RDI")
                                    {
                                        val.CopyTo(_RDIGrade[i], 0);
                                    }
                                    else if (node.Name == "MTD")
                                    {
                                        val.CopyTo(_MTDGrade[i], 0);
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
                                            _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][3] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WSRI"));
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
            bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsMaxRut)
        {
            bool IRIRes = true, RutRes = true, MTDRes = true, GPSRes = true;
            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }

            _RoadPart = new List<MilePart>();
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
            _RoadPart1M.Add(spart);

            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);

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
                if (IsMeanIRI) IRIRes = GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, _Setting.IsWarning);
            }
            else
            {
                IRIRes = true;
            }
            if (prjinfo._IsRut)
            {
                if (IsMeanRut) RutRes = GlobalExcel.GetRutMeanVal(prjinfo, prjdir, _RoadPart, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, _Setting.IsWarning);
                if (IsMaxRut) RutRes = GlobalExcel.GetRutMaxVal(prjinfo, prjdir, _RoadPart, ref _LRutMaxVal, ref _RRutMaxVal, ref _SRutMaxVal);
            }
            else
            {
                RutRes = true;
            }
            if (IsMeanMTD) MTDRes = GlobalExcel.GetMTDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, _Setting.IsWarning);
            
            if (_Setting.ExcelType == 4) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);
            
            return IRIRes && RutRes && MTDRes && GPSRes;
        }

        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\路面平整度评价等级记录表.xlsx",
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

            WriteIRI2Xls(_Worksheet, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal)
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

                vallist[i, 6] = String.Format("=ROUND(100/(1+{0}*EXP({1}*F{2})),5)",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 4);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
            }

            destrange = _Worksheet.get_Range(String.Format("A4:I{0}", len + 3));
            destrange.Value2 = vallist;

            WriteIRIStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:I{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 9, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteIRIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("Q3:U5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 16 + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 10] = "=CONCATENATE(\"路面平整度评价等级“优”率占路段总数\",ROUND(Q4,4)*100,\"%，“良”率占路段总数\",ROUND(R4,4)*100,\"%，“中”率占路段总数\",ROUND(S4,4)*100,\"%，“次”率占路段总数\",ROUND(T4,4)*100,\"%，“差”率占路段总数\",ROUND(U4,4)*100,\"%。\")";
        }

        public static void OutputRut(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\车辙深度评价等级记录表.xlsx",
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
            WriteRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            //_Worksheet = _Workbook.Sheets["Sheet2"] as MSExcel.Worksheet;
            //WriteMaxRut2Xls_orirut(_Worksheet, prjinfo, _RoadPart, _LRutMaxVal, _RRutMaxVal, _SRutMaxVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            if (!prjinfo._IsRut)
                return;

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;
            object[,] vallist = new object[len, 9];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                //if (!(roadpart[i].roadtype == 0 && roadpart[i].roaddegree <= 1))
                //{
                //    vallist[i, 3] = "-";
                //    vallist[i, 4] = "-";
                //    vallist[i, 5] = "-";
                //    vallist[i, 6] = "-";
                //    vallist[i, 7] = "-";
                //    continue;
                //}

                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];
                vallist[i, 5] = SRutVal[i];
                vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
                        i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);

                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                
                //if (roadpart[i].roadtype == 0)
                //{
                //    //vallist[i, 5] = string.Format("=ROUND(MAX(D{0},E{0}),2)", i + 4);
                //    vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
                //        i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                //    vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                //        i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);
                //}
                //else
                //{
                //    vallist[i, 6] = "-";
                //    vallist[i, 7] = "-";
                //}
            }

            destrange = _Worksheet.get_Range(String.Format("A4:I{0}", len + 3));
            destrange.Value2 = vallist;

            WriteRutStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:I{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 9, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteMaxRut2Xls_orirut(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            if (!prjinfo._IsRut)
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

                //if (!(roadpart[i].roadtype == 0 && roadpart[i].roaddegree <= 1))
                //{
                //    vallist[i, 3] = "-";
                //    vallist[i, 4] = "-";
                //    vallist[i, 5] = "-";
                //    vallist[i, 6] = "-";
                //    vallist[i, 7] = "-";
                //    continue;
                //}

                vallist[i, 3] = LRutVal[i];
                vallist[i, 4] = RRutVal[i];
                vallist[i, 5] = SRutVal[i];
                //vallist[i, 5] = string.Format("=ROUND(MAX(D{0},E{0}),2)", i + 4);
                vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
                        i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);

                vallist[i, 8] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                //if (roadpart[i].roadtype == 0)
                //{
                //    vallist[i, 6] = string.Format("=IF(F{0}<{1},{2}-{3}*F{0},IF(F{0}<{4},{5}-{6}*(F{0}-{1}),0))",
                //        i + 4, _RDIRD[0][1], _RDIRD[0][0], _RDIa[0], _RDIRD[1][1], _RDIRD[1][0], _RDIa[1]);
                //    vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                //        i + 4, _RDIGrade[roadpart[i].roaddegree][0], _RDIGrade[roadpart[i].roaddegree][1], _RDIGrade[roadpart[i].roaddegree][2], _RDIGrade[roadpart[i].roaddegree][3]);

                //}
                //else
                //{
                //    vallist[i, 6] = "-";
                //    vallist[i, 7] = "-";
                //}
            }

            destrange = _Worksheet.get_Range(String.Format("A4:I{0}", len + 3));
            destrange.Value2 = vallist;

            WriteRutStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:I{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 9, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteRutStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("Q3:U5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 16 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 10] = "=CONCATENATE(\"沥青路面车辙深度评价等级“优”率占路段总数\",ROUND(Q4,4)*100,\"%，“良”率占路段总数\",ROUND(R4,4)*100,\"%，“中”率占路段总数\",ROUND(S4,4)*100,\"%，“次”率占路段总数\",ROUND(T4,4)*100,\"%，“差”率占路段总数\",ROUND(U4,4)*100,\"%。\")";
        }

        public static void OutputMTD(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\路面构造深度评价等级记录表.xlsx",
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

            WriteMTD2Xls(_Worksheet, prjinfo, _RoadPart, _LMTDMeanVal, _RMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteMTD2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo,
            List<MilePart> roadpart, double[] LMTDVal, double[] RMTDVal)
        {
            if (!prjinfo._IsIRIMTD)
            {
                return;
            }

            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 7];
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
                vallist[i, 6] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
            }
            destrange = _Worksheet.get_Range(String.Format("A4:G{0}", len + 3));
            destrange.Value2 = vallist;

            destrange = _Worksheet.get_Range(String.Format("A1:G{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 7, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }

        public static void OutputCPMSDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            disval *= 10;
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\CPMS路面病害调查表.xlsx",
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
            WriteZJGTDisHZTJ2Xls(_Worksheet_sndc, _Worksheet_lqdc, prjinfo, prjdir, _RoadPart, _RoadDisList, disval);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, _RoadDisList,_RoadPart);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            WriteDisHZTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_snhz, _Worksheet_lqhz,
                prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteDisLB2Xls_roadpart(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist, List<MilePart> roadpart)
        {
            int len = dislist.Length, i = 0, troadtype = -1;
            if (len < 1)
                return;

            MSExcel.Range destrange;
            object[,] val = new object[len, 13];
            string[] s;
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
                            val[i, 12] = tdis.RoadType;
                            ++i;
                            troadtype = -1;
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
                GlobalExcel.Reflection(_Worksheet, 3, 1, 13, true);
            }
        }

        private static void WriteDisLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            int len = dislist.Length, i = 0;
            if (len < 1)
                return;

            MSExcel.Range destrange;
            object[,] val = new object[len, 12];
            string[] s;
            foreach (Disease tdis in dislist)
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
            destrange = _Worksheet.get_Range(String.Format("A3:L{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:L{0}", len + 2));
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
                GlobalExcel.Reflection(_Worksheet, 3, 1, 10, true);
            }
        }
        private static void WriteDisHZTJ2Xls(MSExcel.Worksheet worksheet_sntj, MSExcel.Worksheet worksheet_lqtj,
            MSExcel.Worksheet worksheet_snhz, MSExcel.Worksheet worksheet_lqhz,
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

            destrange = worksheet_lqhz.get_Range(String.Format("A1:AB{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet_snhz.get_Range(String.Format("A1:AA{0}", rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, 53);

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
                worksheet_lqhz.Delete();
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
                worksheet_snhz.Delete();
                worksheet_sntj.Delete();
            }
        }

        /// <summary>
        /// 计算区间段内路面破损率DR
        /// </summary>
        /// <param name="disarea"></param>
        /// <param name="roadtype"></param>
        /// <param name="startidx"></param>
        /// <param name="partarea"></param>
        /// <returns></returns>
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
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\路面破损评价等级记录表.xlsx",
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

            WritePCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePCI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 7];

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
                //  =IF(G3="沥青", 100-15*POWER(D3,0.412),100-10.66*POWER(D3,0.461))  

                //vallist[i, 4] = string.Format("=100-{0}*POWER(D{1},{2})",
                //    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                //    i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                vallist[i, 4] = string.Format("=IF(G{0}=\"沥青\", 100-{1}*POWER(D{0},{2}),100-{3}*POWER(D{0},{4})) ", i + 3,
                _PCIa[roadpart[i].roaddegree][0][0],
                _PCIa[roadpart[i].roaddegree][0][1],
                _PCIa[roadpart[i].roaddegree][1][0],
                _PCIa[roadpart[i].roaddegree][1][1]);

                vallist[i, 5] = string.Format("=IF(E{0}>={1},\"优\",IF(E{0}>={2},\"良\",IF(E{0}>={3},\"中\",IF(E{0}>={4},\"次\",\"差\"))))",
                    i + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);
                vallist[i, 6] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
            }

            destrange = worksheet.get_Range(String.Format("A3:G{0}", len + 2));
            destrange.Value2 = vallist;
            WritePCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:G{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 7, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePCIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("O3:S5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(F:F,\"{0}\",A:A)-SUMIF(F:F,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 14 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 8] = "=CONCATENATE(\"路面PCI评价等级“优”率占路段总数\",ROUND(O4,4)*100,\"%，“良”率占路段总数\",ROUND(P4,4)*100,\"%，“中”率占路段总数\",ROUND(Q4,4)*100,\"%，“次”率占路段总数\",ROUND(R4,4)*100,\"%，“差”率占路段总数\",ROUND(S4,4)*100,\"%。\")";
        }

        public static void OutputPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\路面综合评价等级记录表.xlsx",
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

            WritePQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, rutval = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 12];

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
                vallist[rowcnt, colcnt++] = pcival;
                vallist[rowcnt, colcnt++] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);

                if (prjinfo._IsDIRIMTD)
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5,5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                vallist[rowcnt, colcnt] = trqival;

                colcnt++;
                vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3]);

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
                //  =IF(L3="沥青", ROUND((0.35*D3+0.4*F3+0.15*IF(EXACT(H3,"-"),0,H3))/(0.35+0.4+0.15),5),ROUND((0.5*D3+0.4*F3+0*IF(EXACT(H3,"-"),0,H3))/(0.5+0.4+0),5))
                if (roadpart[i].roaddegree <= 1)
                {
                    //vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0}))/({1}+{2}+{3}),5)",
                    //        rowcnt + 3,
                    //        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    //        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                    //        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);

                    vallist[rowcnt, colcnt++] = string.Format("=IF(L{0}=\"沥青\", ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0}))/({1}+{2}+{3}),5),ROUND(({4}*D{0}+{5}*F{0}+{6}*IF(EXACT(H{0},\"-\"),0,H{0}))/({4}+{5}+{6}),5))", rowcnt + 3,
                     _PQIW[roadpart[i].roaddegree][0][0],
                     _PQIW[roadpart[i].roaddegree][0][1],
                     _PQIW[roadpart[i].roaddegree][0][2],
                     _PQIW[roadpart[i].roaddegree][1][0],
                     _PQIW[roadpart[i].roaddegree][1][1],
                     _PQIW[roadpart[i].roaddegree][1][2]);
                }
                else
                {
                    //vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),5)",
                    //        rowcnt + 3,
                    //        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    //        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                    vallist[rowcnt, colcnt++] = string.Format("=IF(L{0}=\"沥青\",ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),5),ROUND(({3}*D{0}+{4}*F{0})/({3}+{4}),5))", rowcnt + 3,
                          _PQIW[roadpart[i].roaddegree][0][0],
                          _PQIW[roadpart[i].roaddegree][0][1],
                          _PQIW[roadpart[i].roaddegree][1][0],
                          _PQIW[roadpart[i].roaddegree][1][1]);
                }

                vallist[rowcnt, colcnt++] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                vallist[rowcnt, colcnt++] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:L{0}", rowcnt + 2));
            destrange.Value2 = vallist;
            WritePQI2Statistics(worksheet);
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 12, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }

        private static void WritePQI2Statistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("T3:X5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(K:K,\"{0}\",A:A)-SUMIF(K:K,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 19 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 13] = "=CONCATENATE(\"路面PQI评价等级“优”率占路段总数\",ROUND(T4,4)*100,\"%，“良”率占路段总数\",ROUND(U4,4)*100,\"%，“中”率占路段总数\",ROUND(V4,4)*100,\"%，“次”率占路段总数\",ROUND(W4,4)*100,\"%，“差”率占路段总数\",ROUND(X4,4)*100,\"%。\")";
        }

        #region 导出水泥病害
        public static void OutputBkDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir )
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\板块病害列表.xlsx",
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
            int bknum = SnbkSetForm.bknum ;
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
                if (prjinfo._Direction > 0 && (smile + j * bklength) > prjinfo._EndMile) 
                {
                    val[i, 3] = prjinfo._EndMile;
                }
                else if (prjinfo._Direction < 0 && (smile + j * bklength) > prjinfo._StartMile) 
                {
                    val[i, 3] = prjinfo._StartMile;
                }
                val[i, 4] = prjinfo._Direction > 0 ? "上行" : "下行";
                temp++; bknum++;
                double ssmile=smile + (j - 1) * bklength;
                if (judgeroadtype(ssmile, _RoadPart,prjinfo) == 0)
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
        private static int judgeroadtype(double mile, List<MilePart> roadpart ,ProjectInfo prjinfo)
        {
            for (int i = 0; i < roadpart.Count - 1; ++i)
            {
                if ((prjinfo._Direction > 0 && mile >= roadpart[i].mile && mile < roadpart[i + 1].mile) 
                    || (prjinfo._Direction < 0 && mile < roadpart[i].mile && mile >= roadpart[i + 1].mile))
                {
                    return roadpart[i].roadtype;
                }
             
            }
            return -1;
        }
        #endregion
        //private static void WriteBkDisLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist) //只统计有病害的板块
        //{
        //    #region
        //    int len = dislist.Length, i = 0,temp=1;
        //    if (len < 1)
        //        return;

        //    MSExcel.Range destrange;
        //    string[] s;
        //    bool fg = false;
        //    int bknum = SnbkSetForm.bknum - 1;
        //    double bklength=Disease.bklength;
        //    int smile = prjinfo._Direction > 0 ? prjinfo._StartMile : prjinfo._EndMile;
        //    int emile=prjinfo._Direction < 0 ? prjinfo._StartMile : prjinfo._EndMile;
        //    int cnt = Convert.ToInt32((emile - smile) / Disease.bklength) + 1;
        //    len = len + Convert.ToInt32((emile - smile) / bklength);
        //    object[,] val = new object[len , 11];
        //    List<Disease> tdis = new List<Disease>();
        //    List<Disease> udis = new List<Disease>();
        //    for (int j = temp; j <=cnt; ++j) //遍历每个板块
        //    {
        //        temp++;
        //        for (int k = 0; k <dislist.Length; ++k) //遍历所有病害
        //        {
        //            if(dislist[k].m_mile >= smile + (j - 1) * bklength && dislist[k].m_mile < smile + j * bklength)//统计一个板块内的病害,然后再计算
        //            {
        //                tdis.Add(dislist[k]);   //一个板块内所有病害tdis
        //            }
        //            else 
        //            {
        //                if (tdis.Count <= 0) continue;
        //                bknum++;
        //                foreach (Disease t in tdis)
        //                {
        //                    if (t.computetype==5)
        //                    {
        //                        udis.Add(t);
        //                        break;
        //                    }
        //                }
        //                if (udis.Count > 0)//如果该板块内病害含有破碎板
        //                {
        //                    s = udis[0].RoadDisType.Split('.');

        //                    val[i, 0] = bknum;
        //                    val[i, 1] = i + 1;
        //                    val[i, 2] = smile + (j - 1) * bklength;
        //                    val[i, 3] = smile + j * bklength;
        //                    val[i, 4] = prjinfo._RoadNum;
        //                    if (udis[0].RoadType == "水泥")
        //                    {
        //                        val[i, 5] = s[0];
        //                        if (s.Length > 1)
        //                        {
        //                            val[i, 6] = s[1];
        //                        }
        //                        else
        //                        {
        //                            val[i, 6] = " ";
        //                        }
        //                    }

        //                    val[i, 7] = Disease.bklength * Disease.bkwidth;
        //                    val[i, 8] = Disease.bklength;
        //                    val[i, 9] = Disease.bkwidth;
        //                    val[i, 10] = udis[0].RoadType;
        //                    ++i;
        //                }
        //                else
        //                {
        //                    foreach (Disease t in tdis)
        //                    {
        //                        s = t.RoadDisType.Split('.');

        //                        val[i, 0] = bknum;
        //                        val[i, 1] = i + 1;
        //                        val[i, 2] = smile + (j - 1) * bklength;
        //                        val[i, 3] = smile + j * bklength;
        //                        val[i, 4] = prjinfo._RoadNum;
        //                        if (t.RoadType == "水泥")
        //                        {
        //                            val[i, 5] = s[0];
        //                            if (s.Length > 1)
        //                            {
        //                                val[i, 6] = s[1];
        //                            }
        //                            else
        //                            {
        //                                val[i, 6] = " ";
        //                            }
        //                        }

        //                        val[i, 7] = Disease.bklength * Disease.bkwidth;
        //                        val[i, 8] = Disease.bklength;
        //                        val[i, 9] = Disease.bkwidth;
        //                        val[i, 10] = t.RoadType;
        //                        ++i;
        //                    }
        //                }
        //                tdis.Clear();
        //                udis.Clear();
        //            }
        //        }
        //    }

        //    destrange = _Worksheet.get_Range(String.Format("A2:K{0}", len + 2));
        //    destrange.Value2 = val;
        //    int tlen = 0;
        //    for (int k = 0; k < len; ++k)
        //    {
        //        if (val[k, 0] == null)
        //        {
        //            break;
        //        }
        //        tlen++;
        //    }
        //    destrange = _Worksheet.get_Range(String.Format("A2:K{0}", tlen + 1));
        //    GlobalExcel.SetBorderLine(destrange, 53);

        //    #endregion
        //}



        /////////////////////////////////////////////////////////////////////////////////////////////////
        public static void OutputRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\综合报表模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
            WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, _RoadDisList,_RoadPart);

            MSExcel.Worksheet _Worksheet_lqtj = _Workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sntj = _Workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lqhz = _Workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_snhz = _Workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
            WriteDisHZTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_snhz, _Worksheet_lqhz,
                prjinfo, prjdir, _RoadPart, _RoadDisList);

            WriteAll2Xls(_Workbook, prjinfo, prjdir, _RoadPart, _RoadDisList,
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal);


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
            double[] LMTDVal, double[] RMTDVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, mtdval = 0, tpcival = 0;

            object[,] mxlist = new object[len, 21];
            object[,] yhlist = new object[len, 21];
            int yhi = 0;
            string lenstr = "0";
            int tlen = len;
            while ((tlen = tlen / 10) > 0)
            {
                lenstr += "0";
            }

            int typeidx = 0;
            bool res = false;
			
            string errlog = prjdir.FullName + "\\errlog.txt";
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
                mxlist[i, 11] = GlobalExcel._RoadTypeExcelStr[roadpart[i].roadtype];

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
                mxlist[i, 12] = drval;
                mxlist[i, 8] = string.Format("=100-{0}*POWER(M{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 17] = string.Format("=IF(I{0}>={1},\"优\",IF(I{0}>={2},\"良\",IF(I{0}>={3},\"中\",IF(I{0}>={4},\"次\",\"差\"))))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI                
                if (prjinfo._IsDIRIMTD)
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                mxlist[i, 13] = irival;
                mxlist[i, 9] = String.Format("=ROUND(100/(1+{0}*EXP({1}*N{2})),5)", _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 3);
                mxlist[i, 18] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    i + 3, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2], _RQIGrade[roadpart[i].roaddegree][3]);

                //MTD               
                if (prjinfo._IsDIRIMTD)
                    mtdval = Math.Round((LMTDVal[i] + RMTDVal[i]) * 0.5, 5);
                else
                    mtdval = Math.Round(LMTDVal[i], 5);
                mxlist[i, 15] = mtdval;

                //Rut
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    double rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);
                    mxlist[i, 14] = rutval;

                    mxlist[i, 10] = string.Format("=IF(O{0}<{1},{2}-{3}*O{0},IF(O{0}<{4},{5}-{6}*(O{0}-{1}),0))",
                        i + 3,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                    mxlist[i, 19] = string.Format("=IF(K{0}>={1},\"优\",IF(K{0}>={2},\"良\",IF(K{0}>={3},\"中\",IF(K{0}>={4},\"次\",\"差\"))))",
                        i + 3,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);
                }

                if (roadpart[i].roaddegree < 2 )
                {
                    mxlist[i, 20] = string.Format("=ROUND(({1}*I{0}+{2}*J{0}+{3}*IF(EXACT(K{0},\"-\"),0,K{0}))/({1}+{2}+{3}),5)",
                        i + 3,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                }
                else
                {
                    mxlist[i, 20] = string.Format("=ROUND(({1}*I{0}+{2}*J{0})/({1}+{2}),5)",
                        i + 3,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 16] = string.Format("=IF(U{0}>={1},\"优\",IF(U{0}>={2},\"良\",IF(U{0}>={3},\"中\",IF(U{0}>={4},\"次\",\"差\"))))",
                    i + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                mxlist[i, 7] = String.Format("=CONCATENATE(TEXT(U{0},\"0.00\"),\"(\",Q{0},\")\")", i + 3);
                if (_Setting.YHType == 0)
                {
                    if (!(trqival > 85 && tpcival > 85))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!J{0}<=70,\"大修\",IF(技术状况明细表!I{0}>85,IF(技术状况明细表!J{0}>85,\"日常养护\",\"中修\"),\"中修\"))", i + 3);
                        yhi++;
                    }
                }
                else if (_Setting.YHType == 1)
                {
                    if (!(trqival > 85 && tpcival > 70))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!I{0}>=70,IF(技术状况明细表!I{0}>=85,\"日常养护\",IF(技术状况明细表!I{0}>=75,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\"))),IF(技术状况明细表!I{0}>=60,IF(技术状况明细表!I{0}>=85,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\")),IF(技术状况明细表!I{0}>=40,IF(技术状况明细表!I{0}>=75,\"中修\",\"大修\"),\"大修\")))", i + 3);
                        yhi++;
                    }
                }
                else if (_Setting.YHType == 2)
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
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!I{0}<60,\"大修\",\"中修\")", i + 3);
                        yhi++;
                    }
                }
            }

            MSExcel.Worksheet worksheet = workbook.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:U{0}", len + 2));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (prjinfo._IsRut) destrange = worksheet.get_Range(string.Format("F2:F{0}, I2:K{0},U2:U{0}", len + 2));
            else destrange = worksheet.get_Range(string.Format("F2:F{0}, I2:J{0},U2:U{0}", len + 2));
            MSExcel.ChartObject chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            MSExcel.Chart chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, true, "", Type.Missing, Type.Missing, Type.Missing);
            chart.Legend.Position = MSExcel.XlLegendPosition.xlLegendPositionTop;

            destrange = worksheet.get_Range(string.Format("F2:F{0}, M2:M{0}", len + 2));
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(2);
            chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "破损率DR(%)", Type.Missing, Type.Missing, Type.Missing);

            destrange = worksheet.get_Range(string.Format("F2:F{0}, N2:N{0}", len + 2));
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(3);
            chart = chartobj.Chart;
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "平整度IRI", Type.Missing, Type.Missing, Type.Missing);

            if (prjinfo._IsRut)
            {
                destrange = worksheet.get_Range(string.Format("F2:F{0}, O2:O{0}", len + 2));
                chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(4);
                chart = chartobj.Chart;
                chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "车辙Rut", Type.Missing, Type.Missing, Type.Missing);
            }

            worksheet = workbook.Sheets["养护需求建议表"] as MSExcel.Worksheet;
            destrange = worksheet.get_Range(String.Format("A3:G{0}", yhi + 2));
            destrange.Value2 = yhlist;
            GlobalExcel.SetBorderLine(destrange, 53);

            object[,] tjlist = new object[4, 1];
            worksheet = workbook.Sheets["分项指标统计表"] as MSExcel.Worksheet;
            tjlist[0, 0] = String.Format("==SUMPRODUCT(技术状况明细表!G3:G{1}, 技术状况明细表!{0}3:{0}{1})/SUM(技术状况明细表!G3:G{1})", 'U', len + 2);
            for (int i = 1; i < 4; ++i)
            {
                tjlist[i, 0] = String.Format("=SUMPRODUCT(技术状况明细表!G3:G{1}, 技术状况明细表!{0}3:{0}{1})/SUM(技术状况明细表!G3:G{1})", 
                    GlobalExcel.GetCol(((char)('H' + i))), len + 2);
            }
            destrange = worksheet.get_Range("B3:B6");
            destrange.Value2 = tjlist;
            if (prjinfo._IsRut) destrange = worksheet.get_Range("L2:Q6");
            else destrange = worksheet.get_Range("L2:Q5");
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            chart = chartobj.Chart;
            chart.SetSourceData(destrange);
        }
        private static void WriteRoadInfo(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string degreeinfo = File.ReadAllText(prjdir.FullName + "\\DegreeInfo.txt").Replace(" ", Environment.NewLine);
            string roadtypeinfo = File.ReadAllText(prjdir.FullName + "\\RoadTypeInfo.txt").Replace(" ", Environment.NewLine); ;
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
            valobj[5, 0] = Math.Max(prjinfo._StartMile * 0.001, prjinfo._EndMile * 0.001);
            destrange.Value2 = valobj;
        }

        //中南安环--评定汇总表    
        public static void OutputZNRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中南安环\综合报表模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
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
            WriteZNHZ2Xls(_Worksheet_hz, _Worksheet_tj, sheetIRI, sheetRUT, sheetDR,
                prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZNHZ2Xls(MSExcel.Worksheet worksheet, MSExcel.Worksheet worksheet2,
            MSExcel.Worksheet sheetIRI, MSExcel.Worksheet sheetRUT, MSExcel.Worksheet sheetDR,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal)
        {
            worksheet.Cells[1, 1] = prjinfo._RoadCode + prjinfo._RoadName
                + prjinfo._StartMile.ToString("K0+000") + "~"
                + prjinfo._EndMile.ToString("K0+000") + "段\r\n路面使用性能指数评定汇总表";

            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0, tpcival = 0;

            object[,] mxlist = new object[len, 15];
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
                mxlist[i, 13] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 14] = roadpart[i].degreestr.Replace("公路", "");

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
                tpcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 6] = drval;
                mxlist[i, 7] = string.Format("=100-{0}*POWER(G{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 4, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 11] = repair > 0 ? Math.Round(repair * 100 / ksumarea, 5) : 0;

                //IRI
                if (prjinfo._IsDIRIMTD)
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                else
                    irival = Math.Round(LIRIVal[i], 5);
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                mxlist[i, 4] = irival;
                mxlist[i, 5] = String.Format("=ROUND(100/(1+{0}*EXP({1}*E{2})),5)", _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 4);

                //Rut
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
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
                if (roadpart[i].roaddegree < 2 )
                {
                    mxlist[i, 10] = string.Format("=ROUND(({1}*H{0}+{2}*F{0}+{3}*IF(EXACT(J{0},\" \"),0,J{0}))/({1}+{2}+{3}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                }
                else
                {
                    mxlist[i, 10] = string.Format("=ROUND(({1}*H{0}+{2}*F{0})/({1}+{2}),5)",
                        i + 4,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 12] = string.Format("=IF(K{0}>={1},\"优\",IF(K{0}>={2},\"良\",IF(K{0}>={3},\"中\",IF(K{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
            }

            int datarow = len + 3;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A4:O{0}", datarow));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 63);
            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 4, 1, 15, true);
                GlobalExcel.Reflection(worksheet, 4, 1, 2, false);
            }

            //汇总部分
            MSExcel.Range srcrange = worksheet2.get_Range("A1:O10");
            destrange = worksheet.get_Range(String.Format("A{0}", datarow + 1));
            srcrange.Copy(destrange);

            worksheet.Cells[datarow + 1, 4] = string.Format("=SUM(D4:D{0})", datarow);
            for (int i = 0; i < 8; i++)
            {
                worksheet.Cells[datarow + 1, 5 + i] = string.Format("=SUMPRODUCT(D4:D{1},{0}4:{0}{1})/SUM(D4:D{1})", (char)('E' + i), datarow);
            }

            worksheet.Cells[datarow + 1, 13] = string.Format("=IF(K{0}>={1},\"优\",IF(K{0}>={2},\"良\",IF(K{0}>={3},\"中\",IF(K{0}>={4},\"次\",\"差\"))))",
                    datarow + 1,
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2],
                     _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);


            worksheet.Cells[datarow + 3, 5] = string.Format("=SUMIF(F4:F{0},\">={1}\",D4:D{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 5] = string.Format("=SUMIFS(D4:D{0},F4:F{0},\">={1}\",F4:F{0},\"<{2}\")/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 5] = string.Format("=SUMIF(F4:F{0},\"<{1}\",D4:D{0})/1000", datarow,
                _RQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            worksheet.Cells[datarow + 3, 10] = string.Format("=SUMIF(H4:H{0},\">={1}\",D4:D{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 10] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 10] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 10] = string.Format("=SUMIFS(D4:D{0},H4:H{0},\">={1}\",H4:H{0},\"<{2}\")/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 10] = string.Format("=SUMIF(H4:H{0},\"<{1}\",D4:D{0})/1000", datarow,
                _PCIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            worksheet.Cells[datarow + 3, 14] = string.Format("=SUMIF(K4:K{0},\">={1}\",D4:D{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 4, 14] = string.Format("=SUMIFS(D4:D{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][0]);
            worksheet.Cells[datarow + 5, 14] = string.Format("=SUMIFS(D4:D{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][1]);
            worksheet.Cells[datarow + 6, 14] = string.Format("=SUMIFS(D4:D{0},K4:K{0},\">={1}\",K4:K{0},\"<{2}\")/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3], _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][2]);
            worksheet.Cells[datarow + 7, 14] = string.Format("=SUMIF(K4:K{0},\"<{1}\",D4:D{0})/1000", datarow,
                _PQIGrade[_RoadGradeDict[prjinfo._RoadGrade]][3]);

            for (int i = 0; i < 5; ++i)
            {
                worksheet.Cells[datarow + 3 + i, 6] = string.Format("=E{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 11] = string.Format("=J{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
                worksheet.Cells[datarow + 3 + i, 15] = string.Format("=N{0}/D{1}*100000", datarow + 3 + i, datarow + 1);
            }

            //将数值复制进去
            srcrange = worksheet.get_Range(string.Format("A4:B{0}", datarow));
            destrange = sheetIRI.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range("A10");
            srcrange.Copy(destrange);

            string prjname = string.Format("{0}-{1}", prjinfo._RoadCode, prjinfo._RoadName);
            sheetIRI.Cells[7, 3] = prjname;
            sheetRUT.Cells[7, 3] = prjname;
            sheetDR.Cells[7, 3] = prjname;

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
        }
              //原始数据 RQI PCI RDI 
        public static void OutputZNDataRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中南安环\原始数据记录表.xlsx",
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

            WriteIRI2Xls(_Worksheet_RQI, prjinfo, _RoadPart, _LIRIMeanVal, _RIRIMeanVal);
            WriteRut2Xls_orirut(_Worksheet_RDI, prjinfo, _RoadPart, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal);
            WritePCI2Xls(_Worksheet_PCI, prjinfo, prjdir, _RoadPart, _RoadDisList);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        public static void OutputZNRoadDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中南安环\病害统计表.xlsx",
                           System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}.xlsx", path, prjdir.Name,  "病害统计表");
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

            string disname=null;
            string disgrade=null;
            int colidx = 3;
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
            int len = roadpart.Count - 1, dlen = dislist.Length;
            object[,] disinfo = new object[dlen, 8];
            for (int i = 0; i < dlen; i++)//i区间索引，j病害索引
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
                    }
                    else
                    {
                        _Worksheet_hz.Cells[m, colidx] = string.Format("=COUNTIF(病害列表！D:D,\"{0}\")", distype[n]);
                        _Worksheet_hz.Cells[m, colidx + 1] = string.Format("=SUMIF(病害列表!D:D,\"{0}\",病害列表!H:H)", distype[n]);
                    }
                }
            }
           
            if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = _Worksheet_mx.get_Range(String.Format("A3:H{0}", dlen + 2));
                sortrange = _Worksheet_mx.get_Range(String.Format("A3:A{0}", dlen + 2));
                GlobalExcel.ReflectionColnum(_Worksheet_mx, destrange, sortrange);
            }
        }

        //中交国通报表模板
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
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\001-路面平整度报告模板.xlsx",
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
                    vallist[i, datacol] = Math.Round((LIRIVal[i] + RIRIVal[i]) / 2, 5);
                else
                    vallist[i, datacol] = Math.Round(LIRIVal[i], 5);

                vallist[i, datacol + 1] = String.Format("=ROUND(100/(1+{0}*EXP({1}*{3}{2})),5)",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 11, (char)('A' + datacol));
                vallist[i, datacol + 2] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    i + 11,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3],
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
                        _RQIa[roadpart[0].roaddegree][0], _RQIa[roadpart[0].roaddegree][1], len, (char)('A' + datacol));
                _Worksheet.Cells[len, datacol + 3] = string.Format("=IF({5}{0}>={1},\"优\",IF({5}{0}>={2},\"良\",IF({5}{0}>={3},\"中\",IF({5}{0}>={4},\"次\",\"差\"))))",
                    len,
                    _RQIGrade[roadpart[0].roaddegree][0],
                    _RQIGrade[roadpart[0].roaddegree][1],
                    _RQIGrade[roadpart[0].roaddegree][2],
                    _RQIGrade[roadpart[0].roaddegree][3],
                     (char)('A' + datacol + 1));

                GlobalExcel.WriteExcel(++len, 1, 1, 3, "备注", _Worksheet, 63);
                GlobalExcel.WriteExcel(len, 4, 1, 6, "---", _Worksheet, 63);
                destrange = _Worksheet.get_Range(String.Format("A11:I{0}", len));
            }
            _Worksheet.Cells[11, 10] = "起点";
            _Worksheet.Cells[len + 10, 10] = "终点";
            GlobalExcel.SetBorderLine(destrange, 63);
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
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\003-路面车辙报告模板.xlsx",
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
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\013-路面车辙原始.xlsx",
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
                vallist[i, datacol + 1] = string.Format("=IF({7}{0}<{1},{2}-{3}*{7}{0},IF({7}{0}<{4},{5}-{6}*({7}{0}-{1}),0))",
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
                    _Worksheet.Cells[len, 8] = string.Format("=IF(G{0}<{1},{2}-{3}*G{0},IF(G{0}<{4},{5}-{6}*(G{0}-{1}),0))",
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
                if (!(roadpart[i].roadtype == 0 && roadpart[i].roaddegree <= 1))
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
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\004-路面抗滑性能报告模板.xlsx",
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
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\002-路面损坏报告模板.xlsx",
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
                WriteZJGTDisHZTJ2Xls(_Worksheet_snhz, _Worksheet_lqhz, prjinfo, prjdir, _RoadPart, _RoadDisList, disval * 10);
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
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            MSExcel.Range srcrange, destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            int sn_tablerow = _Setting.cmop_rows;
            int lq_tablerow = _Setting.cmop_rows;
            //const int sn_tablerow = 29;
            //const int lq_tablerow = 30;
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

        public static void OutputZJGTPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\011-综合大表.xlsx",
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
                _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteZJGTAll2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal,
            double[] LRutVal, double[] RRutVal, double[] SRutVal, double[] LMTDVal, double[] RMTDVal)
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
            double trqival = 0, irival = 0, mtdval = 0, tpcival = 0;

            object[,] mxlist = new object[len, 25];
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
                mxlist[i, 21] = milelength;

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
                mxlist[i, 18] = drval;
                mxlist[i, 19] = string.Format("=100-{0}*POWER(S{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 2, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 20] = string.Format("=IF(T{0}>={1},\"优\",IF(T{0}>={2},\"良\",IF(T{0}>={3},\"中\",IF(T{0}>={4},\"次\",\"差\"))))",
                    i + 2, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI                
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }
                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                mxlist[i, 12] = irival;

                mxlist[i, 13] = string.Format("=ROUND(100/(1+{0}*EXP({1}*M{2})),5)", _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 2);
                mxlist[i, 14] = string.Format("=IF(N{0}>={1},\"优\",IF(N{0}>={2},\"良\",IF(N{0}>={3},\"中\",IF(N{0}>={4},\"次\",\"差\"))))",
                    i + 2, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2], _RQIGrade[roadpart[i].roaddegree][3]);

                //MTD
                if (prjinfo._IsDIRIMTD)
                {
                    mtdval = Math.Round((LMTDVal[i] + RMTDVal[i]) * 0.5, 5);
                }
                else
                {
                    mtdval = Math.Round(LMTDVal[i], 5);
                }
                mxlist[i, 11] = mtdval;

                //Rut
                if (prjinfo._IsRut)
                {
                    //double rutval = Math.Max(LRutVal[i], RRutVal[i]);
                    double rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);
                    mxlist[i, 15] = rutval;

                    mxlist[i, 16] = string.Format("=IF(P{0}<{1},{2}-{3}*P{0},IF(P{0}<{4},{5}-{6}*(P{0}-{1}),0))",
                        i + 2,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);
                    mxlist[i, 17] = string.Format("=IF(Q{0}>={1},\"优\",IF(Q{0}>={2},\"良\",IF(Q{0}>={3},\"中\",IF(Q{0}>={4},\"次\",\"差\"))))",
                        i + 2,
                        _RDIGrade[roadpart[i].roaddegree][0],
                        _RDIGrade[roadpart[i].roaddegree][1],
                        _RDIGrade[roadpart[i].roaddegree][2],
                        _RDIGrade[roadpart[i].roaddegree][3]);
                }

                //PQI
                if (roadpart[i].roaddegree < 2)
                {
                    mxlist[i, 22] = string.Format("=ROUND(({1}*T{0}+{2}*N{0}+{3}*IF(EXACT(Q{0},\"-\"),0,Q{0}))/({1}+{2}+{3}),5)",
                        i + 2,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                }
                else
                {
                    mxlist[i, 22] = string.Format("=ROUND(({1}*T{0}+{2}*N{0})/({1}+{2}),5)",
                        i + 2,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                mxlist[i, 23] = string.Format("=IF(W{0}>={1},\"优\",IF(W{0}>={2},\"良\",IF(W{0}>={3},\"中\",IF(W{0}>={4},\"次\",\"差\"))))",
                    i + 2,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A2:Y{0}", len + 1));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 63);

            worksheet.Cells[2, 26] = "起点";
            worksheet.Cells[len + 1, 26] = "终点";
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
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\005-路面高程、GPS.xlsx",
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
                string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中交国通\014-路面材质.xlsx",
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
        private static void WritePrj2CPMSXls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo)
        {
            _Worksheet.Cells[3, 2] = prjinfo._RoadCode;
            if (prjinfo._Direction > 0)
            {
                _Worksheet.Cells[3, 4] = "上行";
            }
            else
            {
                _Worksheet.Cells[3, 4] = "下行";
            }
            _Worksheet.Cells[3, 8] = prjinfo._DataDate;
            _Worksheet.Cells[4, 8] = prjinfo._StartMile.ToString("K0+000");
            _Worksheet.Cells[4, 13] = prjinfo._EndMile.ToString("K0+000");
        }

        //带GPS重庆招商局报表模板
        public static void OutputGPSRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\综合报表模板GPS.xlsx",
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
            MSExcel.Worksheet _Worksheet_RDI = _Workbook.Sheets["RDI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_DR = _Workbook.Sheets["DR"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_IRI = _Workbook.Sheets["IRI"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_MTD = _Workbook.Sheets["MTD"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_Rut = _Workbook.Sheets["RD"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sn = _Workbook.Sheets["水泥病害"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_lq = _Workbook.Sheets["沥青病害"] as MSExcel.Worksheet;
            if (prjinfo._Direction > 0)
            {
                WriteGPSAll2XlsUp(_Worksheet_PQI, _Worksheet_PCI, _Worksheet_RQI, _Worksheet_RDI,
                    _Worksheet_MTD, _Worksheet_DR, _Worksheet_IRI, _Worksheet_Rut,
                    prjinfo, prjdir, _RoadPart, _RoadDisList,
                    _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _GPSInfo);

            }
            else
            {
                WriteGPSAll2XlsDown(_Worksheet_PQI, _Worksheet_PCI, _Worksheet_RQI, _Worksheet_RDI,
                                  _Worksheet_MTD, _Worksheet_DR, _Worksheet_IRI, _Worksheet_Rut,
                                  prjinfo, prjdir, _RoadPart, _RoadDisList,
                                  _LIRIMeanVal, _RIRIMeanVal, _LRutMeanVal, _RRutMeanVal, _SRutMeanVal, _LMTDMeanVal, _RMTDMeanVal, _GPSInfo);
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
            ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal, 
            double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] IDObj = new object[len, 1];
            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 3];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];

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
                GPSStartObj[rowcnt, 2] = GPSInfo[i]._longitude; ;
                GPSStartObj[rowcnt, 3] = GPSInfo[i]._latitude;

                GPSEndObj[rowcnt, 0] = GPSInfo[i + 1]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i + 1]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i + 1]._latitude;

                //病害相关
                double drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                if (drval > 100) drval = 0;
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
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                    IRIObj[rowcnt, 1] = RIRIVal[i];
                    IRIObj[rowcnt, 2] = string.Format("=(J{0}+K{0})/2", rowcnt + 2);
                }
                else
                {
                    IRIObj[rowcnt, 0] = LIRIVal[i];
                    IRIObj[rowcnt, 2] = string.Format("=J{0}", rowcnt + 2);
                }
                RQIObj[rowcnt, 0] = string.Format("=100/(1+{0}*EXP(IRI!L{2}*{1}))",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], rowcnt + 2);

                //构造深度相关              
                if (prjinfo._IsDIRIMTD)
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 1] = RMTDVal[i];
                    MTDObj[rowcnt, 2] = string.Format("=(J{0}+K{0})/2", rowcnt + 2);
                }
                else
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 2] = string.Format("=J{0}", rowcnt + 2);
                }

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
                    PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})+{3}*IF(EXACT((RDI!J{0}),\"-\"),0,(RDI!J{0})))/({1}+{2}+{3}),5)",
                            rowcnt + 2,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
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
            MSExcel.Worksheet[] tsheet = { worksheetIRI,worksheetDR, worksheetRut, worksheetMTD, worksheetPCI,worksheetRQI, worksheetRDI, worksheetPQI };
            object[] tobj = { IRIObj, DRObj, RutObj, MTDObj, PCIObj, RQIObj, RDIObj, PQIObj };
            char[] valnum = { 'L','J', 'L', 'L', 'J',  'J', 'J', 'J' };
            for (int i = 0; i < tsheet.Length; ++i)
            {
                destrange = tsheet[i].get_Range(String.Format("A2:A{0}", len + 1));
                destrange.Value2 = IDObj;
                
                if (i <= 4)
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
                if (i > 4) 
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }
            }
        }

        // 下行
        private static void WriteGPSAll2XlsDown(
           MSExcel.Worksheet worksheetPQI, MSExcel.Worksheet worksheetPCI,
           MSExcel.Worksheet worksheetRQI, MSExcel.Worksheet worksheetRDI,
           MSExcel.Worksheet worksheetMTD, MSExcel.Worksheet worksheetDR,
           MSExcel.Worksheet worksheetIRI, MSExcel.Worksheet worksheetRut,
           ProjectInfo prjinfo, DirectoryInfo prjdir,
           List<MilePart> roadpart, Disease[] arrdis,
           double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal,
           double[] LMTDVal, double[] RMTDVal, ExcelGPS[] GPSInfo)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] GPSStartObj = new object[len, 4];
            object[,] GPSEndObj = new object[len, 4];
            object[,] IDObj = new object[len, 1];
            object[,] PQIObj = new object[len, 1];
            object[,] PCIObj = new object[len, 1];
            object[,] RQIObj = new object[len, 1];
            object[,] RDIObj = new object[len, 1];
            object[,] MTDObj = new object[len, 3];
            object[,] DRObj = new object[len, 1];
            object[,] IRIObj = new object[len, 3];
            object[,] RutObj = new object[len, 3];

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;

            int typeidx = 0;
            bool res = false;
			
            for (int i = len-1, j = dlen-1; i >=0; i--)//i区间索引，j病害索引
            {
                int smile = roadpart[i+1].mile;
                int emile = roadpart[i  ].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                RoadDiseaseTypes.Clear();
                while (j >=0 && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
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
                IDObj[rowcnt, 0] = len-i ;
                GPSStartObj[rowcnt, 0] = GPSInfo[i+1]._utctime;
                GPSStartObj[rowcnt, 1] = smile;
                GPSStartObj[rowcnt, 2] = GPSInfo[i+1]._longitude; ;
                GPSStartObj[rowcnt, 3] = GPSInfo[i+1]._latitude;

                GPSEndObj[rowcnt, 0] = GPSInfo[i ]._utctime;
                GPSEndObj[rowcnt, 1] = emile;
                GPSEndObj[rowcnt, 2] = GPSInfo[i ]._longitude;
                GPSEndObj[rowcnt, 3] = GPSInfo[i ]._latitude;

                //病害相关
                double drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                DRObj[rowcnt, 0] = drval;
                PCIObj[rowcnt, 0] = string.Format("=100-{0}*POWER(DR!J{1},{2})",
                              _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                              rowcnt + 2, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

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
                RQIObj[rowcnt, 0] = string.Format("=100/(1+{0}*EXP(IRI!L{2}*{1}))",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], rowcnt + 2);

                //构造深度相关              
                if (prjinfo._IsDIRIMTD)
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 1] = RMTDVal[i];
                    MTDObj[rowcnt, 2] = string.Format("=(J{0}+K{0})/2", rowcnt + 2);
                }
                else
                {
                    MTDObj[rowcnt, 0] = LMTDVal[i];
                    MTDObj[rowcnt, 2] = string.Format("=J{0}", rowcnt + 2);
                }

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
                    PQIObj[rowcnt, 0] = string.Format("=ROUND(({1}*(PCI!J{0})+{2}*(RQI!J{0})+{3}*IF(EXACT((RDI!J{0}),\"-\"),0,(RDI!J{0})))/({1}+{2}+{3}),5)",
                            rowcnt + 2,
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                            _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
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
            MSExcel.Worksheet[] tsheet = { worksheetIRI, worksheetDR, worksheetRut, worksheetMTD, worksheetPCI, worksheetRQI, worksheetRDI, worksheetPQI };
            object[] tobj = { IRIObj, DRObj, RutObj, MTDObj, PCIObj, RQIObj, RDIObj, PQIObj };
            char[] valnum = { 'L', 'J', 'L', 'L', 'J', 'J', 'J', 'J' };
            for (int i = 0; i < tsheet.Length; ++i)
            {
                destrange = tsheet[i].get_Range(String.Format("A2:A{0}", len + 1));
                destrange.Value2 = IDObj;

                if (i <= 4)
                {
                    destrange = tsheet[i].get_Range(String.Format("J2:{0}{1}", valnum[i], len + 1));
                    destrange.Value2 = tobj[i];
                }

                if (_Setting.IsExcelSort && prjinfo._Direction < 0)
                {
                    destrange = tsheet[i].get_Range(String.Format("B2:E{0}", len + 1));
                    destrange.Value2 = GPSStartObj;  // GPSEndObj
                    destrange = tsheet[i].get_Range(String.Format("F2:I{0}", len + 1));
                    destrange.Value2 = GPSEndObj; // GPSStartObj
                    destrange = tsheet[i].get_Range(String.Format("B2:{0}{1}", valnum[i], len + 1));
                    sortrange = tsheet[i].get_Range(String.Format("C2:C{0}", len + 1));
                    GlobalExcel.ReflectionColnum(tsheet[i], destrange, sortrange);
                }
         
                if (i > 4)
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
            string disname="";
            string disgrade="";
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
                if (gi < (tempinfos.Length-1))
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
            else if(_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_lq.get_Range(String.Format("A2:Q{0}", rowcnt_lq - 1));
                sortrange = worksheet_lq.get_Range(String.Format("C2:C{0}", len + 1));
                GlobalExcel.ReflectionColnum(worksheet_lq, destrange, sortrange);

            }
            if (rowcnt_sn < 3)
            {
                worksheet_sn.Delete();
            }
            else if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            {
                destrange = worksheet_sn.get_Range(String.Format("A2:Q{0}", rowcnt_sn - 1));
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

            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\综合报表模板GPS _景观图像.xlsx",
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
        public static void OutputGPSRoadImg(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            string fname = prjdir.FullName + "\\RoadImg\\Camera0\\Road2Mile.txt";
            if (!File.Exists(fname))
            {      
                MessageBox.Show("工程文件缺少路面图像数据");
                return;
            }
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\综合报表模板GPS _路面图像.xlsx",
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
            string fnamemile0 = prjdir.FullName + "\\" + ImgType + "Img\\Camera0\\" + ImgType + "2Mile.txt";
            string fnamemile1 = prjdir.FullName + "\\" + ImgType + "Img\\Camera1\\" + ImgType + "2Mile.txt";
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
                            dataobj[leftidx[i], colcnt++] = "\\" + ImgType + "Img\\Camera0\\" + leftimgsinfo[i].Substring(temp2, temp - temp2);
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
                                dataobj[rightidx[i], colcnt++] = "\\" + ImgType + "Img\\Camera1\\" + rightimgsinfo[i].Substring(temp2, temp - temp2);
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
                srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\奥路通\001-路面病害模板.xlsx",
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
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt] = string.Format("K{0:000+000}-K{1:000+000}", roadpart[i].mile, roadpart[i + 1].mile);

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
                    ++rowcnt_sn_s;
                }
                else if (roadpart[i].roadtype == 0)//沥青
                {
                    Haslqflag = true;

                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt] = string.Format("K{0:000+000}-K{1:000+000}", roadpart[i].mile, roadpart[i + 1].mile);
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 1];
                    for (int di = 0, kk = 1; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }
                    //修补=条状修补+块状修补
                    disval[0, disnum - 1] = Convert.ToDouble(disval[0, disnum - 1]) + Convert.ToDouble(disval[0, disnum]);
                    disval[0, disnum] = null;
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, 0] = string.Format("=100-{0}*POWER({1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    destrange = worksheet_lqhz.get_Range(string.Format("B{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('B' + disnum))));
                    destrange.Value2 = disval;

                    totallqlen += milelength;
                    ++rowcnt_lq_s;
                }

                if (emile % 1000 == 0)
                {
                    if (roadpart[i].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                        disval = new object[1, disnum + 1];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                        }
                        destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum))));
                        destrange.Value2 = disval;
                        ++rowcnt_sn_s;
                        rowcnt_sn_e = rowcnt_sn_s;

                        if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            disval = new object[1, disnum + 1];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                            }
                            destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum))));
                            destrange.Value2 = disval;
                            ++rowcnt_lq_s;
                            rowcnt_lq_e = rowcnt_lq_s;
                        }
                    }
                    else if (roadpart[i].roadtype == 0)//沥青
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                        disval = new object[1, disnum + 1];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                        }
                        destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum))));
                        destrange.Value2 = disval;
                        ++rowcnt_lq_s;
                        rowcnt_lq_e = rowcnt_lq_s;

                        if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                        {
                            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                            worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            disval = new object[1, disnum + 1];
                            for (int di = 0; di < disnum; di++)
                            {
                                disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                            }
                            destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum))));
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
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                    disval = new object[1, disnum + 1];
                    for (int di = 0; di < disnum; di++)
                    {
                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                    }
                    destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum))));
                    destrange.Value2 = disval;
                    ++rowcnt_sn_s;
                    rowcnt_sn_e = rowcnt_sn_s;

                    if (Haslqflag && rowcnt_lq_e < rowcnt_lq_s)
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                        disval = new object[1, disnum + 1];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                        }
                        destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum))));
                        destrange.Value2 = disval;
                        ++rowcnt_lq_s;
                        rowcnt_lq_e = rowcnt_lq_s;
                    }
                }
                else if (roadpart[len].roadtype == 0)
                {
                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                    disval = new object[1, disnum + 1];
                    for (int di = 0; di < disnum; di++)
                    {
                        disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_e, rowcnt_lq_s - 1);
                    }
                    destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum))));
                    destrange.Value2 = disval;
                    ++rowcnt_lq_s;
                    rowcnt_lq_e = rowcnt_lq_s;

                    if (Hassnflag && rowcnt_sn_e < rowcnt_sn_s)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                        disval = new object[1, disnum + 1];
                        for (int di = 0; di < disnum; di++)
                        {
                            disval[0, di] = string.Format("=SUM({0}{1}:{0}{2})", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_e, rowcnt_sn_s - 1);
                        }
                        destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum))));
                        destrange.Value2 = disval;
                        ++rowcnt_sn_s;
                        rowcnt_sn_e = rowcnt_sn_s;
                    }
                }
            }

     
            //总计
            //水泥
            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0);
            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
            disval = new object[1, disnum + 1];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_sn_s - 1);
            }
            destrange = worksheet_snhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('C' + disnum))));
            destrange.Value2 = disval;

            //沥青
            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
            disval = new object[1, disnum + 1];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}5:{0}{1})/2", GlobalExcel.GetCol((char)('C' + di)), rowcnt_lq_s - 1);
            }
            destrange = worksheet_lqhz.get_Range(string.Format("C{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('C' + disnum))));
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
        private static void WriteALTPQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis,
            double[] LIRIVal, double[] RIRIVal, double[] LRutVal, double[] RRutVal, double[] SRutVal
            )
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double  irival = 0, rutval = 0;
            worksheet.Cells[2, 8] = prjinfo._RoadNum;
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

                vallist[rowcnt, colcnt] = string.Format("K{0:000+000}-K{1:000+000}", roadpart[i].mile, roadpart[i + 1].mile);
                vallist[rowcnt, colcnt + 1] =Math.Abs( emile - smile);

                // DR PCI
                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt + 2] = drval;

                vallist[rowcnt, colcnt + 5] = string.Format("=100-{0}*POWER(C{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 5, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);


                //IRI RQI
                if (prjinfo._IsDIRIMTD)
                {
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) * 0.5, 5);
                }
                else
                {
                    irival = Math.Round(LIRIVal[i], 5);
                }

                vallist[rowcnt, colcnt + 3] = Math.Round(irival, 5);
                vallist[rowcnt, colcnt + 6] = string.Format("=ROUND(100/(1+{0}*EXP({1}*D{2})),5)",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 5);


                if (prjinfo._IsRut)
                {
                    rutval = SRutVal[i];
                    rutval = Math.Round(rutval, 5);

                    vallist[rowcnt, colcnt + 4] = Math.Round(rutval, 5);
                    vallist[rowcnt, colcnt + 7] = string.Format("=IF(E{0}<{1},{2}-{3}*E{0},IF(E{0}<{4},{5}-{6}*(E{0}-{1}),0))",
                        i + 5,
                        _RDIRD[0][1],
                        _RDIRD[0][0],
                        _RDIa[0],
                        _RDIRD[1][1],
                        _RDIRD[1][0],
                        _RDIa[1]);

                }

               // PQI
                if (roadpart[i].roaddegree <= 1)
                {

                    // 参数依次为 PCI RQI RDI
                    vallist[rowcnt, colcnt+8] = string.Format("=ROUND(({1}*F{0}+{2}*G{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0}))/({1}+{2}+{3}),5)",
                           i + 5,
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                           _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                }
                else
                {
                    vallist[rowcnt, colcnt + 8] = string.Format("=ROUND(({1}*F{0}+{2}*G{0})/({1}+{2}),5)",
                    i + 5,
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                }
                vallist[rowcnt, colcnt + 9] = roadpart[i].degreestr;
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A5:J{0}", rowcnt + 4));
            destrange.Value2 = vallist;
            GlobalExcel.SetBorderLine(destrange, 53);

            //if (_Setting.IsExcelSort && prjinfo._Direction < 0)
            //{
            //    destrange = worksheet.get_Range(string.Format("A5:J{0}", rowcnt + 4));
            //    MSExcel.Range sortrange = worksheet.get_Range(string.Format("A5:A{0}", rowcnt + 4));
            //    GlobalExcel.ReflectionColnum(worksheet, destrange, sortrange);
            //}

            int chartlen = len + 4;
            MSExcel.ChartObject chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            MSExcel.Chart chart = chartobj.Chart;
            destrange = worksheet.get_Range(string.Format("A4:A{0},F3:F{0}, G3:G{0}, H3:H{0}, I3:I{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, true, "PQI各项指标", Type.Missing, Type.Missing, Type.Missing);

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
            destrange = worksheet.get_Range(string.Format("A4:A{0}, E3:E{0}", chartlen));
            chart.ChartWizard(destrange, MSExcel.XlChartType.xlLine, 2, MSExcel.XlRowCol.xlColumns, 1, 1, false, "RD", Type.Missing, Type.Missing, Type.Missing);
        }
        #endregion
    }
}
