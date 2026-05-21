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
    /// 北京农村路报表
    /// </summary>
    class MyExcelBJDegree
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

        private static double[][][] _RQIa;//公路等级 参数序号
        private static double[][][] _PCIa;//公路等级 路面材质 参数序号
        private static double[][][] _PQIW;//公路等级 路面材质 参数序号
        private static double[][] _WeightParm;//0-沥青，1-水泥
        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        private static Dictionary<string, int> _RoadGradeDict;

        private static List<MilePart> _RoadPart = null;
        public static List<MilePart> _RoadPart1M = null;//1米桩号分段
        private static Disease[] _RoadDisList = null;
        private static Disease[] _RoadRepairList = null;
        private static double[] _LIRIMeanVal = null;
        private static double[] _RIRIMeanVal = null;
        private static double[] _SRutDisVal = null;
        private static int[] _SRutDisMile = null;
        private static double[] _rutThresh = new double[2];

        private static void InitXlsParm()
        {
            int len = _RoadGradeStr.Length;

            _RQIGrade = new double[len][];
            _RDIGrade = new double[len][];
            _MTDGrade = new double[len][];
            _PCIGrade = new double[len][];
            _PQIGrade = new double[len][];

            _RQIa = new double[len][][];
            _PCIa = new double[len][][];
            _PQIW = new double[len][][];

            for (int i = 0; i < len; i++)
            {
                _RQIGrade[i] = new double[5];
                _RDIGrade[i] = new double[5];
                _MTDGrade[i] = new double[5];
                _PCIGrade[i] = new double[5];
                _PQIGrade[i] = new double[5];

                _RQIa[i] = new double[2][];
                _PCIa[i] = new double[2][];
                _PQIW[i] = new double[2][];
                for (int j = 0; j < 2; j++)
                {
                    _RQIa[i][j] = new double[2];
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
                                        foreach (XmlNode nnode in node.ChildNodes)
                                        {
                                            _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a0"));
                                            _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a1"));
                                        }
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

        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval, bool IsDis, bool IsIRI)
        {
            bool IRIRes = true;
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
                spart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade] };
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
                GlobalExcel.GetRutDisVal(prjinfo, prjdir, _RoadPart1M, ref _SRutDisVal, ref _SRutDisMile);
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, _RoadGradeDict, _SRutDisVal, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh,_RoadPart);
            }
            if (IsIRI) IRIRes = GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, _Setting.IsWarning);

            return IRIRes;
        }

        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\北京农村公路\路面平整度评价等级记录表.xlsx",
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
        private static void WriteIRI2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, List<MilePart> roadpart, double[] LIRIVal, double[] RIRIVal)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            object[,] vallist = new object[len, 8];
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                vallist[i, 3] = LIRIVal[i];
                if (prjinfo._IsDIRIMTD)
                {
                    vallist[i, 4] = RIRIVal[i];
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,2)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),2)", i + 4);
                }

                vallist[i, 6] = String.Format("=ROUND(100/(1+{0}*EXP({1}*F{2})),2)",
                        _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1], i + 4);
                vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                    i + 4,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3]);
            }

            destrange = _Worksheet.get_Range(String.Format("A4:H{0}", len + 3));
            destrange.Value2 = vallist;

            WriteIRIStatistics(_Worksheet);
            destrange = _Worksheet.get_Range(String.Format("A1:H{0}", len + 3));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 1, 8, true);
                GlobalExcel.Reflection(_Worksheet, 4, 1, 2, false);
            }
        }
        private static void WriteIRIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("P3:T5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 15 + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 9] = "=CONCATENATE(\"路面平整度评价等级“优”率占路段总数\",ROUND(P4,4)*100,\"%，“良”率占路段总数\",ROUND(Q4,4)*100,\"%，“中”率占路段总数\",ROUND(R4,4)*100,\"%，“次”率占路段总数\",ROUND(S4,4)*100,\"%，“差”率占路段总数\",ROUND(T4,4)*100,\"%。\")";
        }

        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\北京农村公路\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
            WriteDisLB2Xls(_Worksheet_lb, prjinfo, _RoadDisList);

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
        private static void WriteDisLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0;
            object[,] val = new object[len, 9];
            foreach (Disease tdis in dislist)
            {
                val[i, 0] = tdis.m_mile;
                val[i, 1] = prjinfo._RoadNum;
                val[i, 2] = tdis.RoadDisType;
                val[i, 3] = tdis.realheight;
                val[i, 4] = tdis.realwidth;
                val[i, 5] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                val[i, 6] = tdis.Area;
                val[i, 7] = tdis.calcheight;
                val[i, 8] = tdis.calcwidth;
                ++i;
            }
            destrange = _Worksheet.get_Range(String.Format("A3:I{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:I{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 9, true);
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

            int rowcnt_sn_s = 4;
            int rowcnt_sn_e = 4;//小计起始的计算范围
            int rowcnt_lq_s = 4;
            int rowcnt_lq_e = 4;

            int totalsnlen = 0;//水泥路段总长度
            int totallqlen = 0;//沥青路段总长度

            int typeidx = 0;
            bool res = false;
            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count, dlen = arrdis.Length;
            for (int i = 0, j = 0; i < len - 1; i++)//i区间索引，j病害索引
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

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count - 1;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        if (kk == disnum - 1)
                        {
                            disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di + 1].totalarea;
                        }
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER(N{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_sn_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    disval[0, disnum + 2] = string.Format("=IF({0}{1}>={2},\"优\",IF({0}{1}>={3},\"良\",IF({0}{1}>={4},\"中\",IF({0}{1}>={5},\"次\",\"差\"))))",
                        GlobalExcel.GetCol((char)('D' + disnum + 1)),
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
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = prjinfo._RoadNum;

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count - 1;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                        if (kk == disnum - 1)
                        {
                            disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea + RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di + 1].totalarea;
                        }
                    }
                    drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = drval;
                    disval[0, disnum + 1] = string.Format("=100-{0}*POWER(M{1},{2})",
                        _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        rowcnt_lq_s, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                    disval[0, disnum + 2] = string.Format("=IF({0}{1}>={2},\"优\",IF({0}{1}>={3},\"良\",IF({0}{1}>={4},\"中\",IF({0}{1}>={5},\"次\",\"差\"))))",
                        GlobalExcel.GetCol((char)('D' + disnum + 1)),
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

                if (emile % 1000 == 0)
                {
                    if (roadpart[i].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
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
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
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
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
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
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
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
            if (roadpart[len - 1].mile % 1000 != 0)
            {
                if (roadpart[len - 1].roadtype == 1)
                {
                    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                    worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
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
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
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
                else if (roadpart[len - 1].roadtype == 0)
                {
                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                    worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                    disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
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
                        disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
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
            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}4:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_sn_s - 1);
            }
            destrange = worksheet_snhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_sn_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
            destrange.Value2 = disval;
            destrange = worksheet_snhz.get_Range(String.Format("A1:{0}{1}", GlobalExcel.GetCol((char)('D' + disnum + 2)), rowcnt_sn_s));
            GlobalExcel.SetBorderLine(destrange, 53);

            //沥青
            GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "总计", worksheet_lqhz, 0);
            worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
            disval = new object[1, disnum];
            for (int di = 0; di < disnum; di++)
            {
                disval[0, di] = string.Format("=SUM({0}4:{0}{1})/2", GlobalExcel.GetCol((char)('D' + di)), rowcnt_lq_s - 1);
            }
            destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum - 1))));
            destrange.Value2 = disval;
            destrange = worksheet_lqhz.get_Range(String.Format("A1:{0}{1}", GlobalExcel.GetCol((char)('D' + disnum + 2)), rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, 53);

            RoadDiseaseTypes.Clear();
            if (Haslqflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                worksheet_lqtj.Cells[2, 2] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len - 1].mile);
                worksheet_lqtj.Cells[2, 4] = _RoadConfig.DetectWidth;
                worksheet_lqtj.Cells[4, 4] = string.Format("={0}*D2", Math.Abs(roadpart[0].mile - roadpart[len - 1].mile));
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
                worksheet_lqhz.Delete();
                worksheet_lqtj.Delete();
            }

            if (Hassnflag)
            {
                disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                worksheet_sntj.Cells[2, 2] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len - 1].mile);
                worksheet_sntj.Cells[2, 4] = _RoadConfig.DetectWidth;
                worksheet_sntj.Cells[4, 4] = string.Format("={0}*D2", Math.Abs(roadpart[0].mile - roadpart[len - 1].mile));
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
                worksheet_snhz.Delete();
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
            string srcxls = string.Format(@"{0}\报表模板\北京农村公路\路面破损评价等级记录表.xlsx",
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

                vallist[i, 0] = smile;
                vallist[i, 1] = emile;
                vallist[i, 2] = prjinfo._RoadNum;

                drval = ComputPCI(RoadDiseaseTypes.roaddis, roadpart[i].roadtype, _RoadConfig.DetectWidth * milelength);
                vallist[i, 3] = drval;
                vallist[i, 4] = string.Format("=100-{0}*POWER(D{1},{2})",
                    _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                    i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[i, 5] = string.Format("=IF(E{0}>={1},\"优\",IF(E{0}>={2},\"良\",IF(E{0}>={3},\"中\",IF(E{0}>={3},\"次\",\"差\"))))",
                    i + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);
            }

            destrange = worksheet.get_Range(String.Format("A3:F{0}", len + 2));
            destrange.Value2 = vallist;
            WritePCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:F{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 6, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePCIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("N3:R5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(F:F,\"{0}\",A:A)-SUMIF(F:F,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 13 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 7] = "=CONCATENATE(\"路面PCI评价等级“优”率占路段总数\",ROUND(N4,4)*100,\"%，“良”率占路段总数\",ROUND(O4,4)*100,\"%，“中”率占路段总数\",ROUND(P4,4)*100,\"%，“次”率占路段总数\",ROUND(Q4,4)*100,\"%，“差”率占路段总数\",ROUND(R4,4)*100,\"%。\")";
        }

        public static void OutputPQI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\北京农村公路\路面综合评价等级记录表.xlsx",
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

            WritePQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, _LIRIMeanVal, _RIRIMeanVal);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WritePQI2Xls(MSExcel.Worksheet worksheet, ProjectInfo prjinfo,
            DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, double[] LIRIVal, double[] RIRIVal)
        {
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0;
            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 11];

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
                vallist[rowcnt, colcnt++] = Math.Round(pcival, 2);
                vallist[rowcnt, colcnt++] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);

                if (prjinfo._IsDIRIMTD)
                    irival = Math.Round((LIRIVal[i] + RIRIVal[i]) / 2, 2);
                else
                    irival = Math.Round(LIRIVal[i], 2);

                trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][roadpart[i].roadtype][1] * irival));
                vallist[rowcnt, colcnt] = trqival;

                colcnt++;
                vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3]);

                vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*D{0}+{2}*F{0})/({1}+{2}),2)",
                        rowcnt + 3,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1]);

                vallist[rowcnt, colcnt++] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:I{0}", rowcnt + 2));
            destrange.Value2 = vallist;
            WritePQI2Statistics(worksheet);
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 9, true);
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePQI2Statistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            MSExcel.Range destrange = _Worksheet.get_Range("R3:V5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(I:I,\"{0}\",A:A)-SUMIF(I:I,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 17 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 13] = "=CONCATENATE(\"路面PQI评价等级“优”率占路段总数\",ROUND(R4,4)*100,\"%，“良”率占路段总数\",ROUND(S4,4)*100,\"%，“中”率占路段总数\",ROUND(T4,4)*100,\"%，“次”率占路段总数\",ROUND(U4,4)*100,\"%，“差”率占路段总数\",ROUND(V4,4)*100,\"%。\")";
        }

        ////////////////////////////////////////////// CPMS报表模板 /////////////////////////////////////////////  
        public static void OutputCPMSDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\北京农村公路\CPMS路面损坏调查表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_CPMS病害调查表_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lqdc = _Workbook.Sheets["CPMS_沥青路面损坏调查表"] as MSExcel.Worksheet;
            MSExcel.Worksheet _Worksheet_sndc = _Workbook.Sheets["CPMS_水泥路面损坏调查表"] as MSExcel.Worksheet;
            WriteCPMSDisDC2Xls(_Worksheet_sndc, _Worksheet_lqdc, prjinfo, prjdir, _RoadPart, _RoadDisList, disval);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void WriteCPMSDisDC2Xls(MSExcel.Worksheet worksheet_sndc, MSExcel.Worksheet worksheet_lqdc,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            MSExcel.Range srcrange, destrange;
            int disnum = 0;
            object[,] disval;
            bool Haslqflag = false;//有沥青路段标志
            bool Hassnflag = false;//有水泥路段标志

            const int tablerow = 28;
            int tcnt_sn = 0;
            int tcnt_lq = 0;

            int sn_csmile = 0, sn_cemile = 0;
            int lq_csmile = 0, lq_cemile = 0;
            bool sn_flag = false, lq_flag = false;

            int typeidx = 0;
            bool res = false;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int len = roadpart.Count - 1, dlen = arrdis.Length;
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
                    if(res)
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
                if (roadpart[i].roadtype == 1)
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
                    destrange = worksheet_sndc.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('D' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        tablerow * tcnt_sn + 7,
                        tablerow * tcnt_sn + 6 + disnum));
                    destrange.Value2 = disval;
                    sn_cemile = emile;
                    if (!sn_flag)
                    {
                        sn_flag = true;
                        sn_csmile = smile;
                    }
                }
                else if (roadpart[i].roadtype == 0)
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
                    destrange = worksheet_lqdc.get_Range(string.Format("{0}{1}:{0}{2}",
                        GlobalExcel.GetCol((char)('D' + Math.Min(smile, emile) % (xlslen) * 10 / xlslen)),
                        tablerow * tcnt_lq + 7,
                        tablerow * tcnt_lq + 6 + disnum));
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
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 1] = "路线名称：" + prjinfo._RoadName;
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 2] = "调查方向:" + (prjinfo._Direction > 0 ? "上行" : "下行");
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 6] = prjinfo._DataDate;
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 11] = prjinfo._DataPerson;
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 4, 6] = sn_csmile;
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 4, 11] = sn_cemile;
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 5, 6] = Math.Abs(sn_csmile - sn_cemile);
                        worksheet_sndc.Cells[tablerow * tcnt_sn + 5, 11] = _RoadConfig.DetectWidth;
                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[1].Count - 1;
                            srcrange = worksheet_sndc.get_Range(String.Format("A{0}:O{1}", tablerow * tcnt_sn + 1, tablerow * (++tcnt_sn) + 1));
                            destrange = worksheet_sndc.get_Range(String.Format("A{0}", tablerow * tcnt_sn + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_sndc.get_Range(String.Format("D{0}:M{1}", tablerow * tcnt_sn + 7, tablerow * tcnt_sn + 6 + disnum));
                            destrange.ClearContents();
                        }
                        sn_flag = false;
                        sn_csmile = sn_cemile;
                    }
                    if (lq_csmile != lq_cemile)
                    {
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 1] = "路线名称：" + prjinfo._RoadName;
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 2] = "调查方向:" + (prjinfo._Direction > 0 ? "上行" : "下行");
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 6] = prjinfo._DataDate;
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 11] = prjinfo._DataPerson;
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 4, 6] = lq_csmile;
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 4, 11] = lq_cemile;
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 5, 6] = Math.Abs(lq_csmile - lq_cemile);
                        worksheet_lqdc.Cells[tablerow * tcnt_lq + 5, 11] = _RoadConfig.DetectWidth;

                        if (emile != roadpart[len].mile)
                        {
                            disnum = RoadDiseaseTypes.DiseaseTypeDict[0].Count - 1;
                            srcrange = worksheet_lqdc.get_Range(String.Format("A{0}:O{1}", tablerow * tcnt_lq + 1, tablerow * (++tcnt_lq) + 1));
                            destrange = worksheet_lqdc.get_Range(String.Format("A{0}", tablerow * tcnt_lq + 1));
                            srcrange.Copy(destrange);
                            destrange = worksheet_lqdc.get_Range(String.Format("D{0}:M{1}", tablerow * tcnt_lq + 7, tablerow * tcnt_lq + 6 + disnum));
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
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 1] = "路线名称：" + prjinfo._RoadName;
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 2] = "调查方向:" + (prjinfo._Direction > 0 ? "上行" : "下行");
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 6] = prjinfo._DataDate;
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 3, 11] = prjinfo._DataPerson;
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 4, 6] = sn_csmile;
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 4, 11] = roadpart[len].mile;
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 5, 6] = Math.Abs(sn_csmile - roadpart[len].mile);
                    worksheet_sndc.Cells[tablerow * tcnt_sn + 5, 11] = _RoadConfig.DetectWidth;
                }
                if (lq_csmile != lq_cemile)
                {
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 1] = "路线名称：" + prjinfo._RoadName;
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 2] = "调查方向:" + (prjinfo._Direction > 0 ? "上行" : "下行");
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 6] = prjinfo._DataDate;
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 3, 11] = prjinfo._DataPerson;
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 4, 6] = lq_csmile;
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 4, 11] = roadpart[len].mile;
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 5, 6] = Math.Abs(lq_csmile - roadpart[len].mile);
                    worksheet_lqdc.Cells[tablerow * tcnt_lq + 5, 11] = _RoadConfig.DetectWidth;
                }
            }
            if (!Hassnflag)
            {
                worksheet_sndc.Delete();
            }
            if (!Haslqflag)
            {
                worksheet_lqdc.Delete();
            }

            RoadDiseaseTypes.Clear();
        }
    }
}
