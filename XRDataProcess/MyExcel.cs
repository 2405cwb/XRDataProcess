using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms;
using RutDataView;
using MyGlobal;

namespace XRDataProcess
{
    public static class MyExcel
    {
        public static void OutputRoad(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路\综合报表模板.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            WriteDisLB2Xls(_Workbook, prjinfo, _RoadDisList);
            WriteDisHZTJ2Xls(_Workbook, prjinfo, _RoadPart, _RoadDisList, prjdir);
            WritePQI2Xls(_Workbook, prjinfo, prjdir, _RoadPart, _RoadDisList, disval);
            WriteRoadInfo(_Workbook, prjinfo, prjdir);

            _Workbook.Save();
            _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

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

        private static double _RoadRealWidth = 3.75;
        private static Dictionary<string, int> _RoadTypeDict;
        private static Dictionary<string, int> _RoadGradeDict;

        private static string _DutyUnit;
        private static string _RoadSideType;
        private static List<MilePart> _RoadPart = null;
        private static Disease[] _RoadDisList = null;

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
            _RoadGradeDict.Add("高速公路", 0);
            _RoadGradeDict.Add("一级公路", 1);
            _RoadGradeDict.Add("二级公路", 2);
            _RoadGradeDict.Add("三级公路", 3);
            _RoadGradeDict.Add("四级公路", 4);

            _RoadTypeDict = new Dictionary<string, int>();
            _RoadTypeDict.Add("沥青", 0);
            _RoadTypeDict.Add("水泥", 1);

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
                    if (rootchild.Name == Global.g_ParmStyles[MainForm._ParmStyle])
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
                    if (rootchild.Name == Global.g_ParmStyles[MainForm._ParmStyle])
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
                                            _PCIa[i][_RoadTypeDict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a0"));
                                            _PCIa[i][_RoadTypeDict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a1"));
                                        }
                                    }
                                    else if (node.Name == "PQI")
                                    {
                                        val.CopyTo(_PQIGrade[i], 0);
                                        foreach (XmlNode nnode in node.ChildNodes)
                                        {
                                            _PQIW[i][_RoadTypeDict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPCI"));
                                            _PQIW[i][_RoadTypeDict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRQI"));
                                            _PQIW[i][_RoadTypeDict[nnode.Name]][2] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRDI"));
                                            _PQIW[i][_RoadTypeDict[nnode.Name]][3] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WSRI"));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //读取计算RDI的系数
            xmlNodes = Elem.GetElementsByTagName("RDI系数"); //获取Provinces子节点集合
            foreach (XmlNode node in xmlNodes)
            {
                _RDIRD[0][0] = double.Parse(((XmlElement)node).GetAttribute("车辙常数a"));
                _RDIRD[1][0] = double.Parse(((XmlElement)node).GetAttribute("车辙常数b"));
                _RDIRD[0][1] = double.Parse(((XmlElement)node).GetAttribute("车辙RDa"));
                _RDIRD[1][1] = double.Parse(((XmlElement)node).GetAttribute("车辙RDb"));
                _RDIa[0] = double.Parse(((XmlElement)node).GetAttribute("车辙a0"));
                _RDIa[1] = double.Parse(((XmlElement)node).GetAttribute("车辙a1"));
            }
        }

        private class CityRoadDis
        {
            public string _DisType = null;
            public string _DisName = null;
            public double _UseWidth = 0.0;
            public double _Weight = 0.0;
        }

        public static void InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval, bool IsDis)
        {
            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }
            _RoadPart = new List<MilePart>();
            MilePart spart = new MilePart() { roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade] };

            _RoadPart.Add(spart);
            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, ref _RoadPart, _RoadTypeDict, _RoadGradeDict);

            if (IsDis)
            {
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo._Direction, ref _RoadDisList);
            }
            IniFiles inisetting = new IniFiles(Application.StartupPath + @"\RoadConfig.ini");
            _RoadRealWidth = double.Parse(inisetting.ReadString("ImageInfo", "RoadWidth", "3.75"));

            inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\XRSetting.ini");
            _DutyUnit = inisetting.ReadString("UI", "DutyUnit", "交通局").Replace("\0", "");
            _RoadSideType = inisetting.ReadString("UI", "RoadSideType", "双向双车道").Replace("\0", "");
        }

        private static void WriteDisLB2Xls(MSExcel.Workbook workbook, ProjectInfo prjinfo, Disease[] dislist)
        {
            MSExcel.Worksheet worksheet = workbook.Sheets["病害列表"] as MSExcel.Worksheet;

            MSExcel.Range destrange;
            int len = dislist.Length, i = 0;
            object[,] val = new object[len, 8];
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
                val[i, 4] = tdis.rect.Height * Disease.heightscale;
                val[i, 5] = tdis.rect.Width * Disease.widthscale;
                val[i, 6] = (tdis.rect.Width / 2 + tdis.rect.X) * Disease.widthscale;
                val[i, 7] = tdis.Area;
                ++i;
            }
            destrange = worksheet.get_Range(String.Format("A3:H{0}", len + 2));
            destrange.Value2 = val;

            destrange = worksheet.get_Range(String.Format("A1:H{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);
        }
        private static void WriteDisHZTJ2Xls(MSExcel.Workbook workbook, ProjectInfo prjinfo, List<MilePart> roadpart, Disease[] arrdis, DirectoryInfo prjdir)
        {
            MSExcel.Worksheet worksheet_lqtj = workbook.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_sntj = workbook.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_lqhz = workbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet worksheet_snhz = workbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;

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
            int len = roadpart.Count, dlen = arrdis.Length;
            for (int i = 0, j = 0; i < len - 1; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                //统计位于这个区域的病害
                DiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    try
                    {
                        int typeidx = DiseaseTypes.DiseaseIdx[roadpart[i].roadtype][string.Format("{0}.{1}",
                            GlobalExcel._RoadTypeStr[roadpart[i].roadtype], arrdis[j].RoadDisType)];
                        DiseaseTypes.roaddis[typeidx].totalarea += arrdis[j].Area;
                    }
                    catch
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    j++;
                }

                //病害汇总表
                int colcnt = 1;
                if (roadpart[i].roadtype == 1)//水泥
                {
                    Hassnflag = true;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = smile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = emile;
                    worksheet_snhz.Cells[rowcnt_sn_s, colcnt++] = prjinfo._RoadNum;

                    disnum = DiseaseTypes.DiseaseIdx[1].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = DiseaseTypes.DiseaseIdx[0].Count, kk = 0; di < DiseaseTypes.roaddis.Length; ++di, ++kk)
                    {
                        disval[0, kk] = DiseaseTypes.roaddis[di].totalarea;
                    }
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, DiseaseTypes.DiseaseIdx[0].Count, _RoadRealWidth * milelength);
                    disval[0, disnum] = Math.Round(drval, 2);
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
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = smile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = emile;
                    worksheet_lqhz.Cells[rowcnt_lq_s, colcnt++] = prjinfo._RoadNum;

                    disnum = DiseaseTypes.DiseaseIdx[0].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < DiseaseTypes.DiseaseIdx[0].Count; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = DiseaseTypes.roaddis[di].totalarea;
                    }
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, 0, _RoadRealWidth * milelength);
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

                if (emile % 1000 == 0)
                {
                    if (roadpart[i].roadtype == 1)
                    {
                        GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                        worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                        disnum = DiseaseTypes.DiseaseIdx[1].Count;
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
                    else if (roadpart[i].roadtype == 0)
                    {
                        GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                        worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                        disnum = DiseaseTypes.DiseaseIdx[0].Count;
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
            }

            //最后的一个小计
            if (roadpart[len - 1].mile % 1000 != 0)
            {
                if (roadpart[len - 1].roadtype == 1)
                {
                    GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "小计", worksheet_snhz, 0);
                    worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
                    disnum = DiseaseTypes.DiseaseIdx[1].Count;
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
                else if (roadpart[len - 1].roadtype == 0)
                {
                    GlobalExcel.WriteExcel(rowcnt_lq_s, 1, 1, 2, "小计", worksheet_lqhz, 0);
                    worksheet_lqhz.Cells[rowcnt_lq_s, 3] = prjinfo._RoadNum;
                    disnum = DiseaseTypes.DiseaseIdx[0].Count;
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

            //总计
            //水泥
            GlobalExcel.WriteExcel(rowcnt_sn_s, 1, 1, 2, "总计", worksheet_snhz, 0);
            worksheet_snhz.Cells[rowcnt_sn_s, 3] = prjinfo._RoadNum;
            disnum = DiseaseTypes.DiseaseIdx[1].Count;
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
            disnum = DiseaseTypes.DiseaseIdx[0].Count;
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

            DiseaseTypes.Clear();
            if (Haslqflag)
            {
                disnum = DiseaseTypes.DiseaseIdx[0].Count;
                worksheet_lqtj.Cells[2, 2] = _RoadRealWidth;
                worksheet_lqtj.Cells[2, 6] = Math.Abs(roadpart[0].mile - roadpart[len - 1].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=沥青病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_lq_s);
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
                disnum = DiseaseTypes.DiseaseIdx[1].Count;
                worksheet_sntj.Cells[2, 2] = _RoadRealWidth;
                worksheet_sntj.Cells[2, 6] = Math.Abs(roadpart[0].mile - roadpart[len - 1].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=水泥病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_sn_s);
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

        private static double ComputPCI(DiseaseType[] disarea, int roadtype, int startidx, double partarea)
        {
            double sumarea = 0;
            int len = _RoadSocre[roadtype].Keys.Count;

            for (int i = startidx; i < len + startidx; i++)
            {
                sumarea += disarea[i].totalarea * disarea[i].weight;
            }
            return 100 * sumarea / partarea;
        }
        private static void WritePQI2Xls(MSExcel.Workbook workbook,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            bool IsHasRut = false;

            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\IRI_{1}m.txt", prjdir.FullName, 10);
            string LMTDfrname = string.Format(@"{0}\IRIMTD\Laser0\MTD_{1}m.txt", prjdir.FullName, 10);

            string[] LIRIsr = null;
            string[] RIRIsr = null;
            string[] LRutsr = null;
            string[] RRutsr = null;
            string[] LMTDsr = null;
            string[] RMTDsr = null;

            int LIRIsridx = 0;
            int RIRIsridx = 0;
            int LRutsridx = 0;
            int RRutsridx = 0;
            int LMTDsridx = 0;
            int RMTDsridx = 0;

            if (File.Exists(LIRIfrname))
            {
                LIRIsr = File.ReadAllLines(LIRIfrname);
            }
            else
            {
                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度数据!\r\n请检查数据完整性！");
                return;
            }
            if (File.Exists(LMTDfrname))
            {
                LMTDsr = File.ReadAllLines(LMTDfrname);
            }
            else
            {
                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧构造深度数据!\r\n请检查数据完整性！");
                return;
            }

            bool IsDIRIMTD = false;
            string RIRIfrname = null;
            string RMTDfrname = null;
            IniFiles iniset = new IniFiles(prjdir.FullName + @"\Setting.ini");
            if (IsDIRIMTD = iniset.ReadBool("工作模式", "DIRIMTD", false))
            {
                RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\IRI_{1}m.txt", prjdir.FullName, 10);
                if (File.Exists(RIRIfrname))
                {
                    RIRIsr = File.ReadAllLines(RIRIfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度数据!\r\n请检查数据完整性！");
                    return;
                }

                RMTDfrname = string.Format(@"{0}\IRIMTD\Laser1\MTD_{1}m.txt", prjdir.FullName, 10);
                if (File.Exists(RIRIfrname))
                {
                    RMTDsr = File.ReadAllLines(RMTDfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧构造深度数据!\r\n请检查数据完整性！");
                    return;
                }
            }

            int rutdis = 10, rutvalnum = 1;
            //rutdis = Convert.ToInt32(iniset.ReadInteger("Parm", "RUT_Dis", 0));
            rutvalnum = Convert.ToInt32(iniset.ReadInteger("工作模式", "ValNum", 0));
            bool IsRut = false;
            string LRutfrname = null, RRutfrname = null;
            if (IsRut = iniset.ReadBool("工作模式", "Rut", false))
            {
                LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
                if (File.Exists(LRutfrname))
                {
                    LRutsr = File.ReadAllLines(LRutfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
                    return;
                }
                if (rutvalnum < 2)
                {
                    RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
                    if (File.Exists(RRutfrname))
                    {
                        RRutsr = File.ReadAllLines(RRutfrname);
                    }
                    else
                    {
                        MessageBox.Show(prjdir.FullName + "\r\n缺少右侧车辙深度数据!\r\n请检查数据完整性！");
                        return;
                    }
                }
            }

            string LRutstrline, RRutstrline;
            double Lrutoldval = 0, Rrutoldval = 0;
            double Lrutcurval = 0, Rrutcurval = 0;
            string LIRIstrline, RIRIstrline;
            double Lirioldval = 0, Ririoldval = 0;
            double Liricurval = 0, Riricurval = 0;
            string LMTDstrline, RMTDstrline;
            double Lmtdoldval = 0, Rmtdoldval = 0;
            double Lmtdcurval = 0, Rmtdcurval = 0;
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

            string errlog = prjdir.FullName + "\\errlog.txt";
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

                mxlist[i, 0] = prjinfo._RoadCode;
                mxlist[i, 1] = prjinfo._District + "交通运输局";
                mxlist[i, 2] = _RoadSideType;
                mxlist[i, 3] = prjinfo._Direction > 0 ? "上行" : "下行";
                mxlist[i, 4] = smile;
                mxlist[i, 5] = emile;
                mxlist[i, 6] = milelength;
                mxlist[i, 11] = GlobalExcel._RoadTypeStr[roadpart[i].roadtype];

                //统计位于这个区域的病害
                DiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    try
                    {
                        int typeidx = DiseaseTypes.DiseaseIdx[roadpart[i].roadtype][string.Format("{0}.{1}",
                            GlobalExcel._RoadTypeStr[roadpart[i].roadtype], arrdis[j].RoadDisType)];
                        DiseaseTypes.roaddis[typeidx].totalarea += arrdis[j].Area;
                    }
                    catch
                    {
                        string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    j++;
                }
                //PCI
                if (roadpart[i].roadtype == 1)
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, DiseaseTypes.DiseaseIdx[0].Count, _RoadRealWidth * milelength);
                }
                else
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, 0, _RoadRealWidth * milelength);
                }
                tpcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 12] = Math.Round(drval, 2);
                mxlist[i, 8] = string.Format("=100-{0}*POWER(M{1},{2})", _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0], i + 3, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                mxlist[i, 17] = string.Format("=IF(I{0}>={1},\"优\",IF(I{0}>={2},\"良\",IF(I{0}>={3},\"中\",IF(I{0}>={4},\"次\",\"差\"))))",
                    i + 3, _PCIGrade[roadpart[i].roaddegree][0], _PCIGrade[roadpart[i].roaddegree][1], _PCIGrade[roadpart[i].roaddegree][2], _PCIGrade[roadpart[i].roaddegree][3]);

                //IRI
                int iricnt = 0;
                double sumiril = 0, sumirir = 0;
                string[] tiri;
                int irixlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) / 10;
                int lirivalnum = 0, ririvalnum = 0;
                while (iricnt++ < irixlsnum)
                {
                    if (LIRIsridx < LIRIsr.Length)
                    {
                        LIRIstrline = LIRIsr[LIRIsridx++];
                        tiri = LIRIstrline.Split(' ');
                        try
                        {
                            Liricurval = double.Parse(tiri[1]);
                        }
                        catch
                        {
                            Liricurval = Lirioldval;
                        }
                        Lirioldval = MainForm._ErrorVal == 1 && Liricurval > MainForm._ErrorIRI ?
                            MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Liricurval;
                        sumiril += Lirioldval;
                        ++lirivalnum;
                    }
                    if (IsDIRIMTD)
                    {
                        if (RIRIsridx < RIRIsr.Length)
                        {
                            RIRIstrline = RIRIsr[RIRIsridx++];
                            tiri = RIRIstrline.Split(' ');
                            try
                            {
                                Riricurval = double.Parse(tiri[1]);
                            }
                            catch
                            {
                                Riricurval = Ririoldval;
                            }
                            Ririoldval = MainForm._ErrorVal == 1 && Riricurval > MainForm._ErrorIRI ?
                                MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Riricurval;
                            sumirir += Ririoldval;
                            ++ririvalnum;
                        }
                    }
                }
                if (lirivalnum > 0)
                {
                    if (IsDIRIMTD)
                    {
                        irival = Math.Round((sumiril + sumirir) * 0.5 / lirivalnum, 2);
                    }
                    else
                    {
                        irival = Math.Round(sumiril / lirivalnum, 2);
                    }
                    trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                    mxlist[i, 13] = irival;
                }
                else
                {
                    mxlist[i, 13] = i > 0 ? mxlist[i - 1, 13] : 0;
                }
                mxlist[i, 9] = String.Format("=ROUND(100/(1+{0}*EXP({1}*N{2})),2)", _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 3);
                mxlist[i, 18] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    i + 3, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2], _RQIGrade[roadpart[i].roaddegree][3]);

                //MTD
                int mtdcnt = 0;
                double summtdl = 0, summtdr = 0;
                int lmtdvalnum = 0, rmtdvalnum = 0;
                string[] tmtd;
                int mtdxlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) / 10;
                while (mtdcnt++ < mtdxlsnum)
                {
                    if (LMTDsridx < LMTDsr.Length)
                    {
                        LMTDstrline = LMTDsr[LMTDsridx++];
                        tmtd = LMTDstrline.Split(' ');
                        try
                        {
                            Lmtdcurval = double.Parse(tmtd[1]);
                        }
                        catch
                        {
                            Lmtdcurval = Lmtdoldval;
                        }
                        Lmtdoldval = MainForm._ErrorVal == 1 && Lmtdcurval > MainForm._ErrorIRI ?
                            MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Lmtdcurval;
                        summtdl += Lmtdoldval;
                        ++lmtdvalnum;
                    }
                    if (IsDIRIMTD)
                    {
                        if (RMTDsridx < RMTDsr.Length)
                        {
                            RMTDstrline = RMTDsr[RMTDsridx++];
                            tmtd = RMTDstrline.Split(' ');
                            try
                            {
                                Rmtdcurval = double.Parse(tmtd[1]);
                            }
                            catch
                            {
                                Rmtdcurval = Rmtdoldval;
                            }
                            Rmtdoldval = MainForm._ErrorVal == 1 && Rmtdcurval > MainForm._ErrorIRI ?
                                MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Rmtdcurval;
                            summtdr += Rmtdoldval;
                            ++rmtdvalnum;
                        }
                    }
                }
                if (lmtdvalnum > 0)
                {
                    if (IsDIRIMTD)
                    {
                        mtdval = Math.Round((summtdl + summtdr) * 0.5 / lmtdvalnum, 2);
                    }
                    else
                    {
                        mtdval = Math.Round(summtdl / lmtdvalnum, 2);
                    }
                    mxlist[i, 15] = mtdval;
                }
                else
                {
                    mxlist[i, 15] = i > 0 ? mxlist[i - 1, 15] : 0;
                }

                //Rut
                if (IsRut)
                {
                    IsHasRut = true;
                    int rutxlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) * 100 / rutdis;
                    int rutcnt = 0;
                    double sumrutl = 0, sumrutr = 0;
                    int lrutcnt = 0, rrutcnt = 0;
                    List<double> tlruts = new List<double>();
                    List<double> trruts = new List<double>();
                    string[] trut;
                    while (rutcnt++ < rutxlsnum)
                    {
                        if (rutvalnum < 2)
                        {
                            if (LRutsridx < LRutsr.Length)
                            {
                                LRutstrline = LRutsr[LRutsridx++];
                                trut = LRutstrline.Split(',');
                                try
                                {
                                    Lrutcurval = double.Parse(trut[1]);
                                }
                                catch
                                {
                                    Lrutcurval = Lrutoldval;
                                }
                                if (!(MainForm._ErrorVal == 1 && Lrutcurval > MainForm._ErrorRut))
                                {
                                    Lrutoldval = Lrutcurval;
                                    tlruts.Add(Lrutoldval);
                                    sumrutl += Lrutoldval;
                                    ++lrutcnt;
                                }
                            }
                            if (RRutsridx < RRutsr.Length)
                            {
                                RRutstrline = RRutsr[RRutsridx++];
                                trut = RRutstrline.Split(',');
                                try
                                {
                                    Rrutcurval = double.Parse(trut[1]);
                                }
                                catch
                                {
                                    Rrutcurval = Rrutoldval;
                                }
                                if (!(MainForm._ErrorVal == 1 && Rrutcurval > MainForm._ErrorRut))
                                {
                                    Rrutoldval = Rrutcurval;
                                    trruts.Add(Rrutoldval);
                                    sumrutr += Rrutoldval;
                                    ++rrutcnt;
                                }
                            }
                        }
                        else
                        {
                            if (LRutsridx < LRutsr.Length)
                            {
                                LRutstrline = LRutsr[LRutsridx++];
                                trut = LRutstrline.Split(',');
                                try
                                {
                                    Lrutcurval = double.Parse(trut[1]);
                                    Rrutcurval = double.Parse(trut[3]);
                                }
                                catch
                                {
                                    Lrutcurval = Lrutoldval;
                                    Rrutcurval = Rrutoldval;
                                }
                                if (!(MainForm._ErrorVal == 1 && Lrutcurval > MainForm._ErrorRut))
                                {
                                    Lrutoldval = Lrutcurval;
                                    tlruts.Add(Lrutoldval);
                                    sumrutl += Lrutoldval;
                                    ++lrutcnt;
                                }
                                if (!(MainForm._ErrorVal == 1 && Rrutcurval > MainForm._ErrorRut))
                                {
                                    Rrutoldval = Rrutcurval;
                                    trruts.Add(Rrutoldval);
                                    sumrutr += Rrutoldval;
                                    ++rrutcnt;
                                }
                            }
                        }
                    }
                    if (lrutcnt > 0 && rrutcnt > 0)
                    {
                        if (MainForm._IsThrRut)
                        {
                            double mlrut = sumrutl / lrutcnt;
                            double minlrut = mlrut * MainForm._MinThrRut;
                            double maxlrut = mlrut * MainForm._MaxThrRut;
                            foreach (double tval in tlruts)
                            {
                                if (tval < minlrut || tval > maxlrut)
                                {
                                    sumrutl -= tval;
                                    --lrutcnt;
                                }
                            }
                            double mrrut = sumrutr / rrutcnt;
                            double minrrut = mrrut * MainForm._MinThrRut;
                            double maxrrut = mrrut * MainForm._MaxThrRut;
                            foreach (double tval in trruts)
                            {
                                if (tval < minrrut || tval > maxrrut)
                                {
                                    sumrutr -= tval;
                                    --rrutcnt;
                                }
                            }
                        }

                        sumrutl = sumrutl / lrutcnt;
                        sumrutr = sumrutr / rrutcnt;
                    }
                    double rutval = Math.Max(sumrutl, sumrutr);
                    rutval = Math.Round(rutval, 2);
                    mxlist[i, 14] = rutval;

                    if (roadpart[i].roaddegree < 2 && roadpart[i].roadtype == 0)
                    {
                        mxlist[i, 10] = string.Format("=IF(O{0}<{1},{2}-{3}*O{0},IF(O{0}<{4},{5}-{6}*(O{0}-{1}),0))",
                            i + 4,
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
                    else
                    {
                        mxlist[i, 14] = "-";
                        mxlist[i, 10] = "-";
                        mxlist[i, 19] = "-";
                    }
                }
                else
                {
                    mxlist[i, 14] = "-";
                    mxlist[i, 10] = "-";
                    mxlist[i, 19] = "-";
                }
                mxlist[i, 20] = string.Format("=ROUND(({1}*I{0}+{2}*J{0}+{3}*IF(EXACT(K{0},\"-\"),0,K{0}))/({1}+{2}+{3}),2)",
                    i + 3, 
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0], 
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1], 
                    _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);
                mxlist[i, 16] = string.Format("=IF(U{0}>={1},\"优\",IF(U{0}>={2},\"良\",IF(U{0}>={3},\"中\",IF(U{0}>={4},\"次\",\"差\"))))",
                    i + 3, 
                    _PQIGrade[roadpart[i].roaddegree][0], 
                    _PQIGrade[roadpart[i].roaddegree][1], 
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                mxlist[i, 7] = String.Format("=CONCATENATE(TEXT(U{0},\"0.00\"),\"(\",Q{0},\")\")", i + 3);
                if (MainForm._YHType == 0)
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
                else if (MainForm._YHType == 1)
                {
                    if (!(trqival > 85 && tpcival > 70))
                    {
                        yhlist[yhi, 0] = string.Format("{0}_{1}", prjinfo._RoadCode, (yhi + 1).ToString(lenstr));
                        yhlist[yhi, 1] = prjinfo._District + "交通运输局";
                        yhlist[yhi, 2] = _RoadGradeStr[roadpart[i].roaddegree];
                        yhlist[yhi, 3] = mxlist[i, 4];
                        yhlist[yhi, 4] = mxlist[i, 5];
                        yhlist[yhi, 5] = mxlist[i, 6];
                        yhlist[yhi, 6] = String.Format("=IF(技术状况明细表!I{0}>=70,IF(技术状况明细表!I{0}>=85,\"日常养护\",IF(技术状况明细表!I{0}>=75,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\"))),IF(技术状况明细表!I{0}>=60,IF(技术状况明细表!I{0}>=85,\"预防性养护\",IF(技术状况明细表!I{0}>=65,\"中修\",\"大修\")),IF(技术状况明细表!I{0}>=40,IF(技术状况明细表!I{0}>=75,\"中修\",\"大修\"),\"大修\")))", i+3);
                        yhi++;
                    }
                }
            }

            MSExcel.Worksheet worksheet = workbook.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:U{0}", len + 2));
            destrange.Value2 = mxlist;
            GlobalExcel.SetBorderLine(destrange, 53);

            if (IsHasRut) destrange = worksheet.get_Range(string.Format("F2:F{0}, I2:K{0},U2:U{0}", len + 2));
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

            if (IsHasRut)
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
            tjlist[0, 0] = String.Format("=AVERAGE(技术状况明细表!{0}3:{0}{1})", 'U', len + 2);
            for (int i = 1; i < 4; ++i)
            {
                tjlist[i, 0] = String.Format("=AVERAGE(技术状况明细表!{0}3:{0}{1})", GlobalExcel.GetCol(((char)('H' + i))), len + 2);
            }
            destrange = worksheet.get_Range("B3:B6");
            destrange.Value2 = tjlist;
            if (IsHasRut) destrange = worksheet.get_Range("L2:Q6");
            else destrange = worksheet.get_Range("L2:Q5");
            chartobj = (MSExcel.ChartObject)worksheet.ChartObjects(1);
            chart = chartobj.Chart;
            chart.SetSourceData(destrange);
        }

        private static void WriteRoadInfo(MSExcel.Workbook workbook, ProjectInfo prjinfo, DirectoryInfo prjdir)
        {
            IniFiles inisetting = new IniFiles(Application.StartupPath + @"\RoadConfig.ini");
            string degreeinfo = File.ReadAllText(prjdir.FullName + "\\DegreeInfo.txt").Replace(" ", Environment.NewLine);
            string roadtypeinfo = File.ReadAllText(prjdir.FullName + "\\RoadTypeInfo.txt").Replace(" ", Environment.NewLine);;
            MSExcel.Worksheet worksheet = workbook.Sheets["路线信息表"] as MSExcel.Worksheet;
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
            valobj[5, 0] = prjinfo._District+"交通运输局";
            destrange.Value2 = valobj;

            destrange = worksheet.get_Range("B15:B20");
            valobj[0, 0] = inisetting.ReadString("ImageInfo", "DetectWidth", "3.75");
            valobj[1, 0] = prjinfo._Direction > 0 ? "上行" : "下行";
            valobj[2, 0] = "=IF(分项指标统计表!B3>=90,\"优\",IF(分项指标统计表!B3>=80,\"良\",IF(分项指标统计表!B3>=70,\"中\",IF(分项指标统计表!B3>=60,\"次\",\"差\"))))";
            valobj[3, 0] = degreeinfo;
            valobj[4, 0] = roadtypeinfo;
            valobj[5, 0] = Math.Max(prjinfo._StartMile*0.001, prjinfo._EndMile*0.001);
            destrange.Value2 = valobj;
        }
    }
}
