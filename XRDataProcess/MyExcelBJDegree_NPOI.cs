using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms;
using RutDataView;
using MyGlobal;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.POIFS.FileSystem;

namespace XRDataProcess
{
    class MyExcelBJDegree_NPOI
    {
        public static double[][] _RQIGrade;//道路等级 等级区间
        public static double[][] _RDIGrade;
        public static double[][] _MTDGrade;
        public static double[][] _PCIGrade;
        public static double[][] _PQIGrade;
        public static double[][] _RDIRD;
        public static double[] _RDIa;

        public static double[][] _RQIa;//公路等级 参数序号
        public static double[][][] _PCIa;//公路等级 路面材质 参数序号
        public static double[][][] _PQIW;//公路等级 路面材质 参数序号
        public static double[][] _WeightParm;//0-沥青，1-水泥
        public static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        public static double _RoadRealWidth = 3.75;
        public static Dictionary<string, int> _RoadTypeDict;
        public static Dictionary<string, int> _RoadGradeDict;

        public static List<MilePart> _RoadPart = null;
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

        public class CityRoadDis
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
        }

        public static void OutputIRI(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路\路面平整度评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_IRI_{2}m.xlsx", path, prjdir.Name, disval);
            XSSFWorkbook _Workbook;
            using (FileStream file = new FileStream(srcxls, FileMode.Open, FileAccess.Read))  //路径，打开权限，读取权限
            {
                _Workbook = new XSSFWorkbook(file);
                file.Close();
            }
            XSSFSheet _Worksheet  = _Workbook.GetSheet("Sheet1") as XSSFSheet;
            WriteIRI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, disval);
            using (FileStream files = new FileStream(Destxls, FileMode.Create))
            {
                _Workbook.Write(files);
                files.Close();
            }
        }
        private static void WriteIRI2Xls(XSSFSheet _Worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, int xlslen)
        {
            MSExcel.Range destrange;
            int len = roadpart.Count - 1;

            string[] LIRIsr = null;
            string[] RIRIsr = null;
            int LIRIsridx = 0, RIRIsridx = 0;

            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\IRI_{1}m.txt", prjdir.FullName, 10);
            string RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\IRI_{1}m.txt", prjdir.FullName, 10);
            if (File.Exists(LIRIfrname))
            {
                LIRIsr = File.ReadAllLines(LIRIfrname);
            }
            else
            {
                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度构造深度数据!\r\n请检查数据完整性！");
                return;
            }

            bool IsDIRIMTD = false;
            IniFiles iniset = new IniFiles(prjdir.FullName + @"\Setting.ini");
            if (IsDIRIMTD = iniset.ReadBool("工作模式", "DIRIMTD", false))
            {
                if (File.Exists(RIRIfrname))
                {
                    RIRIsr = File.ReadAllLines(RIRIfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度构造深度数据!\r\n请检查数据完整性！");
                    return;
                }
            }

            object[,] vallist = new object[len, 8];
            string LIRIstrline, RIRIstrline;
            double Loldval = 0, Roldval = 0;
            double Lcurval = 0, Rcurval = 0;
            for (int i = 0; i < len; i++)
            {
                vallist[i, 0] = roadpart[i].mile;
                vallist[i, 1] = roadpart[i + 1].mile;
                vallist[i, 2] = prjinfo._RoadNum;

                int irixlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) / 10;
                int iricnt = 0;
                double suml = 0, sumr = 0;
                string[] tmtd;
                int lvalnum = 0, rvalnum = 0;
                while (iricnt++ < irixlsnum)
                {
                    if (LIRIsridx < LIRIsr.Length)
                    {
                        LIRIstrline = LIRIsr[LIRIsridx++];
                        tmtd = LIRIstrline.Split(' ');
                        try
                        {
                            Lcurval = double.Parse(tmtd[1]);
                        }
                        catch
                        {
                            Lcurval = Loldval;
                        }

                        Loldval = MainForm._ErrorVal == 1 && Lcurval > MainForm._ErrorIRI ?
                            MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Lcurval;
                        suml += Loldval;
                        ++lvalnum;
                    }
                    if (IsDIRIMTD)
                    {
                        if (RIRIsridx < RIRIsr.Length)
                        {
                            RIRIstrline = RIRIsr[RIRIsridx++];
                            tmtd = RIRIstrline.Split(' ');
                            try
                            {
                                Rcurval = double.Parse(tmtd[1]);
                            }
                            catch
                            {
                                Rcurval = Roldval;
                            }

                            Roldval = MainForm._ErrorVal == 1 && Rcurval > MainForm._ErrorIRI ?
                                MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Rcurval;
                            sumr += Roldval;
                            ++rvalnum;
                        }
                    }
                }
                if (lvalnum > 0 && rvalnum > 0)
                {
                    vallist[i, 3] = suml / lvalnum;
                    if (IsDIRIMTD)
                    {
                        vallist[i, 4] = sumr / rvalnum;
                    }
                }
                else
                {
                    if (i > 0)
                    {
                        vallist[i, 3] = vallist[i - 1, 3];
                        if (IsDIRIMTD)
                        {
                            vallist[i, 4] = vallist[i - 1, 4];
                        }
                    }
                    else
                    {
                        vallist[i, 3] = 0;
                        if (IsDIRIMTD)
                        {
                            vallist[i, 4] = 0;
                        }
                    }
                }
                if (IsDIRIMTD)
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,2)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),2)", i + 4);
                }

                vallist[i, 6] = String.Format("=ROUND(100/(1+{0}*EXP({1}*F{2})),2)",
                    _RQIa[roadpart[i].roaddegree][0], _RQIa[roadpart[i].roaddegree][1], i + 4);
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

            if (MainForm._IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 4, 8, true);
                GlobalExcel.Reflection(_Worksheet, 4, 2, false);
            }
        }
        private static void WriteIRIStatistics(XSSFSheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            int scol = 15, srow = 2;
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                _Worksheet.GetRow(srow).GetCell(scol+i).SetCellValue(string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]));
                _Worksheet.GetRow(srow+2).GetCell(scol+i).SetCellValue("=ABS(SUM(A:A)-SUM(B:B))");
                _Worksheet.GetRow(srow+1).GetCell(scol+i).SetCellValue(string.Format("={0}3/{0}5", (char)('A' + 15 + i)));
            }
            _Worksheet.GetRow(1).GetCell(8).SetCellValue("=CONCATENATE(\"路面平整度评价等级“优”率占路段总数\",ROUND(P4,4)*100,\"%，“良”率占路段总数\",ROUND(Q4,4)*100,\"%，“中”率占路段总数\",ROUND(R4,4)*100,\"%，“次”率占路段总数\",ROUND(S4,4)*100,\"%，“差”率占路段总数\",ROUND(T4,4)*100,\"%。\")");
        }

        public static void OutputDis(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            XSSFWorkbook _Workbook;
            using (FileStream file = new FileStream(srcxls, FileMode.Open, FileAccess.Read))  //路径，打开权限，读取权限
            {
                _Workbook = new XSSFWorkbook(file);
                file.Close();
            }
            XSSFSheet _Worksheet_lb = _Workbook.GetSheet("病害列表") as XSSFSheet;
            WriteDisLB2Xls(_Worksheet_lb, prjinfo, _RoadDisList);

            XSSFSheet _Worksheet_lqtj = _Workbook.GetSheet("沥青路面病害统计表") as XSSFSheet;
            XSSFSheet _Worksheet_sntj = _Workbook.GetSheet("水泥路面病害统计表") as XSSFSheet;
            XSSFSheet _Worksheet_lqhz = _Workbook.GetSheet("沥青路面病害区间汇总表") as XSSFSheet;
            XSSFSheet _Worksheet_snhz = _Workbook.GetSheet("水泥路面病害区间汇总表") as XSSFSheet;
            WriteDisHZTJ2Xls(_Worksheet_sntj, _Worksheet_lqtj, _Worksheet_snhz, _Worksheet_lqhz,
                prjinfo, prjdir, _RoadPart, _RoadDisList, disval);

            using (FileStream files = new FileStream(Destxls, FileMode.Create))
            {
                _Workbook.Write(files);
                files.Close();
            }
        }
        private static void WriteDisLB2Xls(XSSFSheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
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
            destrange = _Worksheet.get_Range(String.Format("A3:H{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:H{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (MainForm._IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 8, true);
            }
        }
        private static void WriteDisHZTJ2Xls(XSSFSheet worksheet_sntj, XSSFSheet worksheet_lqtj,
            XSSFSheet worksheet_snhz, XSSFSheet worksheet_lqhz,
            ProjectInfo prjinfo, DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
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
                worksheet_lqtj.Cells[2, 3] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len - 1].mile);
                worksheet_lqtj.Cells[2, 5] = _RoadRealWidth;
                worksheet_lqtj.Cells[4, 6] = string.Format("={0}*E2", Math.Abs(roadpart[0].mile - roadpart[len - 1].mile));
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=SUMIF(沥青路面病害区间汇总表!{0}:{0},\"<>\",沥青路面病害区间汇总表!{0}:{0})/3", Convert.ToChar('D' + i), rowcnt_lq_s);
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
                worksheet_sntj.Cells[2, 3] = string.Format("{0:K0+000} - {1:K0+000}", roadpart[0].mile, roadpart[len - 1].mile);
                worksheet_sntj.Cells[2, 5] = _RoadRealWidth;
                worksheet_sntj.Cells[4, 6] = string.Format("={0}*E2", Math.Abs(roadpart[0].mile - roadpart[len - 1].mile));
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    disval[i, 0] = string.Format("=SUMIF(水泥路面病害区间汇总表!{0}5:{0}{1},\"<>\",水泥路面病害区间汇总表!{0}5:{0}{1})/3", Convert.ToChar('D' + i), rowcnt_sn_s);
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
        public static double ComputPCI(DiseaseType[] disarea, int roadtype, int startidx, double partarea)
        {
            double sumarea = 0;
            int len = _RoadSocre[roadtype].Keys.Count;

            for (int i = startidx; i < len + startidx; i++)
            {
                sumarea += disarea[i].totalarea * disarea[i].weight;
            }
            return 100 * sumarea / partarea;
        }

        public static void OutputPCI(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路\路面破损评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PCI_{2}m.xlsx", path, prjdir.Name, disval);
            XSSFWorkbook _Workbook;
            using (FileStream file = new FileStream(srcxls, FileMode.Open, FileAccess.Read))  //路径，打开权限，读取权限
            {
                _Workbook = new XSSFWorkbook(file);
                file.Close();
            }
            XSSFSheet _Worksheet = _Workbook.GetSheet("Sheet1") as XSSFSheet;

            WritePCI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, disval);

            using (FileStream files = new FileStream(Destxls, FileMode.Create))
            {
                _Workbook.Write(files);
                files.Close();
            }
        }
        private static void WritePCI2Xls(XSSFSheet worksheet, ProjectInfo prjinfo, DirectoryInfo prjdir,
            List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            string errlog = prjdir.FullName + "\\errlog.txt";
            MSExcel.Range destrange;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            object[,] vallist = new object[len, 6];
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double drval = 0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                int milelength = Math.Abs(smile - emile);

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

                vallist[i, 0] = smile;
                vallist[i, 1] = emile;
                vallist[i, 2] = prjinfo._RoadNum;
                if (roadpart[i].roadtype == 1)
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, DiseaseTypes.DiseaseIdx[0].Count, _RoadRealWidth * milelength);
                }
                else
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, 0, _RoadRealWidth * milelength);
                }
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

            if (MainForm._IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 6, true);
                GlobalExcel.Reflection(worksheet, 3, 2, false);
            }
        }
        private static void WritePCIStatistics(XSSFSheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            int scol = 13, srow = 2;
            for (int i = 0; i < 5; ++i)
            {
                _Worksheet.GetRow(srow).GetCell(scol+i).SetCellValue(string.Format("=ABS(SUMIF(F:F,\"{0}\",A:A)-SUMIF(F:F,\"{0}\",B:B))", degstr[i]));
                _Worksheet.GetRow(srow+2).GetCell(scol+i).SetCellValue("=ABS(SUM(A:A)-SUM(B:B))");
                _Worksheet.GetRow(srow+1).GetCell(scol+i).SetCellValue(string.Format("={0}3/{0}5", (char)('A' + 13 + i)));
            }
            _Worksheet.GetRow(1).GetCell(6).SetCellValue("=CONCATENATE(\"路面PCI评价等级“优”率占路段总数\",ROUND(N4,4)*100,\"%，“良”率占路段总数\",ROUND(O4,4)*100,\"%，“中”率占路段总数\",ROUND(P4,4)*100,\"%，“次”率占路段总数\",ROUND(Q4,4)*100,\"%，“差”率占路段总数\",ROUND(R4,4)*100,\"%。\")");
        }

        public static void OutputPQI(string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路\路面综合评价等级记录表.xlsx",
                System.Windows.Forms.Application.StartupPath, disval);
            string Destxls = string.Format(@"{0}\{1}_PQI_{2}m.xlsx", path, prjdir.Name, disval);
            XSSFWorkbook _Workbook;
            using (FileStream file = new FileStream(srcxls, FileMode.Open, FileAccess.Read))  //路径，打开权限，读取权限
            {
                _Workbook = new XSSFWorkbook(file);
                file.Close();
            }
            XSSFSheet _Worksheet = _Workbook.GetSheet("Sheet1") as XSSFSheet;

            WritePQI2Xls(_Worksheet, prjinfo, prjdir, _RoadPart, _RoadDisList, disval);

            using (FileStream files = new FileStream(Destxls, FileMode.Create))
            {
                _Workbook.Write(files);
                files.Close();
            }
        }
        private static void WritePQI2Xls(XSSFSheet worksheet, ProjectInfo prjinfo,
            DirectoryInfo prjdir, List<MilePart> roadpart, Disease[] arrdis, int xlslen)
        {
            string[] LIRIsr = null;
            string[] RIRIsr = null;
            int LIRIsridx = 0, RIRIsridx = 0;
            string LIRIfrname = string.Format(@"{0}\IRIMTD\DAQ0\IRI_{1}m.txt", prjdir.FullName, 10);
            string RIRIfrname = string.Format(@"{0}\IRIMTD\DAQ1\IRI_{1}m.txt", prjdir.FullName, 10);
            if (File.Exists(LIRIfrname))
            {
                LIRIsr = File.ReadAllLines(LIRIfrname);
            }
            else
            {
                MessageBox.Show(prjdir.FullName + "\r\n缺少左侧平整度构造深度数据!\r\n请检查数据完整性！");
                return;
            }

            bool IsDIRIMTD = false;
            IniFiles iniset = new IniFiles(prjdir.FullName + @"\Setting.ini");
            if (IsDIRIMTD = iniset.ReadBool("工作模式", "DIRIMTD", false))
            {
                if (File.Exists(RIRIfrname))
                {
                    RIRIsr = File.ReadAllLines(RIRIfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少右侧平整度构造深度数据!\r\n请检查数据完整性！");
                    return;
                }
            }

            //int rutdis  = Convert.ToInt32(iniset.ReadInteger("Parm", "RUT_Dis", 0));
            int rutdis = 10;

            bool IsRut = false;
            String[] LRutsr = null;
            String[] RRutsr = null;
            int LRutsridx = 0, RRutsridx = 0;
            string LRutfrname = string.Format(@"{0}\Rut\camera0\orirut.txt", prjdir.FullName);
            string RRutfrname = string.Format(@"{0}\Rut\camera1\orirut.txt", prjdir.FullName);
            if (IsRut = iniset.ReadBool("工作模式", "Rut", false))
            {
                if (File.Exists(LRutfrname))
                {
                    LRutsr = File.ReadAllLines(LRutfrname);
                }
                else
                {
                    MessageBox.Show(prjdir.FullName + "\r\n缺少左侧车辙深度数据!\r\n请检查数据完整性！");
                    return;
                }

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

            string LRutstrline, RRutstrline;
            double Lrutoldval = 0, Rrutoldval = 0;
            double Lrutcurval = 0, Rrutcurval = 0;
            string LIRIstrline, RIRIstrline;
            double Loldval = 0, Roldval = 0;
            double Lcurval = 0, Rcurval = 0;
            int len = roadpart.Count - 1, dlen = arrdis.Length;
            double trqival = 0, irival = 0;

            string errlog = prjdir.FullName + "\\errlog.txt";
            int rowcnt = 0;
            object[,] vallist = new object[len, 11];
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0, drval;
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

                //计算IRI
                int iricnt = 0;
                double suml = 0, sumr = 0;
                int lirivalnum = 0, ririvalnum = 0;
                string[] tmtd;
                int irixlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) / 10;
                while (iricnt++ < irixlsnum)
                {
                    if (LIRIsridx < LIRIsr.Length)
                    {
                        LIRIstrline = LIRIsr[LIRIsridx++];
                        tmtd = LIRIstrline.Split(' ');
                        try
                        {
                            Lcurval = double.Parse(tmtd[1]);
                        }
                        catch
                        {
                            Lcurval = Loldval;
                        }
                        Loldval = MainForm._ErrorVal == 1 && Lcurval > MainForm._ErrorIRI ?
                            MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Lcurval;
                        suml += Loldval;
                        ++lirivalnum;
                    }
                    if (IsDIRIMTD)
                    {
                        if (RIRIsridx < RIRIsr.Length)
                        {
                            RIRIstrline = RIRIsr[RIRIsridx++];
                            tmtd = RIRIstrline.Split(' ');
                            try
                            {
                                Rcurval = double.Parse(tmtd[1]);
                            }
                            catch
                            {
                                Rcurval = Roldval;
                            }
                            Roldval = MainForm._ErrorVal == 1 && Rcurval > MainForm._ErrorIRI ?
                                MainForm._ErrorIRI - MainForm.rdval.Next(100) * 0.001 : Rcurval;
                            sumr += Roldval;
                            ++ririvalnum;
                        }
                    }
                }

                //病害汇总表
                int colcnt = 0;
                vallist[rowcnt, colcnt++] = smile;
                vallist[rowcnt, colcnt++] = emile;
                vallist[rowcnt, colcnt++] = prjinfo._RoadNum;
                if (roadpart[i].roadtype == 1)
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, DiseaseTypes.DiseaseIdx[0].Count, _RoadRealWidth * milelength);
                }
                else
                {
                    drval = ComputPCI(DiseaseTypes.roaddis, roadpart[i].roadtype, 0, _RoadRealWidth * milelength);
                }
                pcival = 100 - _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][0] * Math.Pow(drval, _PCIa[roadpart[i].roaddegree][roadpart[i].roadtype][1]);
                vallist[rowcnt, colcnt++] = Math.Round(pcival, 2);
                vallist[rowcnt, colcnt++] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PCIGrade[roadpart[i].roaddegree][0],
                    _PCIGrade[roadpart[i].roaddegree][1],
                    _PCIGrade[roadpart[i].roaddegree][2],
                    _PCIGrade[roadpart[i].roaddegree][3]);

                if (lirivalnum > 0 && ririvalnum > 0)
                {
                    if (IsDIRIMTD)
                    {
                        irival = Math.Round((suml + sumr) * 0.5 / lirivalnum, 2);
                    }
                    else
                    {
                        irival = Math.Round(suml / iricnt, 2);
                    }
                    trqival = 100 / (1 + _RQIa[roadpart[i].roaddegree][0] * Math.Exp(_RQIa[roadpart[i].roaddegree][1] * irival));
                    vallist[rowcnt, colcnt] = trqival;
                }
                else
                {
                    if (rowcnt > 3)
                    {
                        vallist[rowcnt, colcnt] = vallist[rowcnt - 1, colcnt];
                    }
                    else
                    {
                        vallist[rowcnt, colcnt] = 0;
                    }
                }

                colcnt++;
                vallist[rowcnt, colcnt++] = string.Format("=IF(F{0}>={1},\"优\",IF(F{0}>={2},\"良\",IF(F{0}>={3},\"中\",IF(F{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _RQIGrade[roadpart[i].roaddegree][0],
                    _RQIGrade[roadpart[i].roaddegree][1],
                    _RQIGrade[roadpart[i].roaddegree][2],
                    _RQIGrade[roadpart[i].roaddegree][3]);

                //Rut
                if (IsRut)
                {
                    int rutxlsnum = Math.Abs(roadpart[i].mile - roadpart[i + 1].mile) * 100 / rutdis;
                    int rutcnt = 0;
                    double sumrutl = 0, sumrutr = 0;
                    int lrutcnt = 0, rrutcnt = 0;
                    List<double> tlruts = new List<double>();
                    List<double> trruts = new List<double>();
                    string[] trut;
                    while (rutcnt++ < rutxlsnum)
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
                    if (rutval <= _RDIRD[0][1])
                    {
                        rutval = _RDIRD[0][0] - _RDIa[0] * rutval;
                    }
                    else if (rutval <= _RDIRD[1][1])
                    {
                        rutval = _RDIRD[1][0] - _RDIa[1] * (rutval - _RDIRD[0][1]);
                    }
                    else
                    {
                        rutval = 0;
                    }
                    if (roadpart[i].roadtype == 0)
                    {
                        vallist[rowcnt, colcnt++] = rutval;
                        vallist[rowcnt, colcnt++] = string.Format("=IF(H{0}>={1},\"优\",IF(H{0}>={2},\"良\",IF(H{0}>={3},\"中\",IF(H{0}>={4},\"次\",\"差\"))))",
                            rowcnt + 3,
                            _RDIGrade[roadpart[i].roaddegree][0],
                            _RDIGrade[roadpart[i].roaddegree][1],
                            _RDIGrade[roadpart[i].roaddegree][2],
                            _RDIGrade[roadpart[i].roaddegree][3]);
                    }
                    else
                    {
                        vallist[rowcnt, colcnt++] = "-";
                        vallist[rowcnt, colcnt++] = "-";
                    }
                }
                else
                {
                    vallist[rowcnt, colcnt++] = "-";
                    vallist[rowcnt, colcnt++] = "-";
                }

                vallist[rowcnt, colcnt++] = string.Format("=ROUND(({1}*D{0}+{2}*F{0}+{3}*IF(EXACT(H{0},\"-\"),0,H{0}))/({1}+{2}+{3}),2)",
                        rowcnt + 3,
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][0],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][1],
                        _PQIW[roadpart[i].roaddegree][roadpart[i].roadtype][2]);

                vallist[rowcnt, colcnt++] = string.Format("=IF(J{0}>={1},\"优\",IF(J{0}>={2},\"良\",IF(J{0}>={3},\"中\",IF(J{0}>={4},\"次\",\"差\"))))",
                    rowcnt + 3,
                    _PQIGrade[roadpart[i].roaddegree][0],
                    _PQIGrade[roadpart[i].roaddegree][1],
                    _PQIGrade[roadpart[i].roaddegree][2],
                    _PQIGrade[roadpart[i].roaddegree][3]);
                ++rowcnt;
            }

            MSExcel.Range destrange = worksheet.get_Range(String.Format("A3:K{0}", rowcnt + 2));
            destrange.Value2 = vallist;
            WritePQI2Statistics(worksheet);
            GlobalExcel.SetBorderLine(destrange, 53);

            if (MainForm._IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 11, true);
                GlobalExcel.Reflection(worksheet, 3, 2, false);
            }
        }

        private static void WritePQI2Statistics(XSSFSheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次", "差" };
            int scol = 19, srow = 2;
            for (int i = 0; i < 5; ++i)
            {
                _Worksheet.GetRow(srow).GetCell(scol+i).SetCellValue(string.Format("=ABS(SUMIF(K:K,\"{0}\",A:A)-SUMIF(K:K,\"{0}\",B:B))", degstr[i]));
                _Worksheet.GetRow(srow+2).GetCell(scol+i).SetCellValue("=ABS(SUM(A:A)-SUM(B:B))");
                _Worksheet.GetRow(srow+1).GetCell(scol+i).SetCellValue(string.Format("={0}3/{0}5", (char)('A' + 19 + i)));
            }
            _Worksheet.GetRow(1).GetCell(12).SetCellValue("=CONCATENATE(\"路面PQI评价等级“优”率占路段总数\",ROUND(T4,4)*100,\"%，“良”率占路段总数\",ROUND(U4,4)*100,\"%，“中”率占路段总数\",ROUND(V4,4)*100,\"%，“次”率占路段总数\",ROUND(W4,4)*100,\"%，“差”率占路段总数\",ROUND(X4,4)*100,\"%。\")");
        }
    }
}
