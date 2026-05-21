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
    /// 2001发行等级公路规范，
    /// JTJ 073.2-2001 公路沥青路面养护技术规范
    /// JTJ 073.1-2001 公路水泥混凝土路面养护技术规范
    /// </summary>
    class MyExcelDegree2001
    {
        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

        private static double[] _PCIGrade;
        private static double[] _snRQIGrade;// 等级区间
        private static double[] _lqRQIGrade;// 等级区间
        private static double[] _DBLGrade;
        private static double[][] _WeightParm;//0-沥青，1-水泥
        private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
        public static string[] _RoadQuailityStr = { "水泥PCI", "水泥RQI", "沥青PCI", "沥青RQI",};
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

        private static Dictionary<string, int> _RoadGradeDict;
        private static Dictionary<string, int>[] _DisType;

        private static ExcelGPS[] _GPSInfo = null;
        private static double[][] _PCIa;
        private static double[][] _RQIa;
        private static List<MilePart> _RoadPart = null;
        private static Disease[] _RoadDisList = null;
        private static Disease[] _RoadRepairList = null;
        private static double[] _LIRIMeanVal = null;
        private static double[] _RIRIMeanVal = null;
        private static double[] _LMTDMeanVal = null;
        private static double[] _RMTDMeanVal = null;
        private static double[] _CMTDMeanVal = null;
        private static string[][] _SNCorrectVal = null;
        private static string[] _SNCorrThredVa = null;
        private static string[][] _DBLCorrectVal = null;

        private static void InitXlsParm()
        {
            _PCIGrade = new double[5];
            _DBLGrade=new double[8];
            _snRQIGrade = new double[5];
            _lqRQIGrade = new double[5];
            _SNCorrectVal = new string[4][];
            _DBLCorrectVal = new string[3][];
            _PCIa = new double[2][];
            _RQIa = new double[2][];
            for (int i = 0; i < 4; i++)
            {
                _SNCorrectVal[i] = new string[3];
            }
            for (int i = 0; i < 3; i++)
            {
                _DBLCorrectVal[i] = new string[3];
            }
            _RoadSocre = new Dictionary<string, CityRoadDis>[2];
            for (int i = 0; i < 2; i++)
            {
                _RoadSocre[i] = new Dictionary<string, CityRoadDis>();
                _PCIa[i] = new double[2];
                _RQIa[i] = new double[2];
            }

            _WeightParm = new double[2][];

            _DisType = new Dictionary<string, int>[2];
            for (int i = 0; i < 2; i++)
            {
                _DisType[i] = new Dictionary<string, int>();
            }
            _DisType[0].Add("裂缝类", 0);
            _DisType[0].Add("松散类", 1);
            _DisType[0].Add("变形类", 2);
            _DisType[0].Add("其他类", 3);
            _DisType[1].Add("断裂类", 0);
            _DisType[1].Add("竖向位移类", 1);
            _DisType[1].Add("接缝类", 2);
            _DisType[1].Add("表层类", 3);

            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }
        }
        /// <summary>
        ///提取 公式系数 扣分值与系数 等级区间 
        /// </summary>
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
                            if (i == 0)//沥青
                            {
                                if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面病害类型")
                                {
                                    foreach (XmlNode node in subnode.ChildNodes)
                                    {
                                        CityRoadDis roaddis = new CityRoadDis();
                                        roaddis._DisName = node.Name;
                                        roaddis._DisType = ((XmlElement)node).GetAttribute("损坏类型");
                                        roaddis.weight =double.Parse( ((XmlElement)node).GetAttribute("权重"));
                                        roaddis.usedwidth = double.Parse(((XmlElement)node).GetAttribute("影响宽度"));
                                        _RoadSocre[i].Add(roaddis._DisName, roaddis);
                                    }
                                }

                                //读取计算RQI的系数
                                else if (i == 0 && subnode.Name == "沥青RQI")
                                {
                                    _RQIa[0][0] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a0"));
                                    _RQIa[0][1] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a1"));
                                }
                            }
                            else //水泥
                            {
                                if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面病害类型")
                                {
                                    foreach (XmlNode node in subnode.ChildNodes)
                                    {
                                        CityRoadDis roaddis = new CityRoadDis();
                                        roaddis._DisName = node.Name;
                                        roaddis.A = double.Parse(((XmlElement)node).GetAttribute("A系数"));
                                        roaddis.B = double.Parse(((XmlElement)node).GetAttribute("B系数"));
                                        _RoadSocre[i].Add(roaddis._DisName, roaddis);
                                    }
                                }

                                //读取计算RQI的系数
                                else if (subnode.Name == "水泥RQI")
                                {
                                    _RQIa[1][0] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a0"));
                                    _RQIa[1][1] = Convert.ToDouble(((XmlElement)subnode).GetAttribute("a1"));
                                }
                                else if (subnode.Name == "水泥路面修正系数")
                                {
                                    string w1 = ((XmlElement)subnode).GetAttribute("w1");
                                    string w2 = ((XmlElement)subnode).GetAttribute("w2");
                                    string w3 = ((XmlElement)subnode).GetAttribute("w3");
                                    string w4 = ((XmlElement)subnode).GetAttribute("w4");
                                    string[] t2 = w2.Split(' ');
                                    string[] t3 = w3.Split(' ');
                                    string[] t4 = w4.Split(' ');
                                    _SNCorrectVal[0][0] = w1;
                                    _SNCorrectVal[1] = t2;
                                    _SNCorrectVal[2] = t3;
                                    _SNCorrectVal[3] = t4;
                                }
                                else if (subnode.Name == "水泥路面修正阈值")
                                {
                                    string thred = ((XmlElement)subnode).GetAttribute("w1");
                                    string[] t1 = thred.Split(' ');
                                    _SNCorrThredVa = t1;
                                }
                                else if (subnode.Name == "水泥断板率修正权系数")
                                {
                                    string w1 = ((XmlElement)subnode).GetAttribute("交叉裂缝");
                                    string w2 = ((XmlElement)subnode).GetAttribute("角隅断裂");
                                    string w3 = ((XmlElement)subnode).GetAttribute("纵横斜向裂缝");
                                    string[] t1 = w1.Split(' ');
                                    string[] t2 = w2.Split(' ');
                                    string[] t3 = w3.Split(' ');
                                    _DBLCorrectVal[0] = t1;
                                    _DBLCorrectVal[1] = t2;
                                    _DBLCorrectVal[2] = t3;
                               

                                }
                            }
                        }
                    }
                }
            }

            //读取等级区间

            foreach (XmlNode rootchild in Elem.ChildNodes)
            {
                if (rootchild.Name == Global.g_ParmStyles[(int)_Setting.ParmStyle])
                {
                    foreach (XmlNode subnode in rootchild.ChildNodes)
                    {
                        if (subnode.Name == "水泥PCI")
                        {
                            string w1 = ((XmlElement)subnode).GetAttribute("等级区间");
                           
                            string[] t1 = w1.Split(' ');
                           
                            for (int j = 0; j < t1.Length; j++)
                            {
                                _PCIGrade[j] =double.Parse (t1[j]);
                            }
                         
                        }
                        else if (subnode.Name == "水泥DBL")
                        { 
                            string w2 = ((XmlElement)subnode).GetAttribute("等级区间");
                            string[] t2 = w2.Split(' ');
                            for (int j = 0; j < t2.Length; j++)
                            {
                                _DBLGrade[j] = double.Parse(t2[j]);
                            } 
                        }
                        else if (subnode.Name == "水泥RQI")
                        {
                            string w1 = ((XmlElement)subnode).GetAttribute("等级区间");
                            string[] t1 = w1.Split(' ');
                            for (int j = 0; j < t1.Length; j++)
                            {
                                _snRQIGrade[j] = double.Parse(t1[j]);
                            }
                        }
                        else if (subnode.Name == "沥青RQI")
                        {
                            string w1 = ((XmlElement)subnode).GetAttribute("等级区间");
                            string[] t1 = w1.Split(' ');
                            for (int j = 0; j < t1.Length; j++)
                            {
                                _lqRQIGrade[j] = double.Parse(t1[j]);
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
            public double weight = 0;
            public double usedwidth = 0;
            public double A = 0;
            public double B = 0;
        }

        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
            bool IsDis, bool IsMeanIRI)
        {
            bool IRIRes = true,  GPSRes = true;
            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }
            _RoadPart = new List<MilePart>();


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

            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);

            if (IsDis)
            {
                GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, ref _RoadDisList);
            }
            if (IsMeanIRI) IRIRes = GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, _Setting.IsWarning);
        
            return IRIRes  && GPSRes;
        }

        public static void OutputIRI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTJ H20-2001\路面平整度评价等级记录表.xlsx",
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
                    vallist[i, 5] = String.Format("=ROUND((D{0}+E{0})/2,5)", i + 4);
                }
                else
                {
                    vallist[i, 5] = String.Format("=ROUND((D{0}),5)", i + 4);
                }
                if (prjinfo._RoadType == 0) //沥青
                {
                    vallist[i, 6] = String.Format("=IF({0}+{1}*F{2}>=0,{0}+{1}*F{2},0)", _RQIa[0][0], _RQIa[0][1], i + 4);
                    vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                  i + 4, _lqRQIGrade[0], _lqRQIGrade[1], _lqRQIGrade[2], _lqRQIGrade[3]);
                }
                else
                {
                    vallist[i, 6] = String.Format("=IF({0}+{1}*F{2}>=0,{0}+{1}*F{2},0)", _RQIa[1][0], _RQIa[1][1], i + 4);
                    vallist[i, 7] = string.Format("=IF(G{0}>={1},\"优\",IF(G{0}>={2},\"良\",IF(G{0}>={3},\"中\",IF(G{0}>={4},\"次\",\"差\"))))",
                  i + 4, _snRQIGrade[0], _snRQIGrade[1], _snRQIGrade[2], _snRQIGrade[3]);
                }
                //vallist[i, 7] = string.Format("=IF(G{0}>={1},\"A\",IF(G{0}>={2},\"B\",IF(G{0}>={3},\"C\",\"D\")))",
                //    i + 4, _RQIGrade[roadpart[i].roaddegree][0], _RQIGrade[roadpart[i].roaddegree][1], _RQIGrade[roadpart[i].roaddegree][2]);
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
            string[] degstr = { "优", "良", "中", "次","差" };
            MSExcel.Range destrange = _Worksheet.get_Range("P3:T5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < degstr.Length; i++)
            {
                val[0, i] = string.Format("=ABS(SUMIF(H:H,\"{0}\",A:A)-SUMIF(H:H,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 15 + i));
            }
            destrange.Value2 = val;
            _Worksheet.Cells[2, 9] = "=CONCATENATE(\"路面平整度评价等级“优”率占路段总数\",ROUND(P4,4)*100,\"%，“良”率占路段总数\",ROUND(Q4,4)*100,\"%，“中”率占路段总数\",ROUND(R4,4)*100,\"%，“次”率占路段总数\",ROUND(S4,4)*100,\"%，“差”率占路段总数\",ROUND(T4,4)*100,\"%。\")";
        }
        
        public static double _LqComputPCI(RoadDiseaseType[] disarea, double partarea)
        {
            double[] totalareatmp = new double[disarea.Length];
            double DR = 0;
            int len = _RoadSocre[0].Keys.Count;

            for (int i = 0; i < len ; i++)
            {
                totalareatmp[i] = disarea[i].totalarea / partarea;
            }

            for (int i = 0; i < len ; i++)
            {
                disarea[i].weight = _RoadSocre[0][disarea[i].disname].weight;
                DR += disarea[i].weight * totalareatmp[i] * 100;
            }

            return 100 -15*Math.Pow( DR,0.412);
        }
        public static double _SnComputPCI(RoadDiseaseType[] disarea, double snum)
        {
            double DP = 0;
            double R = 0;
            double sumDP = 0;
            double W = 0;
            double DR = 0;
            double tt = 0;
            int lqlen = _RoadSocre[0].Keys.Count;
            int snlen = _RoadSocre[1].Keys.Count;

            for (int i = lqlen; i < lqlen + snlen; i++)
            {
                tt = disarea[i].platenum / snum;
               // disarea[i].platenum /= snum;// 面积按板块个数计算 后面需要修改
                DP = disarea[i].para_A * disarea[i].para_B * tt;
                sumDP += DP;
            }

            for (int i = lqlen; i < lqlen + snlen; i++)
            {
                if (sumDP != 0)
                {
                    tt = disarea[i].platenum / snum;
                    DP = disarea[i].para_A * disarea[i].para_B * tt;
                    R = DP / sumDP;
                }
                if (R < double.Parse(_SNCorrThredVa[0])) //R<0.2
                {
                    W = double.Parse(_SNCorrectVal[0][0]) * R; //2.5*R
                }
                else if (R < double.Parse(_SNCorrThredVa[1])) // 0.2<=R<0.5
                {
                    W = double.Parse(_SNCorrectVal[1][0]) + double.Parse(_SNCorrectVal[1][1]) * R+ double.Parse(_SNCorrectVal[1][2]); //0.5*0.686r--0.1372
                }
                else if (R < double.Parse(_SNCorrThredVa[2])) // 0.5<=R<0.8
                {
                    W = double.Parse(_SNCorrectVal[2][0]) + double.Parse(_SNCorrectVal[2][1]) * R + double.Parse(_SNCorrectVal[2][2]); //2.5*R
                }
                else   //  0.8<=R
                {
                    W = W = double.Parse(_SNCorrectVal[3][0]) + double.Parse(_SNCorrectVal[3][1]) * R + double.Parse(_SNCorrectVal[3][2]); //2.5*R
                }
                DR += DP * W;
            }

            return 100 -DR;
        }
        public static double _SnComputDBL(RoadDiseaseType[] disarea, double snum)
        {
            double DB = 0;
            double tt = 0;
            
            string[] dis = { "交叉裂缝.轻", "交叉裂缝.中", "交叉裂缝.重", "角隅断裂.轻", "角隅断裂.中", "角隅断裂.重", "纵横斜向裂缝.轻", "纵横斜向裂缝.中", "纵横斜向裂缝.重" };
            int lqlen = _RoadSocre[0].Keys.Count;
            int snlen = _RoadSocre[1].Keys.Count;

            for (int i = lqlen; i < lqlen + snlen; i++)
            {
                for (int j = 0; j < dis.Length; ++j)
                {
                    if (disarea[i].disname == dis[j])
                    {
                        int k = 0;
                        if (j > 2 && j < 6)
                        {
                            k = 1;
                            j = j - 3;
                        }
                        else if (j >= 6)
                        {
                            k = 2;
                            j = j - 6;
                        }
                        tt = disarea[i].platenum / snum;
                        //disarea[i].platenum /= snum;// 面积按板块个数计算 后面需要修改
                        DB += tt * double.Parse(_DBLCorrectVal[k][j]);
                        break;
                    }
                }
            }
            return DB*100;
        }
        
        public static void OutputPCI(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTJ H20-2001\路面破损评价等级记录表.xlsx",
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
            int len = roadpart.Count - 1, dlen = arrdis.Length, lastmile = roadpart[0].mile;
            object[,] vallist = new object[len, 6];
            int[] dis = new int[70];

            int typeidx = 0;
            bool res = false;
			
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                double pcival = 0,dblval=0;
                int smile = roadpart[i].mile;
                int emile = roadpart[i + 1].mile;
                double milelength = Math.Abs(smile - emile);

                RoadDiseaseTypes.Clear();
                while (j < dlen && ((prjinfo._Direction > 0 && arrdis[j].m_mile >= smile && arrdis[j].m_mile < emile)
                    || (prjinfo._Direction < 0 && arrdis[j].m_mile <= smile && arrdis[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                        arrdis[j].RoadType, arrdis[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        if (typeidx >= 30)//水泥
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].platenum += 1;
                        }
                        else
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        }
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
                if (roadpart[i].roadtype == 0)
                {
                    pcival = _LqComputPCI(RoadDiseaseTypes.roaddis[roadpart[i].roadtype], _RoadConfig.DetectWidth * milelength);
                    vallist[i, 5] = string.Format("=IF(D{0}>={1},\"优\",IF(D{0}>={2},\"良\",IF(D{0}>={3},\"中\",IF(D{0}>={3},\"次\",\"差\"))))",
                        i + 3, _PCIGrade[0], _PCIGrade[1], _PCIGrade[2], _PCIGrade[3]);
                }
                else if (roadpart[i].roadtype == 1) //水泥
                {
                    if (milelength < 4)
                    {
                        milelength = 4;
                    }
                    pcival = _SnComputPCI(RoadDiseaseTypes.roaddis[roadpart[i].roadtype], milelength / 4);
                    dblval = _SnComputDBL(RoadDiseaseTypes.roaddis[roadpart[i].roadtype], milelength / 4);
                    vallist[i, 5] = string.Format("=IF(AND(D{0}>={1},E{0}<={5}),\"优\",IF(AND(D{0}>={2},E{0}>={6},E{0}<={7}),\"良\",IF(AND(D{0}>={3},E{0}>={8},E{0}<={9}),\"中\",IF(AND(D{0}>={4},E{0}>={10},E{0}<={11}),\"次\",\"差\"))))",
                        i + 3, _PCIGrade[0], _PCIGrade[1], _PCIGrade[2], _PCIGrade[3], _DBLGrade[0], _DBLGrade[1], _DBLGrade[2], _DBLGrade[3], _DBLGrade[4], _DBLGrade[5], _DBLGrade[6]);
                }
                vallist[i, 3] = Math.Round(pcival, 5);
                vallist[i, 4] = Math.Round(dblval, 5);

            }
            destrange = worksheet.get_Range(String.Format("A3:F{0}", len + 2));
            destrange.Value2 = vallist;
            WritePCIStatistics(worksheet);
            destrange = worksheet.get_Range(String.Format("A1:F{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(worksheet, 3, 1, 5, true);       
                GlobalExcel.Reflection(worksheet, 3, 1, 2, false);
            }
        }
        private static void WritePCIStatistics(MSExcel.Worksheet _Worksheet)
        {
            string[] degstr = { "优", "良", "中", "次","差" };
            MSExcel.Range destrange = _Worksheet.get_Range("N3:R5");
            object[,] val = new object[3, 5];
            for (int i = 0; i < 5; ++i)
            {
                val[0, i] = string.Format("=ABS(SUMIF(F:F,\"{0}\",A:A)-SUMIF(F:F,\"{0}\",B:B))", degstr[i]);
                val[2, i] = "=ABS(SUM(A:A)-SUM(B:B))";
                val[1, i] = string.Format("={0}3/{0}5", (char)('A' + 13 + i));
            }

            destrange.Value2 = val;
            _Worksheet.Cells[2, 7] = "=CONCATENATE(\"沥青路面PCI评价等级“优”率占路段总数\",ROUND(N4,4)*100,\"%，“良”率占路段总数\",ROUND(O4,4)*100,\"%，“中”率占路段总数\",ROUND(P4,4)*100,\"%，“次”率占路段总数\",ROUND(Q4,4)*100,\"%，“差”率占路段总数\",ROUND(R4,4)*100,\"%。\")";
        }

        public static void OutputDis(MSExcel.Application excelApp, string path, ProjectInfo prjinfo, DirectoryInfo prjdir, int disval)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTJ H20-2001\路面病害面积统计表.xlsx",
                System.Windows.Forms.Application.StartupPath);
            string Destxls = string.Format(@"{0}\{1}_病害统计_{2}m.xlsx", path, prjdir.Name, disval);
            MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _Worksheet_lb = _Workbook.Sheets["病害列表"] as MSExcel.Worksheet;
           // WriteDisLB2Xls(_Worksheet_lb, prjinfo, _RoadDisList);
            WriteDisLB2Xls_roadpart(_Worksheet_lb, prjinfo, _RoadDisList, _RoadPart);

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
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0, troadtype = -1;
            object[,] val = new object[len, 11];
            foreach (Disease tdis in dislist)
            {
                for (int k = 0; k < roadpart.Count - 1; ++k)
                {
                    if (roadpart[k].mile <= tdis.m_mile && tdis.m_mile < roadpart[k + 1].mile)
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
                            val[i, 8] = tdis.calcwidth;
                            val[i, 9] = tdis.imgname;
                            val[i, 10] = tdis.imgpath;
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
            destrange = _Worksheet.get_Range(String.Format("A3:K{0}", tlen + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", tlen + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Out_roadimg == 0) //不导出路面图像
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

            if (_Setting.IsExcelSort)
            {
                GlobalExcel.Reflection(_Worksheet, 3, 1, 9, true);
            }
        }

        private static void WriteDisLB2Xls(MSExcel.Worksheet _Worksheet, ProjectInfo prjinfo, Disease[] dislist)
        {
            MSExcel.Range destrange;
            int len = dislist.Length, i = 0;
            object[,] val = new object[len, 11];
            foreach (Disease tdis in dislist)
            {
                val[i, 0] = tdis.m_mile;
                val[i, 1] = prjinfo._RoadNum;
                val[i, 2] = tdis.RoadDisType;
                val[i, 3] = tdis.rect.Height * _RoadConfig.HeightScale;
                val[i, 4] = tdis.rect.Width * _RoadConfig.WidthScale;
                val[i, 5] = (tdis.rect.Width / 2 + tdis.rect.X) * _RoadConfig.WidthScale;
                val[i, 6] = tdis.Area;
                val[i, 7] = tdis.calcheight;
                val[i, 8] = tdis.calcwidth;
                val[i, 9] = tdis.imgname;
                val[i, 10] = tdis.imgpath;
                ++i;
            }
            destrange = _Worksheet.get_Range(String.Format("A3:K{0}", len + 2));
            destrange.Value2 = val;

            destrange = _Worksheet.get_Range(String.Format("A1:K{0}", len + 2));
            GlobalExcel.SetBorderLine(destrange, 53);

            if (_Setting.Out_roadimg == 0) //不导出路面图像
            {
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 11]).EntireColumn.Delete();
                ((MSExcel.Range)_Worksheet.Cells[System.Reflection.Missing.Value, 10]).EntireColumn.Delete();
            }

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
                double pcival = 0, dblval = 0;
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
                        if (typeidx >= 30)//水泥
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].platenum += 1;
                        }
                        else
                        {
                            RoadDiseaseTypes.roaddis[roadpart[i].roadtype][typeidx].totalarea += arrdis[j].Area;
                        }
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
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].platenum;
                    }

                    pcival = _SnComputPCI(RoadDiseaseTypes.roaddis[roadpart[i].roadtype], milelength / 4);
                    disval[0, disnum] = Math.Round(pcival, 5);//PCI
                    dblval = _SnComputDBL(RoadDiseaseTypes.roaddis[1], milelength / 4);
                    disval[0, disnum + 1] = Math.Round(dblval, 5); //DBL

                    disval[0, disnum + 2] = string.Format("=IF(AND(AR{0}>={1},AS{0}<={5}),\"优\",IF(AND(AR{0}>={2},AS{0}>={6},AS{0}<={7}),\"良\",IF(AND(AR{0}>={3},AS{0}>={8},AS{0}<={9}),\"中\",IF(AND(AR{0}>={4},AS{0}>={10},AS{0}<={11}),\"次\",\"差\"))))",
                        rowcnt_sn_s,
                        _PCIGrade[0],
                        _PCIGrade[1], 
                        _PCIGrade[2],
                        _PCIGrade[3],
                        _DBLGrade[0], _DBLGrade[1], _DBLGrade[2], _DBLGrade[3], _DBLGrade[4], _DBLGrade[5], _DBLGrade[6]);

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

                    disnum = RoadDiseaseTypes.DiseaseTypeDict[roadpart[i].roadtype].Count;
                    disval = new object[1, disnum + 3];
                    for (int di = 0, kk = 0; di < disnum; ++di, ++colcnt, ++kk)
                    {
                        disval[0, kk] = RoadDiseaseTypes.roaddis[roadpart[i].roadtype][di].totalarea;
                    }

                    pcival = _LqComputPCI(RoadDiseaseTypes.roaddis[roadpart[i].roadtype], _RoadConfig.DetectWidth * milelength);
                    disval[0, disnum] = Math.Round(pcival, 5);
                    disval[0, disnum + 1] = string.Format("=IF(AH{0}>={1},\"优\",IF(AH{0}>={2},\"良\",IF(AH{0}>={3},\"中\",IF(AH{0}>={3},\"次\",\"差\"))))",
                        rowcnt_lq_s, 
                        _PCIGrade[0], _PCIGrade[1], _PCIGrade[2], _PCIGrade[3]);

                    destrange = worksheet_lqhz.get_Range(string.Format("D{0}:{1}{0}", rowcnt_lq_s, GlobalExcel.GetCol((char)('D' + disnum + 1))));
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

            destrange = worksheet_lqhz.get_Range(String.Format("A1:AI{0}", rowcnt_lq_s));
            GlobalExcel.SetBorderLine(destrange, 53);
            destrange = worksheet_snhz.get_Range(String.Format("A1:AT{0}", rowcnt_sn_s));
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
                    if (i > 22)
                    {
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!A{0}:A{0},\"<>\",沥青病害汇总表!A{0}:A{0})/3", Convert.ToChar('A' + i - 23));
                    }
                    else
                    {
                        disval[i, 0] = string.Format("=SUMIF(沥青病害汇总表!{0}:{0},\"<>\",沥青病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    }
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
                worksheet_sntj.Cells[2, 2] = Math.Abs(roadpart[0].mile - roadpart[len].mile) / 4;
                worksheet_sntj.Cells[2, 6] = Math.Abs(roadpart[0].mile - roadpart[len].mile);
                disval = new object[disnum, 1];
                for (int i = 0; i < disnum; i++)
                {
                    //disval[i, 0] = string.Format("=水泥病害汇总表!{0}{1}", Convert.ToChar('D' + i), rowcnt_sn_s);
                    if (i > 22)
                    {
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!A{0}:A{0},\"<>\",水泥病害汇总表!A{0}:A{0})/3", Convert.ToChar('A' + i-23));
                    }
                    else
                    {
                        disval[i, 0] = string.Format("=SUMIF(水泥病害汇总表!{0}:{0},\"<>\",水泥病害汇总表!{0}:{0})/3", Convert.ToChar('D' + i));
                    }
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

    }
}
