//using Farmework.Other;
//using HNRoadFormatConverter.MyConfig;
//using HNRoadFormatConverter.MyEntitys;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using System.Xml; 

//namespace HNRoadFormatConverter.ExcelManager
//{
//    public class MyExcelBase
//    { 
//            static XRSetting _Setting = XRSetting.GetInstance();
//            static RoadConfig _RoadConfig = RoadConfig.GetInstance();

//            private static double[][][] _RQIGrade;//道路等级 路面材质 等级区间
//            private static double[][] _RDIGrade;
//            private static double[][] _PCIGrade;
//            private static double[][] _PQIGrade;
//            private static double[][] _PBIGrade;
//            private static double[][] _PWIGrade;
//            private static double[][] _RDIRD;
//            private static double[] _RDIa;
//            private static double[] _PWIa;
//            private static double[] _PBIThresh;
//            private static double[] _PBIScore;

//            private static double[][][] _RQIa;//公路等级 路面材质 参数序号
//            private static double[][][] _PCIa;//公路等级 路面材质 参数序号

//            /// <summary>
//            /// WPCI WRQI WRDI WPBI WPWI
//            /// </summary>
//            private static double[][][] _PQIW;//公路等级 路面材质 参数序号

//            /// <summary>
//            /// WSCI WPQI WBCI WTCI
//            /// </summary>
//            private static double[] _MQIW;
//            /// <summary>
//            /// MQI、SCI、SRI、PSSI、BCI、TCI指标的优良中次差等级区间
//            /// </summary>
//            private static double[] _MQIGrade;

//            private static double[][] _WeightParm;//0-沥青，1-水泥
//            private static Dictionary<string, CityRoadDis>[] _RoadSocre;//0-沥青，1-水泥
//            public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };

//            public static Dictionary<string, int> _RoadGradeDict;

//            public static List<MilePart> _RoadPart = null;

//            public static List<MilePart> _RoadPart10 = null;//整10米桩号分段
//            private static double[] _SpeedVal10 = null;
//            private static string[] _MarkVal10 = null;

//            public static List<MilePart> _RoadPart1M = null;//1米桩号分段
//        //    private static Disease[] _RoadDisList = null;
//          //  private static Disease[] _RoadRepairList = null;

//            private static double[] _LIRIMeanVal = null;
//            private static double[] _RIRIMeanVal = null;

//            private static double[] _LMTDMeanVal = null;
//            private static double[] _RMTDMeanVal = null;
//            private static double[] _CMTDMeanVal = null;

//            private static double[] _LRutMeanVal = null;
//            private static double[] _RRutMeanVal = null;
//            private static double[] _SRutMeanVal = null;

//            private static double[] _LRutMaxVal = null;
//            private static double[] _RRutMaxVal = null;
//            private static double[] _SRutMaxVal = null;

//            private static double[] _SRutDisVal = null;
//            private static int[] _SRutDisMile = null;
//            private static double[] _rutThresh = new double[2];
//            private static int[][] _PBIVal = null;
//            private static double[] _LDeltaHVal = null;
//            private static double[] _RDeltaHVal = null;
//            private static double[] _LDeltaHVal_1M = null;
//            private static double[] _RDeltaHVal_1M = null;
//            private static double[] _SpeedVal = null;
//            private static string[] _MarkVal = null;

//            private static double[] _DeltaHVal = null;

//            private static double[] _LMPDMeanVal = null;
//            private static double[] _RMPDMeanVal = null;
//            private static double[] _CMPDMeanVal = null;



//            private static double[] _Curvature = null;
//            private static double[] _CrossSlope = null;
//            private static double[] _HeightSlope = null;


//            private static ExcelGPS[] _GPSInfo = null;
//            private static void InitXlsParm()
//            {
//                int len = _RoadGradeStr.Length;

//                _RQIGrade = new double[len][][];
//                _RDIGrade = new double[len][];
//                _PCIGrade = new double[len][];
//                _PQIGrade = new double[len][];
//                _PBIGrade = new double[len][];
//                _PWIGrade = new double[len][];

//                _RQIa = new double[len][][];
//                _PCIa = new double[len][][];
//                _PQIW = new double[len][][];

//                _MQIW = new double[4];
//                _MQIGrade = new double[5];

//                for (int i = 0; i < len; i++)
//                {
//                    _RQIGrade[i] = new double[2][];
//                    _RDIGrade[i] = new double[5];
//                    _PCIGrade[i] = new double[5];
//                    _PQIGrade[i] = new double[5];
//                    _PBIGrade[i] = new double[5];
//                    _PWIGrade[i] = new double[5];

//                    _RQIa[i] = new double[2][];
//                    _PCIa[i] = new double[2][];
//                    _PQIW[i] = new double[2][];
//                    for (int j = 0; j < 2; j++)
//                    {
//                        _RQIGrade[i][j] = new double[5];
//                        _PCIa[i][j] = new double[2];
//                        _PQIW[i][j] = new double[5];
//                        _RQIa[i][j] = new double[2];
//                    }
//                }
//                _PBIThresh = new double[4];
//                _PBIScore = new double[4];
//                _RDIa = new double[2];
//                _PWIa = new double[2];
//                _RDIRD = new double[2][];
//                for (int i = 0; i < 2; i++)
//                {
//                    _RDIRD[i] = new double[2];
//                }

//                _RoadSocre = new Dictionary<string, CityRoadDis>[2];
//                for (int i = 0; i < 2; i++)
//                {
//                    _RoadSocre[i] = new Dictionary<string, CityRoadDis>();
//                }

//                _WeightParm = new double[2][];

//                _RoadGradeDict = new Dictionary<string, int>();
//                for (int i = 0; i < _RoadGradeStr.Length; ++i)
//                {
//                    _RoadGradeDict.Add(_RoadGradeStr[i], i);
//                }
//            }

//            public static void LoadXlsParm()
//            {
//                InitXlsParm();

//                XmlDocument Doc = new XmlDocument();
//                Doc = new XmlDocument();
//                XmlElement Elem;
//                XmlNodeList xmlNodes;


//                //读取病害类型
//                Doc.Load(System.Windows.Forms.Application.StartupPath + "\\ParaVal.xml");    //加载Xml文件  
//                Elem = Doc.DocumentElement;   //获取根节点  
//                xmlNodes = Elem.ChildNodes;

//                for (int i = 0; i < 2; i++)
//                {
//                    foreach (XmlNode rootchild in Elem.ChildNodes)
//                    {
//                        if (rootchild.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
//                        {
//                            foreach (XmlNode subnode in rootchild.ChildNodes)
//                            {
//                                if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面病害类型")
//                                {
//                                    foreach (XmlNode node in subnode.ChildNodes)
//                                    {
//                                        CityRoadDis roaddis = new CityRoadDis();
//                                        roaddis._DisName = node.Name;
//                                        roaddis._UseWidth = Convert.ToDouble(((XmlElement)node).GetAttribute("影响宽度"));
//                                        roaddis._Weight = Convert.ToDouble(((XmlElement)node).GetAttribute("权重"));
//                                        _RoadSocre[i].Add(roaddis._DisName, roaddis);
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }

//                //读取等级区间
//                int dlen = _RoadGradeStr.Length;
//                string strval;
//                string[] s;
//                double[] val;
//                for (int i = 0; i < dlen; i++)
//                {
//                    foreach (XmlNode rootchild in Elem.ChildNodes)
//                    {
//                        if (rootchild.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
//                        {
//                            foreach (XmlNode subnode in rootchild.ChildNodes)
//                            {
//                                if (subnode.Name == _RoadGradeStr[i])
//                                {
//                                    foreach (XmlNode node in subnode.ChildNodes)
//                                    {
//                                        if (node.Name == "RQI")
//                                        {
//                                            foreach (XmlNode nnode in node.ChildNodes)
//                                            {
//                                                strval = ((XmlElement)nnode).GetAttribute("等级区间");
//                                                s = strval.Split(' ');
//                                                val = new double[s.Length];
//                                                for (int j = 0; j < s.Length; j++)
//                                                {
//                                                    val[j] = Convert.ToDouble(s[j]);
//                                                }
//                                                val.CopyTo(_RQIGrade[i][RoadDiseaseTypes.roadtypedict[nnode.Name]], 0);
//                                                _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("w1"));
//                                                _RQIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("w2"));
//                                            }
//                                        }
//                                        else
//                                        {
//                                            strval = ((XmlElement)node).GetAttribute("等级区间");
//                                            s = strval.Split(' ');
//                                            val = new double[s.Length];
//                                            for (int j = 0; j < s.Length; j++)
//                                            {
//                                                val[j] = Convert.ToDouble(s[j]);
//                                            }
//                                            if (node.Name == "RDI")
//                                            {
//                                                val.CopyTo(_RDIGrade[i], 0);
//                                            }
//                                            else if (node.Name == "PWI")
//                                            {
//                                                val.CopyTo(_PWIGrade[i], 0);
//                                            }
//                                            else if (node.Name == "PBI")
//                                            {
//                                                val.CopyTo(_PBIGrade[i], 0);
//                                            }
//                                            else if (node.Name == "PCI")
//                                            {
//                                                val.CopyTo(_PCIGrade[i], 0);
//                                                foreach (XmlNode nnode in node.ChildNodes)
//                                                {
//                                                    _PCIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a0"));
//                                                    _PCIa[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("a1"));
//                                                }
//                                            }
//                                            else if (node.Name == "PQI")
//                                            {
//                                                val.CopyTo(_PQIGrade[i], 0);
//                                                foreach (XmlNode nnode in node.ChildNodes)
//                                                {
//                                                    _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][0] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPCI"));
//                                                    _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][1] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRQI"));
//                                                    _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][2] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WRDI"));
//                                                    _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][3] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPBI"));
//                                                    _PQIW[i][RoadDiseaseTypes.roadtypedict[nnode.Name]][4] = Convert.ToDouble(((XmlElement)nnode).GetAttribute("WPWI"));
//                                                }
//                                            }
//                                        }
//                                    }
//                                }
//                                //读取计算RDI的系数
//                                else if (i == 0 && subnode.Name == "RDI系数")
//                                {
//                                    _RDIRD[0][0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙常数a"));
//                                    _RDIRD[1][0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙常数b"));
//                                    _RDIRD[0][1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙RDa"));
//                                    _RDIRD[1][1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙RDb"));
//                                    _RDIa[0] = double.Parse(((XmlElement)subnode).GetAttribute("车辙a0"));
//                                    _RDIa[1] = double.Parse(((XmlElement)subnode).GetAttribute("车辙a1"));
//                                }
//                                //读取计算PBI的系数
//                                else if (i == 0 && subnode.Name == "PBI系数")
//                                {
//                                    strval = ((XmlElement)subnode).GetAttribute("划分标准");
//                                    s = strval.Split(' ');
//                                    val = new double[s.Length];
//                                    for (int j = 0; j < s.Length; j++)
//                                    {
//                                        val[j] = Convert.ToDouble(s[j]);
//                                    }
//                                    val.CopyTo(_PBIThresh, 0);

//                                    strval = ((XmlElement)subnode).GetAttribute("扣分");
//                                    s = strval.Split(' ');
//                                    val = new double[s.Length];
//                                    for (int j = 0; j < s.Length; j++)
//                                    {
//                                        val[j] = Convert.ToDouble(s[j]);
//                                    }
//                                    val.CopyTo(_PBIScore, 0);
//                                }
//                                //读取计算PWI的系数
//                                else if (i == 0 && subnode.Name == "PWI系数")
//                                {
//                                    _PWIa[0] = double.Parse(((XmlElement)subnode).GetAttribute("a0"));
//                                    _PWIa[1] = double.Parse(((XmlElement)subnode).GetAttribute("a1"));
//                                }
//                                else if (i == 0 && subnode.Name == "MQI系数")
//                                {
//                                    _MQIW[0] = double.Parse(((XmlElement)subnode).GetAttribute("WSCI"));
//                                    _MQIW[1] = double.Parse(((XmlElement)subnode).GetAttribute("WPQI"));
//                                    _MQIW[2] = double.Parse(((XmlElement)subnode).GetAttribute("WBCI"));
//                                    _MQIW[3] = double.Parse(((XmlElement)subnode).GetAttribute("WTCI"));

//                                    strval = ((XmlElement)subnode).GetAttribute("等级区间");
//                                    s = strval.Split(' ');
//                                    for (int j = 0; j < s.Length; j++)
//                                    {
//                                        _MQIGrade[j] = Convert.ToDouble(s[j]);
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }

//            private class CityRoadDis
//            {
//                public string _DisType = null;
//                public string _DisName = null;
//                public double _UseWidth = 0.0;
//                public double _Weight = 0.0;
//            }
//            public static List<MilePartD> _RoadPartF = null;//0.1米桩号分段
//            private static double[] _LiriHVal = null;
//            private static double[] _RiriHVal = null;

//            public static bool InitProDataD(DirectoryInfo prjdir, ProjectInfo prjinfo, double disval,
//           bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
//            {
//                _SpeedVal = null;

//                bool IRIRes = true, RutRes = true, MTDRes = true, PBIRes = true, GPSRes = true, MaxRutRes = true, MPDRes = true;


//                #region 国检转换TP
//                if (_RoadPartF != null)
//                {
//                    _RoadPartF.Clear();
//                    _RoadPartF = null;
//                }
//                _RoadPartF = new List<MilePartD>();
//                MilePartD spartF = null;
//                try
//                {
//                    spartF = new MilePartD() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
//                }
//                catch
//                {
//                    MessageBox.Show(string.Format("【低等级农村公路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
//                    System.Environment.Exit(0);
//                }
//                _RoadPartF.Add(spartF);
//                if (prjinfo._IsIRIMTD)
//                {

//                    GlobalExcel.GetAllMilePartD(prjdir.FullName, prjinfo, disval, prjinfo._DirectionInt, _RoadGradeStr, ref _RoadPartF, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
//                    if (IsSpeed)
//                    {
//                        GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPartF, ref _SpeedVal);
//                    }

//                    if (IsPBI)
//                    {
//                        //平整度原始数据 纵断面高程TP   

//                        GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF, disval, 0, ref _LiriHVal);
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                            GlobalExcel.GetIRIHValF(prjinfo, prjdir, _RoadPartF, disval, 1, ref _RiriHVal);
//                        }
//                    }


//                }
//                else
//                {
//                    IRIRes = true;
//                }

//                if (prjinfo._IsRut)
//                {
//                    if (IsMeanRut)
//                    {
//                        RutRes = GlobalExcel.GetRutMeanVal(prjinfo, prjdir, _RoadPartF, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, _Setting.IsWarning);
//                        MaxRutRes = GlobalExcel.GetRutMaxVal(prjinfo, prjdir, _RoadPartF, ref _LRutMaxVal, ref _RRutMaxVal, ref _SRutMaxVal);
//                    }
//                }
//                else
//                {
//                    RutRes = true;
//                }
//                #endregion
//                if (_RoadPartF[0].roaddegree <= 1)
//                {
//                    return IRIRes && RutRes && MTDRes && GPSRes && MPDRes;
//                }
//                else
//                {
//                    return IRIRes && MPDRes;
//                }
//            }

//            public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
//                bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false, bool IsGeoAlig = false)
//            {
//                _SpeedVal = null;

//                bool IRIRes = true, RutRes = true, MTDRes = true, PBIRes = true, GPSRes = true, MaxRutRes = true, MPDRes = true, GeoAligRes = true;
//                if (_RoadPart != null)
//                {
//                    _RoadPart.Clear();
//                    _RoadPart = null;
//                }
//                _RoadPart = new List<MilePart>();

//                if (_RoadPart10 != null)
//                {
//                    _RoadPart10.Clear();
//                    _RoadPart10 = null;
//                }
//                _RoadPart10 = new List<MilePart>();

//                if (_RoadPart1M != null)
//                {
//                    _RoadPart1M.Clear();
//                    _RoadPart1M = null;
//                }

//                _RoadPart1M = new List<MilePart>();
//                MilePart spart = null;
//                try
//                {
//                    spart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
//                }
//                catch
//                {
//                    MessageBox.Show(string.Format("【等级公路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
//                    System.Environment.Exit(0);
//                }
//                _RoadPart.Add(spart);
//                _RoadPart10.Add(spart);
//                _RoadPart1M.Add(spart);
//                GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._DirectionInt, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
//                GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart, ref _MarkVal);
//                if (IsDis)
//                {
//                    GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 1, prjinfo._DirectionInt, _RoadGradeStr, ref _RoadPart1M, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);

//                    if (_Setting.OutRut == 1 || (_Setting.OutRut == 2 && (_RoadGradeDict[prjinfo._RoadGrade] > 1)))
//                    {
//                        GlobalExcel.GetRutDisVal(prjinfo, prjdir, _RoadPart1M, ref _SRutDisVal, ref _SRutDisMile);
//                    }
//                    GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._DirectionInt, _RoadGradeDict, _SRutDisVal, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh, _RoadPart);
//                }
//                if (prjinfo._IsIRIMTD)
//                {
//                    if (IsSpeed)
//                    {
//                        GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPart, ref _SpeedVal);
//                    }
//                    if (IsMeanIRI)
//                    {
//                        IRIRes = GlobalExcel.GetIRIMeanVal(prjinfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, _Setting.IsWarning);
//                    }
//                    if (IsPBI)
//                    {
//                        GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 10, prjinfo._DirectionInt, _RoadGradeStr, ref _RoadPart10, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
//                       // GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart10, 0, ref _LDeltaHVal);
//                        //GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart, 0, ref _LDeltaHVal_1M);
//                        if (prjinfo._IsDIRIMTD)
//                        {
//                          //  GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart10, 1, ref _RDeltaHVal);
//                            //GlobalExcel.GetDeltaHVal(prjinfo, prjdir, _RoadPart, 1, ref _RDeltaHVal_1M);
//                        }
//                        PBIRes = GlobalExcel.GetPBVal(prjinfo, prjdir, _RoadPart, _RoadPart10, ref _PBIVal, _PBIThresh, _LDeltaHVal, _RDeltaHVal, 0, ref _DeltaHVal);
//                        GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart10, ref _MarkVal10);
//                        GlobalExcel.GetSpeedMeanVal(prjinfo, prjdir, _RoadPart10, ref _SpeedVal10);
//                    }
//                    if (IsMeanMTD && !_Setting.isGDIriCalculate)
//                    {
//                        MTDRes = GlobalExcel.GetMTDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, _Setting.IsWarning);

//                    }
//                    if (IsMeanMPD && !_Setting.isGDIriCalculate)
//                    {
//                        MPDRes = GlobalExcel.GetMPDMeanVal(prjinfo, prjdir, _RoadPart, ref _LMPDMeanVal, ref _RMPDMeanVal, ref _CMPDMeanVal, _Setting.IsWarning);

//                    }
//                }
//                else
//                {
//                    IRIRes = true;
//                }

//                if (prjinfo._IsRut)
//                {
//                    if (IsMeanRut)
//                    {
//                        RutRes = GlobalExcel.GetRutMeanVal(prjinfo, prjdir, _RoadPart, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, _Setting.IsWarning);
//                        MaxRutRes = GlobalExcel.GetRutMaxVal(prjinfo, prjdir, _RoadPart, ref _LRutMaxVal, ref _RRutMaxVal, ref _SRutMaxVal);
//                    }
//                    if (IsGeoAlig)
//                    {
//                        GeoAligRes = GlobalExcel.GetGeoAligVal(prjinfo, prjdir, _RoadPart, ref _Curvature, ref _CrossSlope, ref _HeightSlope, _Setting.IsWarning);
//                    }
//                }
//                else
//                {
//                    RutRes = true;
//                }

//                if (_Setting.ExcelType == 4 || _Setting.ExcelType == 18 || _Setting.ExcelType == 15) GPSRes = GlobalExcel.GetGPSInfo(prjinfo, prjdir, _RoadPart, ref _GPSInfo);

//                if (_RoadPart[0].roaddegree <= 1)
//                {
//                    return IRIRes && RutRes && MTDRes && GPSRes && MPDRes && GeoAligRes;
//                }
//                else
//                {
//                    return IRIRes && MPDRes && GeoAligRes;
//                }
//            }

            
        
//    }
//}
