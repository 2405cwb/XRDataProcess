 using Farmework.Office.Excel;
using Farmework.Other; 
using HNRoadFormatConverter.Commons;
using HNRoadFormatConverter.Entitys;
using HNRoadFormatConverter.Entitys.Excel; 
using System; 
using System.Collections.Generic; 
using System.IO;
using System.Linq; 
using System.Text; 
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Windows.Media;

namespace HNRoadFormatConverter.MyEntitys
{
    public class ProjectInfo
    {
        #region 工程参数
        private bool m_readExcelResult = true;
        /// <summary>
        /// 道路方向 A:上行 B：下行
        /// </summary>
        public string _Direction { get; set; }
        public int _DirectionInt { get; set; }

       public short _RoadType { get; set; }

        /// <summary>
        /// 表格数据方向 1 -从小到大  -1 -从大到小
        /// </summary>
        public int DirectionIntForm { get; set; }

        /// 
        /// <summary>
        /// 工程路径
        /// </summary>
        private string _PrjPath;

        public DirectoryInfo _DataDir;
        public string RoadNum { get; set; }

        public string _Province;//省

        public string _City;//市

        public string _District;//县

        private string prj;

        public string _RoadCode;//路线代码

        public string _RoadName;//路线名

        public int _StartMile;//起点桩号

        public string _RoadGrade;//公路等级

        public string _RoadNum;//车道

        public int _EndMile;//终点桩号

        public int _EndDmi;//总里程

        public string _DataDate;
        public string _DataTime;
        public string _DataPerson;
        public string _DataWeather;
        public string _RoadWidth;
        public string _RoadSurfaceName;
        public string _MeasureUnit;
        bool IsComplete = false;

        /// <summary>
        /// 县级行政区划代码
        /// </summary>
        public string _CityCode = "";//

        /// <summary>
        /// 转换中间文件夹
        /// </summary>
        public DirectoryInfo ConvertPath { get; set; }
        /// <summary>
        /// 转换工程名称
        /// </summary>
        public string ConvertProName { get; set; }
        /// <summary>
        /// 转换完成标志
        /// </summary>

        private int _DmiMileLen;//里程桩号关联数组个数
        private double[,] _DmiMile;//里程桩号关联数组
        private double[] _D2MScale;//里程/桩号的系数

        public bool _IsIRIMTD = false;//是否采集了平整度构造深度
        public bool _IsDIRIMTD = false;//是否是双平整度构造深度，true-双，false-单

        /// <summary>
        /// 是否采集了中间构造深度，true-是，false-否
        /// </summary>
        public bool _IsMMTD = false;

        /// <summary>
        /// 是否采集了车辙，true-是，false-否
        /// </summary>
        public bool _IsRut = false;
        /// <summary>
        /// 车辙模块模式，0-2D单车辙模块，1-2D双车辙模块，2-3D车辙模块
        /// </summary>
        public int _RutMode = 0;
        /// <summary>
        /// 几何线形工作模式，0-不采集几何线形数据，1-采集几何线形数据
        /// </summary>
        public int _GeoAlig = 0;

        public bool _IsDStreet = false;//是否是双景观，true-双，false-单
        public bool _IsRoad = false;//是否采集了路面
        public bool _IsStreet = false;//是否采集了景观
        public bool _IsPano = false;//是否采集了全景

        public int _RutDis = 50;//车辙出值间距
        public int _RoadImgDis = 2;//路面图像采集间距
        public int _StreetImgDis = 20;//景观图像采集间距

        public int _StreetImgDis_Left = 20;//景观图像采集间距
        public int _StreetImgDis_Right = 20;//景观图像采集间距
        public int _PanoImgDis = 20;//景观图像采集间距


        public bool _Is23dProject = false;

        /// <summary>
        /// 起始二三维工程的里程
        /// </summary>
        public int _23dStartDmi = 0;

        public double _DMIScale = 1.0; //编码器相关系数
        public bool ConvertOk
        {
            get;
            set;
        }
        public string ConvertSourcePath { get; set; }


        private StandardParmType m_standard;

        public StandardParmType Standard
        {
            get
            {
                return m_standard;
            }

        }


        private int m_drawType;

        public int DrawType
        {
            get
            { 
                return m_drawType;
            }
        }


        /// <summary>
        /// 23d工程报表文件夹
        /// </summary>
        public string Result3dExcelPath { get; set; }
        #endregion


        #region 工程数据获取

        public string DateDay { get; set; }
        public string DateMin { get; set; }
        public string RoadPicPath { get; set; }
        public string StreetPicPath { get; set; }
        private List<double[]> AnalysisResampleDatas(int side)
        {
            List<double[]> iriDatas  = new List<double[]>();
            string iriTxtPath = prj + $"\\IRIMTD\\DAQ{side}\\resample.txt";
            if (!File.Exists(iriTxtPath))
            {
                iriTxtPath = prj + $"\\IRIMTD\\DAQ{side}\\Resample.txt";
                if (!File.Exists(iriTxtPath))
                {
                    return  new List<double[]>();
                }
            } 
           List<string> txts =  File.ReadAllLines(iriTxtPath).ToList();

          
            for (int i = 0; i < txts.Count; i++)
            {
                string[] txtSplit = txts[i].Split('\t');
                double[] data = new double[txtSplit.Length];
                 
                for (int j = 0; j < txtSplit.Length; j++)
                {
                    data[j] = Convert.ToDouble(txtSplit[j]);
                }
                iriDatas.Add(data);
            }
            return iriDatas;
        }

        /// <summary>
        /// 获得左侧平整度原始数据
        /// </summary>
        /// <returns></returns>
        public List<double> getLeftResampleDatas()
        {  
            List<double> data = new List<double>();
            if (!_IsIRIMTD)
            {
                return data;
            }
            List<double[]> ResampleDatas =  AnalysisResampleDatas(0);

            for (int i = 0;i < ResampleDatas.Count;i++)
            {
                data.Add(ResampleDatas[i][2]);
            }
             return data;
        }

        /// <summary>
        /// 获得右侧平整度原始数据
        /// </summary>
        /// <returns></returns>
        public List<double> getRightResampleDatas()
        {
            List<double> data = new List<double>();
            if (!_IsDIRIMTD)
            {
                return data;
            }
            List<double[]> ResampleDatas = AnalysisResampleDatas(0);

            for (int i = 0;i < ResampleDatas.Count;i++)
            {
                data.Add(ResampleDatas[i][2]);
            }
            return data;
        }
        private List<double[]> AnalysisIriDatas(int side)
        {
            List<double[]> iriDatas = new List<double[]>();
            string iriTxtPath = prj + $"\\IRIMTD\\DAQ{side}\\resample.txt";
            if (!File.Exists(iriTxtPath))
            {
                iriTxtPath = prj + $"\\IRIMTD\\DAQ{side}\\Resample.txt";
                if (!File.Exists(iriTxtPath))
                {
                    return new List<double[]>();
                }
            }
            List<string> txts = File.ReadAllLines(iriTxtPath).ToList();


            for (int i = 0; i < txts.Count; i++)
            {
                string[] txtSplit = txts[i].Split('\t');
                double[] data = new double[txtSplit.Length];

                for (int j = 0; j < txtSplit.Length; j++)
                {
                    data[j] = Convert.ToDouble(txtSplit[j]);
                }
                iriDatas.Add(data);
            }
            return iriDatas;
        }

        /// <summary>
        /// 获取左平整度数据
        /// </summary>
        /// <returns></returns>
        public List<double> getLeftIriDatas()
        {
            List<double> data = new List<double>();
            if (!_IsIRIMTD)
            {
                return data;
            }
            List<double[]>  Datas = AnalysisIriDatas(0);

            for (int i = 0; i < Datas.Count; i++)
            {
                data.Add(Datas[i][1]);
            }
            return data;
        }

        /// <summary>
        /// 获取右侧平整度数据
        /// </summary>
        /// <returns></returns>
        public List<double> getRightIriDatas()
        {
            List<double> data = new List<double>();
            if (!_IsDIRIMTD)
            {
                return data;
            }
            List<double[]> Datas = AnalysisIriDatas(1);

            for (int i = 0; i < Datas.Count; i++)
            {
                data.Add(Datas[i][1]);
            }
            return data;
        }

        //获取病害信息
        public List<ExcelInportDisease> GetAllDisease()
        {
            List<ExcelInportDisease> datas = ExcelHelper_NPOI.ImportFromExcel<ExcelInportDisease>(
              Path.Combine
              (Result3dExcelPath,
              "__下行_1_贵州省_遵义市_红花岗区_20231122_100228_路面病害面积统计表_10m.xlsx"
              )
              , "病害列表", 2, false);

            return datas;
        }

        public static Dictionary<string, int> _RoadGradeDict;

        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };
        public List<MilePart> getMiles()
        {

            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }
            List<MilePart> parts = new List<MilePart>();

            MilePart spart = null;
            try
            {
                spart = new MilePart() { dmi = 0, roadtype =  _RoadType, mile =  _StartMile, roaddegree = _RoadGradeDict[ _RoadGrade], degreestr = _RoadGrade };
            }
            catch
            {
                MessageBox.Show(string.Format("【等级公路】不包含【{0}】请检查工程数据！", _RoadGrade));
                System.Environment.Exit(0);
            }
            parts.Add(spart);

        //    GlobalExcel.GetAllMilePart(prj, this, 10, _DirectionInt, _RoadGradeStr, ref parts, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);


            return parts; 
        }

        //public List<Unit> GetMiles()
        //{ 
        
                        
        //}

        #endregion

        /// <summary>
        /// 项目数据
        /// </summary>
        public ProjectInfo(bool readExcelResult, string projectPath)
        {
            this.m_readExcelResult = readExcelResult;
            this._PrjPath = projectPath;
            this.prj = projectPath;
            SetProjectInfo(projectPath);

        }

        /// <summary>
        /// 设置工程信息 
        /// 路线编号 上下行，道路起始及终点
        /// </summary>
        /// <param name="proPath"></param>
        private void SetProjectInfo(string proPath)
        {

            string proInfoTxtPath = Path.Combine(proPath, "ProjectInfo.txt");
            _DataDir = new DirectoryInfo(proPath);
            if (File.Exists(proInfoTxtPath))
            {
                //解析工程数据
                AnalysisProject(proInfoTxtPath);
                //转换中间文件夹
                GetConvertPath();
                //破损图片路径
                GetRoadPicPathAndCreatIndexFile();
                //景观图片路径
                GetStreetPicPathAndCreatIndexFile();
                //获得时间
                GetTime();



            }
            else
            {
                MessageBox.Show("缺少工程文件ProjectInfo.txt，请检查数据完整性");
                Application.Exit();
            }
        }


        private void SetPara(string[] strings)
        {
            foreach (string LineStr in strings)
            {
                var s = LineStr.Split('：');
                switch (s[0])
                {

                    case "省": _Province = s[1]; break;
                    case "市": _City = s[1]; break;
                    case "县": _District = s[1]; break;
                    case "工程起点道路编号": _RoadCode = s[1]; break;
                    case "工程起点道路名称": _RoadName = s[1]; break;
                    case "工程起点桩号":
                        _StartMile = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                     Convert.ToInt32(s[1].Split('+')[1]); break;
                    case "行车方向":
                        {
                            if (s[1] == "上行")
                            {
                                _DirectionInt = 1; _Direction = "A";
                            }

                            else if (s[1] == "下行")
                            {
                                _Direction = "B";
                                _DirectionInt = -1;
                            }

                        }
                        break;
                    case "公路等级": _RoadGrade = s[1]; break;
                    case "道路宽度": _RoadWidth = s[1]; break;

                    case "工程终点道路标识桩号":
                        _EndMile = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                                   Convert.ToInt32(s[1].Split('+')[1]); break;
                    case "路面材质":
                        {
                            _RoadSurfaceName = s[1];
                            try
                            {
                                _RoadType = (short)RoadDiseaseTypes.roadtypedict.Where(t => t.Key.Contains(s[1])).First().Value;
                            }
                            catch (Exception)
                            {

                                MessageBox.Show($"{_PrjPath}项目数据下projectinfo文件中路面材质列填写错误！");
                            }


                        }
                        break;
                    case "工程总里程数":
                        {
                            IsComplete = true;
                            _EndDmi = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                                      Convert.ToInt32(s[1].Split('+')[1]); break;

                        }
                    case "工程开始时刻": _DataTime = s[1]; DateMin = s[1]; break;
                    case "检测员": _DataPerson = s[1]; break;
                    case "操作员": _DataPerson = s[1]; break;
                    case "检测天气": _DataWeather = s[1]; break;
                    case "天气": _DataWeather = s[1]; break;
                    case "测量单位": _MeasureUnit = s[1]; break;
                    case "车道":
                        _RoadNum = s[1];
                        RoadNum = s[1];
                        break;

                    case "采集日期": _DataDate = s[1]; DateDay = s[1]; break;
                    case "工程日期": _DataDate = s[1]; DateDay = s[1]; break;

                    case "县级行政区划代码":
                        {
                            _CityCode = s[1];
                        }
                        break;
                }
            }
            if (!IsComplete)
            {
                if (Directory.Exists(prj + "\\RoadImg\\Camera0"))
                {
                    int imgnum = 0;
                    string[] dirs = Directory.GetDirectories(prj + "\\RoadImg\\Camera0");
                    Array.Sort(dirs);
                    foreach (string dir in dirs)
                    {
                        imgnum += Directory.GetFiles(dir, "*.jpg").Length;
                    }
                    _EndDmi = imgnum * _RoadImgDis;
                }
                else if (File.Exists(prj + "\\IRIMTD\\DAQ0\\Resample.txt"))
                {
                    string[] ttstr = File.ReadAllLines(prj + "\\IRIMTD\\DAQ0\\Resample.txt");
                    _EndDmi = (int)Math.Floor(ttstr.Length * 0.05);
                }
                else if (File.Exists(prj + "\\camera0\\rut.txt"))
                {
                    string[] ttstr = File.ReadAllLines(prj + "\\camera0\\rut.txt");
                    _EndDmi = (int)Math.Floor(ttstr.Length * _RutDis * 0.01);
                }
                else if (Directory.Exists(prj + "\\RoadImg\\Camera0"))
                {
                    int imgnum = 0;
                    string[] dirs = Directory.GetDirectories(prj + "\\StreetImg\\Camera0");
                    Array.Sort(dirs);
                    foreach (string dir in dirs)
                    {
                        imgnum += Directory.GetFiles(dir, "*.jpg").Length;
                    }
                    _EndDmi = imgnum * _StreetImgDis;
                }
                _EndMile = _StartMile + _DirectionInt * _EndDmi;
                string[] appstrs = new string[3];
                appstrs[0] = "工程终点道路标识桩号：" + _EndMile.ToString("K0000+000");
                appstrs[1] = "工程总里程数：" + _EndDmi.ToString("K0000+000");
                File.AppendAllLines(prj + @"\ProjectInfo.txt", appstrs, Encoding.UTF8);
            }
            TranDmi2Mile();
            string tmppath = _PrjPath + @"\ProjectInfo.txt";
            tmppath = _PrjPath + @"\Setting.ini";

            Encoding detectedEnc = EncodingDetector.GetType(tmppath);
            string text;
            // 用检测到的编码读取
            try
            {
                text = File.ReadAllText(tmppath, detectedEnc);
                Console.WriteLine($"检测到编码: {detectedEnc.EncodingName}");  // 调试输出
            }
            catch (Exception ex)
            {
                Console.WriteLine($"setting.ini文件读取失败: {ex.Message}");
                text = "";  // 或 fallback 到 UTF-8
            }  
            if (File.Exists(tmppath))
            {
                if (text.Contains("WorkMode"))
                {
                    IniFiles iniset = new IniFiles(tmppath);
                    _IsStreet = iniset.ReadBool("WorkMode", "Street", false);
                    _IsRoad = iniset.ReadBool("WorkMode", "Road", false);
                    _IsIRIMTD = iniset.ReadBool("WorkMode", "IRIMTD", false);
                    _IsDIRIMTD = iniset.ReadBool("WorkMode", "DIRIMTD", false);
                    _IsMMTD = iniset.ReadBool("WorkMode", "MMTD", false);

                    _GeoAlig = iniset.ReadInteger("WorkMode", "GeoAlig", 0);
                    _IsRut = iniset.ReadBool("WorkMode", "Rut", false);
                    _IsRut = iniset.ReadBool("WorkMode", "OnlyShowRut3d", false);
                    bool IsSRut = iniset.ReadBool("WorkMode", "SRut", false);
                    bool IsDRut = iniset.ReadBool("WorkMode", "DRut", false);
                    bool Is3DRut = iniset.ReadBool("WorkMode", "3DRut", false);
                    if (IsSRut)
                        _RutMode = 0;
                    if (IsDRut)
                        _RutMode = 1;
                    if (Is3DRut)
                        _RutMode = 2;

                    _IsDStreet = iniset.ReadBool("WorkMode", "DStreet", false);

                    _IsPano = iniset.ReadBool("WorkMode", "Pano", false);
                    _DMIScale = Convert.ToDouble(iniset.ReadString("WorkMode", "DMIScale", "1000")) * 0.001;
                    _RutDis = iniset.ReadInteger("Parm", "RUT_Dis", 50);
                    _RoadImgDis = iniset.ReadInteger("Parm", "RoadDis", 2);
                    _StreetImgDis = iniset.ReadInteger("Parm", "StreetDis", 20);
                    _StreetImgDis_Left = iniset.ReadInteger("Parm", "StreetDis", 20);
                    _StreetImgDis_Right = iniset.ReadInteger("Parm", "StreetDis2", 0);
                    if (_StreetImgDis_Right == 0)
                    {
                        _StreetImgDis_Right = _StreetImgDis_Left;
                    }
                    _PanoImgDis = iniset.ReadInteger("Parm", "PanoDis", 20);
                    _Is23dProject = iniset.ReadBool("WorkMode", "3dRoad", false);
                }
                else if (text.Contains("工作模式"))
                {
                    IniFiles iniset = new IniFiles(tmppath);
                    _IsStreet = iniset.ReadBool("工作模式", "Street", false);
                    _IsRoad = iniset.ReadBool("工作模式", "Road", false);
                    _IsIRIMTD = iniset.ReadBool("工作模式", "IRIMTD", false);
                    _IsDIRIMTD = iniset.ReadBool("工作模式", "DIRIMTD", false);
                    //_IsJgAndGd = iniset.ReadBool("工作模式", "IsIRIAndGd", false);

                    _IsMMTD = iniset.ReadBool("工作模式", "MMTD", false);

                    _GeoAlig = iniset.ReadInteger("工作模式", "GeoAlig", 0);
                    _IsRut = iniset.ReadBool("工作模式", "Rut", false);
                    bool IsSRut = iniset.ReadBool("工作模式", "SRut", false);
                    bool IsDRut = iniset.ReadBool("工作模式", "DRut", false);
                    bool Is3DRut = iniset.ReadBool("工作模式", "3DRut", false);
                    if (IsSRut)
                        _RutMode = 0;
                    if (IsDRut)
                        _RutMode = 1;
                    if (Is3DRut)
                        _RutMode = 2;

                    _IsDStreet = iniset.ReadBool("工作模式", "DStreet", false);
                    _DMIScale = Convert.ToDouble(iniset.ReadString("工作模式", "DMIScale", "1000")) * 0.001;
                    _RoadImgDis = iniset.ReadInteger("Parm", "RoadDis", 2);
                    _StreetImgDis = iniset.ReadInteger("Parm", "StreetDis", 20);
                    _StreetImgDis_Left = iniset.ReadInteger("Parm", "StreetDis", 20);
                   _StreetImgDis_Right = iniset.ReadInteger("Parm", "StreetDis2", 0);
                    if (_StreetImgDis_Right == 0)
                    {
                        _StreetImgDis_Right = _StreetImgDis_Left;
                    }
                    _PanoImgDis = iniset.ReadInteger("Parm", "PanoDis", 20);
                    _RutDis = iniset.ReadInteger("Parm", "RUT_Dis", 50);
                    _IsPano = iniset.ReadBool("工作模式", "Pano", false);
                    _Is23dProject = iniset.ReadBool("工作模式", "3dRoad", false);

                }
                else
                {
                    MessageBox.Show($"{_PrjPath}\\setting.ini文件读取失败:");
                }

                if (_Is23dProject)
                {
                    string dmi2mileFile = _PrjPath + @"\Dmi2Mile.txt";

                    if (File.Exists(dmi2mileFile))
                    {
                        string[] dmi2MileTxts = File.ReadAllLines(dmi2mileFile);
                        if (dmi2MileTxts.Length > 0)
                        {
                            int dmi = int.Parse(dmi2MileTxts[0].Split(' ')[0]);
                            
                            _23dStartDmi = dmi;
                        }
                    }
                }

            }
            else
            {
                MessageBox.Show(_PrjPath + "缺少工程配置文件Setting.ini，请检查数据是否完整！");
                Application.Exit();
            }
        }
         
        public void Analysis23DExcelData(IProgress<int> progress)
        { 
            //读取平整度数据
            var iriDatas = ExcelHelper_NPOI.ImportFromExcel<IRI>(
                Path.Combine
                (Result3dExcelPath,
                "__下行_1_贵州省_遵义市_红花岗区_20231122_100228_路面平整度评价等级记录表_10m.xlsx"
                )
                ,0,3,false,progress,0);
            if (iriDatas != null & iriDatas.Count>0 )
            {
                if (iriDatas[0].SMile < iriDatas[0].EMile)
                {
                    //从小到大
                    DirectionIntForm = 1;
                }
                else
                {
                    //从大到小
                    DirectionIntForm = -1;
                } 
            }
            var pciDatas = ExcelHelper_NPOI.ImportFromExcel<PCI>(
                Path.Combine
                (Result3dExcelPath,
                "__下行_1_贵州省_遵义市_红花岗区_20231122_100228_路面破损评价等级记录表_10m.xlsx"
                )
                , 0, 3, false, progress,100);


            //读取平整度原始数据

            //读取车辙数据





        }


        /// <summary>
        /// 解析工程数据
        /// </summary>
        /// <param name="proInfoTxtPath"></param>
        private void AnalysisProject(string proInfoTxtPath)
        {
            SetPara(File.ReadAllLines(proInfoTxtPath));
            

        }


        

        /// <summary>
        /// 图片及桩号结构体列表
        /// </summary>
        public List<PicAndMile> GetPicAndMiles(bool road, CityModelItem stanard)
        {
            List<PicAndMile> picAndMiles = new List<PicAndMile>();
            bool isNational2026 = stanard == CityModelItem.农养国省道路况检测数据提交格式_2026年;

            int GetImageIndexMile(int imageIndex, int imageInterval, out bool shouldStop)
            {
                // 图像 fileindex 的桩号按工程实际起点、终点和采集间隔生成，不使用 2Mile.txt 的校桩值。
                // 上行从起点按间隔递增到终点；下行从起点按间隔递减到终点。
                // 如果最后一步越过边界，则本张写边界值，后续图片不再导出。
                int direction = _DirectionInt == 0 ? 1 : _DirectionInt;
                int mile = _StartMile + direction * (imageIndex + 1) * imageInterval;

                shouldStop = false;
                if (direction > 0 && mile >= _EndMile)
                {
                    shouldStop = true;
                    return _EndMile;
                }

                if (direction < 0 && mile <= _EndMile)
                {
                    shouldStop = true;
                    return Math.Max(0, _EndMile);
                }

                if (mile < 0)
                {
                    shouldStop = true;
                    return 0;
                }

                return mile;
            }

            int GetNational2026BeforeCalibrationMile(int imageIndex, int imageInterval, out bool shouldStop)
            {
                // 2026 规范不生成 fileindex.txt，图片文件名本身必须写实际桩号。
                int direction = _DirectionInt == 0 ? 1 : _DirectionInt;
                int mile = _StartMile + direction * (imageIndex + 1) * imageInterval;

                shouldStop = false;
                if (direction > 0 && mile >= _EndMile)
                {
                    shouldStop = true;
                    return _EndMile;
                }

                if (direction < 0 && mile <= _EndMile)
                {
                    shouldStop = true;
                    return Math.Max(0, _EndMile);
                }

                if (mile < 0)
                {
                    shouldStop = true;
                    return 0;
                }

                return mile;
            }

            int GetNational2026AfterCalibrationMile(int startMileFrom2Mile, int imageInterval)
            {
                // Road2Mile.txt / Street2Mile.txt 记录的是图片起点桩号；2026 客户样例要求图片终点桩号。
                int direction = _DirectionInt == 0 ? 1 : _DirectionInt;
                int mile = startMileFrom2Mile + direction * imageInterval;

                if (direction > 0 && mile >= _EndMile)
                {
                    return _EndMile;
                }

                if (direction < 0 && mile <= _EndMile)
                {
                    return Math.Max(0, _EndMile);
                }

                if (mile < 0)
                {
                    return 0;
                }

                return mile;
            }

            int ConvertAbsoluteMileToRelativeAfterCalibration(int absoluteMile, int imageInterval)
            {
                int direction = _DirectionInt == 0 ? 1 : _DirectionInt;
                int relativeMile = (absoluteMile - _StartMile) * direction + imageInterval;
                return Math.Max(0, relativeMile);
            }

            string GetResultPicName(PicAndMile picAndMile)
            {
                if (stanard == CityModelItem.甘肃省单位一定制)
                {
                    var direction = ConvertProName.Last();
                    string RoadName = ConvertProName.Substring(0, ConvertProName.Length - 1);
                    return $"{RoadName}_{(picAndMile.Mile * 0.001).ToString("f3")}_{DateDay}";
                }
                else
                {
                    return DateDay + "_" + ConvertProName + "_" + (picAndMile.Mile * 0.001).ToString("f3");
                }
            }
            if (road)
            {


                string picBasePath = prj + "\\RoadImg\\Camera0";
                string indexTxt = picBasePath + "\\Road2Mile.txt";
                if (!File.Exists(indexTxt))
                {
                    MessageBox.Show(prj + "路面图像索引文件不存在请检查！");
                    System.Environment.Exit(0);
                }
                string[] strs = File.ReadAllLines(indexTxt);
                for (int i = 0; i < strs.Length; i++)
                {
                    var str = strs[i];
                    PicAndMile _picAndMile = new PicAndMile();
                    var sp = str.Split(' ');
                    int afterCalibrationMile = int.Parse(sp[0]);
                    bool shouldStop = false;
                    int imageMile = isNational2026
                        ? GetNational2026BeforeCalibrationMile(i, _RoadImgDis, out shouldStop)
                        : GetImageIndexMile(i, _RoadImgDis, out shouldStop);
                    _picAndMile.BeforeCalibrationMile = imageMile;
                    _picAndMile.AfterCalibrationMile = isNational2026
                        ? GetNational2026AfterCalibrationMile(afterCalibrationMile, _RoadImgDis)
                        : ConvertAbsoluteMileToRelativeAfterCalibration(afterCalibrationMile, _RoadImgDis);
                    _picAndMile.Mile = imageMile;
                    _picAndMile.PicPath = picBasePath + sp[1];
                    _picAndMile.sourceTxt = str;

                    _picAndMile.ResultPicName = GetResultPicName(_picAndMile);
                    picAndMiles.Add(_picAndMile);
                    if (shouldStop)
                    {
                        break;
                    }
                }

            }
            else
            { //路面图像
                string pciBasePath = prj + "\\StreetImg\\Camera0";
                string indexTxt = pciBasePath + "\\Street2Mile.txt";
                if (!File.Exists(indexTxt))
                {
                    MessageBox.Show(prj + "景观图像索引文件不存在请检查！");
                    System.Environment.Exit(0);
                }
                string[] strs = File.ReadAllLines(indexTxt);
                for (int i = 0; i < strs.Length; i++)
                {
                    var str = strs[i];
                    PicAndMile _picAndMile = new PicAndMile();
                    var sp = str.Split(' ');
                    int afterCalibrationMile = int.Parse(sp[0]);
                    bool shouldStop = false;
                    int imageMile = isNational2026
                        ? GetNational2026BeforeCalibrationMile(i, _StreetImgDis, out shouldStop)
                        : GetImageIndexMile(i, _StreetImgDis, out shouldStop);
                    _picAndMile.BeforeCalibrationMile = imageMile;
                    _picAndMile.AfterCalibrationMile = isNational2026
                        ? GetNational2026AfterCalibrationMile(afterCalibrationMile, _StreetImgDis)
                        : ConvertAbsoluteMileToRelativeAfterCalibration(afterCalibrationMile, _StreetImgDis);
                    _picAndMile.Mile = imageMile;
                    _picAndMile.PicPath = pciBasePath + sp[1];
                    _picAndMile.sourceTxt = str;
                    _picAndMile.ResultPicName = GetResultPicName(_picAndMile);
                    picAndMiles.Add(_picAndMile);
                    if (shouldStop)
                    {
                        break;
                    }
                }


            }
            return picAndMiles;

        }


        public List<PicAndMile> GetPicAndMilesHuNan(bool road)
        {
            List<PicAndMile> picAndMiles = new List<PicAndMile>();
            int GetImageIndexMile(int imageIndex, int imageInterval, out bool shouldStop)
            {
                // 湖南模板同样按工程实际起点、终点和采集间隔生成图片索引桩号。
                int direction = _DirectionInt == 0 ? 1 : _DirectionInt;
                int mile = _StartMile + direction * (imageIndex + 1) * imageInterval;

                shouldStop = false;
                if (direction > 0 && mile >= _EndMile)
                {
                    shouldStop = true;
                    return _EndMile;
                }

                if (direction < 0 && mile <= _EndMile)
                {
                    shouldStop = true;
                    return Math.Max(0, _EndMile);
                }

                if (mile < 0)
                {
                    shouldStop = true;
                    return 0;
                }

                return mile;
            }

            if (road)
            {


                string pciBasePath = prj + "\\RoadImg\\Camera0";
                string indexTxt = pciBasePath + "\\Road2Mile.txt";
                if (!File.Exists(indexTxt))
                {
                    MessageBox.Show(prj + "路面图像索引文件不存在请检查！");
                    System.Environment.Exit(0);
                }
                string[] strs = File.ReadAllLines(indexTxt);
                for (int i = 0; i < strs.Length; i++)
                {
                    var str = strs[i];
                    PicAndMile _picAndMile = new PicAndMile();
                    var sp = str.Split(' ');
                    bool shouldStop;
                    _picAndMile.Mile = GetImageIndexMile(i, _RoadImgDis, out shouldStop);
                    _picAndMile.PicPath = pciBasePath + sp[1];
                    _picAndMile.sourceTxt = str;


                    string mile = Form1.ConvertIntToFormattedString(_picAndMile.Mile);
                    _picAndMile.ResultPicName = ConvertProName + "-" + mile + "-" + mile;
                    picAndMiles.Add(_picAndMile);
                    if (shouldStop)
                    {
                        break;
                    }
                }

            }
            else
            { //路面图像
                string pciBasePath = prj + "\\StreetImg\\Camera0";
                string indexTxt = pciBasePath + "\\Street2Mile.txt";
                if (!File.Exists(indexTxt))
                {
                    MessageBox.Show(prj + "景观图像索引文件不存在请检查！");
                    System.Environment.Exit(0);
                }
                string[] strs = File.ReadAllLines(indexTxt);
                for (int i = 0; i < strs.Length; i++)
                {
                    var str = strs[i];
                    PicAndMile _picAndMile = new PicAndMile();
                    var sp = str.Split(' ');
                    bool shouldStop;
                    _picAndMile.Mile = GetImageIndexMile(i, _StreetImgDis, out shouldStop);
                    _picAndMile.PicPath = pciBasePath + sp[1];
                    _picAndMile.sourceTxt = str;
                    string mile = Form1.ConvertIntToFormattedString(_picAndMile.Mile);
                    _picAndMile.ResultPicName = ConvertProName + "-" + mile + "-" + mile;
                    // _picAndMile.ResultPicName = DateDay + "_" + ConvertProName + "_" + (_picAndMile.Mile * 0.001).ToString("f3");
                    picAndMiles.Add(_picAndMile);
                    if (shouldStop)
                    {
                        break;
                    }
                }


            }
            return picAndMiles;

        }
        private (StandardParmType, int) getStandardAndDrawType(string result23dExcelPath)
        {
            StandardParmType ParmStyle = StandardParmType.DegreeRoad2018;
            int SelectDrawDis = 0;
            string reslutFolder = result23dExcelPath;
            string proName = result23dExcelPath;
            DirectoryInfo resultDir = new DirectoryInfo(reslutFolder);
            if (resultDir.GetDirectories().Length == 1)
            {
                string standardFolder = resultDir.GetDirectories()[0].Name;

                switch (standardFolder)
                {
                    case "等级公路 JTG H20-2018":
                        {
                            ParmStyle = StandardParmType.DegreeRoad2018;

                        }
                        break;
                    case "低等级农村公路":
                        {
                            ParmStyle = StandardParmType.RuralRoadlowLevel;

                        }
                        break;
                    default:
                        break;
                }
                DirectoryInfo standardDir = new DirectoryInfo(resultDir.GetDirectories()[0].FullName);


                if (standardDir.GetDirectories().Length == 1)
                {
                    string drawTypeFolder = standardDir.GetDirectories()[0].Name;

                    switch (drawTypeFolder)
                    {
                        case "自动化模式":
                            SelectDrawDis = 1;
                            break;
                        case "人工模式":
                            SelectDrawDis = 0;
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    MessageBox.Show($"{0}位置未检测到合法的结果表格，请检查！", proName);
                    Application.Exit();
                }
            }
            else
            {
                MessageBox.Show($"{0}位置未检测到合法的结果表格，请检查！", proName);
                Application.Exit();
            }

            return (ParmStyle, SelectDrawDis);


        }
        private void GetConvertPath ()
        {
            if (m_readExcelResult)
            {
                string[] paths = Directory.GetDirectories(this._PrjPath, "结果表格");
                if (paths.Length == 0)
                {
                    MessageBox.Show($"{this._PrjPath}不具有二三维中间结果。\n请先使用[二三维内业数据处理软件/数据输出/计算国检中间数据]功能");
                    Application.Exit();
                }
                Result3dExcelPath = paths.First();

                (m_standard, m_drawType) = getStandardAndDrawType(Result3dExcelPath);

                Result3dExcelPath = Path.Combine(Result3dExcelPath, m_standard.ToString() == "DegreeRoad2018" ? "等级公路 JTG H20-2018" : "低等级农村公路",
                    m_drawType == 1 ? "自动化模式" : "人工模式");
            }
            else
            {
                string[] paths = Directory.GetDirectories(this._PrjPath, "ConverSource");
                if (paths.Length == 0)
                {
                    MessageBox.Show($"{this._PrjPath}不具有转换中间数据文件夹ConverSource请检查");
                    Application.Exit();
                }
                string path = paths[0];
                DirectoryInfo temp = new DirectoryInfo(path);
                path = temp.GetDirectories().First().FullName;
                ConvertSourcePath = path;
                ConvertProName = Path.GetFileNameWithoutExtension(path);
                DirectoryInfo directory = new DirectoryInfo(path);
                ConvertPath = directory;
            }
        }
   


        private void Image2Mile(string ImgSource, int CamIdx, int ImgDis, string imgpath, string folderLastName = "Img")
        {
            //return;

            /////////////////////////////////////////////////////
            if (!Directory.Exists(string.Format("{0}\\{1}{3}\\Camera{2}", _DataDir.FullName, ImgSource, CamIdx, folderLastName)))
            {
                return;
            }

            //获取所有的图像名
            string fname, subfolder;
            List<string> ImgsList = new List<string>();
            int dirnum =  _EndDmi / ImgDis / 1000 + 1;
            int imgcnt = 0;//数的图像张数，会丢帧
            int imgtrigcnt = 0;//触发计数，从开始
            string tstr;
            string[] tstrs;
            for (int i = 0; i < dirnum; ++i)
            {
                string dirname = string.Format("{0}\\{1}{4}\\Camera{2}\\Image_{3:0000}", _DataDir.FullName, ImgSource, CamIdx, i, folderLastName);
                try
                {
                    if (Directory.Exists(dirname))
                    {
                        subfolder = Path.GetFileName(dirname);
                        string[] imgsname = Directory.GetFiles(dirname, "*." + imgpath);
                        Array.Sort(imgsname);
                        foreach (string imgname in imgsname)
                        {
                            fname = Path.GetFileName(imgname);
                            tstr = string.Format("\\{0}\\{1}", subfolder, fname);

                            tstrs = tstr.Split('_');
                            imgtrigcnt = int.Parse(tstrs[1].Replace("\\", ""));

                            //图像数据丢帧，用上一张替代
                            while (imgtrigcnt > imgcnt)
                            {
                                ImgsList.Add(tstr);
                                imgcnt++;
                            }
                            if (imgtrigcnt == imgcnt)
                            {
                                ImgsList.Add(tstr);
                                imgcnt++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // MessageBox.Show(string.Format("打开文件【{0}】出错！！\r\n{1}", dirname, ex.Message));
                    MessageBox.Show(string.Format("打开文件【{0}】出错！！\r\n{1}\r\n{2}", dirname, ex.Message, imgcnt));
                    return;
                }
            }

            //桩号 图像 要处理的数据存盘
            int tdmi = 0, tmile = 0;
            int imgnum = ImgsList.Count;
            if (imgnum < 1)
                return;
            //TODO 读取trigger.txt文件  来获取图像间隔  
            FileStream fw = new FileStream(string.Format("{0}\\{1}{3}\\Camera{2}\\{1}2Mile.txt", _DataDir.FullName, ImgSource, CamIdx, folderLastName), FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < imgnum; ++i)
            {
                tmile =  Dmi2Mile(tdmi);
                if (tmile <= 0 && _DirectionInt < 0
                    || tmile >=  _EndMile && _DirectionInt > 0)
                {
                    break;
                }
                sw.WriteLine(string.Format("{0} {1}", tmile, ImgsList[i]), Encoding.UTF8);
                tdmi = tdmi + ImgDis;
            }
             
            sw.Close();
            fw.Close();
        }


        private void GetRoadPicPathAndCreatIndexFile()
        {
             if (Directory.Exists(prj + "\\RoadImg\\Camera0"))
            {
                RoadPicPath = prj + "\\RoadImg\\Camera0";
            }

            string txt = RoadPicPath + "\\Road2Mile.txt";
            if (!File.Exists(txt))
            {
                Image2Mile("Road", 0, _RoadImgDis, "jpg");

            } 
        }
        private void GetStreetPicPathAndCreatIndexFile()
        {
            if (Directory.Exists(prj + "\\StreetImg\\Camera0"))
            {
                StreetPicPath = prj + "\\StreetImg\\Camera0";
            }
            string txt = StreetPicPath + "\\Street2Mile.txt";
            if (!File.Exists(txt))
            {
                Image2Mile("Street", 0, _StreetImgDis, "jpg");


            }
        }
        /// <summary>
        /// 获得时间
        /// </summary>
        private void GetTime()
        {
          /*  DirectoryInfo dr = new DirectoryInfo(ConvertSourcePath);
           DirectoryInfo info =  dr.GetDirectories("景观图像", SearchOption.AllDirectories).First();
            var timePath = info.GetDirectories().First();
            DateDay = timePath.Name;
            DateMin = timePath.GetDirectories().First().Name.Split('_').LastOrDefault();
          */



        }
        

         
        /// <summary>
        /// 将编码器的 里程值 转换成路边 桩号
        /// </summary>
        /// <param name="prj"></param>
        private void TranDmi2Mile()
        { 
            List<DmiMile> ListDM = new List<DmiMile>();
            string fname = _PrjPath + @"\MileStoneCaliInfo.txt";
            ListDM.Add(new DmiMile(0, _StartMile));
            if (File.Exists(fname))
            {
                string[] sinfo = File.ReadAllLines(fname);
                foreach (string s in sinfo)
                {
                    string[] str = s.Split(' ');
                    if (str.Length > 1)
                    {

                        int dmi = int.Parse(str[0]);
                        int mile = int.Parse(str[1]);

                        bool isexist = false;
                        for (int ttt = 0; ttt < ListDM.Count; ttt++)
                        {
                            var curDmiMile = ListDM[ttt];
                            if (Math.Abs( mile - curDmiMile._Mile) <= 1)
                            {
                                //以及存在了
                                isexist = true;
                                break;
                            }
                        }
                        if (!isexist)
                            ListDM.Add(new DmiMile(dmi, mile));
                    }
                }

                 
            }
            else
            {
                int temp = _StartMile + _EndDmi * _DirectionInt;
                if (temp < 0 && _EndMile > 0)
                {
                    _EndMile = 0;
                    _EndDmi = _StartMile;
                    string[] sinfo = File.ReadAllLines(_PrjPath + @"\ProjectInfo.txt");
                    for (int i = 0; i < sinfo.Length; ++i)
                    {
                        //if (sinfo[i].Contains("工程终点道路实际桩号："))
                        //{
                        //    sinfo[i] = "工程终点道路实际桩号：" + _EndMile.ToString("K0000+000");
                        //}
                        //else 
                        if (sinfo[i].Contains("工程终点道路标识桩号："))
                        {
                            sinfo[i] = "工程终点道路标识桩号：" + _EndMile.ToString("K0000+000");
                        }
                        else if (sinfo[i].Contains("工程总里程数："))
                        {
                            sinfo[i] = "工程总里程数：" + _EndDmi.ToString("K0000+000");
                        }
                    }
                    File.WriteAllLines(_PrjPath + @"\ProjectInfo.txt", sinfo, Encoding.UTF8);
                }
            }

            ListDM.Add(new DmiMile(_EndDmi, _EndMile));
            int dmlen = ListDM.Count;
            for (int i = 0; i < dmlen - 1; ++i)
            {
                if (Math.Abs( ListDM[i]._Mile - ListDM[i + 1]._Mile) <= 1)
                {
                    ListDM.RemoveAt(i);
                    dmlen--;
                    i--;
                }
            }

            if (_DirectionInt > 0)//升序
            {
                ListDM.Sort(delegate (DmiMile x, DmiMile y) { return x._Mile.CompareTo(y._Mile); });
            }
            else if (_DirectionInt < 0)//降序
            {
                ListDM.Sort(delegate (DmiMile x, DmiMile y) { return y._Mile.CompareTo(x._Mile); });
            }

            _DmiMileLen = ListDM.Count;
            _DmiMile = new double[_DmiMileLen, 2];//里程，桩号
            _D2MScale = new double[_DmiMileLen];
            _D2MScale[0] = _DirectionInt;


          //  FileStream fw = new FileStream(_PrjPath + @"\Dmi2Mile.txt", FileMode.Create);
          //  StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < _DmiMileLen; ++i)
            {
                _DmiMile[i, 0] = ListDM[i]._Dmi;
                _DmiMile[i, 1] = ListDM[i]._Mile;
                if (i > 0) _D2MScale[i] = (_DmiMile[i, 0] - _DmiMile[i - 1, 0]) / (_DmiMile[i, 1] - _DmiMile[i - 1, 1]);
                //sw.WriteLine(string.Format("{0} {1}", ListDM[i]._Dmi, ListDM[i]._Mile), Encoding.UTF8);
            }
           // sw.Close();
            //fw.Close();

            string tfname = _PrjPath + "\\RoadStatuMarkInfo.txt";
            if (File.Exists(tfname))
            {
                List<string> strlist = new List<string>();
                string[] strs = File.ReadAllLines(tfname, Encoding.UTF8);
                for (int i = 0; i < strs.Length; ++i)
                {
                    strs[i] = strs[i].Replace("K", "");
                    strs[i] = strs[i].Replace("k", "");
                    strs[i] = strs[i].Replace("+", "");
                    string[] sstrs = strs[i].Split(' ');
                    if (sstrs.Length < 2)
                        continue;
                    int tdmi = (int)Convert.ToDouble(sstrs[2]);
                    //int tmile = (int)Convert.ToDouble(sstrs[0]);
                    int tmile = (int)Dmi2Mile(tdmi);
                    for (int j = 0; j < ListDM.Count - 1; ++j)
                    {
                        if (ListDM[j]._Dmi <= tdmi && ListDM[j + 1]._Dmi > tdmi)
                        {
                            strs[i] = string.Format("{0} {0} {1} {2}", tmile, tdmi, sstrs[sstrs.Length - 1]);
                            break;
                        }
                    }
                    strlist.Add(strs[i]);
                }
               // File.WriteAllLines(tfname, strlist, Encoding.UTF8);
            }
        }

         
       

        /// <summary>
        /// 桩号转换为里程
        /// </summary>
        /// <param name="inmile">桩号</param>
        /// <returns>里程</returns>
        public int Mile2Dmi(int inmile)
        {
            int j = 0;
            for (j = 1; j < _DmiMileLen; ++j)
            {
                if (_DirectionInt > 0)
                {
                    if ((inmile >= _DmiMile[j - 1, 1] && inmile <= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        return Math.Abs((int)Math.Round((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]));
                    }
                }
                else
                {
                    if ((inmile <= _DmiMile[j - 1, 1] && inmile >= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        return Math.Abs((int)Math.Round((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]));
                    }
                }
            }
            return 0;
        }
        public double Mile2Dmi(double inmile)
        {
            int j = 0;
            for (j = 1; j < _DmiMileLen; ++j)
            {
                if (_DirectionInt > 0)
                {
                    if ((inmile >= _DmiMile[j - 1, 1] && inmile <= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        var temp = Math.Abs((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]);
                        return (double)Math.Round(temp, 1);
                    }
                }
                else
                {
                    if ((inmile <= _DmiMile[j - 1, 1] && inmile >= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        return (double)Math.Round(Math.Abs((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]), 1);
                    }
                }
            }
            return 0;
        }
        /// <summary>
        /// 里程转换为桩号
        /// </summary>
        /// <param name="indmi">里程</param>
        /// <returns>桩号</returns>
        public int Dmi2Mile(float indmi)
        {
            for (int j = 1; j < _DmiMileLen; ++j)
            {
                if ((indmi >= _DmiMile[j - 1, 0] && indmi <= _DmiMile[j, 0]) || j == _DmiMileLen - 1)
                {
                    //return Math.Abs((int)Math.Round((indmi - _DmiMile[j - 1, 0]) / _D2MScale[j] + _DmiMile[j - 1, 1]));                    
                    return (int)Math.Round((indmi - _DmiMile[j - 1, 0]) / _D2MScale[j] + _DmiMile[j - 1, 1]);
                }
            }
            return 0;
        }
    }
}
