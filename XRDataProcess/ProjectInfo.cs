using DevExpress.XtraCharts;
using Farmework.Other;
using Framework.Other;
using OpenTK.Graphics.OpenGL;
using OperateIniFile;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UtfUnknown;
using XRDataProcess.Properties;
namespace XRDataProcess
{
    public class ProjectInfo
    {
        /// <summary>
        ///  省
        /// </summary>
        public string _Province;
        
        /// <summary>
        ///市 
        /// </summary>
        public string _City;
       
        /// <summary>
        /// 县
        /// </summary>
        public string _District;
        /// <summary>
        /// 县级行政区划代码
        /// </summary>
        public string _CityCode = "";//
        //
        /// <summary>
        /// 区划代码
        /// </summary>
        public string _AreaCode = "";
        /// <summary>
        /// 路线代码
        /// </summary>
        public string _RoadCode;//
        //
        /// <summary>
        /// 路线代码前四位
        /// </summary>
        public string _RoadCodePart;
        /// <summary>
        /// 路线名
        /// </summary>
        public string _RoadName;//
        /// <summary>
        /// 单元编号
        /// </summary>
        public int _UnitNum = 10000;//
        /// <summary>
        /// 公路等级
        /// </summary>
        public string _RoadGrade;//

        /// <summary>
        /// 车道
        /// </summary>
        public string _RoadNum;//

        /// <summary>
        /// 路面材质，0--沥青，1--水泥，2-砂石
        /// </summary>
        public short _RoadType;

        /// <summary>
        /// 工程文件中是否已明确指定起始路面材质。不能以 _RoadType 的默认值判断，
        /// 因为 0 同时也是“沥青”的有效值。
        /// </summary>
        public bool _HasInitialRoadType { get; private set; }

        public string _DataDate;

        public string _DataTime;

        public string _DataPerson;
        public string _DataWeather;

        //老设备 20220815cwb
        public float _PlusLength;
        public int DAQSampleFrequency = 8000;

        public int _StartMile;//起点桩号
        /// <summary>
        /// 行车方向，1--上行，-1--下行
        /// </summary>
        public int _Direction;
        public int _EndMile;//终点桩号
        public int _EndDmi;//总里程

        private int _DmiMileLen;//里程桩号关联数组个数
        private double[,] _DmiMile;//里程桩号关联数组
        private double[] _D2MScale;//里程/桩号的系数
        // 三维设备先启动时，采集打标使用的 DMI 可能整体早于二维工程起点。
        // 仅在本次导入中使用，校桩和打标文件落盘后立即清零，避免重复修正。
        private int _Pending3DStartupDmiOffset;

        /// <summary>
        /// 是否采集了平整度构造深度
        /// </summary>
        public bool _IsIRIMTD = false;


        /// <summary>
        /// 是否是双平整度构造深度，true-双，false-单
        /// </summary>
        public bool _IsDIRIMTD = false;


        /// <summary>
        /// 湖南定制设备 激光(/IRIMTD0)+惯导平整度
        /// </summary>
        public bool _IsJgAndGd = false;

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
        public int _StreetImgDis_Left = 20;//景观图像采集间距
        public int _StreetImgDis_Right = 20;//景观图像采集间距
        public int _PanoImgDis = 20;//景观图像采集间距
        public double _DMIScale = 1.0; //编码器相关系数

        /// <summary>
        /// 路面打标信息
        /// </summary>
        public List<MarkInfo> _MarkInfo = new List<MarkInfo>();

        /// <summary>
        /// 工程路径
        /// </summary>
        public string _PrjPath;


        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();

        /// <summary>
        /// 加载工程信息
        /// </summary>
        /// <param name="prj"></param>
        public ProjectInfo(string prj)
        {
            _PrjPath = prj;
            _PlusLength = 0;
            string tmppath = _PrjPath + @"\ProjectInfo.txt";
            if (!File.Exists(tmppath))
            {
                MessageBox.Show(_PrjPath + "缺少工程文件ProjectInfo.txt，请检查数据是否完整！");
                Application.Exit();
            }
            Encoding encoding;

            encoding = OtherHelper.GetFileEncodeType(tmppath);
            string text;
            /* if (encoding != Encoding.UTF8)
             {
                
             }*/
            text = File.ReadAllText(tmppath, Encoding.GetEncoding("gb2312"));
            if ( text.Contains("工程"))
            {
                //将文件以gb-2312保存
                File.WriteAllText(tmppath, text, Encoding.UTF8);
            }

            string[] sinfo = File.ReadAllLines(tmppath);
            string[] s;
            bool IsComplete = false;
             
         
            foreach (string str in sinfo)
            {
                s = str.Split('：',':');
                switch (s[0])
                {
                    case "省": _Province = s[1]; break;
                    case "市": _City = s[1]; break;
                    case "县": _District = s[1]; break;
                    case "工程起点道路编号":
                        {
                            _RoadCode = s[1];
                            if (_RoadCode.Length >= 4)
                            {
                                _RoadCodePart = _RoadCode.Substring(0, 4);
                                _AreaCode = "";
                            }

                            else if (_RoadCode.Length >= 10)
                            {
                                _RoadCodePart = _RoadCode.Substring(0, 4);
                                _AreaCode = _RoadCode.Substring(4, 6);
                            }

                        }
                        break;
                    case "工程起点道路名称": _RoadName = s[1]; break;
                    case "工程起点桩号":
                        _StartMile = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                                Convert.ToInt32(s[1].Split('+')[1]); break;
                    case "行车方向":
                        {
                            if (s[1] == "上行") _Direction = 1;
                            else if (s[1] == "下行") _Direction = -1;
                        }
                        break;
                    case "公路等级": _RoadGrade = s[1]; break;
                    case "车道": _RoadNum = s[1]; break;
                    case "采集日期": _DataDate = s[1]; break;
                    case "工程开始时刻": _DataTime = s[1]; break;
                    case "检测员": _DataPerson = s[1]; break;
                    case "检测天气": _DataWeather = s[1]; break;
                    case "编码器分频后脉冲距离":
                        {
                            //说明是老设备
                            _PlusLength = float.Parse(s[1]);

                            break;
                        }
                    case "采集卡采样率：":
                        DAQSampleFrequency = int.Parse(s[1]);
                        break;
                    case "路面材质":
                        {
                            try
                            {
                                _RoadType = (short)RoadDiseaseTypes.roadtypedict.Where(t => t.Key.Contains(s[1])).First().Value;
                                _HasInitialRoadType = true;
                            }
                            catch (Exception)
                            {
                                _HasInitialRoadType = false;
                                MessageBox.Show($"{_PrjPath}项目数据下ProjectInfo.txt的【路面材质】填写错误，请在工程信息中重新选择后再绘制病害！");
                            }


                        }
                        break;
                    case "工程终点道路标识桩号":
                        _EndMile = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                         Convert.ToInt32(s[1].Split('+')[1]); break;
                    case "工程总里程数":
                        {
                            IsComplete = true;
                            _EndDmi = Convert.ToInt32(s[1].Split('+')[0].Replace("K", "")) * 1000 +
                            Convert.ToInt32(s[1].Split('+')[1]); break;

                        }
                    case "县级行政区划代码":
                        {
                            _CityCode = s[1];
                            break;
                        }
                     
                }
            }
            //   _RoadNum = _RoadCode + "_" + _RoadName+"_"+ _RoadNum;
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
                    _EndDmi = imgnum * _StreetImgDis_Left;
                }
                _EndMile = _StartMile + _Direction * _EndDmi;
                string[] appstrs = new string[2];
                appstrs[0] = "工程终点道路标识桩号：" + _EndMile.ToString("K0000+000");
                appstrs[1] = "工程总里程数：" + _EndDmi.ToString("K0000+000");
                File.AppendAllLines(_PrjPath + @"\ProjectInfo.txt", appstrs, Encoding.UTF8);


            }
            {
                // 1. 准备要更新/添加的数据字典
                Dictionary<string, string> updateDict = new Dictionary<string, string>
                {
                 { "道路类型", Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle].ToString() },
                 { "道路宽度", _RoadConfig.DetectWidth.ToString() },
                 { "病害绘制模式", _Setting.SelectDrawDis.ToString() }
                };
                 
                List<string> lines = File.ReadAllLines(tmppath, Encoding.UTF8).ToList();
                 
                HashSet<string> updatedKeys = new HashSet<string>();
                 
                for (int i = 0; i < lines.Count; i++)
                {
                    string currentLine = lines[i].Trim();
                    foreach (var key in updateDict.Keys)
                    {
                        // 匹配“键：”或“键:”
                        if (currentLine.StartsWith(key + "：") || currentLine.StartsWith(key + ":"))
                        {
                            lines[i] = $"{key}：{updateDict[key]}";
                            updatedKeys.Add(key);
                            break;
                        }
                    }
                } 
                foreach (var item in updateDict)
                {
                    if (!updatedKeys.Contains(item.Key))
                    {
                        lines.Add($"{item.Key}：{item.Value}");
                    }
                }
                 
                File.WriteAllLines(tmppath, lines, Encoding.UTF8);
            }
            // 三维设备可能比二维设备提前启动。必须在生成 Dmi2Mile.txt 前完成一次性归一化。
            TryCorrect3DStartupDmiOffset();

            //补充写入工程信息
            TranDmi2Mile(); 
            tmppath = _PrjPath + @"\Setting.ini";
            if (File.Exists(tmppath))
            { 
                ////DetectionResult result = CharsetDetector.DetectFromFile(tmppath);
                ////Encoding defaultEncoding = result.Detected.Encoding; // 自动识别UTF-8/GB2312/GBK等
                ////text = File.ReadAllText(tmppath, encoding);
                Encoding detectedEnc = EncodingDetector.GetType(tmppath);
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
                    _StreetImgDis_Left = iniset.ReadInteger("Parm", "StreetDis", 20);
                    _StreetImgDis_Right = iniset.ReadInteger("Parm", "StreetDis2", 0);
                    if (_StreetImgDis_Right == 0)
                    {
                        _StreetImgDis_Right = _StreetImgDis_Left;
                    }
                    _PanoImgDis = iniset.ReadInteger("Parm", "PanoDis", 20);
               
                }
                else if(text.Contains("工作模式"))
                {
                    IniFiles iniset = new IniFiles(tmppath);
                    _IsStreet = iniset.ReadBool("工作模式", "Street", false);
                    _IsRoad = iniset.ReadBool("工作模式", "Road", false);
                    _IsIRIMTD = iniset.ReadBool("工作模式", "IRIMTD", false);
                    _IsDIRIMTD = iniset.ReadBool("工作模式", "DIRIMTD", false);
                    _IsJgAndGd = iniset.ReadBool("工作模式", "IsIRIAndGd", false);

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
                    _StreetImgDis_Left = iniset.ReadInteger("Parm", "StreetDis", 20);
                    _StreetImgDis_Right = iniset.ReadInteger("Parm", "StreetDis2", 0);
                    if (_StreetImgDis_Right == 0)
                    {
                        _StreetImgDis_Right = _StreetImgDis_Left;
                    }
                    _PanoImgDis = iniset.ReadInteger("Parm", "PanoDis", 20);
                    _RutDis = iniset.ReadInteger("Parm", "RUT_Dis", 50);
                    _IsPano = iniset.ReadBool("工作模式", "Pano", false);
                }
                else
                {
                    MessageBox.Show($"{_PrjPath}\\setting.ini文件读取失败:");
                }
            }
            else
            {
                MessageBox.Show(_PrjPath + "缺少工程配置文件Setting.ini，请检查数据是否完整！");
                Application.Exit();
            }
        }

        public void SavePrjInfo()
        {
            string projectInfoFile = Path.Combine(_PrjPath, "ProjectInfo.txt");
            string pavement = _RoadType == 0 ? "沥青" : _RoadType == 1 ? "水泥" : "砂石";
            Dictionary<string, string> values = new Dictionary<string, string>
            {
                { "省", _Province }, { "市", _City }, { "县", _District },
                { "工程起点道路编号", _RoadCode }, { "工程起点道路名称", _RoadName },
                { "工程起点桩号", _StartMile.ToString("K0000+000") },
                { "行车方向", _Direction > 0 ? "上行" : "下行" }, { "公路等级", _RoadGrade },
                { "车道", _RoadNum }, { "采集日期", _DataDate }, { "工程开始时刻", _DataTime },
                { "检测员", _DataPerson }, { "检测天气", _DataWeather }, { "路面材质", pavement },
                { "工程终点道路标识桩号", _EndMile.ToString("K0000+000") },
                { "工程总里程数", _EndDmi.ToString("K0000+000") }
            };
            List<string> lines = File.ReadAllLines(projectInfoFile, Encoding.UTF8).ToList();
            HashSet<string> written = new HashSet<string>();
            for (int i = 0; i < lines.Count; i++)
            {
                int separator = lines[i].IndexOfAny(new[] { '：', ':' });
                if (separator < 0) continue;
                string key = lines[i].Substring(0, separator).Trim();
                string value;
                if (values.TryGetValue(key, out value))
                {
                    lines[i] = key + "：" + (value ?? string.Empty);
                    written.Add(key);
                }
            }
            // 旧工程可能没有这一行；保存时补齐，之后病害绘制才有明确的起始材质。
            foreach (KeyValuePair<string, string> item in values)
            {
                if (!written.Contains(item.Key)) lines.Add(item.Key + "：" + (item.Value ?? string.Empty));
            }
            WriteAllLinesAtomically(projectInfoFile, lines);
            _HasInitialRoadType = true;
            TranDmi2Mile(); // 同步 Dmi2Mile.txt 以及打标文件中由 DMI 推导出的桩号。
        }

        private static void WriteAllLinesAtomically(string path, IEnumerable<string> lines)
        {
            string temporaryPath = path + ".tmp";
            File.WriteAllLines(temporaryPath, lines, new UTF8Encoding(false));
            if (File.Exists(path))
                File.Replace(temporaryPath, path, path + ".bak", true);
            else
                File.Move(temporaryPath, path);
        }

        /// <summary>
        /// 识别三维设备先于二维设备启动造成的固定 DMI 偏移，并将工程文件归一化到二维 DMI。
        /// 只有全部校验通过才会写入文件，防止把比例误差或异常校桩误当作固定偏移。
        /// </summary>
        private void TryCorrect3DStartupDmiOffset()
        {
            string settingFile = Path.Combine(_PrjPath, "Setting.ini");
            string caliFile = Path.Combine(_PrjPath, "MileStoneCaliInfo.txt");
            if (!File.Exists(settingFile) || !File.Exists(caliFile))
            {
                return;
            }

            bool is3DRoad;
            try
            {
                is3DRoad = new IniFiles(settingFile).ReadBool("WorkMode", "3dRoad", false);
            }
            catch
            {
                return;
            }
            if (!is3DRoad)
            {
                return;
            }

            List<DmiMile> calibration = new List<DmiMile>();
            try
            {
                foreach (string line in File.ReadAllLines(caliFile))
                {
                    string[] values = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (values.Length < 2)
                    {
                        continue;
                    }
                    calibration.Add(new DmiMile(int.Parse(values[0]), int.Parse(values[1])));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("检测到三维数据，但无法读取 MileStoneCaliInfo.txt，未修正 DMI。\r\n" + ex.Message);
                return;
            }

            if (calibration.Count < 2)
            {
                return;
            }

            DmiMile first = calibration[0];
            DmiMile last = calibration[calibration.Count - 1];
            // 已经归一化的工程不再提示，也不重复生成备份或重复扣减。
            if (first._Dmi == 0)
            {
                return;
            }
            bool valid = first._Dmi > 0
                && first._Mile == _StartMile
                && last._Mile == _EndMile
                && last._Dmi - first._Dmi == _EndDmi;

            for (int i = 1; valid && i < calibration.Count; i++)
            {
                bool dmiIncreasing = calibration[i]._Dmi > calibration[i - 1]._Dmi;
                bool mileInDirection = _Direction > 0
                    ? calibration[i]._Mile > calibration[i - 1]._Mile
                    : calibration[i]._Mile < calibration[i - 1]._Mile;
                valid = dmiIncreasing && mileInDirection;
            }

            if (!valid)
            {
                MessageBox.Show(string.Format(
                    "检测到三维数据但不能确认固定 DMI 偏移，未修改工程文件。\r\n首 DMI：{0}，末 DMI：{1}，工程总里程：{2}，DMI 差值：{3}",
                    first._Dmi, last._Dmi, _EndDmi, last._Dmi - first._Dmi));
                return;
            }

            int offset = first._Dmi;
            try
            {
                CreateReadOnlyBackup(caliFile, offset);
                string markFile = Path.Combine(_PrjPath, "RoadStatuMarkInfo.txt");
                if (File.Exists(markFile))
                {
                    CreateReadOnlyBackup(markFile, offset);
                }
                List<string> normalized = calibration
                    .Select(item => string.Format("{0} {1}", item._Dmi - offset, item._Mile))
                    .ToList();
                File.WriteAllLines(caliFile, normalized, Encoding.UTF8);
                _Pending3DStartupDmiOffset = offset;
            }
            catch (Exception ex)
            {
                MessageBox.Show("三维 DMI 偏移检测成功，但写入校桩修正失败，工程未完成修正。\r\n" + ex.Message);
            }
        }

        private static void CreateReadOnlyBackup(string sourceFile, int offset)
        {
            string backupFile = sourceFile + string.Format(".dmi-offset-{0}.bak", offset);
            if (File.Exists(backupFile))
            {
                return;
            }

            File.Copy(sourceFile, backupFile);
            File.SetAttributes(backupFile, File.GetAttributes(backupFile) | FileAttributes.ReadOnly);
        }

        private void NormalizeAndRecalculateRoadStatusMarks(int dmiOffset)
        {
            string markFile = Path.Combine(_PrjPath, "RoadStatuMarkInfo.txt");
            if (!File.Exists(markFile))
            {
                return;
            }

            string[] sourceLines = File.ReadAllLines(markFile, Encoding.UTF8);
            List<string> rewrittenLines = new List<string>();
            bool hasInvalidMark = false;
            if (dmiOffset > 0)
            {
                CreateReadOnlyBackup(markFile, dmiOffset);
            }

            foreach (string sourceLine in sourceLines)
            {
                string line = sourceLine.Replace("K", "").Replace("k", "").Replace("+", "");
                string[] values = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                int dmi;
                if (values.Length < 4 || !int.TryParse(values[2], out dmi))
                {
                    rewrittenLines.Add(sourceLine);
                    hasInvalidMark = sourceLine.Trim().Length > 0;
                    continue;
                }

                dmi -= dmiOffset;
                if (dmi < 0 || dmi > _EndDmi)
                {
                    rewrittenLines.Add(sourceLine);
                    hasInvalidMark = true;
                    continue;
                }

                int mile = Dmi2Mile(dmi);
                string markText = string.Join(" ", values.Skip(3));
                rewrittenLines.Add(string.Format("{0} {0} {1} {2}", mile, dmi, markText));
            }

            File.WriteAllLines(markFile, rewrittenLines, Encoding.UTF8);
            if (hasInvalidMark)
            {
                MessageBox.Show("部分 RoadStatuMarkInfo.txt 打标格式或 DMI 范围异常，已保留原行；请人工核对这些记录。");
            }
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
                        ListDM.Add(new DmiMile(int.Parse(str[0]), int.Parse(str[1])));
                    }
                }
            }
            else
            {
                int temp = _StartMile + _EndDmi * _Direction;
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
                if (ListDM[i]._Mile == ListDM[i + 1]._Mile)
                {
                    ListDM.RemoveAt(i);
                    dmlen--;
                    i--;
                }
            }

            if (_Direction > 0)//升序
            {
                ListDM.Sort(delegate(DmiMile x, DmiMile y) { return x._Mile.CompareTo(y._Mile); });
            }
            else if (_Direction < 0)//降序
            {
                ListDM.Sort(delegate(DmiMile x, DmiMile y) { return y._Mile.CompareTo(x._Mile); });
            }

            _DmiMileLen = ListDM.Count;
            _DmiMile = new double[_DmiMileLen, 2];//里程，桩号
            _D2MScale = new double[_DmiMileLen];
            _D2MScale[0] = _Direction;
            FileStream fw = new FileStream(_PrjPath + @"\Dmi2Mile.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < _DmiMileLen; ++i)
            {
                _DmiMile[i, 0] = ListDM[i]._Dmi;
                _DmiMile[i, 1] = ListDM[i]._Mile;
                if (i > 0) _D2MScale[i] = (_DmiMile[i, 0] - _DmiMile[i - 1, 0]) / (_DmiMile[i, 1] - _DmiMile[i - 1, 1]);
                sw.WriteLine(string.Format("{0} {1}", ListDM[i]._Dmi, ListDM[i]._Mile), Encoding.UTF8);
            }
            sw.Close();
            fw.Close();

            NormalizeAndRecalculateRoadStatusMarks(_Pending3DStartupDmiOffset);
            _Pending3DStartupDmiOffset = 0;
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
                if (_Direction > 0)
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
                if (_Direction > 0)
                {
                    if ((inmile >= _DmiMile[j - 1, 1] && inmile <= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        var temp =Math.Abs((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]);
                        return (double)Math.Round( temp,1);
                    }
                }
                else
                {
                    if ((inmile <= _DmiMile[j - 1, 1] && inmile >= _DmiMile[j, 1]) || j == _DmiMileLen - 1)
                    {
                        return (double)Math.Round(Math.Abs((inmile - _DmiMile[j - 1, 1]) * _D2MScale[j] + _DmiMile[j - 1, 0]),1);
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

    public class DmiMile
    {
        /// <summary>
        /// 里程
        /// </summary>
        public int _Dmi;

        /// <summary>
        /// 桩号
        /// </summary>
        public int _Mile;

        public DmiMile(int d, int m)
        {
            _Dmi = d;
            _Mile = m;
        }
        public DmiMile(DataGridViewRow row)
        {
            _Mile = Convert.ToInt32(row.Cells[0].Value);
            _Dmi = Convert.ToInt32(row.Cells[1].Value);
        }

        public override string ToString()
        {
            return string.Format("{0} {1}", _Dmi, _Mile);
        }
    }

    public class MarkInfo
    {
        /// <summary>
        /// 桩号
        /// </summary>
        public int _Mile;

        /// <summary>
        /// 打标类型
        /// </summary>
        public string _Type;

        /// <summary>
        /// 打标信息
        /// </summary>
        public string _Info;

        public MarkInfo()
        {
            _Mile = 0;
            _Type = null;
            _Info = null;
        }
        public MarkInfo(string info)
        {
            string[] str = info.Split(' ');
            _Mile = int.Parse(str[0].Replace("K", "").Replace("+", ""));

            str = str[str.Length - 1].Split(':');
            _Type = str[0];
            _Info = str[1];
        }
        public MarkInfo(DataGridViewRow row)
        {
            _Mile = Convert.ToInt32(row.Cells[0].Value);
            _Type = Convert.ToString(row.Cells[1].Value);
            _Info = Convert.ToString(row.Cells[2].Value);
        }
    }
}
