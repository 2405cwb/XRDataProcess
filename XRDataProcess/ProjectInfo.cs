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
                            }
                            catch (Exception)
                            {

                                MessageBox.Show($"{_PrjPath}项目数据下projectinfo文件中路面材质列填写错误！");
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
            string[] sinfo = File.ReadAllLines(_PrjPath + @"\ProjectInfo.txt", Encoding.UTF8);
            int len = sinfo.Length;
            for (int i = 0; i < len; i++)
            {
                string[] s = sinfo[i].Split('：');
                switch (s[0])
                {
                    case "省": sinfo[i] = string.Format("{0}：{1}", s[0], _Province); break;
                    case "市": sinfo[i] = string.Format("{0}：{1}", s[0], _City); break;
                    case "县": sinfo[i] = string.Format("{0}：{1}", s[0], _District); break;
                    case "工程起点道路编号": sinfo[i] = string.Format("{0}：{1}", s[0], _RoadCode); break;
                    case "工程起点道路名称": sinfo[i] = string.Format("{0}：{1}", s[0], _RoadName); break;
                    case "工程起点桩号": sinfo[i] = string.Format("{0}：{1:K0000+000}", s[0], _StartMile); break;
                    case "行车方向": sinfo[i] = string.Format("{0}：{1}", s[0], _Direction > 0 ? "上行" : "下行"); break;
                    case "公路等级": sinfo[i] = string.Format("{0}：{1}", s[0], _RoadGrade); break;
                    case "车道": sinfo[i] = string.Format("{0}：{1}", s[0], _RoadNum); break;
                    case "采集日期": sinfo[i] = string.Format("{0}：{1}", s[0], _DataDate); break;
                    case "工程开始时刻": sinfo[i] = string.Format("{0}：{1}", s[0], _DataTime); break;
                    case "检测员": sinfo[i] = string.Format("{0}：{1}", s[0], _DataPerson); break;
                    case "检测天气": sinfo[i] = string.Format("{0}：{1}", s[0], _DataWeather); break;
                    case "路面材质": sinfo[i] = string.Format("{0}：{1}", s[0], _RoadType == 0 ? "沥青" : "水泥"); break;
                    case "工程终点道路标识桩号": sinfo[i] = string.Format("{0}：{1:K0000+000}", s[0], _EndMile); break;
                    case "工程总里程数": sinfo[i] = string.Format("{0}：{1:K0000+000}", s[0], _EndDmi); break;
                }
            }
            File.WriteAllLines(_PrjPath + @"\ProjectInfo.txt", sinfo, Encoding.UTF8);
            TranDmi2Mile();
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
                File.WriteAllLines(tfname, strlist, Encoding.UTF8);
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
