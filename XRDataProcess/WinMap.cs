using Framework;
using Framework.Log;
using Framework.Other.MyGlobal;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms; // WebView2 
using Ookii.Dialogs.WinForms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks; // 新增: IObjectForScripting
using System.Windows.Forms;
namespace XRDataProcess
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]  // 新增：暴露公共成员为 COM 接口，支持 JS 调用
    public partial class WinMap : Form
    {

        // Win32 API：设置透明样式
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int GWL_EXSTYLE = -20;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private void SetPanelTransparent(Panel panel)
        {
            var style = GetWindowLong(panel.Handle, GWL_EXSTYLE);
            SetWindowLong(panel.Handle, GWL_EXSTYLE, style | WS_EX_TRANSPARENT);
        }



        public static double Lat = 30.541093;
        public static double Log = 114.360734;
        public static string MileStr = null;
        public static int _stopflag = 1;
        public static int _PointNum = 0;
        private int _PointCnt = 0;
        MyLogger log = new MyLogger(typeof(WinMap));
        public static int _PointMileNum = 0;
        private int _PointMileCnt = 0;
        /// <summary>外部通过pMap调用UpdateGPS</summary>
        private List<MyGPS> _GPS;
        public ProjectInfo _ProjInfo;
        public string _ProjPath;
        private List<MyGPS> _GPSMile;
        private double _curlat;
        private double _curlng;
        public WinMap(ProjectInfo prjinfo, string ppath)
        {
            InitializeComponent();
            _ProjInfo = prjinfo;
            _ProjPath = ppath;
            _GPS = new List<MyGPS>();
            LoadGPS();
            //_GPS.Clear();
            //_GPS.Add(new MyGPS() { _Lat = 42.308898, _Long = 118.872518 });
            //_GPS.Add(new MyGPS() { _Lat = 42.308754, _Long = 118.872751 });
            //_GPS.Add(new MyGPS() { _Lat = 42.30861, _Long = 118.872984 });
            //_GPS.Add(new MyGPS() { _Lat = 42.308466, _Long = 118.873217 });
            //_GPS.Add(new MyGPS() { _Lat = 42.308322, _Long = 118.87345 });
            //_GPS.Add(new MyGPS() { _Lat = 42.308179, _Long = 118.873683 });
            //_GPS.Add(new MyGPS() { _Lat = 42.308035, _Long = 118.873916 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307891, _Long = 118.874149 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307747, _Long = 118.874382 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307603, _Long = 118.874614 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307459, _Long = 118.874847 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307315, _Long = 118.87508 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307171, _Long = 118.875313 });
            //_GPS.Add(new MyGPS() { _Lat = 42.307027, _Long = 118.875546 });
            //_GPS.Add(new MyGPS() { _Lat = 42.306884, _Long = 118.87578 });
            //_GPS.Add(new MyGPS() { _Lat = 42.306741, _Long = 118.876015 });
            //_GPS.Add(new MyGPS() { _Lat = 42.306598, _Long = 118.876249 });
            _PointNum = _GPS.Count;
            _GPSMile = new List<MyGPS>();
            LoadGPSMile();
            _PointMileNum = _GPSMile.Count;
            if (_GPS.Count > 0)
            {
                Lat = _GPS[0]._Lat;
                Log = _GPS[0]._Long;
            }
            _stopflag = 1;
            button_ShowStart.BringToFront();
            //SetPanelTransparent(panel2);
           // BringAllButtonsToFront();
        }
        private void BringAllButtonsToFront()
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is Button || ctrl is Label || ctrl is ComboBox) // 你想要的控件类型
                {
                    ctrl.BringToFront();
                }
            }
        }

        private async void InitializeAsync()
        {
            await webView2_map.EnsureCoreWebView2Async(null);
            webView2_map.CoreWebView2.AddHostObjectToScript("external", this);
             
            // 加载 HTML 并初始化
            //string str_url = Application.StartupPath + "\\Map\\tianditu\\tianditu.html";
            string str_url = Application.StartupPath + "\\Map\\BaiduMap.html";
            webView2_map.Source = new Uri(str_url);
            webView2_map.WebMessageReceived += (sender, args) =>
            {
                string message = args.TryGetWebMessageAsString();
                MessageBox.Show("JS日志: " + message);
            };
        }

      

      

        

        // 向 JavaScript 发送消息的 C# 方法
        public void SendMessageToJavaScript(string message)
        {
            if (webView2_map?.CoreWebView2 != null)
            {
                webView2_map.CoreWebView2.PostWebMessageAsString(message);
            }
        }
        public void SaveProgress(string str)
        {
            Console.WriteLine("SaveProgress 调用成功: " + str); 
        }

        /// <summary>
        /// 将当前工程的 GPS 轨迹打包成 JSON 格式的字符串返回给 JS
        /// </summary>
        public string GetTrajectoryJson()
        {
            if (_GPS == null || _GPS.Count == 0) return "[]";

            // 格式化为 JS 数组：[[lng, lat], [lng, lat], ...]
            var points = _GPS
                .Where(p => p._Lat > 0 && p._Long > 0) // 简单过滤一下明显的脏数据
                .Select(p => $"[{p._Long},{p._Lat}]");

            return "[" + string.Join(",", points) + "]";
        }



        // 初始化事件：添加宿主对象后，手动调用 init() 
        private void LoadGPSMile()
        {
            if (!File.Exists(_ProjPath + @"\GPS2Mile.txt"))
            {
                return;
            }
            string[] gpsstr = File.ReadAllLines(_ProjPath + @"\GPS2Mile.txt");
            string[] s;
            int oldmile = 0;
            int mile = 0;
            int cnt = 0;
            int len = gpsstr.Length - 1;
            foreach (string str in gpsstr)
            {
                s = str.Split(' ');
                if (s.Length == 6)
                {
                    mile = int.Parse(s[5]);
                    if (cnt == 0)
                    {
                        MyGPS tmp = new MyGPS();
                        tmp._Lat = double.Parse(s[2]);
                        tmp._Long = double.Parse(s[1]);
                        tmp._Mile = string.Format("起点 K{0}+{1}", mile / 1000, mile % 1000);
                        _GPSMile.Add(tmp);
                    }
                    else if (cnt == len)
                    {
                        MyGPS tmp = new MyGPS();
                        tmp._Lat = double.Parse(s[2]);
                        tmp._Long = double.Parse(s[1]);
                        tmp._Mile = string.Format("终点 K{0}+{1}", mile / 1000, mile % 1000);
                        _GPSMile.Add(tmp);
                    }
                    else
                    {
                        if (_ProjInfo._Direction > 0)
                        {
                            int tt1000 = (oldmile + 500) / 1000 * 1000;
                            if (oldmile < tt1000 && mile >= tt1000)
                            {
                                MyGPS tmp = new MyGPS();
                                tmp._Lat = double.Parse(s[2]);
                                tmp._Long = double.Parse(s[1]);
                                tmp._Mile = "K" + (tt1000 / 1000).ToString();
                                _GPSMile.Add(tmp);
                            }
                        }
                        else
                        {
                            int tt1000 = (oldmile + 500) / 1000 * 1000;
                            if (oldmile > tt1000 && mile <= tt1000)
                            {
                                MyGPS tmp = new MyGPS();
                                tmp._Lat = double.Parse(s[2]);
                                tmp._Long = double.Parse(s[1]);
                                tmp._Mile = "K" + (tt1000 / 1000).ToString();
                                _GPSMile.Add(tmp);
                            }
                        }
                    }
                    oldmile = mile;
                    ++cnt;
                }
            }
        }
        /// <summary>
        /// 20250611 增加对于GGA格式gps文件的支持
        /// </summary>
        private void LoadGPS()
        {
              
            string gpsFilePath =  findGpsFile();
            if (string.IsNullOrEmpty(gpsFilePath))
            {
                return;
            }

            string[] gpsLines = File.ReadAllLines(gpsFilePath);
            MyGPS lastGps = null; // 记录上一条有效数据
            double maxJump = 0.1; // 最大允许的经纬度跳跃（单位：度，约11公里）
            foreach (string line in gpsLines)
            {
                if (string.IsNullOrWhiteSpace(line) || !line.StartsWith("$"))
                {
                    continue;
                }
                string[] fields = line.Split(',');
                if (fields.Length < 6)
                {
                    continue;
                }
                MyGPS tgps = new MyGPS();
                bool isValid = false;
                if (fields[0].Contains("RMC"))
                {
                    if (fields.Length < 10 || fields[2] != "A")
                    {
                        continue;
                    }
                    if (!double.TryParse(fields[1], out tgps._UTC)) continue;
                    if (!ConvertDegreesToDigital(fields[3], 2, out tgps._Lat)) continue;
                    if (fields[4] == "S") tgps._Lat = -tgps._Lat;
                    if (!ConvertDegreesToDigital(fields[5], 3, out tgps._Long)) continue;
                    if (fields[6] == "W") tgps._Long = -tgps._Long;
                    isValid = true;
                }
                else if (fields[0].Contains("GGA"))
                {
                    if (fields.Length < 10 || fields[6] == "0")
                    {
                        continue;
                    }
                    if (!double.TryParse(fields[1], out tgps._UTC)) continue;
                    if (!ConvertDegreesToDigital(fields[2], 2, out tgps._Lat)) continue;
                    if (fields[3] == "S") tgps._Lat = -tgps._Lat;
                    if (!ConvertDegreesToDigital(fields[4], 3, out tgps._Long)) continue;
                    if (fields[5] == "W") tgps._Long = -tgps._Long;
                    isValid = true;
                }
                if (isValid)
                {
                    // 基本范围检查
                    if (tgps._Lat < 18 || tgps._Lat > 54 || tgps._Long < 73 || tgps._Long > 136)
                    {
                        Console.WriteLine($"Invalid coordinate: Lat={tgps._Lat}, Long={tgps._Long}, Line={line}");
                        continue;
                    }
                    // 连续性检查
                    if (lastGps != null)
                    {
                        double latDiff = Math.Abs(tgps._Lat - lastGps._Lat);
                        double lonDiff = Math.Abs(tgps._Long - lastGps._Long);
                        if (latDiff > maxJump || lonDiff > maxJump)
                        {
                            Console.WriteLine($"Jump detected: Lat={tgps._Lat}, Long={tgps._Long}, Line={line}");
                            continue;
                        }
                    }
                    _GPS.Add(tgps);
                    lastGps = tgps; // 更新上一条有效数据
                }
            }
        }


        private string findGpsFile()
        {
            string path = "";
            string[] subpath = { "\\RoadImg\\SYN\\gps.txt", "\\StreetImg\\SYN\\gps.txt",
                                       "\\IRIMTD\\SYN0\\gps.txt","\\IRIMTD\\SYN1\\gps.txt","\\camera0\\gps.txt","\\camera1\\gps.txt"};

            for (int i = 0; i < subpath.Length; i++)
            {
                path = _ProjPath + subpath[i];
                if (File.Exists(path))
                {
                    return path;
                }
            }
            return path;
        }
        
        /// <summary>
        /// 地图API获取纬度接口函数
        /// </summary>
        /// <returns>解析的纬度</returns>
        public double GetLat()
        {
            if (_PointCnt < _GPS.Count)
            {
                Lat = _GPS[_PointCnt]._Lat;
                ++_PointCnt;
            }
            return (Lat);
        }
        /// <summary>
        /// 地图API获取经度接口函数
        /// </summary>
        /// <returns>解析的经度</returns>
        public double GetLog()
        {

            if (_PointCnt < _GPS.Count)
            {
                Log = _GPS[_PointCnt]._Long;
            }
            return (Log);
        }
        public double StopFlag()
        {
            return (_stopflag);
        }
        public int GPSPointNum()
        {
            return _PointNum;
        }
        public double GetMileLat()
        {
            if (_PointMileCnt < _PointMileNum)
            {
                Lat = _GPSMile[_PointMileCnt]._Lat;
            }
            return (Lat);
        }
        public double GetMileLog()
        {
            if (_PointMileCnt < _PointMileNum)
            {
                Log = _GPSMile[_PointMileCnt]._Long;
            }
            return (Log);
        }
        public string GetMileStr()
        {
            if (_PointMileCnt < _PointMileNum)
            {
                MileStr = _GPSMile[_PointMileCnt]._Mile;
                ++_PointMileCnt;
            }
            return (MileStr);
        }
        public int GetMilePointNum()
        {
            return _PointMileNum;
        }
        public void SetCurPoint(double lat, double lng)
        {
            _curlat = lat;
            _curlng = lng;
            textBox1.Text = string.Format("{0:0.0000000},{1:0.0000000}", _curlng, _curlat);
        }
        /// <summary>
        /// 将度分转换为度
        /// </summary>
        /// <param name="inval">待转换的度分-字符串</param>
        /// <param name="sidx">度分所占字符个数</param>
        /// <param name="outval">转换后的度分-浮点型</param>
        /// <returns></returns>
        private bool ConvertDegreesToDigital(string inval, int sidx, out double outval)
        {
            bool res = false;
            outval = 0;
            if (inval == string.Empty)
            {
                return res;
            }
            try
            {
                int du = int.Parse(inval.Substring(0, sidx));
                outval = double.Parse(inval.Substring(sidx));
                outval = outval / 60;
                outval += du;
                res = true;
            }
            catch (Exception ex)
            {
                Framework.Other.MyGlobal.Global.g_mutex.WaitOne();
                // Global.g_log.Write(ex, "转换GPS坐标出现异常-->inval=" + inval);
                log.Error("转换GPS坐标出现异常-->inval=" + inval + ex.StackTrace);
                Global.g_mutex.ReleaseMutex();
            }
            return res;
        }
        //RMC,<1>,<2>,<3>,<4>,<5>,<6>,<7>,<8>,<9>,<10>,<11>,<12>*hh
        //<1> UTC时间，hhmmss.sss(时分秒.毫秒)格式
        //<2> 定位状态，A=有效定位，V=无效定位
        //<3> 纬度ddmm.mmmm(度分)格式(前面的0也将被传输)
        //<4> 纬度半球N(北半球)或S(南半球)
        //<5> 经度dddmm.mmmm(度分)格式(前面的0也将被传输)
        //<6> 经度半球E(东经)或W(西经)
        public class MyGPS
        {
            public double _UTC;//utc时间
            public double _Lat;//纬度
            public double _Long;//经度
            public string _Mile;//桩号
            public MyGPS()
            {
                _UTC = 0;
                _Lat = 0;
                _Long = 0;
                _Mile = null;
            }
        }
        private void timer_update_Tick(object sender, EventArgs e)
        {
            if (_PointCnt < _PointNum)
            {
                Lat = _GPS[_PointCnt]._Lat;
                Log = _GPS[_PointCnt]._Long;
                _PointCnt++;
            }
            else
            {
                timer_update.Stop();
                _stopflag = 0;
            }
        }
        //private void button_ShowRoad_Click(object sender, EventArgs e)
        //{
        //    this.Cursor = Cursors.WaitCursor;
        //     button_Clear.Enabled = true; // 根据需要调整按钮状态
        //     button_ShowRoad.Enabled = false;
        //    _PointCnt = 1;
        //    webView2_map.CoreWebView2.ExecuteScriptAsync("ShowRoad();");
        //    this.Cursor = Cursors.Default;
        //}

        private void button_ShowRoad_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            button_Clear.Enabled = true;
            button_ShowRoad.Enabled = false;

            // 1. 拿到全量 JSON 数据
            string trajectoryData = GetTrajectoryJson();

            // 2. 将数据推给 JS，让 JS 瞬间完成渲染
            string script = $"LoadNewProjectTrajectory({trajectoryData});";
            webView2_map.CoreWebView2.ExecuteScriptAsync(script);

            this.Cursor = Cursors.Default;
        }
        private void button_ShowMile_Click(object sender, EventArgs e)
        {
            if (!File.Exists(_ProjPath + @"\GPS2Mile.txt"))
            {
                MessageBox.Show("请先进行GPS桩号匹配！");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
             button_Clear.Enabled = true;
             button_ShowMile.Enabled = false;
            _PointMileCnt = 0;
            if (_PointMileNum < 2)
            {
                _GPSMile.Clear();
                LoadGPSMile();
                _PointMileNum = _GPSMile.Count;
            }
            webView2_map.CoreWebView2.ExecuteScriptAsync("ShowMile();");
            this.Cursor = Cursors.Default;
        }
        private void button_Clear_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
             button_ShowRoad.Enabled = true;
             button_ShowMile.Enabled = true;
            webView2_map.CoreWebView2.ExecuteScriptAsync("ClearAll();");
            this.Cursor = Cursors.Default;
        }
        private void button_add_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == null || string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("请输入标签内容！");
                return;
            }
            // 使用当前点击点或默认点添加标记
            string script = $"AddMark({_curlat}, {_curlng}, '{textBox2.Text.Replace("'", "\\'")}');";
            webView2_map.CoreWebView2.ExecuteScriptAsync(script);
        }
        private void button_delete_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == null || string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("请输入要删除的标签内容！");
                return;
            }
            string script = $"DeleteMark('{textBox2.Text.Replace("'", "\\'")}');";
            webView2_map.CoreWebView2.ExecuteScriptAsync(script);
        }
        private void button_ShowStart_Click(object sender, EventArgs e)
        {
            if (_GPSMile.Count > 0)
            {
                Lat = _GPSMile[0]._Lat;
                Log = _GPSMile[0]._Long;
            }
            string script = $"ShowStart({Lat}, {Log});";
            webView2_map.CoreWebView2.ExecuteScriptAsync(script);
        }
        //private async Task InitializeWebView2Safely()
        //{
        //    try
        //    {
        //        // 先检查是否已经初始化
        //        if (webView2_map.CoreWebView2 != null)
        //        {
        //            Console.WriteLine("WebView2 已经初始化");
        //            return;
        //        }

        //        // 设置用户数据文件夹，避免冲突
        //        var userDataFolder = Path.Combine(Path.GetTempPath(), "WebView2Cache", Guid.NewGuid().ToString());

        //        var environment = await CoreWebView2Environment.CreateAsync(
        //            browserExecutableFolder: null,
        //            userDataFolder: userDataFolder,
        //            options: null);

        //        await webView2_map.EnsureCoreWebView2Async(environment);

        //        // 注册宿主对象
        //        webView2_map.CoreWebView2.AddHostObjectToScript("external", this);

        //        // 加载地图页面
        //       // string str_url = Application.StartupPath + "\\Map\\BaiduMap.html";
        //        string str_url = Application.StartupPath + "\\Map\\tianditu\\tianditu.html";
        //        webView2_map.Source = new Uri(str_url);

        //        Console.WriteLine("WebView2 初始化成功");
        //    }
        //    catch (Exception ex)
        //    {
        //        //MessageBox.Show($"WebView2 初始化失败: {ex.Message}\n\n建议：重启应用程序或检查网络连接",
        //        //               "初始化错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        // 记录详细错误信息
        //        log.Error($"WebView2初始化异常: {ex}");
        //    }
        //}

        private async Task InitializeWebView2Safely()
        {
            try
            {
                if (webView2_map.CoreWebView2 != null) return;

             
                // ✅ 替换为固定路径：把地图瓦片永久存在客户电脑的 AppData 目录下
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string userDataFolder = Path.Combine(appDataPath, "XRDataProcess", "MapCache");

                // 如果文件夹不存在，自动创建
                if (!Directory.Exists(userDataFolder))
                {
                    Directory.CreateDirectory(userDataFolder);
                }

                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
                await webView2_map.EnsureCoreWebView2Async(environment);

            
               webView2_map.CoreWebView2.AddHostObjectToScript("external", this);

                // 加载地图页面
                // string str_url = Application.StartupPath + "\\Map\\BaiduMap.html";
                string str_url = Application.StartupPath + "\\Map\\tianditu\\tianditu.html";
                webView2_map.Source = new Uri(str_url);

                Console.WriteLine("WebView2 初始化成功");
            }
            catch (Exception ex)
            {
                log.Error($"WebView2初始化异常: {ex}");
            }
        }
        private async void WinMap_Shown(object sender, EventArgs e)
        {
            await InitializeWebView2Safely();
           
        }
        private string GetUniqueFilePath(string filePath)
        {
            if (!File.Exists(filePath))
                return filePath;

            string directory = Path.GetDirectoryName(filePath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            int count = 1;
            string newPath;
            do
            {
                string numberedName = string.Format("{0}_{1}{2}", fileNameWithoutExt, count++, extension);
                newPath = Path.Combine(directory, numberedName);
            } while (File.Exists(newPath));

            return newPath;
        }
        private string ShowImageFormatDialog()
        {
            Form form = null;
            ComboBox combo = null;
            Button btnOk = null;
            Button btnCancel = null;

            try
            {
                form = new Form
                {
                    Text = "选择图片格式",
                    Width = 280,
                    Height = 140,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    StartPosition = FormStartPosition.CenterParent,
                    MaximizeBox = false,
                    MinimizeBox = false
                };

                combo = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Left = 20,
                    Top = 20,
                    Width = 220
                };
                combo.Items.AddRange(new object[] { "PNG", "JPEG", "BMP" });
                combo.SelectedIndex = 0;

                btnOk = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Left = 100,
                    Top = 60,
                    Width = 75
                };

                btnCancel = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Left = 180,
                    Top = 60,
                    Width = 75
                };

                form.Controls.Add(combo);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);

                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                return form.ShowDialog(this) == DialogResult.OK ? combo.SelectedItem.ToString() : null;
            }
            finally
            {
                form?.Dispose();
                combo?.Dispose();
                btnOk?.Dispose();
                btnCancel?.Dispose();
            }
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            VistaFolderBrowserDialog fd = null;
            try
            {
                fd = new VistaFolderBrowserDialog
                {
                    Description = "选择保存地图截图的文件夹",
                    ShowNewFolderButton = true
                };

                if (fd.ShowDialog(this) != DialogResult.OK || string.IsNullOrEmpty(fd.SelectedPath))
                    return;

                string selectedFormat = ShowImageFormatDialog();
                if (string.IsNullOrEmpty(selectedFormat))
                    return;

                string extension = selectedFormat.ToLower();

                //获取工程文件名称
                string projectName = Path.GetFileName(_ProjInfo._PrjPath);


                string fileName = projectName+"_MapImage." + extension;
                string savePath = Path.Combine(fd.SelectedPath, fileName);
                savePath = GetUniqueFilePath(savePath); // 防重名

                this.Cursor = Cursors.WaitCursor;

                // 格式映射
                CoreWebView2CapturePreviewImageFormat imageFormat;
                if (selectedFormat == "PNG")
                    imageFormat = CoreWebView2CapturePreviewImageFormat.Png;
                else if (selectedFormat == "JPEG")
                    imageFormat = CoreWebView2CapturePreviewImageFormat.Jpeg;
                else
                    imageFormat = CoreWebView2CapturePreviewImageFormat.Png;

                // 关键：使用 Stream 重载（兼容旧版）
                using (FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                {
                    await webView2_map.CoreWebView2.CapturePreviewAsync(imageFormat, fs);
                }

                MessageBox.Show("地图截图已保存：\n" + savePath, "保存成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (log != null)
                    log.Error("地图截图保存失败", ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                if (fd != null)
                    fd.Dispose();
            }
        }
    }
}