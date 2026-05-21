#define 辽宁奔驰综合检测车
//#define 支持无平整度导出国检
using DevExpress.Internal.WinApi;
using DevExpress.Map.Native;
using DevExpress.SpreadsheetSource.Implementation;
using DevExpress.SpreadsheetSource.Xls;
using DevExpress.Utils.Drawing;
using DevExpress.Utils.Serializing.Helpers;
using DevExpress.XtraBars.Docking2010;
using DevExpress.XtraCharts;
using DevExpress.XtraLayout.Customization;
using DevExpress.XtraMap;
using DevExpress.XtraPrinting.Export.Pdf;
using Framework.Log;
using Framework.Other;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using OperateIniFile;
using Org.BouncyCastle.Utilities.Date;
using RoadStreet;
using Spire.Pdf.Exporting.XPS.Schema;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Windows.Forms;
using XRDataProcess.toolForms;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace XRDataProcess
{
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class SingleProject : Form
    {
        private static MyLogger _log = new MyLogger(typeof(SingleProject));
        [DllImport("hnCalcuMethod", EntryPoint = "setParam")]
        static extern void setParam(string strDaqPath, int nImuHz);
        [DllImport("hnCalcuMethod", EntryPoint = "setSaveResamplePath")]
        static extern void setSaveResamplePath(string strSaveResamplePath);
        [DllImport("hnCalcuMethod", EntryPoint = "setSaveIRIPath10")]
        static extern void setSaveIRIPath10(string strSavePath, bool bSaveIRI10);
        [DllImport("hnCalcuMethod", EntryPoint = "setSaveIRIPath100")]
        static extern void setSaveIRIPath100(string strSavePath, bool bSaveIRI100);
        [DllImport("hnCalcuMethod", EntryPoint = "setSaveIRIPath1000")]
        static extern void setSaveIRIPath1000(string strSavePath, bool bSaveIRI1000);
        [DllImport("hnCalcuMethod", EntryPoint = "setIsOnRight")]
        static extern void setIsOnRight(int onRight);
        [DllImport("hnCalcuMethod", EntryPoint = "calcuIRI")]
        static extern bool calcuIRI();

        [DllImport("hnCalcuMethod", EntryPoint = "calcuCelerator")]
        static extern bool calcuCelerator(string savePath);

        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();
        // 放在 SingleProject 类内部、任何方法之外
        private const string _layoutFileName = "SingleLayout.xml";        // 运行时布局
        private const string _layoutDefaultFileName = "SingleDefaultLayout.xml"; // 安装目录默认布局
        private bool _isLayoutSafeToSave = true;
        public DirectoryInfo _DataDir;
        public ProjectInfo _ProjectInfo;
        private WinProj _winproj;
        private WinRoad _winroad;
        private WinStreetImg _winstreet;
        private 采集打标列表 _winmark;
        private WinRoadDisList _winroadis;
        private WinStreetDisList _winstreetdis;
        private WinMap _winmap;
        private WinIRM _winirm;
        private WinPanoImg _winpano;
        private YGView _winYunGuang;
        private WinImgFull _winimgfull;
        public WinProcessBar _Bars;
        public WinPanoProcessBar _PanoBars;

        public void SaveCurDisease()
        {
            if (_winroad is null)
            {

            }
            else
            {
                _winroad.SaveDisease();

            }
        }

        public SingleProject(DirectoryInfo dir)
        {
            InitializeComponent();
            // 1. 确保默认布局已复制（仅第一次需要）
            EnsureDefaultLayoutCopied();
            _ThreadIRI = new Thread[2];
            _ThreadMTD = new Thread[3];
            _ThreadMPD = new Thread[3];

            _DataDir = dir;
            _ProjectInfo = new ProjectInfo(_DataDir.FullName);
        }

        public SingleProject(string dirPath)
        {
            DirectoryInfo dir = new DirectoryInfo(dirPath);
            InitializeComponent();
            EnsureDefaultLayoutCopied();
            _ThreadIRI = new Thread[2];
            _ThreadMTD = new Thread[3];
            _ThreadMPD = new Thread[3];

            _DataDir = dir;
            _ProjectInfo = new ProjectInfo(_DataDir.FullName);
        }

        public void RestoreDefaultLayout()
        {
            string defaultPath = GetDefaultLayoutPath();
            if (File.Exists(defaultPath))
            {
                TryRestoreLayout(defaultPath, "默认项目布局");
                _isLayoutSafeToSave = true;
            }
        }

        public void RestoreSavedLayout()
        {
            _isLayoutSafeToSave = RestoreSavedLayoutSafely();
        }

        private bool RestoreSavedLayoutSafely()
        {
            string userPath = GetUserLayoutPath();
            string defaultPath = GetDefaultLayoutPath();

            if (File.Exists(userPath))
            {
                if (TryRestoreLayout(userPath, "用户项目布局"))
                {
                    return true;
                }

                BackupBadLayoutFile(userPath);
                if (File.Exists(defaultPath) && TryRestoreLayout(defaultPath, "默认项目布局"))
                {
                    return true;
                }

                return false;
            }

            if (File.Exists(defaultPath))
            {
                TryRestoreLayout(defaultPath, "默认项目布局");
            }

            return true;
        }

        private bool TryRestoreLayout(string layoutPath, string layoutName)
        {
            try
            {
                dockManager_main.RestoreLayoutFromXml(layoutPath);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(layoutName + "恢复失败: " + ex.Message);
                return false;
            }
        }

        public void Resize(int height)
        {
            this.Height = height;
        }
         
        /// <summary>
        /// 获取用户专属的 \\Setting\\Project\\Setting.ini
        /// </summary>
        private string GetUserDefaultConfigPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = System.IO.Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");

            return appFolder;
        }


        private void CheckSetting()
        {
            bool IsCopy = false;

            string temp = GetUserDefaultConfigPath() + "\\Setting\\Project";
            Directory.CreateDirectory(temp);
            string rpath = GetUserDefaultConfigPath() + "\\Setting\\Project\\Setting.ini";
            string rDirPath = "";
            string dpath = _DataDir.FullName;
            string fpath = dpath + "\\Setting.ini";
            if (File.Exists(fpath))
            {
                string[] strs = File.ReadAllLines(fpath);
                if (strs.Length > 2)
                {
                    IsCopy = false;
                }
                else
                {
                    IsCopy = true;
                }
            }
            else
            {
                IsCopy = true;
            }
            if (IsCopy)
            {
                if (File.Exists(rpath))
                {
                    File.Copy(rpath, fpath, true);
                }
            }
            else
            {
                File.Copy(fpath, rpath, true);
            }

            string[] subdir = { "IRIMTD\\DAQ", "IRIMTD\\Laser" };
            for (int si = 0; si < subdir.Length; ++si)
            {
                for (int i = 0; i < 2; ++i)
                {
                    IsCopy = false;
                    rDirPath = string.Format("{0}\\Setting\\{1}{2}\\", GetUserDefaultConfigPath(), subdir[si], i);
                    Directory.CreateDirectory(rDirPath);
                    rpath = string.Format("{0}\\Setting\\{1}{2}\\Setting.ini", GetUserDefaultConfigPath(), subdir[si], i);

                    dpath = string.Format("{0}\\{1}{2}", _DataDir.FullName, subdir[si], i);
                    if (Directory.Exists(dpath))
                    {
                        fpath = dpath + "\\Setting.ini";
                        if (File.Exists(fpath))
                        {
                            IsCopy = false;
                        }
                        else
                        {
                            IsCopy = true;
                        }
                        if (IsCopy)
                        {
                            if (File.Exists(rpath))
                            {
                                File.Copy(rpath, fpath, true);
                            }
                        }
                        else
                        {
                            File.Copy(fpath, rpath, true);
                        }
                    }
                }
            }
        }
         
        
     
        /// <summary>
        /// 获取用户专属的 Single 布局文件路径（%LocalAppData%）
        /// </summary>
        private string GetUserLayoutPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = System.IO.Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);
            return System.IO.Path.Combine(appFolder, _layoutFileName);
        }

        /// <summary>
        /// 获取安装目录下的默认布局路径（只读）
        /// </summary>
        private string GetDefaultLayoutPath()
        {
            return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _layoutDefaultFileName);
        }

        /// <summary>
        /// 首次运行时复制默认布局到用户目录
        /// </summary>
        private void EnsureDefaultLayoutCopied()
        {
            string userPath = GetUserLayoutPath(); 

            if (!File.Exists(userPath) && File.Exists(GetDefaultLayoutPath()))
            {
                try { File.Copy(GetDefaultLayoutPath(), userPath); }
                catch { /* 静默忽略 */ }
            }
        }

        private void SaveLayoutSafely()
        {
            if (!_isLayoutSafeToSave)
            {
                System.Diagnostics.Debug.WriteLine("跳过项目布局保存: 本次启动未能恢复有效布局。");
                return;
            }

            string userLayoutPath = GetUserLayoutPath();
            string tempPath = userLayoutPath + ".tmp";

            try
            {
                using (FileStream stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    dockManager_main.SaveLayoutToStream(stream);
                }

                ValidateLayoutFile(tempPath);

                if (File.Exists(userLayoutPath))
                {
                    File.Replace(tempPath, userLayoutPath, null, true);
                }
                else
                {
                    File.Move(tempPath, userLayoutPath);
                }
            }
            finally
            {
                TryDeleteFile(tempPath);
            }
        }

        private static void ValidateLayoutFile(string layoutPath)
        {
            FileInfo fileInfo = new FileInfo(layoutPath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                throw new InvalidDataException("布局文件为空。");
            }

            System.Xml.XmlReaderSettings settings = new System.Xml.XmlReaderSettings
            {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit
            };

            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(layoutPath, settings))
            {
                while (reader.Read())
                {
                }
            }
        }

        private static void BackupBadLayoutFile(string layoutPath)
        {
            try
            {
                if (!File.Exists(layoutPath))
                {
                    return;
                }

                string folder = System.IO.Path.GetDirectoryName(layoutPath);
                string fileName = System.IO.Path.GetFileNameWithoutExtension(layoutPath);
                string extension = System.IO.Path.GetExtension(layoutPath);
                string backupPath = System.IO.Path.Combine(folder, fileName + ".bad" + extension);
                TryDeleteFile(backupPath);
                File.Move(layoutPath, backupPath);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("备份损坏布局失败: " + ex.Message);
            }
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("删除临时布局文件失败: " + ex.Message);
            }
        }

        public void SingleProject_FormClosed(object sender, FormClosedEventArgs e)
        {

            try
            { 
                SaveLayoutSafely();

                // 3. 关闭子窗体（保持原有逻辑）
                if (_winroad != null)
                {
                    _winroad.Close();
                }
            }
            catch (Exception ex)
            {
                // 防止任何异常导致崩溃
                System.Diagnostics.Debug.WriteLine("SingleProject_FormClosed 错误: " + ex.Message);
                // 如使用 log4net：
                // log.Error("SingleProject_FormClosed 异常", ex);
            }


            //dockManager_main.SaveLayoutToXml(_layoutpath);
            //if (_winroad != null)
            //{
            //    _winroad.Close();
            //}
        }


        private void SingleProject_Load(object sender, EventArgs e)
        {
            try
            {
                loadQueue.Enqueue(() =>
                {

                    RestoreSavedLayout();
                    if (LoadProjInfo() == false) return;
                });

                ProcessLoadQueue();
            }
            finally
            {
            }

            //LoadMap();

        }
        private Queue<System.Action> loadQueue = new Queue<System.Action>();
        private bool isProcessingQueue = false;

        /// <summary>
        /// 在批量操作时，优化加载逻辑，减少同时创建的窗口句柄数量。
        /// </summary>
        private void ProcessLoadQueue()
        {
            if (isProcessingQueue) return;
            isProcessingQueue = true;

            while (loadQueue.Count > 0)
            {
                var action = loadQueue.Dequeue();
                action();
                Thread.Sleep(100); // 延迟一段时间，避免资源耗尽
            }

            isProcessingQueue = false;
        }





        public void DisposeResources()
        {
            if (_winproj != null)
            {
                _winproj.Dispose();
                _winproj = null;
            }

            if (_winYunGuang != null)
            {
                _winYunGuang.Dispose();
                _winYunGuang = null;
            }

            if (_winimgfull != null)
            {
                _winimgfull.Dispose();
                _winimgfull = null;
            }

            if (_winroad != null)
            {
                _winroad.Dispose();
                _winroad = null;
            }

            if (_winroadis != null)
            {
                _winroadis.Dispose();
                _winroadis = null;
            }

            if (_winstreet != null)
            {
                _winstreet.Dispose();
                _winstreet = null;
            }

            if (_winstreetdis != null)
            {
                _winstreetdis.Dispose();
                _winstreetdis = null;
            }

            if (_winpano != null)
            {
                _winpano.Dispose();
                _winpano = null;
            }

            if (_winirm != null)
            {
                _winirm.Dispose();
                _winirm = null;
            }

            if (_winmark != null)
            {
                _winmark.Dispose();
                _winmark = null;
            }

            if (_winmap != null)
            {
                _winmap.Dispose();
                _winmap = null;
            }
        }
        /// <summary>
        /// 低等级农村路计算IRI
        /// </summary>
        public bool lowComputeIRI(WinGDProcessBar process)
        {



            _ProjectInfo._IsIRIMTD = true;
            string inifpath = _DataDir.FullName + "\\Setting.ini";
            IniFiles iniset = new IniFiles(inifpath);

            iniset.WriteBool("工作模式", "IRIMTD", true);

            CheckSetting();
            DirectoryInfo di = new DirectoryInfo(_DataDir.FullName);
            FileInfo[] files = di.GetFiles("*.daq");
            if (files.Length > 0)
            {
                //检测是否为长沙理工 惯导+激光一体平整度 
                string coeffFilePath = System.IO.Path.Combine(_DataDir.FullName, "IRIMTD", "DAQ0", "Coeff.dat");
                bool hasJgAndJsd = false;
                if (File.Exists(coeffFilePath))
                {
                    hasJgAndJsd = true;
                }
                DirectoryInfo irPath = null;
                if (hasJgAndJsd || _ProjectInfo._IsJgAndGd)
                {
                    //惯导的平整度结果放到右侧去
                    irPath = di.CreateSubdirectory("IRIMTD\\DAQ1");
                    //设置平整度为双侧
                    _ProjectInfo._IsDIRIMTD = true;
                    //设置该文件
                    string iniSettingFilePath = System.IO.Path.Combine(_DataDir.FullName, "Setting.ini");
                    // 读取INI文件内容
                    List<string> iniTxts = File.ReadAllLines(iniSettingFilePath, Encoding.GetEncoding("gb2312")).ToList();

                    // 查找并修改DIRIMTD的值
                    for (int i = 0; i < iniTxts.Count; i++)
                    {
                        if (iniTxts[i].Contains("DIRIMTD =False"))
                        {
                            iniTxts[i] = "DIRIMTD =True";
                            break; // 找到并修改后退出循环
                        }
                    }
                    // 将修改后的内容写回到文件
                    File.WriteAllLines(iniSettingFilePath, iniTxts, Encoding.GetEncoding("gb2312"));

                }
                else
                {
                    //只有加速度
                    irPath = di.CreateSubdirectory("IRIMTD\\DAQ0");
                }
                string strResamplePath250 = irPath.FullName + "\\ReSample250.txt";
                string strSaveIRIPath10 = irPath.FullName + "\\IRI_10m.txt";
                string strSaveIRIPath100 = irPath.FullName + "\\IRI_100m.txt";
                string strSaveIRIPath1000 = irPath.FullName + "\\IRI_1000m.txt";
                //{149656288595357090_S312_沿黄线_HNJC-A0000_20220318_143921}_2000_2000_0783_1044

                setParam(files[0].FullName, 900);
                setSaveResamplePath(strResamplePath250);
                setSaveIRIPath10(strSaveIRIPath10, true);
                setSaveIRIPath100(strSaveIRIPath100, true);
                setSaveIRIPath1000(strSaveIRIPath1000, true);
                setIsOnRight(1);
                bool result = calcuIRI();

                return result;

            }
            else
            {

                MessageBox.Show("缺少数据文件");
                return false;
            }


        }
        /// <summary>
        /// 直接将daq转为加速度文件
        /// </summary>
        public static void computeDaqToAcce(List<string> daqFiles)
        {
            if (daqFiles.Count < 1)
            {
                return;
            }
            string firstDaqName = System.IO.Path.GetFileNameWithoutExtension(daqFiles[0]);
            string[] nameSplits = firstDaqName.Split('_');
            if (nameSplits.Length < 6)
            {
                MessageBox.Show("文件名称不符合规定，请检查");
                return;
            }
            ConventGetRoadNumberForm form = new ConventGetRoadNumberForm(nameSplits[1]);
            form.ShowDialog();
            if (!form.Ok)
            {
                return;
            }
            string CityCode = form.CityNum;

            foreach (var daq in daqFiles)
            {
                string daqName = System.IO.Path.GetFileNameWithoutExtension(daq);
                //20250811 根据河南交发院需求对输出文件名称进行修改
                //309083419481050114_Y002_睦邻大道_20240728_162458_20240728_2174_2154_0691_0045 =>
                //路线编码(C001)行政代码(410324)上下行(上行A下行B)-a-起点桩号(0.000)-检测时间(20240520082045)
                // C001410324A - a - 0.000 - 20240520082045
                string[] nowNameSplits = daqName.Split('_');

                if (nowNameSplits.Length < 6 || nowNameSplits[1].Length < 4)
                {
                    MessageBox.Show($"{daqName}文件名称不符合规定，请检查", "提示", MessageBoxButtons.OK);
                    continue;
                }
                string temp = nowNameSplits[1].Substring(0, 4);
                string RoadCode = $"{temp}{CityCode}A"; //路线编码+县级行政代码+方向
                string RoadName = "a";
                string sMile = "0.000";
                string time = nowNameSplits[3] + nowNameSplits[4];

                string path = System.IO.Path.GetDirectoryName(daq);
                string fileName = string.Join("_", RoadCode, RoadName, sMile, time);

                string savePath = path + "\\" + fileName + ".txt";
                setParam(daq, 900);
                calcuCelerator(savePath);
                List<string> texts = File.ReadAllLines(savePath).ToList();

                // 处理每一行：将第二列乘以0.001
                List<string> newTexts = texts

     .Select(line =>
     {
         string[] columns = line.Split(',');
         if (columns.Length >= 1 && double.TryParse(columns[0], out double value0))
         {
             columns[0] = value0.ToString("f6");
         }
         if (columns.Length >= 2 && double.TryParse(columns[1], out double value1))
         {
             columns[1] = (value1 * 0.001).ToString("f6"); // 第二列单位转换为m
         }
         if (columns.Length >= 3 && double.TryParse(columns[2], out double value2))
         {
             columns[2] = value2.ToString("f5");
         }
         if (columns.Length >= 4 && double.TryParse(columns[3], out double value3))
         {
             columns[3] = value3.ToString("f5");
         }
         if (columns.Length >= 5 && double.TryParse(columns[4], out double value4))
         {
             columns[4] = value4.ToString("f5");
         }
         return string.Join(",", columns);
     })
     .ToList();

                newTexts.Insert(0, "时长,起点桩号(km),左加速度(m/s²),右加速度(m/s²),速度(m/s)");
                File.WriteAllLines(savePath, newTexts);
            }

            MessageBox.Show("加速度文本文件生成完成!");
        }

        /// <summary>
        /// 计算或读取IRI、MTD和Rut
        /// 函数用多线程实现
        /// </summary>
        public void ComputeIRM(WinProcessBar bars)
        {
            CheckSetting();

            _Bars = bars;
            if (_ProjectInfo._IsIRIMTD && (SelectIRM.irm[0] || SelectIRM.irm[2] || SelectIRM.irm[3]))
            {
                try
                {
                    string dtstr = _ProjectInfo._DataDate + _ProjectInfo._DataTime;
                    DateTime prjdt = DateTime.ParseExact(dtstr, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                    MainForm.rdval = new Random(prjdt.TimeOfDay.Seconds);
                }
                catch (System.Exception)
                { }

                MyIRIMTD.LoadParm(_DataDir.FullName);
                for (int i = 0; i < 2; i++)
                {
                    int t = i;

                    if (Directory.Exists(string.Format(@"{0}\IRIMTD\DAQ{1}", _DataDir.FullName, t)) && SelectIRM.irm[0])
                    {
                        string fpath = string.Format(@"{0}\IRIMTD\DAQ{1}\resample.txt", _DataDir.FullName, t);
                        if (_Setting.Las_Filter)
                        {
                            MyIRIMTD.FilterLaserData(fpath, _Setting.Las_Filter_Thresh0, _Setting.Las_Filter_Thresh1);
                        }
                        else
                        {
                            if (File.Exists(fpath + ".bak"))
                            {
                                File.Delete(fpath);
                                File.Move(fpath + ".bak", fpath);
                            }


                        }

                        StartIRIThread(new ThreadInfo(_DataDir.FullName, t, bars));
                    }

                    if (Directory.Exists(string.Format(@"{0}\IRIMTD\Laser{1}", _DataDir.FullName, t)))
                    {
                        double lasthreshval = 50;
                        if (Directory.Exists(string.Format(@"{0}\IRIMTD\DAQ{1}", _DataDir.FullName, t)))
                        {
                            string fpath = string.Format(@"{0}\IRIMTD\DAQ{1}\resample.txt", _DataDir.FullName, t);
                            lasthreshval = MyIRIMTD.GetLaserThresh(fpath);
                            lasthreshval = Math.Ceiling(lasthreshval / 10) * 10 * _Setting.lasthreshvalFactor;//20230329 cwb  
                            //lasthreshval = Math.Ceiling(lasthreshval / 10) * 10 ;
                        }

                        if (SelectIRM.irm[2])
                            StartMTDThread(new ThreadInfo(_DataDir.FullName, t, bars, null, lasthreshval));
                        if (SelectIRM.irm[3])
                            StartMPDThread(new ThreadInfo(_DataDir.FullName, t, bars, null, lasthreshval));
                    }
                }
                if (_ProjectInfo._IsMMTD)
                {

                    if (Directory.Exists(string.Format(@"{0}\IRIMTD\Laser2", _DataDir.FullName)))
                    {

                        if (SelectIRM.irm[2])
                            StartMTDThread(new ThreadInfo(_DataDir.FullName, 2, bars));
                        if (SelectIRM.irm[3])
                            StartMPDThread(new ThreadInfo(_DataDir.FullName, 2, bars));
                    }
                }
            }
            if (_ProjectInfo._IsRut)
            {
                if (SelectIRM.irm[1])
                {
                    switch (_ProjectInfo._RutMode)
                    {
                        case 0: StartRutThread(new ThreadInfo(_DataDir.FullName, 3, bars, _ProjectInfo)); break;
                        case 1: StartRutThread(new ThreadInfo(_DataDir.FullName, 1, bars, _ProjectInfo)); break;
                        case 2: StartRutThread(new ThreadInfo(_DataDir.FullName, 3, bars, _ProjectInfo)); break;
                        default: break;
                    }
                }

                if (SelectIRM.irm[4])
                {
                    if (_ProjectInfo._GeoAlig == 1)
                    {
                        StartGeoAligThread(new ThreadInfo(_DataDir.FullName, 0, bars));
                    }
                }
            }

            while ((_ThreadIRI[0] != null && _ThreadIRI[0].IsAlive)
                || (_ThreadIRI[1] != null && _ThreadIRI[1].IsAlive)
                || (_ThreadMTD[0] != null && _ThreadMTD[0].IsAlive)
                || (_ThreadMTD[1] != null && _ThreadMTD[1].IsAlive)
                || (_ThreadMTD[2] != null && _ThreadMTD[2].IsAlive)
                || (_ThreadMPD[0] != null && _ThreadMPD[0].IsAlive)
                || (_ThreadMPD[1] != null && _ThreadMPD[1].IsAlive)
                || (_ThreadMPD[2] != null && _ThreadMPD[2].IsAlive)
                || (_ThreadRut != null && _ThreadRut.IsAlive)
                || (_ThreadGeoAlig != null && _ThreadGeoAlig.IsAlive))
            {
                Thread.Sleep(1000);
            }

            //MyIRIMTD.ComputePB(_DataDir.FullName, 0, 10);
        }

        Thread[] _ThreadIRI;
        private void StartIRIThread(ThreadInfo prjinfo)
        {
            _ThreadIRI[prjinfo._id] = new Thread(IRIThreadMethod) { IsBackground = true };
            _ThreadIRI[prjinfo._id].Start(prjinfo);
        }
        private void IRIThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinProcessBar bar = tinfo._bar as WinProcessBar;
            MyIRIMTD.ComputeIRI(tinfo._prjname, tinfo._id, bar, _ProjectInfo);

            //// 2018年公路养护标准 且 双平整度 计算路面跳车
            //if (MainForm._ParmStyle == 3 && _ProjectInfo._IsDIRIMTD && _ProjectInfo._IsIRIMTD)
            //{
            //    MyIRIMTD.ComputePB(_DataDir.FullName, tinfo._id, 10);
            //}
            MyIRIMTD.ComputePB(_DataDir.FullName, tinfo._id, 10, _ProjectInfo);
        }

        Thread[] _ThreadMTD;
        private void StartMTDThread(ThreadInfo prjinfo)
        {
            _ThreadMTD[prjinfo._id] = new Thread(MTDThreadMethod) { IsBackground = true };
            _ThreadMTD[prjinfo._id].Start(prjinfo);
        }
        private void MTDThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinProcessBar bar = tinfo._bar as WinProcessBar;
            MyIRIMTD.ComputeMTD(tinfo._prjname, tinfo._id, 10, tinfo._LasThreshVal, bar);
        }

        Thread _ThreadRut;
        private void StartRutThread(ThreadInfo prjinfo)
        {
            _ThreadRut = new Thread(RutThreadMethod) { IsBackground = true };
            _ThreadRut.Start(prjinfo);
        }

        private void RutThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinProcessBar bar = tinfo._bar as WinProcessBar;
            MyRut.ComputeRut(tinfo._prjname, bar, tinfo._id, tinfo._prjinfo._RutMode);
        }

        Thread _ThreadGeoAlig;
        private void StartGeoAligThread(ThreadInfo prjinfo)
        {
            _ThreadGeoAlig = new Thread(GeoAligThreadMethod) { IsBackground = true };
            _ThreadGeoAlig.Start(prjinfo);
        }
        private void GeoAligThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinProcessBar bar = tinfo._bar as WinProcessBar;
            MyGeoAlig.ComputeGeoalig(tinfo._prjname, bar);
        }

        Thread[] _ThreadMPD;
        private void StartMPDThread(ThreadInfo prjinfo)
        {
            _ThreadMPD[prjinfo._id] = new Thread(MPDThreadMethod) { IsBackground = true };
            _ThreadMPD[prjinfo._id].Start(prjinfo);
        }
        private void MPDThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinProcessBar bar = tinfo._bar as WinProcessBar;
            MyIRIMTD.ComputeMPD(tinfo._prjname, tinfo._id, 10, tinfo._LasThreshVal, bar);

        }

        class ThreadInfo
        {
            public int _id;
            public string _prjname;
            public double _LasThreshVal;
            public Form _bar;
            public ProjectInfo _prjinfo;
            public ThreadInfo(string s, int n, Form b, ProjectInfo p = null, double threshval = 50)
            {
                _id = n;
                _prjname = s;
                _bar = b;
                _prjinfo = p;
                _LasThreshVal = threshval;
            }
        }

        /// <summary>
        /// 加载工程信息，并给相关控件赋值
        /// </summary>
        private bool LoadProjInfo()
        {
            CheckSetting();
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            if (_ProjectInfo == null)
            {
                _ProjectInfo = new ProjectInfo(_DataDir.FullName);
            }

            if (((_ProjectInfo._Direction > 0) && _ProjectInfo._StartMile > _ProjectInfo._EndMile)
                || ((_ProjectInfo._Direction < 0) && _ProjectInfo._StartMile < _ProjectInfo._EndMile)
                || (_ProjectInfo._StartMile < 0 || _ProjectInfo._EndMile < 0))
            {
                MessageBox.Show("工程起点桩号和终点桩号不合法，请检查！");
                return false;
            }
            _winproj = new WinProj(_ProjectInfo);
            _winproj.TopLevel = false;
            _winproj.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            dockPanel_proj.FloatSize = new Size(_winproj.Width + 60, _winproj.Height + 60);
            dockPanel_proj.Controls.Add(_winproj);
            _winproj.Dock = DockStyle.Fill;
            _winproj.EventUpdateProjectInfo += new EventHandler(_winproj_EventUpdateProjectInfo);
            _winproj.Show();

            if (Directory.Exists(_DataDir.FullName + "\\RoadImg\\Camera0"))
            {
                Image2Mile("Road", 0, _ProjectInfo._RoadImgDis, "jpg");


                if (_Setting.IsRename)
                {
                    ChangeImageName("Road", 0);
                }


                _winYunGuang = new YGView();
                dockPanel_YunGuang.Controls.Add(_winYunGuang);
                _winYunGuang.Dock = DockStyle.Fill;
                _winYunGuang.EventUpdateImg += new EventHandler(_winYunGuang_EventUpdateImg);

                _winimgfull = new WinImgFull();
                _winimgfull.TopLevel = false;
                _winimgfull.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                dockPanel_FullImg.FloatSize = new Size(_winimgfull.Width + 60, _winimgfull.Height + 60);
                dockPanel_FullImg.Controls.Add(_winimgfull);
                _winimgfull.Dock = DockStyle.Fill;
                _winimgfull.Show();

                if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.DegreeRoad2018)
                {
                    _winroad = new WinRoadNew(_ProjectInfo, _DataDir.FullName);
                }
                else if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
                {
                    _winroad = new WinRoadNew(_ProjectInfo, _DataDir.FullName);
                }
                else if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                {

                    _winroad = new WinRoadNew(_ProjectInfo, _DataDir.FullName);
                }
                else
                {
                    _winroad = new WinRoadImg(_ProjectInfo, _DataDir.FullName);
                }
                _winroad.TopLevel = false;
                _winroad.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                //cwb
                dockPanel_road.FloatSize = new Size(_winroad.Width, _winroad.Height);
                dockPanel_road.Controls.Add(_winroad);


                _winroad.Dock = DockStyle.Fill;
                _winroad.EventChangeType += new EventHandler(_winroad_EventChangeType);
                _winroad.EventUpdateMile += new EventHandler(_winroad_EventUpdateMile);
                _winroad.EventUpdateDmi += new EventHandler(_winroad_EventUpdateDmi);
                _winroad.EventUpdateDis += new EventHandler(_winroad_EventUpdateDis);
                _winroad.EventUpdateYG += new EventHandler(_winroad_EventUpdateYG);
                _winroad.Enter += new EventHandler(_winroad_Enter);
                _winroad.EventUpdateFullImg += new EventHandler(_winroad_EventUpdateFullImg);
                _winroad.EventUpdateFullPoint += new EventHandler(_winroad_EventUpdateFullPoint);

                _winroadis = new WinRoadDisList(_ProjectInfo, _DataDir.FullName);

                _winroad.Show();




                _winroadis.TopLevel = false;
                _winroadis.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                dockPanel_roaddislist.FloatSize = new Size(_winroadis.Width + 60, _winroadis.Height + 60);
                dockPanel_roaddislist.Controls.Add(_winroadis);
                _winroadis.Dock = DockStyle.Fill;
                _winroadis.EventJump2Dis += new EventHandler(_winroadis_EventJump2Dis);
                _winroadis.EventUpdateDis += new EventHandler(_winroadis_EventUpdateDis);
                _winroadis.Show();
            }

            if (Directory.Exists(_DataDir.FullName + "\\StreetImg"))
            {
                Image2Mile("Street", 0, _ProjectInfo._StreetImgDis_Left, "jpg");
                Image2Mile("Street", 1, _ProjectInfo._StreetImgDis_Right, "jpg");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Street", 0);
                    ChangeImageName("Street", 1);
                }
                //20240926 cwb
                if (Directory.Exists(_DataDir.FullName + "\\StreetImg2"))
                {
                    Image2Mile("Street", 0, _ProjectInfo._StreetImgDis_Right, "jpg", "Img2");
                    Image2Mile("Street", 1, _ProjectInfo._StreetImgDis_Right, "jpg", "Img2");
                    if (_Setting.IsRename)
                    {
                        ChangeImageName("Street", 0, "Img2");
                        ChangeImageName("Street", 1, "Img2");
                    }
                }

                //景观病害
                _winstreetdis = new WinStreetDisList();
                _winstreetdis.TopLevel = false;
                _winstreetdis.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                dockPanel_streetdislist.FloatSize = new Size(_winstreetdis.Width + 60, _winstreetdis.Height + 60);
                dockPanel_streetdislist.Controls.Add(_winstreetdis);
                _winstreetdis.Dock = DockStyle.Fill;
                //_winstreetdis.EventJump2Dis += new EventHandler(_winstreetdis_EventJump2Dis);
                //_winstreetdis.EventUpdateDis += new EventHandler(_winroadis_EventUpdateDis);
                _winstreetdis.Show();

                _winstreet = new WinStreetImg(_ProjectInfo, _DataDir.FullName);
                _winstreet.TopLevel = false;
                _winstreet.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                dockPanel_street.FloatSize = new Size(_winstreet.Width + 60, _winstreet.Height + 60);
                dockPanel_street.Controls.Add(_winstreet);
                _winstreet.Dock = DockStyle.Fill;
                _winstreet.EventUpdateMile += new EventHandler(_winstreet_EventUpdateMile);
                _winstreet.EventUpdateDisList += new EventHandler(_winstreet_EventUpdateDisList);
                _winstreet.EventDeleteDis += new EventHandler(_winstreet_EventDeleteDis);
                _winstreet.EventLoadDisList += new EventHandler(_winstreet_EventLoadDisList);
                _winstreet.EventUpdateFullImg += new EventHandler(_winroad_EventUpdateFullImg);
                _winstreet.EventUpdateFullPoint += new EventHandler(_winroad_EventUpdateFullPoint);
                _winstreet.Enter += new EventHandler(_winstreet_Enter);
                _winstreet.Show();
            }

            if (Directory.Exists(_DataDir.FullName + "\\PanoImg"))
            {
                Image2Mile("Pano", 0, _ProjectInfo._PanoImgDis, "jpeg");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Pano", 0);
                }

                _winpano = new WinPanoImg(_ProjectInfo, _DataDir.FullName);
                _winpano.TopLevel = false;
                _winpano.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                dockPanel_pano.FloatSize = new Size(_winpano.Width + 60, _winpano.Height + 60);
                dockPanel_pano.Controls.Add(_winpano);
                _winpano.Dock = DockStyle.Fill;
                _winpano.EventUpdateMile += new EventHandler(_winpano_EventUpdateMile);
                _winpano.Enter += new EventHandler(_winpano_Enter);
                _winpano.Show();
            }

            if (Directory.Exists(_DataDir.FullName + "\\IRIMTD"))
            {
                try
                {
                    _winirm = new WinIRM(_ProjectInfo, _DataDir.FullName);
                    _winirm.TopLevel = false;
                    _winirm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    dockPanel_IRM.Controls.Add(_winirm);
                    _winirm.Dock = DockStyle.Fill;
                    _winirm.Show();
                }
                catch (Exception)
                {


                }

            }

            _winmark = new 采集打标列表(_ProjectInfo, _DataDir.FullName);
            _winmark.TopLevel = false;
            _winmark.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            dockPanel_mark.Controls.Add(_winmark);
            _winmark.Dock = DockStyle.Fill;
            _winmark.EventJumpMark += new EventHandler(_winmark_EventJumpMark);
            _winmark.EventUpdateRoadPart += new EventHandler(_winmark_EventUpdateRoadPart);
            _winmark.EventUpdateProjectInfo += new EventHandler(_winproj_EventUpdateProjectInfo);
            _winmark.Show();

            _winmap = new WinMap(_ProjectInfo, _DataDir.FullName);
            _winmap.TopLevel = false;
            _winmap.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            dockPanel_Map.Controls.Add(_winmap);
            _winmap.Dock = DockStyle.Fill;
            _winmap.Show();

            this.Cursor = System.Windows.Forms.Cursors.Default;

            return true;
        }

        /// <summary>
        /// 加载工程信息，不给相关控件赋值，只有导出报表报告的时候会用到
        /// </summary>
        private bool LoadProjInfoData()
        {
            CheckSetting();
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;

            if (((_ProjectInfo._Direction > 0) && _ProjectInfo._StartMile > _ProjectInfo._EndMile)
                || ((_ProjectInfo._Direction < 0) && _ProjectInfo._StartMile < _ProjectInfo._EndMile)
                || (_ProjectInfo._StartMile < 0 || _ProjectInfo._EndMile < 0))
            {
                MessageBox.Show("工程起点桩号和终点桩号不合法，请检查！");
                return false;
            }

            if (Directory.Exists(_DataDir.FullName + "\\RoadImg\\Camera0"))
            {
                Image2Mile("Road", 0, _ProjectInfo._RoadImgDis, "jpg");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Road", 0);
                }
            }

            if (Directory.Exists(_DataDir.FullName + "\\StreetImg"))
            {
                Image2Mile("Street", 0, _ProjectInfo._StreetImgDis_Left, "jpg");
                Image2Mile("Street", 1, _ProjectInfo._StreetImgDis_Right, "jpg");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Street", 0);
                    ChangeImageName("Street", 1);
                }
            }
            //20240926 cwb
            if (Directory.Exists(_DataDir.FullName + "\\StreetImg2"))
            {
                Image2Mile("Street", 0, _ProjectInfo._StreetImgDis_Right, "jpg", "Img2");
                Image2Mile("Street", 1, _ProjectInfo._StreetImgDis_Right, "jpg", "Img2");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Street", 0, "Img2");
                    ChangeImageName("Street", 1, "Img2");
                }
            }



            if (Directory.Exists(_DataDir.FullName + "\\PanoImg"))
            {
                Image2Mile("Pano", 0, _ProjectInfo._PanoImgDis, "jpeg");
                if (_Setting.IsRename)
                {
                    ChangeImageName("Pano", 0);
                }
            }

            this.Cursor = System.Windows.Forms.Cursors.Default;
            return true;
        }

        void _winroad_EventUpdateFullPoint(object sender, EventArgs e)
        {
            if (_winimgfull != null)
            {
                _winimgfull.UpdateShowImg(sender);

            }
        }

        void _winroad_EventUpdateFullImg(object sender, EventArgs e)
        {
            if (_winimgfull != null)
                _winimgfull.UpdateImg(sender);
        }

        void _winroad_EventUpdateYG(object sender, EventArgs e)
        {
            if (_winYunGuang != null)
            {
                _winYunGuang.InitNewImg();
            }
        }

        void _winYunGuang_EventUpdateImg(object sender, EventArgs e)
        {
            if (_winroad != null)
                _winroad.UpdateYG(sender);
        }

        void _winroadis_EventUpdateDis(object sender, EventArgs e)
        {
            if (_winroad != null)
                _winroad.UpdateDisType(sender);
        }

        void _winroad_EventUpdateDis(object sender, EventArgs e)
        {
            if (_winroadis != null)
                _winroadis.UpDateCurDis(sender);
            //_winroadis.LoadAllDis();
        }

        void _winproj_EventUpdateProjectInfo(object sender, EventArgs e)
        {
            if (_winproj != null)
            {
                _winproj.Dispose();
                _winproj = null;
            }
            if (_winroad != null)
            {
                _winroad.Dispose();
                _winroad = null;
            }
            if (_winstreet != null)
            {
                _winstreet.Dispose();
                _winstreet = null;
            }
            if (_winmark != null)
            {
                _winmark.Dispose();
                _winmark = null;
            }
            if (_winmap != null)
            {
                _winmap.Dispose();
                _winmap = null;
            }
            if (_winirm != null)
            {
                _winirm.Dispose();
                _winirm = null;
            }

            if (_winYunGuang != null)
            {
                _winYunGuang.Dispose();
                _winYunGuang = null;
            }
            if (_winimgfull != null)
            {
                _winimgfull.Dispose();
                _winimgfull = null;
            }


            if (_winroadis != null)
            {
                _winroadis.Dispose();
                _winroadis = null;
            }



            if (_winstreetdis != null)
            {
                _winstreetdis.Dispose();
                _winstreetdis = null;
            }

            if (_winpano != null)
            {
                _winpano.Dispose();
                _winpano = null;
            }





            if (_ProjectInfo != null)
            {
                _ProjectInfo = null;
            }
            if (LoadProjInfo() == false) return;
        }

        void _winstreet_Enter(object sender, EventArgs e)
        {
            if (_winroad != null) _winroad._IsActivated = false;
            if (_winstreet != null) _winstreet._IsActivated = true;
            if (_winpano != null) _winpano._IsActivated = false;
        }

        void _winpano_Enter(object sender, EventArgs e)
        {
            if (_winroad != null) _winroad._IsActivated = false;
            if (_winstreet != null) _winstreet._IsActivated = false;
            if (_winpano != null) _winpano._IsActivated = true;
        }
        void _winroad_Enter(object sender, EventArgs e)
        {
            if (_winroad != null) _winroad._IsActivated = true;
            if (_winstreet != null) _winstreet._IsActivated = false;
            if (_winpano != null) _winpano._IsActivated = false;
        }

        #region 景观病害
        void _winstreet_EventDeleteDis(object sender, EventArgs e)
        {
            _winstreetdis.DeleteDis(sender, e);
        }

        void _winstreet_EventUpdateDisList(object sender, EventArgs e)
        {
            _winstreetdis.UpdateDisList(sender, e);
        }

        void _winstreet_EventLoadDisList(object sender, EventArgs e)
        {
            _winstreetdis.LoadDisList(sender, e);
        }
        #endregion

        void _winstreet_EventUpdateMile(object sender, EventArgs e)
        {
            if (MainForm._IsLinkShow)
            {
                if (_winroad != null && _winroad._IsInitLoad && _winstreet._IsActivated)
                {
                    _winroad.ShowJumpImg(Convert.ToDouble(sender));
                }
                if (_winpano != null && _winpano._IsInitLoad && _winstreet._IsActivated)
                {
                    _winpano.ShowJumpImg(Convert.ToDouble(sender));
                }
            }
        }

        void _winpano_EventUpdateMile(object sender, EventArgs e)
        {
            if (MainForm._IsLinkShow)
            {
                if (_winroad != null && _winroad._IsInitLoad && _winpano._IsActivated)
                {
                    _winroad.ShowJumpImg(Convert.ToDouble(sender));
                }
                if (_winstreet != null && _winstreet._IsInitLoad && _winpano._IsActivated)
                {
                    _winstreet.ShowJumpImg(Convert.ToDouble(sender));
                }
            }
        }

        void _winroadis_EventJump2Dis(object sender, EventArgs e)
        {
            _winroad.ShowJumpImg(Convert.ToDouble(sender));
        }

        void _winroad_EventUpdateDmi(object sender, EventArgs e)
        {
            if (MainForm._IsLinkShow)
            {
                if (_winirm != null)
                {
                    _winirm.ShowVal(Convert.ToDouble(sender));
                }
            }
        }

        void _winroad_EventUpdateMile(object sender, EventArgs e)
        {
            if (MainForm._IsLinkShow)
            {
                if (_winstreet != null && _winstreet._IsInitLoad && _winroad._IsActivated)
                {
                    _winstreet.ShowJumpImg(Convert.ToDouble(sender));
                }
                if (_winpano != null && _winpano._IsInitLoad && _winroad._IsActivated)
                {
                    _winpano.ShowJumpImg(Convert.ToDouble(sender));
                }
            }
        }

        void _winroad_EventChangeType(object sender, EventArgs e)
        {
            _winmark.LoadAllMark(true);
        }

        //可以共用
        void _winmark_EventUpdateRoadPart(object sender, EventArgs e)
        {
            _winroad.GetTypeMilePart(_DataDir.FullName, _ProjectInfo._Direction);
        }

        void _winmark_EventJumpMark(object sender, EventArgs e)
        {
            if (_winroad._IsInitLoad)
            {
                _winroad.ShowJumpImg(int.Parse(sender.ToString()));
            }
        }
        public void GenerateExcel_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (!_Setting.needSub)
                if (LoadProjInfoData() == false) return;

            switch (_Setting.ParmStyle)
            {
                case StandardParmType.DegreeRoad2018:
                    {
                        switch (_Setting.SelectDrawDis)//1
                        {
                            case 0: GenerateExcel_Degree2018_BigRect_Street(excelApp, xlspath, xlslen, IsOutputxls); break;
                            case 1: GenerateExcel_Degree2018_SmallRect_Street(excelApp, xlspath, xlslen, IsOutputxls); break;
                            default: break;
                        }
                    }
                    break;
                case StandardParmType.RuralRoadGuangxi: GenerateExcel_GuangXi_Street(excelApp, xlspath, xlslen, IsOutputxls); break;
                case StandardParmType.RuralRoadChongqing: GenerateExcel_ChongQing_Street(excelApp, xlspath, xlslen, IsOutputxls); break;
                case StandardParmType.RuralRoadlowLevel:
                    {
                        switch (_Setting.SelectDrawDis)
                        {

                            case 0:

                                GenerateExcel_low_Street(excelApp, xlspath, xlslen, IsOutputxls);

                                break;
                            case 1:

                                GenerateExcel_Village_Street_Small(excelApp, xlspath, xlslen, IsOutputxls);

                                break;
                            default: break;
                        }
                    }
                    break;
                case StandardParmType.RuralRoadHunan:

                    {
                        switch (_Setting.SelectDrawDis)
                        {
                            case 0:
                                GenerateExcel_HuNan_Street(excelApp, xlspath, xlslen, IsOutputxls);

                                break;
                            case 1:

                                GenerateExcel_hn_Street_Small(excelApp, xlspath, xlslen, IsOutputxls);

                                break;
                            default: break;
                        }


                    }
                    break;


                default: break;
            }
        }
        public void GenerateExcel_Degree2018_BigRect_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelDegree2018.InitStreetData(_ProjectInfo, _DataDir);

            #region 孝感
            if (_Setting.ExcelType == 19)
            {
                ///孝感定制
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, true, false, false))
                {
                    MyExcelDegree2018.OutputStreetDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelDegree2018.OutputRoadBedDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

            }

            #endregion

            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, false, false, false))
                    {
                        MyExcelDegree2018.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelDegree2018.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelDegree2018.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelDegree2018.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[6])
            { 
                MyExcelDegree2018.OutputStreetAllDis(excelApp, xlspath, _ProjectInfo, _DataDir);
              
            }

        }
        public static List<char> chars = new List<char> { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'j', 'k', 'L', 'M', 'N', 'O', 'P', 'Q' };
        public void GenerateExcel_Degree2018_SmallRect_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {

            #region 孝感
            if (_Setting.ExcelType == 16)
            {
                ///孝感定制
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, true, false, false))
                {
                    MyExcelDegreeSmall2018.OutputStreetDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelDegreeSmall2018.OutputRoadBedDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

            }

            #endregion

            MyExcelDegreeSmall2018.InitStreetData(_ProjectInfo, _DataDir);
            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, false, false, false))
                    {
                        MyExcelDegreeSmall2018.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelDegreeSmall2018.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelDegreeSmall2018.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelDegreeSmall2018.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[6])
            {
                MyExcelDegreeSmall2018.OutputStreetAllDis(excelApp, xlspath, _ProjectInfo, _DataDir);

            }
        }
        public void GenerateExcel_GuangXi_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelGXDegree.InitStreetData(_ProjectInfo, _DataDir);
            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelGXDegree.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelGXDegree.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelGXDegree.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelGXDegree.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
        }
        public void GenerateExcel_ChongQing_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelCQDegree.InitStreetData(_ProjectInfo, _DataDir);
            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelCQDegree.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelCQDegree.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelCQDegree.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelCQDegree.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
        }

        public void GenerateExcel_low_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelVillageDegree.InitStreetData(_ProjectInfo, _DataDir);
            if (_Setting.ExcelType == 3)
            {
                ///孝感定制
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, true, false, false))
                {
                    MyExcelVillageDegree.OutputStreetDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelVillageDegree.OutputRoadBedDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelVillageDegree.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelVillageDegree.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelVillageDegree.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelVillageDegree.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[4])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], false, false, false, false, false, false))
                    {
                        MyExcelVillageDegree.OutputRoadBedDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[5])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], false, false, false, true, false, false))
                    {
                        MyExcelVillageDegree.OutputStreetDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }

            if (IsOutputxls[6])
            {

                MyExcelVillageDegree.OutputStreetAllDis(excelApp, xlspath, _ProjectInfo, _DataDir);
            }
        }
        public void GenerateExcel_HuNan_Street(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelHNDegree.InitStreetData(_ProjectInfo, _DataDir);

            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelHNDegree.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelHNDegree.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelHNDegree.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelHNDegree.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[4])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], false, false, false, false, false, false))
                    {
                        MyExcelHNDegree.OutputRoadBedDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[5])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], false, false, false, true, false, false))
                    {
                        MyExcelHNDegree.OutputStreetDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
        }

        private void GenerateExcel_Degree2007(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            //各单项出表
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, true))
                            MyExcelDegree2007.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false))
                            MyExcelDegree2007.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false))
                            MyExcelDegree2007.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false))
                            MyExcelDegree2007.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false))
                            MyExcelDegree2007.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, true, false))
                            MyExcelDegree2007.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false))
                            MyExcelDegree2007.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
            //桂兴达报表
            else if (_Setting.ExcelType == 1)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, false))
                        MyExcelDegree2007.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //中南安环报表
            else if (_Setting.ExcelType == 2)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, false))
                    {
                        MyExcelDegree2007.OutputZNRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegree2007.OutputZNRoadDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegree2007.OutputZNDataRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //中交国通报表
            else if (_Setting.ExcelType == 3)
            {
                int[] lenval = { 5, 10, 20, 100, 1000 };
                for (int i = 0; i < lenval.Length; ++i)
                {
                    bool tflag = false;
                    if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000)
                        || IsOutputxls[5] && lenval[i] == 1000)
                    {
                        tflag = true;
                    }
                    if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, lenval[i], tflag, IsOutputxls[2] || IsOutputxls[5],
                         IsOutputxls[1] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5]))
                    {
                        if (IsOutputxls[0] && (lenval[i] == 10 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2007.OutputZJGTRut(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[1] && (lenval[i] == 10 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2007.OutputZJGTMTD(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[2] && (lenval[i] == 20 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2007.OutputZJGTIRI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2007.OutputZJGTDis(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[4] && (lenval[i] == 5 || lenval[i] == 1000))
                        {
                            MyExcelDegree2007.OutputZJGTGPS(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[5] && lenval[i] == 1000)
                        {
                            MyExcelDegree2007.OutputZJGTPQI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        MyExcelDegree2007.OutputZJGTRoadType(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //带GPS的重庆招商局报表模板
            else if (_Setting.ExcelType == 4)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, false))
                    {
                        MyExcelDegree2007.OutputGPSRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegree2007.OutputGPSStreetImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        MyExcelDegree2007.OutputGPSRoadImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

            //奥路通
            else if (_Setting.ExcelType == 5)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, false))
                    {
                        MyExcelDegree2007.OutputALTDIS(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
        }
        private void GenerateExcel_City(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            //带GPS
            if (_Setting.ExcelType == 4)
            {
                if (_Setting.PartType == 0)
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, false, _Setting.PartType, false, true))
                    {
                        MyExcelCity.OutputGPSAll2Xls_2(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, 100, true, true, true, true, false, _Setting.PartType, false, true))
                    {
                        MyExcelCity.OutputGPSAll2Xls_2(excelApp, xlspath, _ProjectInfo, _DataDir, 100);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, true, true, false, _Setting.PartType, false, true))
                    {
                        MyExcelCity.OutputGPSAll2Xls_2(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                else if (_Setting.PartType == 1)
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, true, true, false, _Setting.PartType, false, true))
                    {
                        MyExcelCity.OutputGPSAll2Xls_2_Dmi(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 5)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, _Setting.PartType, false, false))
                    {
                        MyExcelCity.OutputALTDIS(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);

                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 6) //模板1
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, true, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputZYPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            else if (_Setting.ExcelType == 13)
            {
                //rut 0.1m
                if (IsOutputxls[0])
                {
                    if (MyExcelCity.InitProDataD(_DataDir, _ProjectInfo, 0.1f, true, true, true, true, true, true))
                    {
                        MyExcelCity.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, 0.1);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

            }
            else if (_Setting.ExcelType == 12)
            {

                if (IsOutputxls[0])
                {

                    //10
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, 1, true, true, false, false, false, _Setting.PartType, false, false))
                    {
                        MyExcelCity.OutputDis_HPcsv_0(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, true, false, false, false, _Setting.PartType, false, false))
                        {

                            MyExcelCity.OutputDis_HPcsv_IRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
            }
            else
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, true, _Setting.PartType, false, false))
                            MyExcelCity.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, _Setting.PartType, true, false))
                            MyExcelCity.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, true, false, false, _Setting.PartType, false, false))
                            MyExcelCity.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //上海惠普客户 20250401需求
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, xlslen[7][i], false, true, false, false, false, _Setting.PartType, true, false))
                            MyExcelCity.OutputIRI_WithSpeed(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }
        private void GenerateExcel_BeiJin(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelBJDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true))
                            MyExcelBJDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelBJDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false))
                            MyExcelBJDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelBJDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false))
                            MyExcelBJDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelBJDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true))
                            MyExcelBJDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelBJDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i] / 10, true, false))
                            MyExcelBJDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }
        private void GenerateExcel_LiaoNing(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true))
                            MyExcelLNDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false))
                            MyExcelLNDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false))
                            MyExcelLNDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i] / 10, true, false))
                            MyExcelLNDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true))
                            MyExcelLNDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelLNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true))
                            MyExcelLNDegree.OutputALT(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }


        //private void GenerateExcel_JiangXi(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        //{
        //    if (_Setting.ExcelType == 0)
        //    {
        //        if (IsOutputxls[2])
        //        {
        //            for (int i = 0; i < xlslen[2].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
        //                    MyExcelGXDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[3])
        //        {
        //            for (int i = 0; i < xlslen[3].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
        //                    MyExcelGXDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[4])
        //        {
        //            for (int i = 0; i < xlslen[4].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
        //                    MyExcelGXDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[5])
        //        {
        //            for (int i = 0; i < xlslen[5].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
        //                    MyExcelGXDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[7])
        //        {
        //            for (int i = 0; i < xlslen[7].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
        //                    MyExcelGXDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[8])
        //        {
        //            MyExcelGXDegree.InitStreetData(_ProjectInfo, _DataDir);
        //            for (int i = 0; i < xlslen[8].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
        //                    MyExcelGXDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //    }
        //    //桂兴达
        //    else if (_Setting.ExcelType == 1)
        //    {
        //        if (IsOutputxls[0])
        //        {
        //            for (int i = 0; i < xlslen[0].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, true))
        //                    MyExcelGXDegree.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //        if (IsOutputxls[1])
        //        {
        //            for (int i = 0; i < xlslen[1].Length; ++i)
        //            {
        //                if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
        //                    MyExcelGXDegree.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
        //                else
        //                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
        //            }
        //        }
        //    }
        //}

        private void GenerateExcel_GuangXi(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelGXDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelGXDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                            MyExcelGXDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelGXDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                            MyExcelGXDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelGXDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
                            MyExcelGXDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //桂兴达
            else if (_Setting.ExcelType == 1)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, true))
                            MyExcelGXDegree.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelGXDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
                            MyExcelGXDegree.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }


        private void GenerateExcel_JiangXi(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelCQDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelCQDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                            MyExcelCQDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelCQDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                            MyExcelCQDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelCQDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
                            MyExcelCQDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }


        private void GenerateExcel_ChongQing(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelCQDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelCQDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                            MyExcelCQDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelCQDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                            MyExcelCQDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelCQDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelCQDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
                            MyExcelCQDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }




        private void GenerateExcel_VillageSmall(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {


                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {

                        //(5211)路面破损评定汇总表
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegreeSmall.OutputPci_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        //(5211)路面平整度评定汇总表
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, false, false, false, true))
                        {
                            MyExcelVillageDegreeSmall.OutputRqi_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }


                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        try
                        {
                            if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                                MyExcelVillageDegreeSmall.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (System.Exception)
                        {

                        }


                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelVillageDegreeSmall.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false)) //倒数第三个为PBI计算 改为false
                            MyExcelVillageDegreeSmall.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegreeSmall.OutputDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {

                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                        {
                            MyExcelVillageDegreeSmall.OutputPQI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    MyExcelVillageDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false))
                            {
                                MyExcelVillageDegreeSmall.OutputPDMX_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (VillageIndexException ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }


                    }
                }

            }
            else if (_Setting.ExcelType == 1)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, false, false))
                            MyExcelVillageDegreeSmall.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
                            MyExcelVillageDegreeSmall.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            else if (_Setting.ExcelType == 2)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_0(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //沥青破损
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputLQDamage(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //水泥破损
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, false, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputSNDamage(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], false, true, false, false, false, true))
                            MyExcelVillageDegreeSmall.outPutAutoTest_1(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }

                //平整度原始数据
                if (IsOutputxls[4])
                {

                    {

                        if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1F, false, true, false, false, true, true))
                            MyExcelVillageDegreeSmall.outPutAutoTest_2(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //  DR破损率csv检测结果数据表格
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, false, false, false, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_5(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //IRI平整度csv检测结果数据表格
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], true, true, false, false, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_6(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //IRI平整度csv检测原始数据表格
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[7][i], true, true, false, false, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_7(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //空间定位数据csv检测原始数据表格
                if (IsOutputxls[8])
                {
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], false, false, false, false, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_8(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
            else if (_Setting.ExcelType == 4)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, false, false, false))
                            MyExcelVillageDegreeSmall.outPutAutoTest_HN0(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //水泥
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegreeSmall.OutputSNDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //沥青
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegreeSmall.OutputLQDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //病害流水
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegreeSmall.OutputAllDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //重庆 公里指标导入模板
                if (IsOutputxls[4])
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 1000, true, true, false, false, false, true))
                        MyExcelVillageDegreeSmall.OutputChongQingSumExcel(excelApp, xlspath, _ProjectInfo, _DataDir, 1000);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");


                }

                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputChongQingDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }

            else if (_Setting.ExcelType == 7)
            {

                //合肥
                if (IsOutputxls[0])
                {
                    _Setting.hefeiOutExcel2 = true;
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        //首先读取资产配置表
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelVillageDegreeSmall.OutputRoad_Hefei(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.hefeiOutExcel2 = false;
                }

            }


            else if (_Setting.ExcelType == 8)
            {
                if (IsOutputxls[0])
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 2, true, false, false, false, false, false))
                        MyExcelVillageDegreeSmall.OutputDis_TH(excelApp, xlspath, _ProjectInfo, _DataDir, 2);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
            }
            else if (_Setting.ExcelType == 9)
            {


            }
            if (_Setting.ExcelType == 10)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, true, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, true))
                            MyExcelVillageDegreeSmall.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
        }
        /// <summary>
        /// 就地转换xlsx到csv
        /// </summary>
        /// <param name="sourceFile"></param>
        public static void XlsxToCsv(string sourceFile)
        {
            string outFile = System.IO.Path.GetFileNameWithoutExtension(sourceFile) + ".csv";
            string temp = System.IO.Path.GetFileName(sourceFile);
            outFile = sourceFile.Replace(temp, outFile);
             Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(sourceFile);
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
            //if (!File.Exists(outFile))
            // {
            sheet.SaveToFile(outFile, ",", Encoding.UTF8);
            File.Delete(sourceFile);
            //}
            if (File.Exists(sourceFile))
            {
                File.Delete(sourceFile);
            }
        }




        public void Convent_Village_HeBei(MSExcel.Application excelApp, string BumpPath, string riPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 1)
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    if (!_Setting.isGDIriCalculate)
                    {
                        MyExcelVillageDegreeSmall.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    }
                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    // MyExcelVillageDegreeSmall.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1f, true, true, true, true, true, true))
                {
                    // MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //  if (!_Setting.isGDIriCalculate)
                    MyExcelVillageDegree.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true))
                {
                    MyExcelVillageDegree.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }
        public void Convent_Village_JiangXi(MSExcel.Application excelApp, string BumpPath, string riPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //  if (!_Setting.isGDIriCalculate)
                    MyExcelVillageDegree.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Iri_JiangXi(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Damage(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true))
                {
                    MyExcelVillageDegree.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    if (!_Setting.isGDIriCalculate)
                    {
                        MyExcelVillageDegreeSmall.Convent_Bump_JiangXi(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    }

                    MyExcelVillageDegreeSmall.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegreeSmall.Convent_Damage(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri_JiangXi(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    // MyExcelVillageDegreeSmall.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1f, true, true, true, true, true, true))
                {
                    // MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri_Original(excelApp, riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

            }
            excelApp.Quit();
            //定位信息

        }


        public void Convent_Village2023(MSExcel.Application excelApp, string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
         string BumpPath, string mpdPath, string textFile, string tpFile, string proName)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelVillageDegree.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);

                    //平整度
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                    //磨耗
                    MyExcelVillageDegree.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    //MyExcelVillageDegree.Convent_TT2023(excelApp, textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelVillageDegree.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelVillageDegree.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelVillageDegree.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //空间定位
                    MyExcelVillageDegree.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                    }
                    _Setting.banMarkSign = false;
                }

                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    if (_ProjectInfo._IsRut)
                        MyExcelVillageDegree.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    // MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;

            }
            else
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelVillageDegreeSmall.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);

                    //平整度
                    MyExcelVillageDegreeSmall.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    //MyExcelVillageDegreeSmall.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelVillageDegreeSmall.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName);

                    //磨耗
                    MyExcelVillageDegreeSmall.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {

                    //空间定位
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelVillageDegreeSmall.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    if (_ProjectInfo._IsRut)
                        MyExcelVillageDegreeSmall.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    // MyExcelVillageDegreeSmall.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            excelApp.Quit();
            //定位信息

        }
        public void Convent_Village2024_HeBei(string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
       string BumpPath, string textFile, string tpFile, string proName, Encoding encoding)
        {
            if (_Setting.SelectDrawDis == 0)
            {

                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    if (_ProjectInfo._IsRut)
                    {
                        MyExcelVillageDegree.Convent_Rut2024_ChongQing(rdFilePath, _ProjectInfo, _DataDir, proName);

                    }

                    //平整度
                    MyExcelVillageDegree.Convent_Iri2024_ChongQing(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                    //磨耗
                    //MyExcelVillageDegree.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    //MyExcelVillageDegree.Convent_TT2023(excelApp, textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelVillageDegree.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelVillageDegree.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {

                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                    }

                }


                if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    if (_ProjectInfo._IsRut)
                        MyExcelVillageDegree.Convent_TP2024_ChongQing(tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    // MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //空间定位
                    MyExcelVillageDegree.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //破损
                    MyExcelVillageDegree.Convent_Damage2024_ChongQing(bhPath, _ProjectInfo, _DataDir, 10, proName, ".txt", encoding);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }


            //定位信息

        }


        public void Convent_Village2024_ChongQing(string riFilePath, string lbiPath, string bhPath, string iriPath, string proName, Encoding encoding)
        {
            if (_Setting.SelectDrawDis == 0)
            {

                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    ////车辙
                    //if (_ProjectInfo._IsRut)
                    //{
                    //    MyExcelVillageDegree.Convent_Rut2024_ChongQing(rdFilePath, _ProjectInfo, _DataDir, proName);

                    //}

                    //平整度
                    MyExcelVillageDegree.Convent_Iri2024_ChongQing(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                    //磨耗
                    // MyExcelVillageDegree.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    //MyExcelVillageDegree.Convent_TT2023(excelApp, textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelVillageDegree.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    //  MyExcelVillageDegree.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName); 
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {

                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName, ".txt");
                    }

                }


                if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    //if (_ProjectInfo._IsRut)
                    // MyExcelVillageDegree.Convent_TP2024_ChongQing( tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    // MyExcelVillageDegree.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //空间定位
                    MyExcelVillageDegree.Convent_Lbi2024_ChongQing(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //破损
                    MyExcelVillageDegree.Convent_Damage2024_ChongQing(bhPath, _ProjectInfo, _DataDir, 10, proName, ".txt", encoding);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

        }


        public void Convent_Village2023(MSExcel.Application excelApp, string riPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 1)
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                    MyExcelVillageDegree.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {

                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, ".txt");

                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, ".txt");
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {

                    MyExcelVillageDegree.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            excelApp.Quit();
            //定位信息

        }


        public void Convent_Village2024_LiaoNing(MSExcel.Application excelApp, string riPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 1)
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri2024_LiaoNing(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Iri2024_LiaoNing(iriPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegree.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {

                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, ".txt");

                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, ".txt");
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {

                    MyExcelVillageDegree.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");



                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            excelApp.Quit();
            //定位信息

        }



        public void Convent_2018(MSExcel.Application excelApp, string BumpPath, string rdFilePath, string riFilePath, string SFCPath, string SSRPath, string TextPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegree2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegree2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);

                    //破损
                    MyExcelDegree2018.Convent_Damage(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {

                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LP(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);

                    //tp高程  车辙
                    MyExcelDegree2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegreeSmall2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    //MyExcelDegreeSmall2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName); 
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {
                    MyExcelDegreeSmall2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    MyExcelDegreeSmall2018.Convent_LP(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }

        public void Convent_2018_HeBei(MSExcel.Application excelApp, string BumpPath, string rdFilePath, string riFilePath, string SFCPath, string SSRPath, string TextPath, string lbiPath, string bhPath, string iriPath, string proName, Encoding en)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegree2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegree2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);

                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {

                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LP_hebei(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);

                    //tp高程  车辙
                    MyExcelDegree2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegreeSmall2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    //MyExcelDegreeSmall2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName); 
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {
                    MyExcelDegreeSmall2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    MyExcelDegreeSmall2018.Convent_LP(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }
        public void Convent_2018_JiangXi(MSExcel.Application excelApp, string BumpPath, string rdFilePath, string riFilePath, string SFCPath, string SSRPath, string TextPath, string lbiPath, string bhPath, string iriPath, string proName)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {


                    //跳车
                    MyExcelDegree2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);

                    //平整度
                    MyExcelDegree2018.Convent_Iri_JiangXi(excelApp, iriPath, _ProjectInfo, _DataDir, proName);

                    //破损
                    MyExcelDegree2018.Convent_Damage_JiangXi(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);



                    //车辙
                    MyExcelDegree2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);

                    //磨耗
                    MyExcelDegree2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);

                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {

                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LP(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);

                    //tp高程  车辙
                    MyExcelDegree2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true, true))
                {

                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri_JiangXi(excelApp, iriPath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage_JiangXi(excelApp, bhPath, _ProjectInfo, _DataDir, 10, proName);



                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi(excelApp, lbiPath, _ProjectInfo, _DataDir, proName);

                    //磨耗
                    MyExcelDegreeSmall2018.Convent_Mpd(excelApp, TextPath, _ProjectInfo, _DataDir, proName);

                    //加速度
                    //MyExcelDegreeSmall2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName); 
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, true, true, true, true, true))
                {
                    MyExcelDegreeSmall2018.Convent_TP(excelApp, rdFilePath, _ProjectInfo, _DataDir, proName);
                    MyExcelDegreeSmall2018.Convent_LP(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }

        public void Convent_2024_HuNan(string lbiPath, string iriPath, string rdFilePath,
            string BumpPath, string proName, string suff)
        {
            if (_Setting.SelectDrawDis == 0)
            {//大框
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut2024_HuNan(rdFilePath, _ProjectInfo, _DataDir, proName, suff);

                    //平整度
                    MyExcelDegree2018.Convent_Iri2024_HuNan(iriPath, _ProjectInfo, _DataDir, proName, suff);

                    //跳车
                    MyExcelDegree2018.Convent_Bump2024_HuNan(BumpPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 5, false, false, false, false, false, false))
                {   //GPS 5米一出
                    MyExcelDegree2018.Convent_Lbi2024_HuNan(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {//小框
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, true, true, true))
                {
                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut2024_HuNan(rdFilePath, _ProjectInfo, _DataDir, proName, suff);

                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2024_HuNan(iriPath, _ProjectInfo, _DataDir, proName, suff);

                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump2024_HuNan(BumpPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 5, false, false, false, false, false, false))
                {   //GPS 5米一出
                    MyExcelDegreeSmall2018.Convent_Lbi2024_HuNan(lbiPath, _ProjectInfo, _DataDir, proName, ".txt");
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

            //定位信息
        }

        public void Convent_2023(MSExcel.Application excelApp, string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
            string BumpPath, string mpdPath, string textFile, string tpFile, string proName, Encoding en)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegree2018.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    //MyExcelDegree2018.Convent_TT2023(excelApp, textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegree2018.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    //MyExcelDegreeSmall2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump2023(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);

                    //磨耗
                    MyExcelDegreeSmall2018.Convent_Mpd2023(excelApp, mpdPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegreeSmall2018.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }
        public void Convent_2024(string riFilePath, string lbiPath, string bhPath, string iriPath, string proName, Encoding en)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                _Setting.banMarkSign = true;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2024(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //破损
                    MyExcelDegree2018.Convent_Damage2024(bhPath, _ProjectInfo, _DataDir, 10, proName, en);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                _Setting.banMarkSign = true;
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {

                    MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2024(bhPath, _ProjectInfo, _DataDir, 10, proName, en);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //定位信息

        }
        public void Convent_Standard(Dictionary<string, DirectoryInfo> ConverDic, string proName, Encoding en, string suff = ".txt")
        {

        }
        public void Convent_Village2024(string riPath,
            string lbiPath, string bhPath,
            string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {

            {

                _Setting.banMarkSign = true;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, suff);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Damage2024(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

        }


        public void Convent_Village(MSExcel.Application excelApp, string riPath,
           string lbiPath, string bhPath,
           string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {

            {

                _Setting.banMarkSign = true;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, suff);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Damage2024(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }



        public void Convent_Village2024_AnHui2(MSExcel.Application excelApp, string riPath,
            string lbiPath, string bhPath,
            string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {
            {

                _Setting.banMarkSign = true;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }



                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, suff);
                    MyExcelVillageDegree.Convent_Damage2024_AnHui(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }



            excelApp.Quit();
            //定位信息

        }

        private static Dictionary<string, object?> CaptureAllStaticValues(Type type)
        {
            var dict = new Dictionary<string, object?>();

            // 静态字段（包括 private static 和 readonly static）
            var fields = type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                .Where(f => !f.IsLiteral); // 排除 const

            foreach (var field in fields)
            {
                string fullName = $"{type.Name}.{field.Name}";

                // 跳过所有包含 Setting/Config 的字段（避免序列化巨大配置对象）
                if (fullName.Contains("Setting") || fullName.Contains("Config"))
                    continue;

                try
                {
                    var value = field.GetValue(null);

                    // 防止委托、事件、函数指针等无法序列化的东西炸掉
                    if (value is Delegate || value is MulticastDelegate || value?.GetType().IsPointer == true)
                    {
                        dict[fullName] = $"<Delegate/FunctionPointer: {value}>";
                    }
                    else
                    {
                        dict[fullName] = value;
                    }
                }
                catch
                {
                    dict[fullName] = "<无法访问或序列化>";
                }
            }

            // 静态属性（可选，如果你也有重要静态属性也可以加进来）
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                .Where(p => p.CanRead && p.GetGetMethod(true)?.IsStatic == true);

            foreach (var prop in properties)
            {
                string fullName = $"{type.Name}.{prop.Name}";
                if (fullName.Contains("Setting") || fullName.Contains("Config"))
                    continue;

                try
                {
                    dict[fullName] = prop.GetValue(null);
                }
                catch
                {
                    dict[fullName] = "<属性无法读取>";
                }
            }

            return dict;
        }
        private static readonly JsonSerializerOptions JsonOpts = new()
        {
            WriteIndented = true,
            IncludeFields = true,                             // 必须！才能认 [JsonInclude] 的字段
            ReferenceHandler = ReferenceHandler.Preserve,     // 支持循环引用，防止炸
            DefaultIgnoreCondition = JsonIgnoreCondition.Never
        };
        /// <summary>
        /// 最终版：验证计算结果是否被意外修改（带自动生成/严格比对）
        /// </summary>
        public bool VerifyCalculationResults(StandardParmType standard, int drawType, string dirPath)
        {
            _Setting.ReadData();
            _Setting.SelectDrawDis = drawType;
            _Setting.ParmStyle = standard;
            _Setting.zcSplit = false; 
            _Setting.hefeiOutExcel2 = false;
            _Setting.banMarkSign = false;
            _Setting.is5211MergeArea500 = false;
            _Setting.OutRut = 0;
            //_RoadConfig.DetectWidth = 3.2;
            //_RoadConfig.RealWidth = 3.2;
            //_RoadConfig.RealHeight = 2.0;
            _Setting.hasCamsetting = true;
            Type targetType = null;

            if (_Setting.SelectDrawDis == 1)
            {
                RoadDiseaseTypes.LoadAutoDectRoadDisParm();
            }
            else
            {
                RoadDiseaseTypes.LoadRoadDisParm();
            }

            DiseaseTypes.LoadStreetDisParm();
            DiseaseTypes.LoadRoadBedDisParm();

            // 你的初始化逻辑（保持不变）
            switch (drawType)
            {
                case 0: // 大框
                    if (standard == StandardParmType.DegreeRoad2018)
                    {
                        
                        MyExcelDegree2018.LoadXlsParm();
                        MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true);
                        targetType = typeof(MyExcelDegree2018);
                    }
                    else if (standard == StandardParmType.RuralRoadlowLevel)
                    {
                        MyExcelVillageDegree.LoadXlsParm();
                        MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true);
                        targetType = typeof(MyExcelVillageDegree);
                    }
                    else if (standard == StandardParmType.CityRoad)
                    {
                        MyExcelCity.LoadXlsParm();
                        MyExcelCity.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, false, 0, false, true);
                        targetType = typeof(MyExcelCity);

                    }
                     break;

                case 1: // 小框
                    if (standard == StandardParmType.DegreeRoad2018)
                    {
                        MyExcelDegreeSmall2018.LoadXlsParm(); // 如果有的话
                        MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true);
                        targetType = typeof(MyExcelDegreeSmall2018);
                    }
                    else if (standard == StandardParmType.RuralRoadlowLevel)
                    {
                        MyExcelVillageDegreeSmall.LoadXlsParm();
                        MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true);
                        targetType = typeof(MyExcelVillageDegreeSmall);
                    }
                    break;
            }

            if (targetType == null) return false;

            string standardStr = standard.ToString();
            string drawTypeStr = drawType == 0 ? "大框" : "小框";
            string modelPath = System.IO.Path.Combine(dirPath, $"ModelData_{standardStr}_{drawTypeStr}.json");

            // 关键修复：这里必须传 targetType，而不是 typeof(Type)！！！
            var currentValues = CaptureAllStaticValues(targetType);

            // 自动比对或生成快照
            if (File.Exists(modelPath))
            {
                string expectedJson = File.ReadAllText(modelPath);
                string currentJson = System.Text.Json.JsonSerializer.Serialize(currentValues, JsonOpts);

                if (currentJson != expectedJson)
                {
                    // 用文件对比工具可以直接看到差异
                    File.WriteAllText(modelPath + ".current", currentJson); // 方便你 diff

                    throw new InvalidOperationException(
                        $"""
                【严重警告】核心计算参数已变更！禁止提交！
                文件：{modelPath}
                如果你确实要修改参数，请删除上面的 json 文件后重新运行一次生成新快照。

                当前值已临时保存为：{modelPath}.current
                """);
                }

                Console.WriteLine($"校验通过：{System.IO.Path.GetFileName(modelPath)} 未发生变更");
            }
            else
            {
                // 第一次运行：生成基准快照
                Directory.CreateDirectory(dirPath);
                File.WriteAllText(modelPath,
                    System.Text.Json.JsonSerializer.Serialize(currentValues, JsonOpts));
                Console.WriteLine($"已生成参数快照（以后改了就报错）：{modelPath}");
            }

            return true;
        }

        public void Convent_Village2024_AnHui(string riPath,
            string lbiPath, string bhPath,
            string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {
#if 支持无平整度导出国检

            if (_Setting.SelectDrawDis == 1)
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                      
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);
                     
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, false, false, false))
                {
                    MyExcelVillageDegree.Convent_Damage2024_AnHui(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
#else

            if (_Setting.SelectDrawDis == 1)
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegreeSmall.InitProDataD(_DataDir, _ProjectInfo, _Setting.IRIAlgorithmInterval, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    MyExcelVillageDegreeSmall.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                _Setting.banMarkSign = true;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, _Setting.IRIAlgorithmInterval, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }



                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Damage2024_AnHui(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
#endif


        }
        public void Convent_Village2024_GuangDong(string riPath,
         string lbiPath, string bhPath,
         string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {
            {

                _Setting.banMarkSign = true;
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }



                }

                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_1(lbiPath, _ProjectInfo, _DataDir, proName, suff);
                    MyExcelVillageDegree.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Damage2024_GuangDong(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

        }


        public void Convent_Village2024_GanSu(string riPath,
            string lbiPath, string bhPath,
            string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {

            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {


                    MyExcelVillageDegree.Convent_Damage2024_GanSu(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2024_GanSu(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_GanSu(lbiPath, _ProjectInfo, _DataDir, proName, suff);
                    MyExcelVillageDegree.Convent_Iri2024_GanSu(iriPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2024_GanSu(bhPath, _ProjectInfo, _DataDir, 10, proName);

                    MyExcelVillageDegreeSmall.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }



        }

        public void Convent_Village2024_HeNan(string riPath,
           string lbiPath, string bhPath,
           string iriPath, string proName, Farmework.Other.enumTools.hnEnumTools.CityModelItem standard, string suff)
        {

            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {

                    MyExcelVillageDegree.Convent_Iri2024_HeNan(iriPath, _ProjectInfo, _DataDir, proName, suff);
                    MyExcelVillageDegree.Convent_Damage2024_HeNan(bhPath, _ProjectInfo, _DataDir, 10, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName, suff);

                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegree.Convent_LP2024_HeNan(riPath, _ProjectInfo, _DataDir, proName, suff);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {
                    MyExcelVillageDegree.Convent_Lbi2024_HeNan(lbiPath, _ProjectInfo, _DataDir, proName, suff);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }
            else
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {

                    MyExcelVillageDegreeSmall.Convent_Damage2024(bhPath, _ProjectInfo, _DataDir, 10, proName);

                    MyExcelVillageDegreeSmall.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegreeSmall.Convent_GdLP2023(riPath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 1, true, true, false, false, true, true))
                    {
                        MyExcelVillageDegreeSmall.Convent_LP2023(riPath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    _Setting.banMarkSign = false;
                }
                _Setting.banMarkSign = true;
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    MyExcelVillageDegreeSmall.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
            }


        }

        public void Convent_2024_LiaoNing(MSExcel.Application excelApp, string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
         string BumpPath, string mpdPath, string textFile, string tpFile, string proName, Encoding en)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2024_LiaoNing(iriPath, _ProjectInfo, _DataDir, proName);
                    //磨耗
                    MyExcelDegree2018.Convent_Mpd2023(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    MyExcelDegree2018.Convent_TT2023(textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegree2018.Convent_Bump2023(BumpPath, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegreeSmall2018.Convent_Rut2023(rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2024_LiaoNing(iriPath, _ProjectInfo, _DataDir, proName);
                    //加速度
                    //MyExcelDegreeSmall2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegreeSmall2018.Convent_Bump2023(excelApp, BumpPath, _ProjectInfo, _DataDir, proName);

                    //磨耗
                    MyExcelDegreeSmall2018.Convent_Mpd2023(excelApp, mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    MyExcelDegreeSmall2018.Convent_TT2023(textFile, _ProjectInfo, _DataDir, proName);
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegreeSmall2018.Convent_TP2023(excelApp, tpFile, _ProjectInfo, _DataDir, proName);
                    MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息

        }

        public void Convent_Standard(string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
        string BumpPath, string mpdPath, string textFile, string tpFile, string proName)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (_ProjectInfo._IsRut)
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                    {
                        //车辙
                        MyExcelDegree2018.Convent_Rut_Standard(rdFilePath, _ProjectInfo, _DataDir, proName);
                        //空间定位
                        MyExcelDegree2018.Convent_Standard(lbiPath, _ProjectInfo, _DataDir, proName);
                        //平整度
                        MyExcelDegree2018.Convent_Iri2024_Standard(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                        //磨耗
                        MyExcelDegree2018.Convent_Mpd2024_Standard(mpdPath, _ProjectInfo, _DataDir, proName);
                        //磨耗高程数据
                        MyExcelDegree2018.Convent_TT2024_Standard(textFile, _ProjectInfo, _DataDir, proName);
                        //加速度
                        // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                        //跳车
                        MyExcelDegree2018.Convent_Bump2024(BumpPath, _ProjectInfo, _DataDir, proName);

                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }


                    if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelDegree2018.Convent_TP2024_ChongQing(tpFile, _ProjectInfo, _DataDir, proName);

                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, _Setting.IRIAlgorithmInterval, true, true, false, false, true, true))
                    {

                        MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                    {
                        //破损
                        MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                    {

                        //空间定位
                        MyExcelDegree2018.Convent_Standard(lbiPath, _ProjectInfo, _DataDir, proName);
                        //平整度
                        MyExcelDegree2018.Convent_Iri2024_Standard(iriPath, _ProjectInfo, _DataDir, proName, ".txt");


                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }


                    if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        //Lp高程  iri
                        MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                    {
                        //破损
                        MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
            else
            {
                if (_ProjectInfo._IsRut)
                {
                    _Setting.banMarkSign = true;
                    if ((MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true)))
                    {

                        //车辙
                        MyExcelDegreeSmall2018.Convent_Rut_Standard(rdFilePath, _ProjectInfo, _DataDir, proName);

                        //磨耗
                        MyExcelDegreeSmall2018.Convent_Mpd2024_Standard(mpdPath, _ProjectInfo, _DataDir, proName);
                        //磨耗高程数据
                        MyExcelDegreeSmall2018.Convent_TT2024_Standard(textFile, _ProjectInfo, _DataDir, proName);
                        //加速度
                        // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                        //跳车
                        MyExcelDegreeSmall2018.Convent_Bump2024(BumpPath, _ProjectInfo, _DataDir, proName);


                        //空间定位
                        MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                        //平整度
                        MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);


                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {
                        MyExcelDegreeSmall2018.Convent_TP2024_ChongQing(tpFile, _ProjectInfo, _DataDir, proName);

                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, _Setting.IRIAlgorithmInterval, true, true, false, false, true, true))
                    {

                        MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    _Setting.banMarkSign = false;
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                    {
                        //破损
                        MyExcelDegreeSmall2018.Convent_Damage2024_ChongQing(bhPath, _ProjectInfo, _DataDir, 10, proName, Encoding.UTF8);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                    {
                        //空间定位
                        MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                        //平整度
                        MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);


                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                    if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.banMarkSign = false;
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                    {
                        //破损
                        MyExcelDegreeSmall2018.Convent_Damage2024_ChongQing(bhPath, _ProjectInfo, _DataDir, 10, proName, Encoding.UTF8);
                    }
                    else
                    {
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //定位信息

        }
        public void Convent_2024_HeBei(string riFilePath, string lbiPath, string bhPath, string iriPath, string rdFilePath,
        string BumpPath, string mpdPath, string textFile, string tpFile, string proName, Encoding encoding)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                _Setting.banMarkSign = true;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //车辙
                    MyExcelDegree2018.Convent_Rut2024_ChongQing(rdFilePath, _ProjectInfo, _DataDir, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2024_Standard(iriPath, _ProjectInfo, _DataDir, proName, ".txt");
                    //磨耗
                    MyExcelDegree2018.Convent_Mpd2024_chongqing(mpdPath, _ProjectInfo, _DataDir, proName);
                    //磨耗高程数据
                    MyExcelDegree2018.Convent_TT2024_ChongQing(textFile, _ProjectInfo, _DataDir, proName);
                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);
                    //跳车
                    MyExcelDegree2018.Convent_Bump2024(BumpPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_TP2024_ChongQing(tpFile, _ProjectInfo, _DataDir, proName);
                    //Lp高程  iri
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

            //定位信息

        }


        public void Convent_2024_ChongQing(string riFilePath, string lbiPath, string bhPath, string iriPath, string proName, Encoding encoding)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                _Setting.banMarkSign = true;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {

                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2024_ChongQing(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2024_Standard(iriPath, _ProjectInfo, _DataDir, proName, ".txt");

                    //加速度
                    // MyExcelDegree2018.Convent_Acc(excelApp, riFilePath, _ProjectInfo, _DataDir, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, true))
                {
                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else
            {
                _Setting.banMarkSign = true;
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);


                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {

                    MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                _Setting.banMarkSign = false;
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2024_ChongQing(bhPath, _ProjectInfo, _DataDir, 10, proName, encoding);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }

            //定位信息

        }
        public void Convent_2023(MSExcel.Application excelApp, string riFilePath, string lbiPath, string bhPath, string iriPath, string proName, Encoding en)
        {
            if (_Setting.SelectDrawDis == 0)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //破损
                    MyExcelDegree2018.Convent_DamageStandard(bhPath, _ProjectInfo, _DataDir, 10, proName);
                    //空间定位
                    MyExcelDegree2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegree2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);

                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {
                    MyExcelDegree2018.Convent_LPStandard(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


            }
            else
            {
                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                {
                    //空间定位
                    MyExcelDegreeSmall2018.Convent_Lbi2023(lbiPath, _ProjectInfo, _DataDir, proName);
                    //平整度
                    MyExcelDegreeSmall2018.Convent_Iri2023(iriPath, _ProjectInfo, _DataDir, proName);

                    //破损
                    MyExcelDegreeSmall2018.Convent_Damage2023(bhPath, _ProjectInfo, _DataDir, 10, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                {

                    MyExcelDegreeSmall2018.Convent_LP2023(riFilePath, _ProjectInfo, _DataDir, proName);
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            excelApp.Quit();
            //定位信息
        }
        private void GenerateExcel_Village(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {

                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {

                        //(5211)路面破损评定汇总表
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputPci_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        //(5211)路面平整度评定汇总表
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputRqi_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {

                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {

                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        try
                        {
                            if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                                MyExcelVillageDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (System.Exception)
                        {

                        }

                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelVillageDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false))
                            {

                                MyExcelVillageDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (VillageIndexException ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }


                    }
                }

                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], true, false, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[9][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[10].Length; ++i)
                    {

                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[10][i], true, true, false, false, false, true))
                        {
                            MyExcelVillageDegree.OutputPQI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[10][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    MyExcelVillageDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[11].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[11][i], true, true, false, false, false, false))
                            {
                                MyExcelVillageDegree.OutputPDMX_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[11][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (VillageIndexException ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }


                    }
                }
            }
            //桂兴达
            else if (_Setting.ExcelType == 1)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, false, false))
                            MyExcelVillageDegree.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
                            MyExcelVillageDegree.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[2])
                {
                    MyExcelVillageDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, true, false, false, false, false))
                            MyExcelVillageDegree.OutputPDMX_2024Sum(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            else if (_Setting.ExcelType == 2)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_0(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //沥青破损
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, true))
                            MyExcelVillageDegree.OutputLQDamage(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //水泥破损
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, false, false, false, false, true))
                            MyExcelVillageDegree.OutputSNDamage(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], false, true, false, false, false, true))
                            MyExcelVillageDegree.outPutAutoTest_1(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[4])
                {
                    //if (_Setting.hasCamsetting)
                    //{
                    //    MessageBox.Show("该工程设备采集数据不支持出具平整度数据！");
                    //}
                    //else
                    {
                        if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1f, false, true, false, false, true, true))
                            MyExcelVillageDegree.outPutAutoTest_2(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");


                    }

                }


                //  DR破损率csv检测结果数据表格
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, false, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_5(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //IRI平整度csv检测结果数据表格
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], true, true, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_6(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //IRI平整度csv检测原始数据表格
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i], true, true, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_7(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //空间定位数据csv检测原始数据表格
                if (IsOutputxls[8])
                {
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], false, false, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_8(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                //空间定位数据txt检测原始数据表格
                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], false, false, false, false, false, false))
                            MyExcelVillageDegree.outPutAutoTest_9(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
            else if (_Setting.ExcelType == 4)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, false, false))
                            MyExcelVillageDegree.outPutAutoTest_HN0(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //水泥
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegree.OutputSNDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                //沥青
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegree.OutputLQDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //病害流水
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))

                            MyExcelVillageDegree.OutputAllDamage_hn(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                //重庆 公里指标导入模板
                if (IsOutputxls[4])
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1000, true, true, false, false, false, true))
                        MyExcelVillageDegree.OutputChongQingSumExcel(excelApp, xlspath, _ProjectInfo, _DataDir, 1000);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }

                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelVillageDegree.OutputChongQingDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }


                }
            }
            else if (_Setting.ExcelType == 6)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (_ProjectInfo._StreetImgDis_Left > xlslen[0][i])
                        {
                            MessageBox.Show("工程景观图片间隔为" + _ProjectInfo._StreetImgDis_Left + ",分段间距必须大于等于该值");

                        }
                        else
                        {
                            _Setting.shieldMark = true;
                            if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, false, false, false))
                                MyExcelVillageDegree.outPutAutoTest_GuiZhou(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                            _Setting.shieldMark = false;
                        }


                    }

                }
            }
            else if (_Setting.ExcelType == 7)
            {
                //合肥
                if (IsOutputxls[0])
                {
                    _Setting.hefeiOutExcel2 = true;
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        //首先读取资产配置表
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelVillageDegree.OutputRoad_Hefei(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.hefeiOutExcel2 = false;
                }

            }

            else if (_Setting.ExcelType == 8)
            {
                if (IsOutputxls[0])
                {

                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 2, true, false, false, false, false, false))
                        MyExcelVillageDegree.OutputDis_TH(excelApp, xlspath, _ProjectInfo, _DataDir, 2);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
            }
            else if (_Setting.ExcelType == 9)
            {


            }


            else if (_Setting.ExcelType == 10)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, true, false, false, false, true))
                            MyExcelVillageDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, true))
                            MyExcelVillageDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

            }
            else if (_Setting.ExcelType == 11)
            {

                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();


                //空间定位文本
                if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, true, true))
                {

                    MyExcelVillageDegree.Convent_Lbi2024(xlspath, _ProjectInfo, _DataDir, "LBI");
                }
                else
                {
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                //平整度文本 
                if (_Setting.isGDIriCalculate)
                {
                    //惯导
                    MyExcelVillageDegree.Convent_GdLP2024(xlspath, _ProjectInfo, _DataDir, "LP");
                }
                else
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProDataD(_DataDir, _ProjectInfo, 0.1, true, true, false, false, true, true))
                    {

                        MyExcelVillageDegree.Convent_LP2024(xlspath, _ProjectInfo, _DataDir, "LP");
                    }
                    _Setting.banMarkSign = false;
                }
                //激光 
                //加速度计
            }
            // 贵州： 未完成
            else if (_Setting.ExcelType == 12)// 未完成
            {
                //MyExcelVillageDegree.贵州农村公路检测图片
                //MyExcelVillageDegree.贵州农村公路检测数据明细表
                //MyExcelVillageDegree.贵州农村公路检测轨迹

                //贵州农村公路检测图片
                if (IsOutputxls[0])
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                        MyExcelVillageDegree.贵州农村公路检测图片(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                //贵州农村公路检测数据明细表
                if (IsOutputxls[1])
                {

                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, true))
                        MyExcelVillageDegree.贵州农村公路检测数据明细表(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                //贵州农村公路检测轨迹
                if (IsOutputxls[2])
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, false, false, true))
                        MyExcelVillageDegree.贵州农村公路检测轨迹(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 13)
            {
                _Setting.banMarkSign = false;
                if (IsOutputxls[0])
                {
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, false, false, true))
                        MyExcelVillageDegree.江西农村路沥青病害(excelApp, xlspath, _ProjectInfo, _DataDir, _RoadConfig);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[1])
                {

                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, false, false, true))
                        MyExcelVillageDegree.江西农村路水泥病害(excelApp, xlspath, _ProjectInfo, _DataDir, _RoadConfig);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[2])
                {
                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 10, false, true, false, false, false, true))
                        MyExcelVillageDegree.江西农村路平整度(excelApp, xlspath, _ProjectInfo, _DataDir, _RoadConfig);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    _Setting.banMarkSign = false;
                }
            }


            else if (_Setting.ExcelType == 14)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        //10
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false, false))
                        {
                            MyExcelVillageDegree.OutputDis_HPcsv_0(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }

            else if (_Setting.ExcelType == 15)
            {
                if (IsOutputxls[0])
                {

                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1000, false, false, false, false, false, false))
                        MyExcelVillageDegree.outPutAccessory04(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    _Setting.banMarkSign = false;
                }
                if (IsOutputxls[1])
                {

                    {
                        if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1, false, false, false, false, false, false))
                            MyExcelVillageDegree.outPutAccessory05(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {

                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1000, false, false, false, false, false, false))
                        MyExcelVillageDegree.outPutAccessory06(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
                if (IsOutputxls[3])
                {
                    _Setting.userRoadCondition = true;
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1000, false, false, false, false, false, false))
                        MyExcelVillageDegree.outPutAccessory07(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    _Setting.userRoadCondition = false;
                }

                if (IsOutputxls[4])
                {

                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, 1000, true, true, false, false, false, true))
                        MyExcelVillageDegree.outPutAccessory08(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }

                //贵州质安检测 5m景观图像和文件生成
                if (IsOutputxls[5])
                {

                    _Setting.banMarkSign = true;
                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, _ProjectInfo._StreetImgDis_Left, false, false, false, false, false, false))
                        MyExcelVillageDegree.outPutAccessory09(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    _Setting.banMarkSign = false;
                }
                if (IsOutputxls[6])
                {



                    if (MyExcelVillageDegree.InitProData(_DataDir, _ProjectInfo, _ProjectInfo._StreetImgDis_Left, false, false, false, false, false, false))
                        MyExcelVillageDegree.outPutAutoTest_9(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }

            }
        }
        private void GenerateExcel_Village_Street_Small(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelVillageDegreeSmall.InitStreetData(_ProjectInfo, _DataDir);

            if (_Setting.ExcelType == 3)
            {
                ///孝感定制
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, true, false, false))
                {
                    MyExcelVillageDegreeSmall.OutputStreetDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }

                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 10, false, false, false, false, false, false))
                {
                    MyExcelVillageDegreeSmall.OutputRoadBedDis_XG(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }

            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelVillageDegreeSmall.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelVillageDegreeSmall.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelVillageDegreeSmall.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelVillageDegreeSmall.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[4])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], false, false, false, false, false, false))
                    {
                        MyExcelVillageDegreeSmall.OutputRoadBedDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[5])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelVillageDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], false, false, false, true, false, false))
                    {
                        MyExcelVillageDegreeSmall.OutputStreetDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[6])
            {

                MyExcelVillageDegreeSmall.OutputStreetAllDis(excelApp, xlspath, _ProjectInfo, _DataDir);
            }
        }

        private void GenerateExcel_hn_Street_Small(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            MyExcelHNDegreeSmall.InitStreetData(_ProjectInfo, _DataDir);



            if (IsOutputxls[0])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, false))
                    {
                        MyExcelHNDegreeSmall.OutputStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[1])
            {
                if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, true, false, false))
                {
                    MyExcelHNDegreeSmall.OutputCPMSStreetDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, false, false, false, false, false))
                    {
                        MyExcelHNDegreeSmall.OutputRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[3])
            {
                if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, 100, false, false, false, false, false, false))
                {
                    MyExcelHNDegreeSmall.OutputCPMSRoadBedDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                {
                    MessageBox.Show("加载工程数据失败！");
                }
            }
            if (IsOutputxls[4])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], false, false, false, false, false, false))
                    {
                        MyExcelHNDegreeSmall.OutputRoadBedDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
            if (IsOutputxls[5])
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], false, false, false, true, false, false))
                    {
                        MyExcelHNDegreeSmall.OutputStreetDis_JSZK(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                    }
                    else
                    {
                        MessageBox.Show("加载工程数据失败！");
                    }
                }
            }
        }

        private void GenerateExcel_HuNanSmall(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelHNDegreeSmall.OutputIRI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelHNDegreeSmall.OutputPCI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                            MyExcelHNDegreeSmall.OutputDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                            MyExcelHNDegreeSmall.OutputPQI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        try
                        {
                            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                                MyExcelHNDegreeSmall.OutputCPMSDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (System.Exception)
                        {

                        }


                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelHNDegreeSmall.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false)) //倒数第三个为PBI计算 改为false
                            MyExcelHNDegreeSmall.OutputPDMX_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                        {
                            MyExcelHNDegreeSmall.OutputDis_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {

                        if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                        {
                            MyExcelHNDegreeSmall.OutputPQI_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    MyExcelHNDegreeSmall.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false))
                            {
                                MyExcelHNDegreeSmall.OutputPDMX_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (VillageIndexException ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }


                    }
                }

            }

            //if (_Setting.ExcelType == 0)
            //{
            //    if (IsOutputxls[2])
            //    {
            //        for (int i = 0; i < xlslen[2].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
            //                MyExcelHNDegreeSmall.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[3])
            //    {
            //        for (int i = 0; i < xlslen[3].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
            //                MyExcelHNDegreeSmall.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[4])
            //    {
            //        for (int i = 0; i < xlslen[4].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
            //                MyExcelHNDegreeSmall.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[5])
            //    {
            //        for (int i = 0; i < xlslen[5].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
            //                MyExcelHNDegreeSmall.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[7])
            //    {
            //        for (int i = 0; i < xlslen[7].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
            //                MyExcelHNDegreeSmall.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[8])
            //    {
            //        MyExcelHNDegree.InitStreetData(_ProjectInfo, _DataDir);
            //        for (int i = 0; i < xlslen[8].Length; ++i)
            //        {
            //            if (MyExcelHNDegreeSmall.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, false, false))
            //                MyExcelHNDegreeSmall.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //}
        }
        private void GenerateExcel_HuNan(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {

                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputIRI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputPCI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {

                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputPQI_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        try
                        {
                            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                                MyExcelHNDegree.OutputCPMSDis_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (System.Exception)
                        {

                        }

                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelHNDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false))
                            {

                                MyExcelHNDegree.OutputPDMX_2024(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (VillageIndexException ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }


                    }
                }

                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputDis_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {

                        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
                        {
                            MyExcelHNDegree.OutputPQI_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    MyExcelHNDegree.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        //public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval,
                        // bool IsDis, bool IsMeanIRI, bool IsMeanMTD, bool IsMeanRut, bool IsPBI, bool IsSpeed, bool IsMeanMPD = false)
                        try
                        {
                            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, false, false, false, false))
                            {
                                MyExcelHNDegree.OutputPDMX_2024_new(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            }
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        catch (System.Exception)
                        {
                            MessageBox.Show("发生了一个未知错误，请联系管理员。");
                            return;
                        }


                    }
                }
            }
            //if (_Setting.ExcelType == 0)
            //{
            //    if (IsOutputxls[2])
            //    {
            //        for (int i = 0; i < xlslen[2].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
            //                MyExcelHNDegree.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[3])
            //    {
            //        for (int i = 0; i < xlslen[3].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
            //                MyExcelHNDegree.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[4])
            //    {
            //        for (int i = 0; i < xlslen[4].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
            //                MyExcelHNDegree.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[5])
            //    {
            //        for (int i = 0; i < xlslen[5].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, true))
            //                MyExcelHNDegree.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[7])
            //    {
            //        for (int i = 0; i < xlslen[7].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
            //                MyExcelHNDegree.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //    if (IsOutputxls[8])
            //    {
            //        MyExcelHNDegree.InitStreetData(_ProjectInfo, _DataDir);
            //        for (int i = 0; i < xlslen[8].Length; ++i)
            //        {
            //            if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, false, false))
            //                MyExcelHNDegree.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
            //            else
            //                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //        }
            //    }
            //}
            //if (_Setting.ExcelType == 1)
            //{
            //    for (int i = 0; i < xlslen[4].Length; ++i)
            //    {
            //        if (MyExcelHNDegree.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, true))
            //            MyExcelHNDegree.OutputDis_lf(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
            //        else
            //            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            //    }
            //}
        }


        private void GenerateExcel_Degree2018_SmallRect(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {

                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, true))
                            MyExcelDegreeSmall2018.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, true))
                            MyExcelDegreeSmall2018.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelDegreeSmall2018.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelDegreeSmall2018.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, false))
                        {
                            MyExcelDegreeSmall2018.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, true, true, true, true))
                            MyExcelDegreeSmall2018.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, true, true))
                            MyExcelDegreeSmall2018.OutputPBI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                            MyExcelDegreeSmall2018.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    MyExcelDegreeSmall2018.InitStreetData(_ProjectInfo, _DataDir);
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        if (MainForm._IsOutputEmptyExcel)
                        {
                            if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], false, false, false, false, false, false))
                                MyExcelDegreeSmall2018.OutputPDMX_Empty(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        else
                        {
                            if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
                                MyExcelDegreeSmall2018.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                    }
                }
                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], false, false, true, false, false, true))
                            MyExcelDegreeSmall2018.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[9][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[10].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[10][i], false, false, false, false, false, true, true))
                            MyExcelDegreeSmall2018.OutputMPD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[10][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    for (int i = 0; i < xlslen[10].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[10][i], false, false, false, false, false, true, false, true))
                            MyExcelDegreeSmall2018.OutputGeoAlig(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[11][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            else if (_Setting.ExcelType == 1)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelDegreeSmall2018.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
                            MyExcelDegreeSmall2018.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[2])
                {

                

                    int[]  curXlslen= new int[] { 10,100,1000 };
                    string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\公路技术状况评定结果_Small.xlsx",
                     System.Windows.Forms.Application.StartupPath);
                    
                    string roadCode = _ProjectInfo._RoadCode;
                    string startMile = (_ProjectInfo._StartMile * 0.001).ToString("f3");
                    string yearTime = _ProjectInfo._DataDate;
                    string dateTime = _ProjectInfo._DataTime;
                    string dir = _ProjectInfo._Direction > 0 ? "A" : "B";

                 
                    string Destxls = $"公路技术状况评定结果-{roadCode}{dir}-{startMile}-{yearTime}{dateTime}.xlsx";
                    Destxls = string.Format(@"{0}\{1}", xlspath, Destxls);

                    MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, System.Type.Missing,
                         true, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                          System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                          System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);

                    _Workbook.SaveAs(Destxls, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                    for (int i = 0; i < curXlslen.Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, curXlslen[i], true, true, true, true, true, false))
                        {
                            MyExcelDegreeSmall2018.OutputAllInfo(_Workbook, xlspath, _ProjectInfo, _DataDir, curXlslen[i]);
                            if (i == 2)
                            {
                                MSExcel.Worksheet _Worksheet = _Workbook.Sheets["主页"] as MSExcel.Worksheet;
                                _Worksheet.Select();
                                _Workbook.Save();
                                _Workbook.Close(System.Type.Missing, System.Type.Missing, System.Type.Missing);
                                int generation = System.GC.GetGeneration(excelApp);
                                System.GC.Collect(generation);//垃圾回收
                                System.GC.WaitForPendingFinalizers();
                            } 
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                   
                } 
            }
            //中南安环
            else if (_Setting.ExcelType == 2)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, true, false))
                    {
                        MyExcelDegreeSmall2018.OutputZNRoadDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegreeSmall2018.OutputZNRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegreeSmall2018.OutputZNDataRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 3)
            {
                int[] lenval = { 5, 10, 20, 100, 1000 };
                for (int i = 0; i < lenval.Length; ++i)
                {
                    bool tflag = false;
                    if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000)
                        || IsOutputxls[5] && lenval[i] == 1000)
                    {
                        tflag = true;
                    }
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, lenval[i], tflag, IsOutputxls[2] || IsOutputxls[5],
                         IsOutputxls[1] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5], false))
                    {
                        if (IsOutputxls[0] && (lenval[i] == 10 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegreeSmall2018.OutputZJGTRut(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[1] && (lenval[i] == 10))
                        {
                            MyExcelDegreeSmall2018.OutputZJGTPWI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[2] && (lenval[i] == 20 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegreeSmall2018.OutputZJGTIRI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegreeSmall2018.OutputZJGTDis(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[4] && (lenval[i] == 5 || lenval[i] == 1000))
                        {
                            MyExcelDegreeSmall2018.OutputZJGTGPS(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[5] && lenval[i] == 1000)
                        {
                            MyExcelDegreeSmall2018.OutputZJGTPQI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[6] && lenval[i] == 10)
                        {
                            MyExcelDegreeSmall2018.OutputZJGTPBI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }

                        MyExcelDegreeSmall2018.OutputZJGTRoadType(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //带GPS的重庆招商局报表模板
            else if (_Setting.ExcelType == 4)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                    {
                        MyExcelDegreeSmall2018.OutputGPSRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        if (_ProjectInfo._IsRoad)
                        {
                            MyExcelDegreeSmall2018.OutputGPSRoadImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        if (_ProjectInfo._IsStreet)
                        {
                            MyExcelDegreeSmall2018.OutputGPSStreetImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        if (_ProjectInfo._IsPano)
                        {
                            MyExcelDegreeSmall2018.OutputGPSPanoImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 5)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, true, false))
                    {
                        MyExcelDegreeSmall2018.OutputALTDIS(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        //MyExcelDegree2018.OutputCountRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 8)
            {   //0, 1, 2, 3, 5, 6, 9 
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, true))
                            MyExcelDegreeSmall2018.OutputRut_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, true))
                            MyExcelDegreeSmall2018.OutputPWI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelDegreeSmall2018.OutputIRI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelDegreeSmall2018.OutputPCI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, true, true, true, true))
                            MyExcelDegreeSmall2018.OutputPQI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, true, true))
                            MyExcelDegreeSmall2018.OutputPBI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], false, false, true, false, false, true))
                            MyExcelDegreeSmall2018.OutputMTD_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[9][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //河南焦作
            else if (_Setting.ExcelType == 9)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false, false))
                        MyExcelDegreeSmall2018.OutputPQI_HNJZ_ZHPD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 11)
            {
                for (int i = 0; i < 1; ++i)
                {
                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, true, true, true, false, false))
                        MyExcelDegreeSmall2018.OutputDis_CICS(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 12)
            {
                if (IsOutputxls[0])
                {

                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, false, false, false, false))
                    {

                        MyExcelDegreeSmall2018.OutputDis_HPcsv_0(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
                if (IsOutputxls[1])
                {

                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, true, false, false, false, false, false))
                        {

                            MyExcelDegreeSmall2018.OutputDis_HPcsv_1(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }


                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, true, false, false, false, false, false))
                        {

                            MyExcelDegreeSmall2018.OutputDis_HPcsv_2(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, true, false, false, false))
                        {
                            MyExcelDegreeSmall2018.OutputDis_HPcsv_Rut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, true, false, false, false))
                        {
                            MyExcelDegreeSmall2018.outPutDrCsv_WH(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, true, false, false, false))
                        {
                            MyExcelDegreeSmall2018.outPutIriCsv_WH(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }

            }
            else if (_Setting.ExcelType == 13)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_0(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, false, false, false, true))
                            MyExcelDegreeSmall2018.outPutAutoTest_1(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度原始数据
                if (IsOutputxls[2])
                {
                    if (MyExcelDegreeSmall2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, false, true, false, false, true, true))
                        MyExcelDegreeSmall2018.outPutAutoTest_2(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_5(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //IRI平整度csv检测结果数据表格
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, true, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_6(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //IRI平整度csv检测原始数据表格
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_7(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //空间定位数据csv检测原始数据表格
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_8(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                //四川工勘院定制空间定位表格

                if (IsOutputxls[8])
                {
                    for (int i = 0; i < xlslen[8].Length; ++i)
                    {
                        _Setting.shieldMark = true;
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], false, false, false, false, false, false))
                            MyExcelDegreeSmall2018.outPutAutoTest_9(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        _Setting.shieldMark = false;
                    }
                }
            }
            else if (_Setting.ExcelType == 14)
            {

                //窗口供用户选择

                if (IsOutputxls[0])
                {
                    四川公路院出表选择 form = 四川公路院出表选择.getInstance();
                    DialogResult result = form.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        List<int> resultSelect = form.getUserSelect();
                        if (resultSelect.Count != 0)
                        {
                            for (int i = 0; i < xlslen[0].Length; ++i)
                            {
                                if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, true, true))

                                {
                                    ProcessOperator p = new ProcessOperator();

                                    System.Action t1 = () => { MyExcelDegreeSmall2018.OutputRoad_Szechwan(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i], resultSelect, p); };

                                    p.BackgroundWork = t1;
                                    p.BackgroundWorkerCompleted += P_BackgroundWorkerCompleted;
                                    p.Start();


                                }
                                else
                                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                            }
                        }
                    }

                }

            }
            else if (_Setting.ExcelType == 15)
            {
                //合肥
                if (IsOutputxls[0])
                {
                    _Setting.hefeiOutExcel2 = true;
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        //首先读取资产配置表
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelDegreeSmall2018.OutputRoad_Hefei(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.hefeiOutExcel2 = false;
                }

            }

            else if (_Setting.ExcelType == 18)
            {
                if (IsOutputxls[0])
                {

                    if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 2, true, false, false, false, false, false))
                        MyExcelDegreeSmall2018.OutputDis_TH(excelApp, xlspath, _ProjectInfo, _DataDir, 2);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }
            }
            else if (_Setting.ExcelType == 19)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, true, false, false, false, true))
                            MyExcelDegreeSmall2018.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, false))
                            MyExcelDegreeSmall2018.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }

        private void P_BackgroundWorkerCompleted(object sender, EventArgs e)
        {

        }





        //拉框模式
        private void GenerateExcel_Degree2018_BigRect(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, true))
                        {
                            MyExcelDegree2018.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, true))
                            MyExcelDegree2018.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                        {
                            MyExcelDegree2018.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelDegree2018.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, false))
                        {
                            MyExcelDegree2018.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        }


                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }

                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, true, true, true, true))
                            MyExcelDegree2018.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, true, true))
                        {
                            MyExcelDegree2018.OutputPBI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[7])
                {
                    for (int i = 0; i < xlslen[7].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[7][i] / 10, true, false, false, false, false, false))
                            MyExcelDegree2018.OutputCPMSDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[7][i] / 10);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[8])
                {
                    if (MainForm._IsOutputEmptyExcel)
                    {
                        for (int i = 0; i < xlslen[8].Length; ++i)
                        {
                            if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], false, false, false, false, false, false))
                                MyExcelDegree2018.OutputPDMX_Empty(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                    }
                    else
                    {
                        MyExcelDegree2018.InitStreetData(_ProjectInfo, _DataDir);
                        for (int i = 0; i < xlslen[8].Length; ++i)
                        {
                            if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[8][i], true, true, true, true, true, false))
                                MyExcelDegree2018.OutputPDMX(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[8][i]);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                    }
                }
                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], false, false, true, false, false, true))
                            MyExcelDegree2018.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[9][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[10])
                {
                    for (int i = 0; i < xlslen[10].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[10][i], false, false, false, false, false, true, true))
                        {
                            MyExcelDegree2018.OutputMPD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[10][i]);
                        }

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[11])
                {
                    for (int i = 0; i < xlslen[10].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[10][i], false, false, false, false, false, true, false, true))
                            MyExcelDegree2018.OutputGeoAlig(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[11][i]);

                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //桂兴达
            else if (_Setting.ExcelType == 1)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelDegree2018.OutputRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, true, false, false, true))
                            MyExcelDegree2018.OutputGXDIRIMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //中南安环
            else if (_Setting.ExcelType == 2)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, true, true))
                    {
                        MyExcelDegree2018.OutputZNRoadDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegree2018.OutputZNRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        MyExcelDegree2018.OutputZNDataRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 3)
            {
                int[] lenval = { 5, 10, 20, 100, 1000 };
                for (int i = 0; i < lenval.Length; ++i)
                {
                    bool tflag = false;
                    if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000)
                        || IsOutputxls[5] && lenval[i] == 1000)
                    {
                        tflag = true;
                    }
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, lenval[i], tflag, IsOutputxls[2] || IsOutputxls[5],
                         IsOutputxls[1] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5], IsOutputxls[0] || IsOutputxls[5], false))
                    {
                        if (IsOutputxls[0] && (lenval[i] == 10 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2018.OutputZJGTRut(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[1] && (lenval[i] == 10))
                        {
                            MyExcelDegree2018.OutputZJGTPWI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[2] && (lenval[i] == 20 || lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2018.OutputZJGTIRI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[3] && (lenval[i] == 100 || lenval[i] == 1000))
                        {
                            MyExcelDegree2018.OutputZJGTDis(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[4] && (lenval[i] == 5 || lenval[i] == 1000))
                        {
                            MyExcelDegree2018.OutputZJGTGPS(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[5] && lenval[i] == 1000)
                        {
                            MyExcelDegree2018.OutputZJGTPQI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }
                        if (IsOutputxls[6] && lenval[i] == 10)
                        {
                            MyExcelDegree2018.OutputZJGTPBI(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                        }

                        MyExcelDegree2018.OutputZJGTRoadType(excelApp, xlspath, _ProjectInfo, _DataDir, lenval[i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //带GPS的重庆招商局报表模板
            else if (_Setting.ExcelType == 4)
            {

                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                    {
                        MyExcelDegree2018.OutputGPSRoad(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);


                        if (_ProjectInfo._IsRoad)
                        {
                            MyExcelDegree2018.OutputGPSRoadImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                            MyExcelDegree2018.OutputGPSRoadImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        if (_ProjectInfo._IsStreet)
                        {
                            MyExcelDegree2018.OutputGPSStreetImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        if (_ProjectInfo._IsPano)
                        {
                            MyExcelDegree2018.OutputGPSPanoImg(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 5)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsIRIMTD, _ProjectInfo._IsRut, true, false))
                    {
                        MyExcelDegree2018.OutputALTDIS(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        //MyExcelDegree2018.OutputCountRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 8)
            {   // 0, 1, 2, 3, 5, 6, 9 
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, false, true))
                            MyExcelDegree2018.OutputRut_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, true))
                            MyExcelDegree2018.OutputPWI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelDegree2018.OutputIRI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, true))
                            MyExcelDegree2018.OutputPCI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, true, true, true, true))
                            MyExcelDegree2018.OutputPQI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, true, true))
                            MyExcelDegree2018.OutputPBI_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[9])
                {
                    for (int i = 0; i < xlslen[9].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[9][i], false, false, true, false, false, true))
                            MyExcelDegree2018.OutputMTD_XMJH(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[9][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            //河南焦作
            else if (_Setting.ExcelType == 9)
            {
                for (int i = 0; i < xlslen[0].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false, false))
                    {
                        MyExcelDegree2018.OutputPQI_HNJZ_ZHPD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            //广东华路
            else if (_Setting.ExcelType == 10)
            {
                for (int i = 0; i < xlslen[4].Length; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, false, false))
                    {
                        MyExcelDegree2018.OutputDis_GDHL(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                }
            }
            else if (_Setting.ExcelType == 11)
            {
                for (int i = 0; i < 1; ++i)
                {
                    if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, false, false, false, false))
                    {
                        MyExcelDegree2018.OutputDis_CICS(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                    }
                }
            }
            //上海惠普csv三个定制报表
            else if (_Setting.ExcelType == 12)
            {
                if (IsOutputxls[0])
                {

                    {
                        //10
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 1, true, true, true, true, true, false, false))
                        {
                            MyExcelDegree2018.OutputDis_HPcsv_0(excelApp, xlspath, _ProjectInfo, _DataDir, 1);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }


                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, true, false, false, false, false, false))
                        {

                            MyExcelDegree2018.OutputDis_HPcsv_IRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, true, false, false, false, false, false))
                        {

                            MyExcelDegree2018.OutputDis_HPcsv_2(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, true, false, false, false))
                        {

                            MyExcelDegree2018.OutputDis_HPcsv_Rut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, false, false, true, false, false, false))
                        {
                            MyExcelDegree2018.outPutDrCsv_WH(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, true, false, false, false))
                        {
                            MyExcelDegree2018.outPutIriCsv_WH(excelApp, xlspath, _ProjectInfo, _DataDir);
                        }
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
            }
            else if (_Setting.ExcelType == 15)
            {
                //合肥
                if (IsOutputxls[0])
                {
                    _Setting.hefeiOutExcel2 = true;
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        //首先读取资产配置表
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, true, true, true, false))
                            MyExcelDegree2018.OutputRoad_Hefei(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                    _Setting.hefeiOutExcel2 = false;
                }



            }
            else if (_Setting.ExcelType == 13)
            {
                //定位信息
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, false))
                            MyExcelDegree2018.outPutAutoTest_0(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, true, false, false, false, true))
                            MyExcelDegree2018.outPutAutoTest_1(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //平整度原始数据
                if (IsOutputxls[2])
                {

                    if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1, false, true, false, false, true, true))
                        MyExcelDegree2018.outPutAutoTest_2(excelApp, xlspath, _ProjectInfo, _DataDir);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, false))
                            MyExcelDegree2018.outPutAutoTest_5(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }

                }
                //IRI平整度csv检测结果数据表格
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, true, false, false, false, false))
                            MyExcelDegree2018.outPutAutoTest_6(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //IRI平整度csv检测原始数据表格
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, false))
                            MyExcelDegree2018.outPutAutoTest_7(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //空间定位数据csv检测原始数据表格
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, false, false, false, false))
                            MyExcelDegree2018.outPutAutoTest_8(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                //rut 0.1m
                if (IsOutputxls[7])
                {

                    if (MyExcelDegree2018.InitProDataD(_DataDir, _ProjectInfo, 0.1f, true, true, true, true, true, true))
                        MyExcelDegree2018.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, 0.1);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                }

            }
            else if (_Setting.ExcelType == 18)
            {
                if (_Setting.splitExcelDh)
                {
                    if (IsOutputxls[0])
                    {

                        excelApp.Quit();
                        Marshal.FinalReleaseComObject(excelApp);
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 2, true, false, false, false, false, false))
                            MyExcelDegree2018.OutputDis_TH(xlspath, _ProjectInfo, _DataDir, 2);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                    }
                }
                else
                {
                    if (IsOutputxls[0])
                    {

                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 2, true, false, false, false, false, false))
                            MyExcelDegree2018.OutputDis_THSum(excelApp, xlspath, _ProjectInfo, _DataDir, 2);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                    }
                }

            }
            else if (_Setting.ExcelType == 19)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, true, false, false, false, true))
                            MyExcelDegree2018.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, false))
                            MyExcelDegree2018.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
            else if (_Setting.ExcelType == 20)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, true, false, false, false, true))
                            MyExcelDegree2018.OutputChongQingDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }


                }
            }
            else if (_Setting.ExcelType == 21)
            { // 江西车检
                _Setting.banMarkSign = false;
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], true, false, false, false, false, true))
                            MyExcelDegree2018.江西公路沥青病害(excelApp, xlspath, _ProjectInfo, _DataDir, _RoadConfig);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], true, false, false, false, false, true))
                            MyExcelDegree2018.江西公路水泥病害(excelApp, xlspath, _ProjectInfo, _DataDir, _RoadConfig);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                _Setting.banMarkSign = true;
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, true))
                            MyExcelDegree2018.江西平整度(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], false, false, false, false, false, true, IsMeanMPD: true))
                            MyExcelDegree2018.江西磨耗(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], false, false, false, false, true, false))
                            MyExcelDegree2018.江西跳车(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], false, false, false, true, false, true))
                            MyExcelDegree2018.江西车辙(excelApp, xlspath, _ProjectInfo, _DataDir);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                _Setting.banMarkSign = false;
            }
        }
        private void GenerateExcel_Degree2001(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (IsOutputxls[2])
            {
                for (int i = 0; i < xlslen[2].Length; ++i)
                {
                    if (MyExcelDegree2001.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], true, true))
                    {
                        MyExcelDegree2001.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            if (IsOutputxls[3])
            {
                for (int i = 0; i < xlslen[3].Length; ++i)
                {
                    if (MyExcelDegree2001.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, true))
                    {
                        MyExcelDegree2001.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            if (IsOutputxls[4])
            {
                for (int i = 0; i < xlslen[4].Length; ++i)
                {
                    if (MyExcelDegree2001.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, true))
                    {
                        MyExcelDegree2001.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                    }
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
        }
        private void GenerateExcel_City_Dmi(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls, LaneProjectClass laneinfo = null)
        {
            //带GPS
            if (_Setting.ExcelType == 4)
            {
                if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, true, true, false, _Setting.PartType, false, true))
                {
                    MyExcelCity.OutputGPSAll2Xls_2_Dmi(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                }
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            }
            else if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, false, true, true, _Setting.PartType, false, false))
                        MyExcelCity.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[1])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, true, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[2])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, true, false, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[3])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, false, false, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[4])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, false, false, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[5])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, false, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[6])
                {
                    if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, true, false, false, _Setting.PartType, false, false))
                        MyExcelCity.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }


            }
            else if (_Setting.ExcelType == 7)
            {
                if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, true, true, false, _Setting.PartType, false, true))
                    MyExcelCity.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len, laneinfo);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, 10, false, true, true, false, false, 0, true, true))
                    MyExcelCity.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                if (MyExcelCity.InitProData(_DataDir, _ProjectInfo, 1, false, false, false, true, false, 0, false, false))
                    MyExcelCity.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, 1);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            }

        }
        private void GenerateExcel_City_Dmi_SH2013(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls, LaneProjectClass laneinfo = null)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, false, true, true, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[1])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, true, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[2])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, true, false, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[3])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, false, false, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[4])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, false, false, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[5])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, false, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
                if (IsOutputxls[6])
                {
                    if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, false, false, true, false, false, _Setting.PartType, false, false))
                        MyExcelCitySH2013.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len);
                    else
                        MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                }
            }
            else if (_Setting.ExcelType == 7)
            {
                if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, _Setting.PartType_Dmi_Len, true, true, true, true, false, _Setting.PartType, false, true))
                    MyExcelCitySH2013.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, _Setting.PartType_Dmi_Len, laneinfo);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, 10, false, true, true, false, false, 0, true, true))
                    MyExcelCitySH2013.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, 10);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

                if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, 1, false, false, false, true, false, 0, false, false))
                    MyExcelCitySH2013.OutputSHPG2Xls(excelApp, xlspath, _ProjectInfo, _DataDir, 1);
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            }
        }
        private void GenerateExcel_City_SH2013(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls)
        {
            if (_Setting.ExcelType == 0)
            {
                if (IsOutputxls[0])
                {
                    for (int i = 0; i < xlslen[0].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[0][i], false, false, false, true, true, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputRut(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[0][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[1])
                {
                    for (int i = 0; i < xlslen[1].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[1][i], false, false, true, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputMTD(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[1][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[2])
                {
                    for (int i = 0; i < xlslen[2].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[2][i], false, true, false, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputIRI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[2][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[3])
                {
                    for (int i = 0; i < xlslen[3].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[3][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputPCI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[3][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[4])
                {
                    for (int i = 0; i < xlslen[4].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[4][i], true, false, false, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputDis(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[4][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[5])
                {
                    for (int i = 0; i < xlslen[5].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[5][i], true, true, false, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputPQI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[5][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
                if (IsOutputxls[6])
                {
                    for (int i = 0; i < xlslen[6].Length; ++i)
                    {
                        if (MyExcelCitySH2013.InitProData(_DataDir, _ProjectInfo, xlslen[6][i], false, false, true, false, false, _Setting.PartType, false, false))
                            MyExcelCitySH2013.OutputPWI(excelApp, xlspath, _ProjectInfo, _DataDir, xlslen[6][i]);
                        else
                            MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                    }
                }
            }
        }

        /// <summary>
        /// 批量生成报表
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="xlspath">报表放置位置</param>
        /// <param name="xlslen">输出报表长度，二维数值</param>
        /// <param name="IsOutputxls">是否输出该指标，数组</param>
        public void GenerateExcel(MSExcel.Application excelApp, string xlspath, int[][] xlslen, bool[] IsOutputxls, LaneProjectClass laneinfo = null)
        {
            if (!_Setting.needSub)
            {
                if (LoadProjInfoData() == false) return;

            }

            switch (_Setting.ParmStyle)
            {
                //2007等级公路规范
                case StandardParmType.DegreeRoad2007: GenerateExcel_Degree2007(excelApp, xlspath, xlslen, IsOutputxls); break;

                //市政路报表
                case StandardParmType.CityRoad:
                    {
                        switch (_Setting.PartType)
                        {
                            case 0: GenerateExcel_City(excelApp, xlspath, xlslen, IsOutputxls); break;
                            case 1: GenerateExcel_City_Dmi(excelApp, xlspath, xlslen, IsOutputxls, laneinfo); break;
                            default: break;
                        }
                    }
                    break;

                //北京农村路报表
                case StandardParmType.RuralRoadBeijing: GenerateExcel_BeiJin(excelApp, xlspath, xlslen, IsOutputxls); break;

                //2018等级公路规范
                case StandardParmType.DegreeRoad2018:
                    {
                        if (xlslen == null)
                        {
                            //自动化检测报表输出
                            if (MyExcelDegreeSmall2018.InitProData(_DataDir, _ProjectInfo, 10, true, true, false, true, false, true))
                                MyExcelDegreeSmall2018.outPutAutoTest(excelApp, xlspath, _ProjectInfo, _DataDir);
                            else
                                MessageBox.Show("加载IRM数据失败，请先计算IRM！");
                        }
                        else
                        {
                            if (_Setting.SelectDrawDis == 1)//1
                            {
                                // 小方格
                                GenerateExcel_Degree2018_SmallRect(excelApp, xlspath, xlslen, IsOutputxls);
                            }
                            else
                            {
                                // 大方框
                                GenerateExcel_Degree2018_BigRect(excelApp, xlspath, xlslen, IsOutputxls);
                            }
                        }
                    }
                    break;
                case StandardParmType.RuralRoadlowLevel:
                    if (_Setting.SelectDrawDis == 1)
                    {
                        //小方格
                        GenerateExcel_VillageSmall(excelApp, xlspath, xlslen, IsOutputxls);
                    }
                    else
                    {
                        //大方框
                        GenerateExcel_Village(excelApp, xlspath, xlslen, IsOutputxls);
                    }
                    break;
                //2001等级公路规范
                case StandardParmType.DegreeRoad2001: GenerateExcel_Degree2001(excelApp, xlspath, xlslen, IsOutputxls); break;

                case StandardParmType.CityRoadShanghai:
                    {
                        switch (_Setting.PartType)
                        {
                            case 0: GenerateExcel_City_SH2013(excelApp, xlspath, xlslen, IsOutputxls); break;
                            case 1: GenerateExcel_City_Dmi_SH2013(excelApp, xlspath, xlslen, IsOutputxls, laneinfo); break;
                            default: break;
                        }
                    }
                    break;

                //辽宁农村路报表
                case StandardParmType.RuralRoadLiaoning: GenerateExcel_LiaoNing(excelApp, xlspath, xlslen, IsOutputxls); break;

                case StandardParmType.RuralRoadGuangxi: GenerateExcel_GuangXi(excelApp, xlspath, xlslen, IsOutputxls); break;

                case StandardParmType.RuralRoadChongqing: GenerateExcel_ChongQing(excelApp, xlspath, xlslen, IsOutputxls); break;
                case StandardParmType.RuralRoadHunan:
                    if (_Setting.SelectDrawDis == 1)
                    {
                        //小方格
                        GenerateExcel_HuNanSmall(excelApp, xlspath, xlslen, IsOutputxls);
                    }
                    else
                    {
                        //大方框
                        GenerateExcel_HuNan(excelApp, xlspath, xlslen, IsOutputxls);
                    }
                    break;

                default: break;
            }
        }

        public void BkExcel(MSExcel.Application excelApp, string xlspath)
        {
            if (!_Setting.needSub)
                if (LoadProjInfoData() == false) return;
            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007)
            {
                if (MyExcelDegree2007.InitProData(_DataDir, _ProjectInfo, 100, true, false, false, false, false))
                {
                    MyExcelDegree2007.OutputBkDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");

            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018)
            {
                if (MyExcelDegree2018.InitProData(_DataDir, _ProjectInfo, 100, true, false, false, false, false, false))
                {
                    MyExcelDegree2018.OutputBkDis(excelApp, xlspath, _ProjectInfo, _DataDir);
                }
                else
                    MessageBox.Show("加载IRM数据失败，请先计算IRM！");
            }
            else
            {
                MessageBox.Show("该功能仅支持等级公路2007和等级公路2018出表！");
                return;
            }

        }


        //1、路面图像文件重命名
        //2、修改几个路面txt里边的文件名
        //3、景观图像文件重命名
        //4、修改几个景观txt里边的文件名
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
            int dirnum = _ProjectInfo._EndDmi / ImgDis / 1000 + 1;
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
                        subfolder = System.IO.Path.GetFileName(dirname);
                        string[] imgsname = Directory.GetFiles(dirname, "*." + imgpath);
                        Array.Sort(imgsname);
                        foreach (string imgname in imgsname)
                        {
                            fname = System.IO.Path.GetFileName(imgname);
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
                tmile = _ProjectInfo.Dmi2Mile(tdmi);
                if (tmile <= 0 && _ProjectInfo._Direction < 0
                    || tmile >= _ProjectInfo._EndMile && _ProjectInfo._Direction > 0)
                {
                    break;
                }
                sw.WriteLine(string.Format("{0} {1}", tmile, ImgsList[i]), Encoding.UTF8);
                tdmi = tdmi + ImgDis;
            }

            #region 读取Trigger文件 cwb 20240622
            //string fpath = string.Format("{0}\\{1}Img\\SYN\\trigger.txt", _DataDir.FullName, ImgSource);
            //string[] syntrigstrs = File.ReadAllLines(fpath);
            //List<String> tempStrs = new List<string>();
            //syntrigstrs = syntrigstrs.Where(t => !(t.Contains("G") || t.Contains('g') || t.Contains('�'))
            //     && (t.StartsWith("%CDA") || t.StartsWith("%BDA") || t.StartsWith("%XRC"))).ToArray(); //过滤掉非16进制数
            //string number = ""; string preNumber = ""; 
            //int errorRow = 0;
            //if (File.Exists(fpath))
            //{
            //    //同样帧号取最后一个

            //    for (int i = 0; i < syntrigstrs.Length - 1; i++)
            //    {
            //        try
            //        {
            //            errorRow++;
            //            string nowStr = syntrigstrs[i];

            //            number = nowStr.Split(',')[1];

            //            string resultStr = syntrigstrs[i + 1];
            //            preNumber = resultStr.Split(',')[1];
            //            long lTemp0 = long.Parse(nowStr.Split(',')[2]);  //过滤掉异常值
            //            long lTemp1 = long.Parse(resultStr.Split(',')[2]);
            //            if (number == preNumber)
            //            {
            //                continue;
            //            }
            //            else
            //            {
            //                tempStrs.Add(nowStr);

            //            }
            //        }
            //        catch (Exception)
            //        {

            //            continue;
            //        }

            //    }
            //     tempStrs.Add(syntrigstrs.Last());
            //    //对异常数据进行过滤  
            //    List<String> tempStrs1 = new List<string>();
            //    if (_ProjectInfo._PlusLength == 0)
            //    {
            //        if (tempStrs.Count >= 3)
            //        {
            //            tempStrs1.Add(tempStrs[0]);
            //            for (int i = 1; i < tempStrs.Count - 1; i++)
            //            {
            //                string pre = tempStrs[i - 1];
            //                string now = tempStrs[i];
            //                string last = tempStrs[i + 1];

            //                long lTemp0 = long.Parse(now.Split(',')[2]) - long.Parse(pre.Split(',')[2]);
            //                long lTemp1 = long.Parse(last.Split(',')[2]) - long.Parse(now.Split(',')[2]);

            //                if (lTemp0 > 0 && lTemp1 > 0)
            //                {
            //                    tempStrs1.Add(now);
            //                }
            //                else
            //                {
            //                    continue;
            //                }

            //            }
            //            tempStrs1.Add(tempStrs.Last());
            //        }
            //    }
            //    else
            //    {
            //        if (tempStrs.Count >= 3)
            //        {
            //            tempStrs1.Add(tempStrs[0]);
            //            for (int i = 1; i < tempStrs.Count - 1; i++)
            //            {
            //                string pre = tempStrs[i - 1];
            //                string now = tempStrs[i];
            //                string last = tempStrs[i + 1];
            //                tempStrs1.Add(now);
            //            }
            //            tempStrs1.Add(tempStrs.Last());
            //        }
            //    }
            //} 
            //Dictionary<int,double> triggerDic = new Dictionary<int,double>();
            //List<(int,float)> triggerList = new List<(int, float)>();
            //long firstDmi = 0;

            //for (int i = 0; i < tempStrs.Count; i++)
            //{
            //    string[] values = tempStrs[i].Split(',');
            //    int index =int.Parse(values[1], System.Globalization.NumberStyles.HexNumber);
            //      long  dmi =long.Parse(values[3], System.Globalization.NumberStyles.HexNumber);
            //    if (i==0)
            //    {
            //        firstDmi = dmi;
            //    }
            //    triggerDic.Add(index, (dmi - firstDmi)/1000);
            //    triggerList.Add((index, (dmi - firstDmi) / 1000));

            //}
            //数据格式  帧号,里程
            //if (triggerList.Count != ImgsList.Count)
            //{
            //    Console.WriteLine("bug?");
            //}
            //for (int i = 0; i < imgnum; ++i)
            //{
            //    tmile = _ProjectInfo.Dmi2Mile(triggerList[i].Item2);
            //    if (tmile <= 0 && _ProjectInfo._Direction < 0
            //        || tmile >= _ProjectInfo._EndMile && _ProjectInfo._Direction > 0)
            //    {
            //        break;
            //    }
            //    sw.WriteLine(string.Format("{0} {1}", tmile, ImgsList[i]), Encoding.UTF8);
            //    //tdmi = tdmi + ImgDis;
            //}
            #endregion



            sw.Close();
            fw.Close();
        }

        //1、路面图像文件重命名
        //2、修改几个路面txt里边的文件名
        //3、景观图像文件重命名
        //4、修改几个景观txt里边的文件名
        private void Image2MileTrigger(string ImgSource, int CamIdx, int ImgDis, string imgpath)
        {
            //return;

            /////////////////////////////////////////////////////
            if (!Directory.Exists(string.Format("{0}\\{1}Img\\Camera{2}", _DataDir.FullName, ImgSource, CamIdx)))
            {
                return;
            }

            //获取所有的图像名
            string fname, subfolder;
            List<string> ImgsList = new List<string>();
            int dirnum = _ProjectInfo._EndDmi / ImgDis / 1000 + 1;
            int imgcnt = 0;//数的图像张数，会丢帧
            int imgtrigcnt = 0;//触发计数，从开始
            string tstr;
            string[] tstrs;
            for (int i = 0; i < dirnum; ++i)
            {
                string dirname = string.Format("{0}\\{1}Img\\Camera{2}\\Image_{3:0000}", _DataDir.FullName, ImgSource, CamIdx, i);
                try
                {
                    if (Directory.Exists(dirname))
                    {
                        subfolder = System.IO.Path.GetFileName(dirname);
                        string[] imgsname = Directory.GetFiles(dirname, "*." + imgpath);
                        Array.Sort(imgsname);
                        foreach (string imgname in imgsname)
                        {
                            fname = System.IO.Path.GetFileName(imgname);
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

            FileStream fw = new FileStream(string.Format("{0}\\{1}Img\\Camera{2}\\{1}2Mile.txt", _DataDir.FullName, ImgSource, CamIdx), FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < imgnum; ++i)
            {
                tmile = _ProjectInfo.Dmi2Mile(tdmi);
                if (tmile < 0)
                {
                    break;
                }
                sw.WriteLine(string.Format("{0} {1}", tmile, ImgsList[i]), Encoding.UTF8);
                tdmi = tdmi + ImgDis;
            }
            sw.Close();
            fw.Close();
        }

        private void ChangeImageName(string ImgSource, int CamIdx, string folderLastName = "Img")
        {
            string subfolder = string.Format("{0}{2}\\Camera{1}", ImgSource, CamIdx, folderLastName);
            if (!Directory.Exists(string.Format("{0}\\{1}", _DataDir.FullName, subfolder)))
            {
                return;
            }

            FileInfo fileInfo;
            string[] s, ts;
            int mile = 0;
            char[] schar = { ' ', '.' };
            string newname, imgpath, newimgpath, mstack;
            string mfname = string.Format("{0}\\{1}\\{2}2Mile.txt", _DataDir.FullName, subfolder, ImgSource);
            string[] strs = File.ReadAllLines(mfname);

            FileStream fw = new FileStream(mfname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);

            //路面图像文件重命名
            for (int i = 0; i < strs.Length; ++i)
            {
                s = strs[i].Split(schar);
                ts = s[1].Split('_');
                s[1] = string.Format("{0}.{1}", s[1], s[2]);

                mile = (int)Math.Round(Convert.ToDouble(s[0]));
                mstack = string.Format("K{0:0000}+{1:000}", mile / 1000, mile % 1000);
                newname = string.Format("{0}_{1}_{2}_{3}.{4}", ts[0], ts[1], ts[2], mstack, s[2]);

                imgpath = string.Format("{0}\\{1}{2}", _DataDir.FullName, subfolder, s[1]);
                newimgpath = string.Format("{0}\\{1}{2}", _DataDir.FullName, subfolder, newname);

                if (File.Exists(imgpath) && !File.Exists(newimgpath))
                {
                    fileInfo = new FileInfo(imgpath);
                    fileInfo.MoveTo(newimgpath);
                }


                imgpath = string.Format("{0}\\{1}{2}.txt", _DataDir.FullName, subfolder, s[1]);
                if (s[1].Contains(mstack))
                {
                    imgpath = string.Format("{0}\\{1}{2}.txt", _DataDir.FullName, subfolder, s[1].Replace("_" + mstack, ""));
                }
                newimgpath = string.Format("{0}\\{1}{2}.txt", _DataDir.FullName, subfolder, newname);
                if (File.Exists(imgpath) && !File.Exists(newimgpath))
                {
                    fileInfo = new FileInfo(imgpath);
                    fileInfo.MoveTo(newimgpath);
                }

                imgpath = string.Format("{0}\\{1}{2}_PartClass.txt", _DataDir.FullName, subfolder, s[1]);
                if (s[1].Contains(mstack))
                {
                    imgpath = string.Format("{0}\\{1}{2}_PartClass.txt", _DataDir.FullName, subfolder, s[1].Replace("_" + mstack, ""));
                }
                newimgpath = string.Format("{0}\\{1}{2}_PartClass.txt", _DataDir.FullName, subfolder, newname);
                if (File.Exists(imgpath) && !File.Exists(newimgpath))
                {
                    fileInfo = new FileInfo(imgpath);
                    fileInfo.MoveTo(newimgpath);
                }

                strs[i] = strs[i].Replace(s[1], newname);
                sw.WriteLine(strs[i]);



            }
            sw.Close();
            fw.Close();
        }


        ///将GPS时间和里程映射为一个文件，用里程进行插值，获得每2米的utc时间
        ///文件格式
        ///GPS时间 里程
        ///每2米一行，第一行的里程从0开始，中间不足的地方用线性插值获取GPS时间
        private bool GetRoadGPSTime2Dmi(string ImgSource)
        {
            string gpsModelFile = _DataDir.FullName + "\\GPSModel\\gps.txt";
            bool needSub1s = false;
            if (File.Exists(gpsModelFile))
            {
                needSub1s = true;
                if (_Setting.equipType == 1)
                {
                    //二三维设备 不用减去1s
                    needSub1s = false;
                }
            }
            else
            {
                needSub1s = false;
            }


            //获取触发序号、触发时间和里程
            string fpath = string.Format("{0}\\{1}Img\\SYN\\trigger.txt", _DataDir.FullName, ImgSource);
            if (!File.Exists(fpath)) return false;

            string[] syntrigstrs = File.ReadAllLines(fpath);
            List<String> tempStrs = new List<string>();

            syntrigstrs = syntrigstrs.Where(t => !(t.Contains("G") || t.Contains('g') || t.Contains('�'))
            && (t.StartsWith("%CDA") || t.StartsWith("%BDA") || t.StartsWith("%XRC"))).ToArray(); //过滤掉非16进制数

            string number = ""; string preNumber = "";

            int errorRow = 0;


            //同样帧号取最后一个

            for (int i = 0; i < syntrigstrs.Length - 1; i++)
            {
                try
                {
                    errorRow++;
                    string nowStr = syntrigstrs[i];

                    number = nowStr.Split(',')[1];

                    string resultStr = syntrigstrs[i + 1];
                    preNumber = resultStr.Split(',')[1];
                    long lTemp0 = long.Parse(nowStr.Split(',')[2]);  //过滤掉异常值
                    long lTemp1 = long.Parse(resultStr.Split(',')[2]);
                    if (number == preNumber)
                    {
                        continue;
                    }
                    else
                    {
                        tempStrs.Add(nowStr);

                    }
                }
                catch (Exception)
                {

                    continue;
                }

            }
            tempStrs.Add(syntrigstrs.Last());
            //对异常数据进行过滤  
            List<string> tempStrs1 = new List<string>();

            if (_ProjectInfo._PlusLength != 0)
            {
                if (tempStrs.Count >= 3)
                {
                    tempStrs1.Add(tempStrs[0]);
                    for (int i = 1; i < tempStrs.Count - 1; i++)
                    {
                        string pre = tempStrs[i - 1].Replace('o', '9');
                        string now = tempStrs[i].Replace(':', '9'); ;
                        string last = tempStrs[i + 1];
                        long lTemp0 = 0;
                        long lTemp1 = 0;
                        try
                        {
                            //辽宁奔驰综合检测车
#if 辽宁奔驰综合检测车
                            lTemp0 = long.Parse(now.Split(',')[3]) - long.Parse(pre.Split(',')[3]);
                            lTemp1 = long.Parse(last.Split(',')[3]) - long.Parse(now.Split(',')[3]);

#else
                            //其他老设备
                             lTemp0 = long.Parse(now.Split(',')[2]) - long.Parse(pre.Split(',')[3]);
                            lTemp1 = long.Parse(last.Split(',')[2]) - long.Parse(now.Split(',')[3]);
#endif
                        }
                        catch (Exception)
                        {

                            continue;
                        }

                        if (lTemp0 > 0 && lTemp1 > 0)
                        {
                            tempStrs1.Add(now);
                        }
                        else
                        {
                            continue;
                        }

                    }
                    tempStrs1.Add(tempStrs.Last());
                }
            }
            else
            {
                if (tempStrs.Count >= 3)
                {
                    tempStrs1.Add(tempStrs[0]);
                    for (int i = 1; i < tempStrs.Count - 1; i++)
                    {
                        string pre = tempStrs[i - 1];
                        string now = tempStrs[i];
                        string last = tempStrs[i + 1];
                        tempStrs1.Add(now);
                    }
                    tempStrs1.Add(tempStrs.Last());
                }
            }
            errorRow = 0;
            syntrigstrs = tempStrs1.ToArray();
            if (syntrigstrs.Length < 1) return false;
            string gPath = string.Format("{0}\\{1}Img\\SYN\\gps.txt", _DataDir.FullName, ImgSource);
            #region 20250709 修复由于同步版60未进位导致的bug
            int addTime = 0;
            if (File.Exists(gPath))
            {
                string[] gStrs = File.ReadLines(gPath).ToArray();
                if (gStrs.Length > 0)
                {
                    string ansLines = gStrs[gStrs.Length - 2];
                    if (ansLines.Contains("%ANS"))
                    {
                        List<string> oneDatas = ansLines.Split(',').ToList();
                        if (oneDatas.Count >= 10)
                        {
                            int findIndex = oneDatas.FindIndex(t => t == "S");
                            string timeStr = "";
                            if (findIndex == 8)
                            {
                                timeStr = oneDatas[6];
                            }
                            if (findIndex == 10)
                            {
                                timeStr = oneDatas[9];
                            }

                            //判断是否未进位 
                            if (timeStr.Length == 6)
                            {
                                int hour = int.Parse(timeStr.Substring(0, 2));
                                int min = int.Parse(timeStr.Substring(2, 2));
                                int mm = int.Parse(timeStr.Substring(4, 2));

                                if (min == 60)
                                {
                                    addTime = 60 * 61;//
                                }
                                if (mm == 60)
                                {
                                    addTime = 61; //需要加61s
                                }
                            }

                        }


                    }
                }



            }

            #endregion
            #region 千寻

            //cwb 20230321 
            //对千寻设备进行 跳秒问题处理 
            //在gps.txt 文件中进行查找    b562 开头的语句 
            //b56201261800f038c60600000000021202001cba8afd89080700000000034126$GPRMC,073356.00,A,3027.29518,N,11424.01042,E,0.378,,200323,,,A*7C
            //2120->18   
            //0711->17
            //查找授时时间 %ANS,20230320,073430226,20000101,000112,20230320,063633,80BC90,S,T,I,A,R,W,1000,0100,50,50,1000,02,02,0050,1000,XR_220214A,43
            //先获得该文件
            //获得同级目录下gps.txt文件 
            NeedSubOneSecTimeGpsHelp tempGpsHelper = null; string ymd = "";
            if (File.Exists(gPath))
            {
                var gStrs = File.ReadLines(gPath).ToList();
                var dataTmeps = gStrs.Where(t => t.Contains("%ANS")).ToList();
                string oTimeStr = gStrs.FirstOrDefault(t => t.Contains("%ANS"));


                //当前高精度定位有两套类型设备(由于外业版本不同的处理方式)
                //判定条件  最后一条语句   old:找不到 #0711或者#0212  new:找得到#0711 或者#0212
                bool isNewEquip = false;
                if (gStrs.Last().Contains("#0212") || gStrs.Last().Contains("#0711"))
                {
                    isNewEquip = true;
                }

                DateTime oTime = new DateTime();
                if (!string.IsNullOrEmpty(oTimeStr) && needSub1s)
                {
                    string[] splitOtime = oTimeStr.Split(',');

                    string oTimeS = "";
                    if (splitOtime.Length < 26)
                    {
                        oTimeS = splitOtime[5] + splitOtime[6];
                    }
                    else
                    {
                        oTimeS = splitOtime[8] + splitOtime[9];
                    }
                    var longTime = long.Parse(oTimeS);
                    //获得授时时间"yyyyMMddHHmmssfff"
                    oTime = DateTime.ParseExact(oTimeS, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                    ymd = oTime.Year.ToString() + oTime.Month.ToString("00") + oTime.Day.ToString("00");
                    tempGpsHelper = new NeedSubOneSecTimeGpsHelp(isNewEquip, gStrs.ToArray(), oTime, ymd);
                }
            }
            else
            {
                return false;
            }
            #endregion
            bool IsFirst = true;
            int tidx;
            int dmival = 0;
            int linecnt = 0;
            float startdmi0 = 0;
            float dmival0 = 0;
            string curtime = "000000000";
            string curIndex = "000000000";
            string ttstr;
            string[] strs;
            List<SynTrigInfo> syntriglist = new List<SynTrigInfo>();
            SynTrigInfo lastinfo = new SynTrigInfo();
            double firstDmi = 0;
            foreach (string linestr in syntrigstrs)
            {
                errorRow++;
                ++linecnt;
                if (linestr.StartsWith("%CDA") || linestr.StartsWith("%BDA"))
                {
                    try
                    {
                        tidx = linestr.LastIndexOf('%');
                        tidx = tidx > 0 ? tidx : 0;
                        ttstr = linestr.Substring(tidx);
                        ttstr = ttstr.Replace(':', '9');
                        ttstr = ttstr.Replace('o', '9');
                        strs = ttstr.Split(',');
                    }
                    catch
                    {
                        continue;
                    }
                    ;
                    if (strs.Length == 7)
                    {
                        try
                        {
                            if (_ProjectInfo._PlusLength != 0 && IsFirst)
                            {
                                firstDmi = Convert.ToDouble(strs[5]);
                            }
                            if (_ProjectInfo._PlusLength != 0)
                            {

                                double dmiCur = Convert.ToDouble(strs[5]) - firstDmi;
                                SynTrigInfo tinfo = new SynTrigInfo(strs[1], strs[2], strs[3], dmiCur.ToString());
                                if (IsFirst)
                                {
                                    syntriglist.Add(tinfo);
                                    IsFirst = false;
                                }
                                else if (tinfo._trigdmi != lastinfo._trigdmi)
                                {
                                    syntriglist.Add(tinfo);
                                }
                                lastinfo = tinfo;
                            }
                            else
                            {
                                SynTrigInfo tinfo = new SynTrigInfo(strs[1], strs[2], strs[3], strs[5]);
                                if (IsFirst && tinfo._trigdmi > 0)
                                {
                                    syntriglist.Add(tinfo);
                                    IsFirst = false;
                                }
                                else if (tinfo._trigdmi != lastinfo._trigdmi)
                                {
                                    syntriglist.Add(tinfo);
                                }
                                lastinfo = tinfo;
                            }

                        }
                        catch
                        { continue; }
                    }
                }
                else if (linestr.StartsWith("%XRC"))
                {
                    if (linestr.Length < 33)
                        continue;

                    strs = linestr.Split(',');
                    if (strs.Length < 3)
                        continue;

                    if (strs[2].Length == 9)
                    {
                        curtime = strs[2]; //格式为 074646988  hhmmssmmm
                        curIndex = strs[1]; //格式为 074646988  hhmmssmmm
                        string nowStr = curtime;
                        if (addTime != 0)
                        {
                            DateTime now = DateTime.ParseExact(nowStr, "HHmmssfff", System.Globalization.CultureInfo.InvariantCulture);
                            DateTime realNow = now.AddSeconds(addTime);
                            curtime = realNow.ToString("HHmmssfff");
                        }
                        #region 千寻设备 跳秒处理
                        if (tempGpsHelper != null)
                        {
                            //if (tempGpsHelper.NeedHandel) //需要处理-1s问题
                            if (true) //需要处理-1s问题
                            {
                                nowStr = ymd + curtime;
                                DateTime now = DateTime.ParseExact(nowStr, "yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture);

                                //  if (tempGpsHelper.IsNewEquip)
                                if (true)
                                {
                                    //同步版所有时间减去1s
                                    // 减去1秒
                                    DateTime realNow = now.AddSeconds(-1);

                                    // 将结果转换回字符串格式
                                    curtime = realNow.ToString("HHmmssfff");
                                }
                                else
                                {
                                    bool thisNeedSub = false;

                                    {
                                        DateTime sTime = tempGpsHelper.sTime;
                                        DateTime eTime = tempGpsHelper.eTime;
                                        if (now >= sTime && now <= eTime) //如果 now在 stime与etime之间
                                        {
                                            // 减去1秒
                                            DateTime realNow = now.AddSeconds(-1);

                                            // 将结果转换回字符串格式
                                            curtime = realNow.ToString("HHmmssfff");
                                            thisNeedSub = true;
                                        }
                                        if (!thisNeedSub)
                                        {

                                        }
                                    }

                                }
                            }
                        }

                        #endregion
                    }
                    else
                    {
                        continue;
                    }

                    if (strs[3].Length == 8)
                    {
                        try
                        {

                            dmival = Convert.ToInt32(strs[3], 16);
                            if (dmival < 0)
                            {
                                continue;
                            }
                            dmival0 = (float)(dmival * 0.001);

                        }
                        catch
                        {
                            continue;
                        }
                    }
                    else
                    {
                        continue;
                    }
                    if (strs[1].Length == 8)
                    {
                        try
                        {
                            if (Convert.ToInt32(strs[1], 16) == 0 && linecnt == 1)
                            {
                                startdmi0 = dmival0;
                            }
                        }

                        catch (Exception ex)
                        {
                            string errorStr = $"gps桩号匹配功能错误:文件{fpath}解析时在第{errorRow}行出现错误，请提交专业人员检查";
                            _log.Error(errorStr, ex);
                            continue;
                        }

                    }
                    else
                    {
                        continue;
                    }
                    //if (linecnt < 2)
                    //{
                    //    continue;
                    //}

                    dmival0 = dmival0 - startdmi0;
                    try
                    {

                        if (syntriglist.Count > 0)
                        {
                            //if (Math.Abs(syntriglist.Last()._trigdmi - dmival0) > 2000) //前后两个的间隔超过2000m明显不合理
                            //{
                            //    continue;
                            //}
                        }


                    }
                    catch (Exception ex)
                    {
                        string errorStr = $"gps桩号匹配功能错误:文件{fpath}解析时在第*{errorRow}*行出现错误，请提交专业人员检查";
                        _log.Error(errorStr, ex);
                        continue;
                    }

                    {
                        DateTime dtTemp;
                        if (DateTime.TryParseExact(curtime, "HHmmssfff", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dtTemp))
                        {
                            // 解析成功，dtTemp 包含解析后的 DateTime 对象
                            Console.WriteLine("Parsed DateTime: " + dtTemp.ToString());
                        }
                        else
                        {
                            // 解析失败
                            continue;
                        }
                    }
                    SynTrigInfo tinfo = new SynTrigInfo(curIndex, _ProjectInfo._DataDate, curtime, ((int)dmival0).ToString());

                    syntriglist.Add(tinfo);
                }
            }

            int triglen = syntriglist.Count;
            if (triglen < 3) return false;

            for (int i = triglen - 2; i >= 0; --i)
            {
                if (syntriglist[i]._trigdmi == syntriglist[i + 1]._trigdmi)
                {
                    syntriglist.RemoveAt(i + 1);
                }
            }
            triglen = syntriglist.Count;

            try
            {
                SynTrigInfo tempTrigInfo = null;
                List<string> gpstrigstrs = new List<string>();
                for (int i = 0; i < syntriglist[0]._trigdmi; i += 2)
                {
                    tempTrigInfo = new SynTrigInfo(syntriglist[0], syntriglist[1], i);
                    gpstrigstrs.Add(tempTrigInfo.ToString());
                }
                //绕远长度
                double otherDmi = 0;
                //总停车长度
                double sumStopLength = 0;


                //找到停车位置

                for (int i = 1; i < syntriglist.Count; ++i)
                {
                    double preDmi = syntriglist[i - 1]._trigdmi;
                    double nowDmi = syntriglist[i]._trigdmi;

                    otherDmi = (nowDmi - preDmi - 2);
                    if (otherDmi >= 50) //人为丢帧超过50m就是暂停采集了
                    {
                        sumStopLength += otherDmi;
                        for (int d = i; d < syntriglist.Count; d++)
                        {
                            double newDmi = syntriglist[d]._trigdmi - otherDmi;
                            //string newStr = syntriglist[d]._trigdmi + " " + newDmi.ToString();
                            syntriglist[d]._trigdmi = newDmi;
                        }
                    }
                }
                triglen = syntriglist.Count;



                //先对丢帧数据进行补齐
                for (int i = 1; i < triglen; ++i)
                {

                    gpstrigstrs.Add(syntriglist[i - 1].ToString());

                    for (double j = syntriglist[i - 1]._trigdmi + 2; j < syntriglist[i]._trigdmi; j += 2)
                    {

                        //如果 帧号差值*2 = 里程差值  就需要进行补充 

                        tempTrigInfo = new SynTrigInfo(syntriglist[i - 1], syntriglist[i], j);

                        gpstrigstrs.Add(tempTrigInfo.ToString());

                    }
                }

                triglen--;
                string lastMsg = syntriglist[triglen]._trigtime + " " + (syntriglist[triglen]._trigdmi - sumStopLength).ToString();
                gpstrigstrs.Add(lastMsg);
                for (double i = syntriglist[triglen]._trigdmi + 2; i <= _ProjectInfo._EndDmi; i += 2)
                {
                    try
                    {
                        tempTrigInfo = new SynTrigInfo(syntriglist[triglen - 1], syntriglist[triglen], i);
                        gpstrigstrs.Add(tempTrigInfo.ToString());
                    }
                    catch (Exception)
                    {

                        throw;
                    }

                }

                if (File.Exists(gPath))
                {
                    var gStrs = File.ReadLines(gPath).ToList();
                    GPSInfo tgpsinfoFirst = null;
                    for (int i = 0; i < gStrs.Count; i++)
                    {
                        GPSInfo tgpsinfo = new GPSInfo(gStrs[i]);
                        if (tgpsinfo._IsOK)
                        {
                            tgpsinfoFirst = tgpsinfo;
                            break;
                        }
                    }
                    if (tgpsinfoFirst == null)
                    {
                        //没有一条有效信息
                        return false;
                    }
                    if (gpstrigstrs.Count > 0)
                    {
                        //010206833
                        DateTime dt = new DateTime();
                        string time = gpstrigstrs[0].Split(' ').First();

                        dt = dt.AddHours(int.Parse(time.Substring(0, 2)));
                        dt = dt.AddMinutes(int.Parse(time.Substring(2, 2)));
                        dt = dt.AddSeconds(int.Parse(time.Substring(4, 5)) * 0.001);
                        TimeSpan diff = dt - tgpsinfoFirst._utctime; // 计算时间差
                        if (diff.TotalMinutes > 10)
                        {
                            //gps和 同步版 有效信息 时间少于10分钟 有问题
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }

                }

                File.WriteAllLines(string.Format("{0}\\GPSTime2Dmi.txt", _DataDir.FullName), gpstrigstrs.ToArray());

            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool GetIRIGPSTime2Dmi(int side)
        {
            string folderpath = string.Format("{0}\\IRIMTD\\DAQ{1}", _DataDir.FullName, side);
            if (!Directory.Exists(folderpath))
            { return false; }

            bool isfirst = true;
            int plusecnt = 0, tdmi = 0;
            int olddmival = 0, curdmival = 0;
            string linestr = null;
            List<string> gpstrigstrs = new List<string>();
            string[] daqs = Directory.GetFiles(folderpath, "*.daq");
            Array.Sort(daqs);
            if (daqs.Length < 1)
                return false;

            FileStream fr = null;
            StreamReader sr = null;
            foreach (string fname in daqs)
            {
                fr = File.OpenRead(fname);
                sr = new StreamReader(fr);
                while ((linestr = sr.ReadLine()) != null)
                {
                    string[] strs = linestr.Split(',');
                    if (strs.Length != 7)
                        continue;

                    if (isfirst)
                    {
                        isfirst = false;
                        olddmival = int.Parse(strs[4]);
                        gpstrigstrs.Add(string.Format("{0} {1}", strs[1], plusecnt / 40));
                    }
                    else
                    {
                        curdmival = int.Parse(strs[4]);
                        if (curdmival >= 26000 && olddmival < 26000)
                        {
                            plusecnt++;
                            if (plusecnt % 40 == 0)// 2米/50mm分频
                            {
                                tdmi = (plusecnt / 40)*2;
                                gpstrigstrs.Add(string.Format("{0} {1}", strs[1], tdmi));
                                if (tdmi >= _ProjectInfo._EndDmi)
                                {
                                    break;
                                }
                            }

                            //if (plusecnt % 80 == 0)// 2米/50mm分频
                            //{
                            //    tdmi = plusecnt / 80;
                            //    gpstrigstrs.Add(string.Format("{0} {1}", strs[1], tdmi));
                            //    if (tdmi >= _ProjectInfo._EndDmi)
                            //    {
                            //        break;
                            //    }
                            //}
                        }
                        olddmival = curdmival;
                    }
                }
                sr.Close();
                fr.Close();
            }
            if (gpstrigstrs.Count <= 0)
                return false;

            File.WriteAllLines(string.Format("{0}\\GPSTime2Dmi.txt", _DataDir.FullName), gpstrigstrs.ToArray());
            return true;
        }

        private bool GetRutGPSTime2Dmi(int side)
        {
            string fpath = string.Format("{0}\\camera{1}\\trigger.txt", _DataDir.FullName, side);
            if (!File.Exists(fpath)) return false;
            else
            {
                string datestr = ",20190101,";
                bool isfirt = true;
                List<string> gpstrigstrs = new List<string>();
                int dmival = 0, dmival2 = 0, sidx = 0;
                float dmival0 = 0;
                string linestr, olddmistr = "1000000", curdmistr = "0000000", curdmistr2 = "0000000", curtime = "000000000";
                string[] strs;
                FileStream fr = File.OpenRead(fpath);
                StreamReader sr = new StreamReader(fr);
                while ((linestr = sr.ReadLine()) != null)
                {
                    if (linestr.StartsWith("%BDA"))
                    {
                        if (linestr.Length < 52)
                            continue;

                        if (isfirt)
                        {
                            strs = linestr.Split(',');
                            if (strs.Length < 2)
                                continue;

                            if (strs[2] == "20000101")
                                return false;
                            else
                            {
                                if (strs.Length < 6)
                                    continue;

                                if (strs[2].Length == 8)
                                {
                                    isfirt = false;
                                    datestr = string.Format(",{0},", strs[2]);
                                    curdmistr = strs[5];
                                    if (curdmistr != olddmistr)
                                    {
                                        curdmistr2 = strs[6];
                                        try { dmival = int.Parse(curdmistr); dmival2 = int.Parse(curdmistr2); }
                                        catch { continue; }
                                        dmival0 = (float)(dmival * 2 + dmival2 * 0.001);
                                        gpstrigstrs.Add(string.Format("{0} {1}", strs[3], dmival0));
                                    }
                                    olddmistr = strs[5];
                                }
                            }
                        }
                        else
                        {
                            sidx = linestr.IndexOf(datestr);
                            if (sidx > 0)
                            {
                                curdmistr = linestr.Substring(sidx + 22, 7);
                                curtime = linestr.Substring(sidx + 10, 9);
                                if (curdmistr != olddmistr)
                                {
                                    curdmistr2 = linestr.Substring(sidx + 30, 7);
                                    try { dmival = int.Parse(curdmistr); dmival2 = int.Parse(curdmistr2); }
                                    catch { continue; }
                                    dmival0 = (float)(dmival * 2 + dmival2 * 0.001);
                                    gpstrigstrs.Add(string.Format("{0} {1}", curtime, dmival0));
                                    if (dmival >= _ProjectInfo._EndDmi)
                                    {
                                        break;
                                    }
                                }
                                olddmistr = curdmistr;
                            }
                            else
                            {
                                isfirt = true;
                            }
                        }
                    }
                    else if (linestr.StartsWith("%XRC"))
                    {
                        if (linestr.Length < 33)
                            continue;

                        strs = linestr.Split(',');
                        if (strs.Length < 3)
                            continue;

                        if (strs[2].Length == 9)
                        {
                            curtime = strs[2];
                        }
                        else
                        {
                            continue;
                        }

                        if (strs[3].Length == 8)
                        {
                            dmival = Convert.ToInt32(strs[3], 16);
                            dmival0 = (float)(dmival * 0.001);
                        }
                        else
                        {
                            continue;
                        }

                        gpstrigstrs.Add(string.Format("{0} {1}", curtime, dmival0));
                    }
                }
                sr.Close();
                fr.Close();

                if (gpstrigstrs.Count <= 0)
                    return false;

                File.WriteAllLines(string.Format("{0}\\GPSTime2Dmi.txt", _DataDir.FullName), gpstrigstrs.ToArray());
                return true;
            }
        }

        /// <summary>
        /// 根据gps的原始数据 和 GPSTime2Dmi.txt 生成UTC时间、经度、纬度、高程、里程、桩号映射文件
        /// 采样GPS时间关联原始GPS数据经纬度高程 和 桩号里程，用一次函数线性插值
        /// 文件格式：
        /// UTC时间 经度 纬度 高程 里程 桩号
        /// </summary>
        /// <param name="gpsfile"></param>
        private bool GetGPSMileMapping(string gps_fname)
        {
            int i = 0, j = 0;
            int hour = 0;
            int minute = 0;
            double second = 0;

            //解析GPS数据
            List<GPSInfo> GPSInfoList = new List<GPSInfo>();
            List<string> GPSInfoStrs = new List<string>();
            string[] gps_strs = File.ReadAllLines(gps_fname);
            int gps_len = gps_strs.Length;
            double elevation = 0;
            for (i = 0; i < gps_len; ++i)
            {
   
                GPSInfo tgpsinfo = new GPSInfo(gps_strs[i]);

                if (tgpsinfo._IsOK)
                {
                    if (tgpsinfo._elevation < 0)
                    {

                    }
                    if (tgpsinfo._elevation != 0)
                    {
                        elevation = tgpsinfo._elevation;
                    }
                    if (GPSInfoList.Count > 0)
                    {
                        if (tgpsinfo._utctime != GPSInfoList[GPSInfoList.Count - 1]._utctime)
                        {
                            GPSInfoList.Add(tgpsinfo);
                            GPSInfoStrs.Add(tgpsinfo.ToString());
                        }
                        if (tgpsinfo._utctime == GPSInfoList[GPSInfoList.Count - 1]._utctime)
                        {
                            //高程
                            if (GPSInfoList[GPSInfoList.Count - 1]._elevation == 0)
                            {
                                GPSInfoList[GPSInfoList.Count - 1]._elevation = elevation;
                            }

                        }
                    }
                    else
                    {
                        GPSInfoList.Add(tgpsinfo);
                        GPSInfoStrs.Add(tgpsinfo.ToString());
                    }
                    //if (tgpsinfo.ToString()== "034629000 103.7594128 27.0869155 0.00")
                    //{

                    //}
                }
            }
            gps_len = GPSInfoList.Count;
            if (gps_len < 5)
            {
                //MessageBox.Show("处理失败！" + gps_fname + "文件有效GPS位置信息少于5条,点击确定将尝试\n自动使用其他gps.txt文件进行计算");
                return false;
            }
            else
            {
                File.WriteAllLines(string.Format("{0}\\GPSInfo.txt", _DataDir.FullName), GPSInfoStrs.ToArray());
            }

            // 跨天的时候，日期要加1
            System.TimeSpan t2;
            for (i = 1; i < gps_len; ++i)
            {
                t2 = GPSInfoList[i]._utctime - GPSInfoList[i - 1]._utctime;
                if (t2.TotalSeconds < -72000.0f)
                {
                    break;
                }
            }
            for (; i < gps_len; ++i)
            {
                GPSInfoList[i]._utctime = GPSInfoList[i]._utctime.AddDays(1);
            }

            List<string> GPSMileStrList = new List<string>();
            List<MapGPSMile> GPSMileList = new List<MapGPSMile>();
            string utc_dmi_fname = string.Format("{0}\\GPSTime2Dmi.txt", _DataDir.FullName);
            string[] utc_dmi_strs = File.ReadAllLines(utc_dmi_fname);
            int utc_dmi_len = utc_dmi_strs.Length;
            for (i = 0; i < utc_dmi_len; ++i)
            {

                string[] strs = utc_dmi_strs[i].Split(' ');
                MapGPSMile tmap = new MapGPSMile();
                tmap._gpsinfo = new GPSInfo();
                tmap._gpsinfo._utctime = new DateTime();

                try
                {
                    hour = int.Parse(strs[0].Substring(0, 2));
                    minute = int.Parse(strs[0].Substring(2, 2));
                    second = int.Parse(strs[0].Substring(4, 5)) * 0.001;
                }
                catch (System.Exception)
                {
                    continue;
                }

                tmap._gpsinfo._utctime = tmap._gpsinfo._utctime.AddHours(hour);
                tmap._gpsinfo._utctime = tmap._gpsinfo._utctime.AddMinutes(minute);
                tmap._gpsinfo._utctime = tmap._gpsinfo._utctime.AddSeconds(second);

                tmap._dmi = float.Parse(strs[1]);
                tmap._mile = _ProjectInfo.Dmi2Mile(tmap._dmi);

                GPSMileList.Add(tmap);
            }

            // 跨天的时候，日期要加1
            System.TimeSpan t3;
            utc_dmi_len = GPSMileList.Count;
            for (i = 1; i < utc_dmi_len; ++i)
            {
                //  两个时间相减 。默认得到的是 两个时间之间的天数   得到：365.00:00:00  
                t3 = GPSMileList[i]._gpsinfo._utctime - GPSMileList[i - 1]._gpsinfo._utctime;
                if (t3.TotalSeconds < -72000.0f)
                {
                    break;
                }

                // 出现了跳秒
                else if (t3.TotalSeconds < 0)
                {
                    if (t3.TotalSeconds > -1)
                    {
                        GPSMileList[i]._gpsinfo._utctime = GPSMileList[i]._gpsinfo._utctime.AddSeconds(1);
                    }
                    else if (t3.TotalSeconds > -10 && t3.TotalSeconds < -9)
                    {
                        GPSMileList[i]._gpsinfo._utctime = GPSMileList[i]._gpsinfo._utctime.AddSeconds(10);
                    }
                    //else if (t3.TotalSeconds > -60 && t3.TotalSeconds < -59)
                    //{
                    //    GPSMileList[i]._gpsinfo._utctime = GPSMileList[i]._gpsinfo._utctime.AddSeconds(60);
                    //}
                    //else if (t3.TotalSeconds > -600 && t3.TotalSeconds < -599)
                    //{
                    //    GPSMileList[i]._gpsinfo._utctime = GPSMileList[i]._gpsinfo._utctime.AddSeconds(600);
                    //}
                }
            }
            for (; i < utc_dmi_len; ++i)
            {
                GPSMileList[i]._gpsinfo._utctime = GPSMileList[i]._gpsinfo._utctime.AddDays(1);
            }

            //查找触发同步里面的utc时间 和 在gps信息里面utc时间的上一时刻和下一时刻
            //开始进行插值
            bool isfind = false;
            List<DataIdx> SynIdx = new List<DataIdx>();
            for (i = 0; i < utc_dmi_len; ++i)
            {
                isfind = false;
                for (j = 1; j < gps_len; ++j)
                {
                    if (DateTime.Compare(GPSInfoList[j - 1]._utctime, GPSMileList[i]._gpsinfo._utctime) <= 0
                        && DateTime.Compare(GPSInfoList[j]._utctime, GPSMileList[i]._gpsinfo._utctime) > 0)
                    {
                        DateTime temp = GPSMileList[i]._gpsinfo._utctime;
                        GPSMileList[i]._gpsinfo = new GPSInfo(GPSInfoList[j - 1], GPSInfoList[j], temp);
                        GPSMileStrList.Add(GPSMileList[i].ToString());

                        isfind = true;
                        break;
                    }
                }
                if (!isfind)
                {
                    if (DateTime.Compare(GPSInfoList[0]._utctime, GPSMileList[i]._gpsinfo._utctime) > 0)
                    {
                        DateTime temp = GPSMileList[i]._gpsinfo._utctime;
                        GPSMileList[i]._gpsinfo = new GPSInfo(GPSInfoList[0], GPSInfoList[1], temp);
                        GPSMileStrList.Add(GPSMileList[i].ToString());
                    }

                    if (DateTime.Compare(GPSInfoList[GPSInfoList.Count - 1]._utctime, GPSMileList[i]._gpsinfo._utctime) <= 0)
                    {
                        DateTime temp = GPSMileList[i]._gpsinfo._utctime;
                        GPSMileList[i]._gpsinfo = new GPSInfo(GPSInfoList[GPSInfoList.Count - 2], GPSInfoList[GPSInfoList.Count - 1], temp);
                        GPSMileStrList.Add(GPSMileList[i].ToString());
                    }
                }
                //if (GPSMileList[i].ToString() == "034628354 103.7594347 27.0869667 1402.98 3706 3661")
                //{

                //}
            }

            if (GPSMileStrList.Count == GPSMileList.Count)
            {
                File.WriteAllLines(string.Format("{0}\\GPS2Mile.txt", _DataDir.FullName), GPSMileStrList.ToArray());
            }
            else
            {
                MessageBox.Show("处理失败！+" + gps_fname + "文件存在 GPS数据 和 同步数据 没有搜索到关联时间！,点击确定将尝试\n自动使用其他gps.txt文件进行计算");
                return false;
            }
            return true;
        }

        public void MappingGPS2Mile()
        {
            if (!_Setting.needSub)
                if (LoadProjInfoData() == false) return;

            bool IsIRIMap = false;
            bool IsRoadMap = false;
            bool IsRutMap = false;
            bool IsStreetMap = false;

            //将GPS时间和里程关联起来
            //string fileGps = _DataDir.FullName + "\\gps.dat";
            ////使用gpsModel的设备 时间上需要减去1秒  20240308 cwb


            //bool isHas = false;
            //if (File.Exists(fileGps))
            //{
            //    isHas = true;
            //}
            //else
            //    isHas = false;
            if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
            {
                IsRoadMap = GetRoadGPSTime2Dmi("Road");
                //采用左侧车辙的数据进行同步
                if (!IsRoadMap)
                {
                    IsRutMap = GetRutGPSTime2Dmi(0);
                    if (!IsRutMap)
                    {
                        IsRutMap = GetRutGPSTime2Dmi(1);
                        //采用左侧平整度的数据进行同步
                        if (!IsRutMap)
                        {
                            IsIRIMap = GetIRIGPSTime2Dmi(0);
                            if (!IsIRIMap)
                            {
                                IsIRIMap = GetIRIGPSTime2Dmi(1);
                                //采用景观的数据进行同步
                                if (!IsIRIMap)
                                {
                                    IsStreetMap = GetRoadGPSTime2Dmi("Street");
                                }
                            }
                        }
                    }
                }
            }
            else
            {

                if (this._ProjectInfo._PlusLength != 0)
                {

                    IsRutMap = GetRutGPSTime2Dmi(0);
                    if (!IsRutMap)
                    {
                        IsRutMap = GetRutGPSTime2Dmi(1);
                        //采用左侧平整度的数据进行同步
                        if (!IsRutMap)
                        {
                            IsRoadMap = GetRoadGPSTime2Dmi("Road");
                            //采用左侧车辙的数据进行同步
                            if (!IsRoadMap)
                            {
                                IsIRIMap = GetIRIGPSTime2Dmi(0);
                                if (!IsIRIMap)
                                {
                                    IsIRIMap = GetIRIGPSTime2Dmi(1);
                                    //采用景观的数据进行同步
                                    if (!IsIRIMap)
                                    {
                                        IsStreetMap = GetRoadGPSTime2Dmi("Street");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // IsStreetMap = GetRoadGPSTime2Dmi("Street");
                    IsRutMap = GetRutGPSTime2Dmi(0);
                    if (!IsRutMap)
                    {
                        IsRutMap = GetRutGPSTime2Dmi(1);
                        //采用左侧平整度的数据进行同步
                        if (!IsRutMap)
                        {
                            IsRoadMap = GetRoadGPSTime2Dmi("Road");
                            //采用左侧车辙的数据进行同步
                            if (!IsRoadMap)
                            {
                                IsIRIMap = GetIRIGPSTime2Dmi(0);
                                if (!IsIRIMap)
                                {
                                    IsIRIMap = GetIRIGPSTime2Dmi(1);
                                    //采用景观的数据进行同步
                                    if (!IsIRIMap)
                                    {
                                        IsStreetMap = GetRoadGPSTime2Dmi("Street");
                                    }
                                }
                            }
                        }
                    }
                }

            }

            ///生成了 GPSTime2Dmi.txt GPS时间和里程映射文件
            ///进行GPS经纬度和里程、桩号映射
            bool ok = false;
            if (IsRoadMap || IsRutMap || IsIRIMap || IsStreetMap)
            {
                string modelName = "";
                if (IsRoadMap)
                {
                    modelName = "\\RoadImg\\SYN";
                }
                if (IsRutMap)
                {
                    modelName = "\\camera0\\gps.txt";
                }
                if (IsIRIMap)
                {
                    modelName = "\\IRIMTD\\SYN0";
                }
                if (IsStreetMap)
                {
                    modelName = "\\StreetImg\\SYN";

                }
                string[] subpath = {"\\GPSModel\\gps.txt","\\gps.dat", "\\RoadImg\\SYN\\gps.txt", "\\StreetImg\\SYN\\gps.txt",
                                       "\\IRIMTD\\SYN0\\gps.txt","\\IRIMTD\\SYN1\\gps.txt","\\camera0\\gps.txt","\\camera1\\gps.txt"};
                string gpsfile = null;
                int i = 0;
                List<string> gpsResultMsg = new List<string>();//记录gps生成信息
                string gpsResultMsgFilePath = _DataDir.FullName + "\\GenerateGPSMessage.txt";

                for (i = 0; i < subpath.Length; i++)
                {
                    gpsfile = _DataDir.FullName + subpath[i];
                    if (File.Exists(gpsfile) && !ok)
                    {
                        ok = GetGPSMileMapping(gpsfile);
                        if (!ok)
                        {
                            continue;
                        }
                        else
                        {


                            // 判断gps2Mile长度是否在5%范围内
                            string gps2milefile = _DataDir.FullName + "\\GPSInfo.txt";
                            double sumLength = calGpsLength(gps2milefile);
                            gps2milefile = _DataDir.FullName + "\\GPS2Mile.txt";
                            double gps2MileLength = calGpsLength(gps2milefile);
                            if (sumLength == 0 || gps2MileLength == 0)
                            {
                                continue;
                            }
                            double wucha = Math.Round(((gps2MileLength - sumLength) / sumLength) * 100, 2);
                            string gpsMsgLine = $"根据{modelName}下文件解析得到,GPS轨迹里程：{sumLength},匹配桩号后轨迹里程：{gps2MileLength}。误差大小{wucha}%。";
                            gpsResultMsg.Add(gpsMsgLine);
                            if (sumLength * 0.90 <= gps2MileLength && gps2MileLength <= sumLength * 1.1)
                            {
                                //符合要求
                            }
                            else
                            {
                                //修改该文件夹名称
                                if (IsRoadMap || IsStreetMap)
                                {
                                    gpsResultMsg.Clear();
                                    string gpsDirPath = _DataDir.FullName + modelName;
                                    RenameFolder(gpsDirPath, "SYN_ERROR");
                                    MappingGPS2Mile();
                                }
                            }

                        }
                        break;
                    }
                }
                if (gpsResultMsg.Count > 0)
                {

                    File.WriteAllLines(gpsResultMsgFilePath, gpsResultMsg);
                }

                if (i < subpath.Length)
                {

                }
                else
                {
                    //  MessageBox.Show(_ProjectInfo._RoadName + "工程没有gps文件！");
                }

                if (!ok)
                {
                    MessageBox.Show("处理失败！" + _DataDir.FullName + "所有gps文件有效GPS位置信息少于5条,gps计算失败！");
                }
            }
            else
            {
                MessageBox.Show($"{_DataDir.FullName}生成 GPSTime2Dmi.txt GPS时间和里程映射文件 失败！");
            } 
                string gpsFile = _DataDir.FullName + "\\GPSModel\\gps.txt"; 
                if (File.Exists(gpsFile))
                {
                    bool isOk = HighAccuracyPositioning.writeHighAccPicture(_DataDir.FullName, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight, _Setting.equipType);

                }

            
        }

        private static double calGpsLength(string fileName)
        {
            double gps2MileLength = 0;

            if (File.Exists(fileName))
            {
                string[] gps2mileStrs = File.ReadAllLines(fileName);
                List<double[]> validPoints = new List<double[]>();

                // 第一步：解析所有有效点
                foreach (var line in gps2mileStrs)
                {
                    string[] strs = line.Split(' ');
                    if (strs.Length < 4) continue;

                    try
                    {
                        double lon = double.Parse(strs[1]);
                        double lat = double.Parse(strs[2]);
                        // ele 未使用但保留解析
                        double ele = double.Parse(strs[3]);
                        validPoints.Add(new double[2] { lon, lat });
                    }
                    catch (FormatException)
                    {
                        // 忽略格式错误的数据点
                        continue;
                    }
                }

                // 第二步：计算累积距离
                if (validPoints.Count >= 2)
                {
                    for (int d = 1; d < validPoints.Count; d++)
                    {
                        double[] prevPoint = validPoints[d - 1];
                        double[] currPoint = validPoints[d];

                        double distance = CalculateHaversineDistance(
                            prevPoint[1], prevPoint[0], // 前一点纬度，经度
                            currPoint[1], currPoint[0]  // 当前点纬度，经度
                        );

                        gps2MileLength += distance;
                    }
                }
            }

            return Math.Round(gps2MileLength, 2);
        }
        private static void RenameFolder(string folderPath, string newName)
        {
            try
            {
                // 获取父目录
                string parentDir = System.IO.Path.GetDirectoryName(folderPath);
                // 组合新的路径
                string newPath = System.IO.Path.Combine(parentDir, newName);

                // 检查原文件夹是否存在
                if (Directory.Exists(folderPath))
                {
                    // 检查新路径是否已经存在
                    if (Directory.Exists(newPath))
                    {
                        throw new InvalidOperationException("目标文件夹已存在。");
                    }

                    Directory.Move(folderPath, newPath);
                    Console.WriteLine("重命名成功！");
                }
                else
                {
                    Console.WriteLine("原文件夹不存在。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
            }
        }

        // 使用Haversine公式计算两点间距离（单位：米）
        private static double CalculateHaversineDistance(double lat1, double lon1, double lat2, double lon2)
        {
            const double R = 6371000; // 地球平均半径（米）
            var dLat = (lat2 - lat1) * Math.PI / 180;
            var dLon = (lon2 - lon1) * Math.PI / 180;

            var a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
                    Math.Cos(lat1 * Math.PI / 180) *
                    Math.Cos(lat2 * Math.PI / 180) *
                    Math.Sin(dLon / 2) * Math.Sin(dLon / 2);

            var c = 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
            return R * c;
        }

        /// 拼接全景图像
        /// 函数用多线程实现
        public void StitchPanoImg(WinPanoProcessBar bars)
        {
            _PanoBars = bars;
            if (IsExitPanoData())
            {
                StartPanoThread(new ThreadInfo(_DataDir.FullName, 0, _PanoBars, _ProjectInfo));
            }
            while (_ThreadPano != null && _ThreadPano.IsAlive)
            {
                Thread.Sleep(1000);
            }
        }
        Thread _ThreadPano;
        private void StartPanoThread(ThreadInfo prjinfo)
        {
            _ThreadPano = new Thread(PanoThreadMethod) { IsBackground = true };
            _ThreadPano.Start(prjinfo);
        }
        private void PanoThreadMethod(object prj)
        {
            ThreadInfo tinfo = (ThreadInfo)prj;
            WinPanoProcessBar bar = tinfo._bar as WinPanoProcessBar;
            MyPanoStitch.StitchImg(tinfo._prjname, bar);

            // 拼接完成之后，重新加载工程
            Image2Mile("Pano", 0, _ProjectInfo._PanoImgDis, "jpeg");
            if (_Setting.IsRename)
            {
                ChangeImageName("Pano", 0);
            }

            if (_winpano != null)
            {
                _winpano.InitPano();
            }
        }
        private bool IsExitPanoData()
        {
            string inifile = _DataDir.FullName + @"\Setting.ini";
            if (File.Exists(inifile))
            {
                IniFiles inisetting = new IniFiles(inifile);
                return inisetting.ReadBool("工作模式", "Pano", false);
            }
            else
            {
                return false;
            }
        }

        //清除IRM等中间计算结果
        public void CleanIRMVal()
        {
            List<string> fnames = new List<string>();
            if (SelectIRM.irm[0])
            {
                fnames.Add(@"\IRIMTD\DAQ0\IRI_10m.txt");
                fnames.Add(@"\IRIMTD\DAQ1\IRI_10m.txt");
                fnames.Add(@"\IRIMTD\DAQ0\IRI_20m.txt");
                fnames.Add(@"\IRIMTD\DAQ1\IRI_20m.txt");
                fnames.Add(@"\IRIMTD\DAQ0\IRI_100m.txt");
                fnames.Add(@"\IRIMTD\DAQ1\IRI_100m.txt");
                fnames.Add(@"\IRIMTD\DAQ0\IRI_1000m.txt");
                fnames.Add(@"\IRIMTD\DAQ1\IRI_1000m.txt");
                fnames.Add(@"\IRIMTD\DAQ0\PavementBump.txt");
                fnames.Add(@"\IRIMTD\DAQ1\PavementBump.txt");
                fnames.Add(@"\IRIMTD\DAQ0\ReSample250.txt");
                fnames.Add(@"\IRIMTD\DAQ1\ReSample250.txt");
                fnames.Add(@"\IRIMTD\DAQ0\Speed_10m.txt");
                fnames.Add(@"\IRIMTD\DAQ1\Speed_10m.txt");
            }
            if (SelectIRM.irm[1])
            {
                fnames.Add(@"\RUT\camera0\orioldrut.txt");
                fnames.Add(@"\RUT\camera1\orioldrut.txt");
                fnames.Add(@"\RUT\camera0\orirut.txt");
                fnames.Add(@"\RUT\camera1\orirut.txt");
                fnames.Add(@"\RUT\camera0\rut.txt");
                fnames.Add(@"\RUT\camera1\rut.txt");
                fnames.Add(@"\RUT\maxorirut.txt");
            }
            if (SelectIRM.irm[2])
            {
                fnames.Add(@"\IRIMTD\Laser0\MTD_10m.txt");
                fnames.Add(@"\IRIMTD\Laser1\MTD_10m.txt");
                fnames.Add(@"\IRIMTD\Laser2\MTD_10m.txt");
            }
            if (SelectIRM.irm[3])
            {
                fnames.Add(@"\IRIMTD\Laser0\MPD_10m.txt");
                fnames.Add(@"\IRIMTD\Laser1\MPD_10m.txt");
                fnames.Add(@"\IRIMTD\Laser2\MPD_10m.txt");
            }
            if (SelectIRM.irm[4])
            {
                fnames.Add(@"\camera0\imu.hon.csv");
                fnames.Add(@"\camera0\imu.hon.CrossSlope");
                fnames.Add(@"\camera0\imu.hon.Curvature");
                fnames.Add(@"\camera0\imu.hon.HeightSlope");
            }

            foreach (string fname in fnames)
            {
                if (File.Exists(_DataDir.FullName + fname))
                {
                    try
                    {
                        File.Delete(_DataDir.FullName + fname);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("清除计算结果失败！\r\n" + ex.Message);
                    }
                }
            }
        }
        //生成简易文档
        public void CreateSlimProj(string newPath)
        {
            List<string> fnames = new List<string>();
            //string newFullName = _DataDir.Parent.FullName + "_slim//" + _DataDir.Name;
            string newFullName = newPath + "\\" + _DataDir.Name;
            Directory.CreateDirectory(newFullName);

            //此处原始文件可以继续精简
            fnames.Add(@"\RUT\camera0\orirut.txt");
            fnames.Add(@"\RUT\camera1\orirut.txt");

            fnames.Add(@"\IRIMTD\Laser0\MTD_10m.txt");
            fnames.Add(@"\IRIMTD\Laser1\MTD_10m.txt");
            fnames.Add(@"\IRIMTD\Laser2\MTD_10m.txt");

            fnames.Add(@"\IRIMTD\Laser0\MPD_10m.txt");
            fnames.Add(@"\IRIMTD\Laser1\MPD_10m.txt");
            fnames.Add(@"\IRIMTD\Laser2\MPD_10m.txt");

            fnames.Add(@"\IRIMTD\Laser0\Setting.ini");
            fnames.Add(@"\IRIMTD\Laser1\Setting.ini");
            fnames.Add(@"\IRIMTD\Laser2\Setting.ini");

            fnames.Add(@"\IRIMTD\DAQ0\IRI_10m.txt");
            fnames.Add(@"\IRIMTD\DAQ0\Speed_10m.txt");
            fnames.Add(@"\IRIMTD\DAQ0\Setting.ini");
            fnames.Add(@"\IRIMTD\DAQ0\IRI_100.txt");
            fnames.Add(@"\IRIMTD\DAQ0\PavementBump.txt");
            fnames.Add(@"\IRIMTD\DAQ0\ReSample250.txt");
            fnames.Add(@"\IRIMTD\DAQ0\Resample.txt");
            fnames.Add(@"\IRIMTD\DAQ0\resample.txt");

            fnames.Add(@"\IRIMTD\DAQ1\IRI_10m.txt");
            fnames.Add(@"\IRIMTD\DAQ1\Speed_10m.txt");
            fnames.Add(@"\IRIMTD\DAQ1\Setting.ini");
            fnames.Add(@"\IRIMTD\DAQ1\IRI_100.txt");
            fnames.Add(@"\IRIMTD\DAQ1\PavementBump.txt");
            fnames.Add(@"\IRIMTD\DAQ1\ReSample250.txt");
            fnames.Add(@"\IRIMTD\DAQ1\resample.txt");
            fnames.Add(@"\IRIMTD\SYN0\gps.txt");
            fnames.Add(@"\RoadImg\Camera0\Road2Mile.txt");
            fnames.Add(@"\RoadImg\SYN\gps.txt");
            fnames.Add(@"\RoadImg\SYN\trigger.txt");
            fnames.Add(@"\StreetImg\SYN\gps.txt");
            fnames.Add(@"\StreetImg\SYN\trigger.txt");
            fnames.Add(@"\StreetImg\Camera0\Street2Mile.txt");
            fnames.Add(@"\StreetImg\Camera1\Street2Mile.txt");
            fnames.Add(@"\CamSetting.ini");
            fnames.Add(@"\DegreeInfo.txt");
            fnames.Add(@"\Dmi2Mile.txt");
            fnames.Add(@"\GPS2Mile.txt");
            fnames.Add(@"\GPSInfo.txt");
            fnames.Add(@"\GPSTime2Dmi.txt");
            fnames.Add(@"\ProjectInfo.txt");
            fnames.Add(@"\RoadTypeInfo.txt");
            fnames.Add(@"\Setting.ini");
            fnames.Add(@"\RoadStatuMarkInfo.txt");
            fnames.Add(@"\MileStoneCaliInfo.txt");

            foreach (string fname in fnames)
            {
                if (File.Exists(_DataDir.FullName + fname))
                {
                    try
                    {
                        DirectoryInfo Prjdir = new DirectoryInfo(newFullName + fname);
                        if (!Directory.Exists(Prjdir.Parent.FullName))
                        {
                            Directory.CreateDirectory(Prjdir.Parent.FullName);
                        }

                        File.Copy(_DataDir.FullName + fname, newFullName + fname, true);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("生成简易工程失败！\r\n" + ex.Message);
                    }
                }
            }
            #region 复制病害
            //路面病害
            string tdir = _DataDir.FullName + "\\RoadImg\\Camera0\\Image_0000";
            string picBaseDir = _DataDir.FullName + "\\RoadImg\\Camera0";
            DirectoryInfo picDir = new DirectoryInfo(picBaseDir);
            DirectoryInfo[] picPaths = picDir.GetDirectories();
            foreach (var roaddir in picPaths)
            {
                if (Directory.Exists(roaddir.FullName))
                {
                    string tpath = newFullName + "\\RoadImg\\Camera0\\" + roaddir.Name;
                    Directory.CreateDirectory(tpath);

                    FileInfo[] srcfiles = roaddir.GetFiles("*.txt");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string newfilepath = tpath + "\\" + tfile.Name;
                        File.Copy(tfile.FullName, newfilepath, true);
                    }



                }
            }

            //景观病害
            try
            {
                tdir = _DataDir.FullName + "\\StreetImg\\Camera0\\Image_0000";
                DirectoryInfo roaddir = new DirectoryInfo(tdir);
                if (Directory.Exists(roaddir.FullName))
                {
                    string tpath = newFullName + "\\StreetImg\\Camera0\\Image_0000";
                    Directory.CreateDirectory(tpath);

                    FileInfo[] srcfiles = roaddir.GetFiles("*.txt");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string newfilepath = tpath + "\\" + tfile.Name;
                        File.Copy(tfile.FullName, newfilepath, true);
                    }
                    srcfiles = roaddir.GetFiles("*.rbd");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string newfilepath = tpath + "\\" + tfile.Name;
                        File.Copy(tfile.FullName, newfilepath, true);
                    }
                }
            }
            catch
            {
            }
            try
            {

                //双景观病害
                tdir = _DataDir.FullName + "\\StreetImg\\Camera1\\Image_0000";
                DirectoryInfo roaddir = new DirectoryInfo(tdir);
                if (Directory.Exists(roaddir.FullName))
                {
                    string tpath = newFullName + "\\StreetImg\\Camera1\\Image_0000";
                    Directory.CreateDirectory(tpath);

                    FileInfo[] srcfiles = roaddir.GetFiles("*.txt");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string newfilepath = tpath + "\\" + tfile.Name;
                        File.Copy(tfile.FullName, newfilepath, true);
                    }
                    srcfiles = roaddir.GetFiles("*.rbd");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string newfilepath = tpath + "\\" + tfile.Name;
                        File.Copy(tfile.FullName, newfilepath, true);
                    }
                }
            }
            catch
            {
            }
            #endregion

            #region 生成中间结果
            DirectoryInfo prjdir = new DirectoryInfo(_DataDir.FullName);
            string[] _RoadGradeStr = null;
            switch (_Setting.ParmStyle)
            {
                case StandardParmType.CityRoadShanghai:
                case StandardParmType.CityRoad:
                    {
                        _RoadGradeStr = new string[4];
                        _RoadGradeStr[0] = "快速路";
                        _RoadGradeStr[1] = "主干路";
                        _RoadGradeStr[2] = "次干路";
                        _RoadGradeStr[3] = "支路";
                        break;
                    }
                case StandardParmType.DegreeRoad2007:

                case StandardParmType.RuralRoadBeijing:
                case StandardParmType.DegreeRoad2018:
                case StandardParmType.DegreeRoad2001:

                case StandardParmType.RuralRoadLiaoning:
                case StandardParmType.RuralRoadGuangxi:
                case StandardParmType.RuralRoadChongqing:
                case StandardParmType.RuralRoadHunan:
                case StandardParmType.RuralRoadlowLevel:
                default:
                    {
                        _RoadGradeStr = new string[5];
                        _RoadGradeStr[0] = "高速公路";
                        _RoadGradeStr[1] = "一级公路";
                        _RoadGradeStr[2] = "二级公路";
                        _RoadGradeStr[3] = "三级公路";
                        _RoadGradeStr[4] = "四级公路";
                        break;
                    }
            }

            Dictionary<string, int> _RoadGradeDict;
            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }
            MilePart spart = null;
            try
            {
                spart = new MilePart() { dmi = 0, roadtype = _ProjectInfo._RoadType, mile = _ProjectInfo._StartMile, roaddegree = _RoadGradeDict[_ProjectInfo._RoadGrade], degreestr = _ProjectInfo._RoadGrade };
            }
            catch
            {
            }
            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
            {
                return;
            }
            List<MilePart> _RoadPart10 = new List<MilePart>(); _RoadPart10.Add(spart);
            GlobalExcel.GetAllMilePart(prjdir.FullName, _ProjectInfo, 10, _ProjectInfo._Direction, _RoadGradeStr, ref _RoadPart10, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
            double[] _LDeltaHVal = null;
            double[] _RDeltaHVal = null;
            GlobalExcel.GetDeltaHVal(_ProjectInfo, prjdir, _RoadPart10, 0, ref _LDeltaHVal);
            if (_LDeltaHVal != null)
            {
                string[] listLDeltaHVal = new string[_LDeltaHVal.Length];
                for (int i = 0; i < _LDeltaHVal.Length; i++)
                {
                    listLDeltaHVal[i] = string.Format("{0}", _LDeltaHVal[i]);
                }
                File.WriteAllLines(prjdir.FullName + "\\IRIMTD\\DAQ0\\DeltaHVal.txt", listLDeltaHVal.ToArray(), Encoding.UTF8);
                if (File.Exists(prjdir.FullName + "\\IRIMTD\\DAQ0\\DeltaHVal.txt"))
                {
                    File.Copy(prjdir.FullName + "\\IRIMTD\\DAQ0\\DeltaHVal.txt", newFullName + "\\IRIMTD\\DAQ0\\DeltaHVal.txt", true);

                }
            }


            if (_ProjectInfo._IsDIRIMTD)
            {
                GlobalExcel.GetDeltaHVal(_ProjectInfo, prjdir, _RoadPart10, 1, ref _RDeltaHVal);
                if (_RDeltaHVal != null)
                {
                    string[] listRDeltaHVal = new string[_RDeltaHVal.Length];
                    for (int i = 0; i < _LDeltaHVal.Length; i++)
                    {
                        listRDeltaHVal[i] = string.Format("{0}", _RDeltaHVal[i]);
                    }
                    File.WriteAllLines(prjdir.FullName + "\\IRIMTD\\DAQ1\\DeltaHVal.txt", listRDeltaHVal.ToArray(), Encoding.UTF8);
                    File.Copy(prjdir.FullName + "\\IRIMTD\\DAQ1\\DeltaHVal.txt", newFullName + "\\IRIMTD\\DAQ1\\DeltaHVal.txt", true);
                }

            }

            #endregion

        }

        //检查原始数据的完整性
        public void CheckOriDataComplete(ref List<string> errorProjectInfo)
        {
            //农村路遇到一二级进行提醒
            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
            {
                if (_ProjectInfo._RoadGrade.Contains("高速") || _ProjectInfo._RoadGrade.Contains("一级") || _ProjectInfo._RoadGrade.Contains("二级"))
                {
                    string msg = string.Format("提示！检测到低等级农村路标准下工程{0}({1})道路等级为【高速|一级|二级】",
                       _ProjectInfo._RoadName, _ProjectInfo._RoadCode);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
            }

            //检查桩号是否合法、工程数据的完整性
            if (_ProjectInfo._Direction < 0)
            {
                if (_ProjectInfo._StartMile <= _ProjectInfo._EndMile)
                {
                    string msg = string.Format("不合法数据，请检查！\n起点桩号：K{0}+{1:000}\n终点桩号：K{2}+{3:000}\n行车方向：下行",
                        _ProjectInfo._StartMile / 1000, _ProjectInfo._StartMile % 1000, _ProjectInfo._EndMile / 1000, _ProjectInfo._EndMile % 1000);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
            }
            else if (_ProjectInfo._Direction > 0)
            {
                if (_ProjectInfo._StartMile >= _ProjectInfo._EndMile)
                {
                    string msg = string.Format("不合法数据，请检查！\n起点桩号：K{0}+{1:000}\n终点桩号：K{2}+{3:000}\n行车方向：上行",
                        _ProjectInfo._StartMile / 1000, _ProjectInfo._StartMile % 1000, _ProjectInfo._EndMile / 1000, _ProjectInfo._EndMile % 1000);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
            }

            if (!File.Exists(_DataDir.FullName + "\\Setting.ini"))
            {
                string msg = string.Format("缺少工程配置文件：{0}\\Setting.ini\n请从其他同配置工程拷贝同名文件到以上路径", _DataDir.FullName);
                WriteCheckRes(msg, ref errorProjectInfo);
            }
            //else
            //{
            //    IniFiles iniset = new IniFiles(_DataDir.FullName + "\\Setting.ini");
            //    iniset.WriteBool("工作模式", "SIRIMTD", true);
            //    iniset.WriteBool("工作模式", "DIRIMTD", false);
            //    iniset.WriteBool("工作模式", "MMTD", false);
            //}
            //检查车辙数据的完整性
            if (_ProjectInfo._IsRut)
            {
                CheckRutData(_DataDir.FullName, 0, ref errorProjectInfo);
                if (_ProjectInfo._RutMode == 1)
                {
                    CheckRutData(_DataDir.FullName, 1, ref errorProjectInfo);
                }
            }

            //检查平整度构造深度数据的完整性
            if (_ProjectInfo._IsIRIMTD && !_Setting.isGDIriCalculate)
            {
                if (!Directory.Exists(_DataDir.FullName + "\\IRIMTD"))
                {
                    string msg = string.Format("缺少平整度构造深度文件夹：{0}\\IRIMTD", _DataDir.FullName);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    CheckIRIData(_DataDir.FullName, 0, ref errorProjectInfo);
                    CheckMTDData(_DataDir.FullName, 0, ref errorProjectInfo);

                    if (_ProjectInfo._IsDIRIMTD)
                    {
                        CheckIRIData(_DataDir.FullName, 1, ref errorProjectInfo);
                        CheckMTDData(_DataDir.FullName, 1, ref errorProjectInfo);

                        if (_ProjectInfo._IsMMTD)
                        {
                            CheckMTDData(_DataDir.FullName, 2, ref errorProjectInfo);
                        }
                    }
                }
            }

            //检查路面数据的完整性
            if (_ProjectInfo._IsRoad)
            {
                if (!Directory.Exists(_DataDir.FullName + "\\RoadImg"))
                {
                    string msg = string.Format("缺少路面文件夹：{0}\\RoadImg", _DataDir.FullName);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    CheckImgData(_DataDir.FullName + "\\RoadImg", 0, _ProjectInfo._RoadImgDis, ref errorProjectInfo);
                }
            }

            //检查景观数据的完整性
            if (_ProjectInfo._IsStreet)
            {
                if (!Directory.Exists(_DataDir.FullName + "\\StreetImg"))
                {
                    string msg = string.Format("缺少景观文件夹：{0}\\StreetImg", _DataDir.FullName);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    CheckImgData(_DataDir.FullName + "\\StreetImg", 0, _ProjectInfo._StreetImgDis_Left, ref errorProjectInfo);
                    if (_ProjectInfo._IsDStreet)
                    {
                        string path = _DataDir.FullName + "\\StreetImg" + "\\Camera1";
                        if (Directory.Exists(path))
                        {
                            CheckImgData(_DataDir.FullName + "\\StreetImg", 1, _ProjectInfo._StreetImgDis_Right, ref errorProjectInfo);

                        }
                        else
                        {
                            CheckImgData(_DataDir.FullName + "\\StreetImg2", 0, _ProjectInfo._StreetImgDis_Right, ref errorProjectInfo);

                        }

                    }
                }
            }
        }

        private void CheckImgData(string fpath, int idx, int imgdis, ref List<string> errorProjectInfo)
        {
            string campath = fpath + "\\Camera" + idx.ToString();
            if (!Directory.Exists(campath))
            {
                string msg = string.Format("缺少文件夹：{0}", campath);
                WriteCheckRes(msg, ref errorProjectInfo);
            }

            int imgnum = 0;
            DirectoryInfo camdir = new DirectoryInfo(campath);
            DirectoryInfo[] imgdirs = camdir.GetDirectories();
            foreach (DirectoryInfo imgdir in imgdirs)
            {
                FileInfo[] imgfiles = imgdir.GetFiles("*.jpg");
                imgnum = imgnum + imgfiles.Length;
            }

            imgnum = imgnum * imgdis;
            if (imgnum < _ProjectInfo._EndDmi - imgdis * 5)
            {
                if (fpath.Contains("StreetImg"))
                {
                    string msg = string.Format("文件夹：{0}，缺少图像{1}张，设置拍照距离{2}", fpath, (_ProjectInfo._EndDmi / imgdis - imgnum / imgdis), imgdis);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    string msg = string.Format("文件夹：{0}，缺少图像{1}张", fpath, (_ProjectInfo._EndDmi / imgdis - imgnum / imgdis));
                    WriteCheckRes(msg, ref errorProjectInfo);


                }
            }
            else if (imgnum > _ProjectInfo._EndDmi + imgdis * 5)
            {
                if (fpath.Contains("StreetImg"))
                {
                    string msg = string.Format("文件夹：{0}，多采图像{1}张，设置拍照距离{2}", fpath, (imgnum / imgdis - _ProjectInfo._EndDmi / imgdis), imgdis);
                    WriteCheckRes(msg, ref errorProjectInfo);
                    //if (imgdis == 20)
                    //{
                    //    int timgnum = imgnum / 2;
                    //    if (timgnum <= _ProjectInfo._EndDmi + 50)
                    //    {
                    //        string inifpath = _DataDir.FullName + "\\Setting.ini";
                    //        IniFiles iniset = new IniFiles(inifpath);
                    //        iniset.WriteInteger("Parm", "StreetDis", 10);
                    //    }
                    //    else
                    //    {

                    //    }
                    //}
                    //else
                    //{
                    //    string msg = string.Format("文件夹：{0}，多采图像{1}张，设置拍照距离{2}", fpath, (imgnum / imgdis - _ProjectInfo._EndDmi / imgdis), imgdis);
                    //    WriteCheckRes(msg);
                    //}
                }
                else
                {
                    string msg = string.Format("文件夹：{0}，多采图像{1}张，设置拍照距离{2}", fpath, (imgnum / imgdis - _ProjectInfo._EndDmi / imgdis), imgdis);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
            }

            string synpath = fpath + "\\SYN";
            if (!Directory.Exists(synpath))
            {
                string msg = string.Format("缺少文件夹：{0}", synpath);
                WriteCheckRes(msg, ref errorProjectInfo);
            }
            else
            {
                string synfpath = synpath + "\\trigger.txt";
                if (!File.Exists(synfpath))
                {
                    string msg = string.Format("缺少文件：{0}", synfpath);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
            }
        }

        private void CheckRutData(string fpath, int idx, ref List<string> errorProjectInfo)
        {
            FileInfo tfile = null;
            if (!Directory.Exists(string.Format("{0}\\camera{1}", fpath, idx)))
            {
                string msg = string.Format("缺少车辙文件夹：{0}\\camera{1}", fpath, idx);
                WriteCheckRes(msg, ref errorProjectInfo);
            }
            else
            {
                if (!Directory.Exists(string.Format("{0}\\camera{1}\\data", fpath, idx)))
                {
                    string msg = string.Format("缺少车辙文件：{0}\\camera{1}\\data", fpath, idx);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    long dirlen = 0;
                    MainForm.GetDirectoryLength(string.Format("{0}\\camera{1}\\data", fpath, idx), ref dirlen, "*.dat");
                    if (dirlen <= 0)
                    {
                        string msg = string.Format("车辙文件夹：{0}\\camera{1}\\data为空", fpath, idx);
                        WriteCheckRes(msg, ref errorProjectInfo);
                    }
                    else
                    {
                        if (!File.Exists(string.Format("{0}\\camera{1}\\rutcfg.ini", fpath, idx)))
                        {
                            string msg = string.Format("缺少车辙文件：{0}\\camera{1}\\rutcfg.ini", fpath, idx);
                            WriteCheckRes(msg, ref errorProjectInfo);
                        }
                        else
                        {
                            tfile = new FileInfo(string.Format("{0}\\camera{1}\\rutcfg.ini", fpath, idx));
                            if (tfile.Length == 0)
                            {
                                string msg = string.Format("车辙文件：{0}\\camera{1}\\rutcfg.ini为空", fpath, idx);
                                WriteCheckRes(msg, ref errorProjectInfo);
                            }
                        }

                        if (!File.Exists(string.Format("{0}\\camera{1}\\c2cali.c2w", fpath, idx)))
                        {
                            string msg = string.Format("缺少车辙文件：{0}\\camera{1}\\c2cali.c2w", fpath, idx);
                            WriteCheckRes(msg, ref errorProjectInfo);
                        }
                        else
                        {
                            tfile = new FileInfo(string.Format("{0}\\camera{1}\\c2cali.c2w", fpath, idx));
                            if (tfile.Length == 0)
                            {
                                string msg = string.Format("车辙文件：{0}\\camera{1}\\c2cali.c2w为空", fpath, idx);
                                WriteCheckRes(msg, ref errorProjectInfo);
                            }
                        }
                    }
                }
            }
        }

        private void CheckIRIData(string fpath, int idx, ref List<string> errorProjectInfo)
        {
            FileInfo tfile = null;
            if (!Directory.Exists(string.Format("{0}\\IRIMTD\\DAQ{1}", fpath, idx)))
            {
                string msg = string.Format("缺少平整度数据文件夹：{0}\\IRIMTD\\DAQ{1}", fpath, idx);
                WriteCheckRes(msg, ref errorProjectInfo);
            }
            else
            {
                if (!File.Exists(string.Format("{0}\\IRIMTD\\DAQ{1}\\Resample.txt", fpath, idx)) && !_Setting.isGDIriCalculate)
                {
                    string msg = string.Format("缺少平整度文件：{0}\\IRIMTD\\DAQ{1}\\Resample.txt", fpath, idx);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    tfile = new FileInfo(string.Format("{0}\\IRIMTD\\DAQ{1}\\Resample.txt", fpath, idx));

                    if (tfile.Length == 0)
                    {
                        string msg = string.Format("平整度文件：{0}\\IRIMTD\\DAQ{1}\\Resample.txt为空", fpath, idx);
                        WriteCheckRes(msg, ref errorProjectInfo);
                    }

                    else if (!File.Exists(string.Format("{0}\\IRIMTD\\DAQ{1}\\Setting.ini", fpath, idx)))
                    {
                        string msg = string.Format("缺少平整度文件：{0}\\IRIMTD\\DAQ{1}\\Setting.ini", fpath, idx);
                        WriteCheckRes(msg, ref errorProjectInfo);
                    }
                    else
                    {
                        tfile = new FileInfo(string.Format("{0}\\IRIMTD\\DAQ{1}\\Setting.ini", fpath, idx));
                        if (tfile.Length == 0)
                        {
                            string msg = string.Format("平整度文件：{0}\\IRIMTD\\DAQ{1}\\Setting.ini为空", fpath, idx);
                            WriteCheckRes(msg, ref errorProjectInfo);
                        }
                    }
                }
            }
        }

        private void CheckMTDData(string fpath, int idx, ref List<string> errorProjectInfo)
        {
            FileInfo tfile = null;
            if (!Directory.Exists(string.Format("{0}\\IRIMTD\\Laser{1}", fpath, idx)) && _Setting.ParmStyle != StandardParmType.RuralRoadlowLevel)
            {
                string msg = string.Format("缺少构造深度数据文件夹：{0}\\IRIMTD\\Laser{1}", fpath, idx);
                WriteCheckRes(msg, ref errorProjectInfo);
            }
            else
            {
                long laslen = 0;
                MainForm.GetDirectoryLength(string.Format("{0}\\IRIMTD\\Laser{1}", fpath, idx), ref laslen, "*.las");
                if (laslen <= 0)
                {
                    string msg = string.Format("缺少构造深度原始数据：{0}\\IRIMTD\\Laser{1}\\*.las", fpath, idx);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else if (!File.Exists(string.Format("{0}\\IRIMTD\\Laser{1}\\Setting.ini", fpath, idx)))
                {
                    string msg = string.Format("缺少构造深度文件：{0}\\IRIMTD\\Laser{1}\\Setting.ini", fpath, idx);
                    WriteCheckRes(msg, ref errorProjectInfo);
                }
                else
                {
                    tfile = new FileInfo(string.Format("{0}\\IRIMTD\\Laser{1}\\Setting.ini", fpath, idx));
                    if (tfile.Length == 0)
                    {
                        string msg = string.Format("构造深度文件：{0}\\IRIMTD\\Laser{1}\\Setting.ini为空", fpath, idx);
                        WriteCheckRes(msg, ref errorProjectInfo);
                    }
                }
            }
        }

        private void WriteCheckRes(string msg, ref List<string> errorProjectInfo)
        {
            //  MessageBox.Show(msg);
            msg = msg + "\n";
            File.AppendAllText(MainForm.chktxt_fpath + "\\数据检查结果.txt", msg);
        }

        private void dockPanel_Map_Click(object sender, EventArgs e)
        {

        }
    }

    //将GPS时间、桩号与图像进行关联
    public class SynTrigInfo
    {
        public string _trigdate = null;//同步触发时的GPS日期
        public string _trigtime = null;//同步触发时的GPS时间
        public double _trigdmi = 0;//触发时的里程
        public int _FrameIndex = 0;

        public SynTrigInfo() { }

        public SynTrigInfo(string index, string date, string time, string dmi)
        {
            _trigdate = date;
            _trigtime = time;
            _trigdmi = Convert.ToDouble(dmi);
            _FrameIndex = int.Parse(index, System.Globalization.NumberStyles.HexNumber);

        }

        /// <summary>
        /// 用里程进行插值，获得每2米的utc时间
        /// </summary>
        /// <param name="STrig">插值起点的触发信息</param>
        /// <param name="ETrig">插值终点的触发信息</param>
        /// <param name="insertidx">要插入的触发序号</param>
        public SynTrigInfo(SynTrigInfo STrig, SynTrigInfo ETrig, double insertdmi)
        {
            try
            {
                double x = (insertdmi - STrig._trigdmi) * 1.0 / (ETrig._trigdmi - STrig._trigdmi);

                DateTime sdate = DateTime.ParseExact(STrig._trigdate + STrig._trigtime, "yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture);
                DateTime edate = DateTime.ParseExact(ETrig._trigdate + ETrig._trigtime, "yyyyMMddHHmmssfff", System.Globalization.CultureInfo.InvariantCulture);

                System.TimeSpan pdate = edate - sdate;
                double pdatemsecond = pdate.TotalMilliseconds * x;
                System.TimeSpan cpdate = new TimeSpan((long)(pdatemsecond * 10000));
                DateTime cdate = sdate.Add(cpdate);
                _trigdate = cdate.ToString("yyyyMMdd");
                _trigtime = cdate.ToString("HHmmssfff");
                _trigdmi = insertdmi;
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        public override string ToString()
        {
            return string.Format("{0} {1}", _trigtime, _trigdmi);
        }
    }

    public class GPSInfo
    {
        /// <summary>
        /// 高程，单位m
        /// </summary>
        public double _elevation;

        /// <summary>
        /// 纬度，单位度
        /// </summary>
        public double _latitude;

        /// <summary>
        /// 经度，单位度
        /// </summary>
        public double _longitude;

        /// <summary>
        /// UTC时间，当天的
        /// </summary>
        public DateTime _utctime;

        /// <summary>
        /// 解析是否成功
        /// </summary>
        public bool _IsOK = false;

        //public string _utcstr;

        public static double _GGACorrection;

        public GPSInfo() { }

        /// <summary>
        /// GGA、FPD数据解析
        /// </summary>
        /// <param name="infostr"></param>
        public GPSInfo(string infostr)
        {

            int signCount = infostr.Count(t => t == '$');
            if (signCount != 1)
            {
                _IsOK = false;
                return;

            }
            string[] strs = infostr.Split(',');
            if (strs[0].Contains("GGA") && strs.Length == 15 && infostr.Length <= 90)
            {
                if (strs[1].Length == 9 && strs[2].Length > 0 && strs[4].Length > 0 && strs[9].Length > 0 && strs[11].Length > 0)
                {
                    try
                    {
                        int hour = int.Parse(strs[1].Substring(0, 2));
                        int minute = int.Parse(strs[1].Substring(2, 2));
                        double second = double.Parse(strs[1].Substring(4, 5));

                        _utctime = new DateTime();
                        _utctime = _utctime.AddHours(hour);
                        _utctime = _utctime.AddMinutes(minute);
                        _utctime = _utctime.AddSeconds(second);
                    }
                    catch
                    {

                        _IsOK = false;
                        return;
                    }


                    if (!ConvertDegreesToDigital(strs[2], 2, out _latitude))
                    {
                        _IsOK = false;
                        return;
                    }

                    if (!ConvertDegreesToDigital(strs[4], 3, out _longitude))
                    {
                        _IsOK = false;
                        return;
                    }
                    #region cwb 20230816 
                    if (3.86 > _latitude || _latitude > 53.55)
                    {
                        _IsOK = false;
                        return;
                    }
                    if (73.66 > _longitude || _longitude > 135.05)
                    {
                        _IsOK = false;
                        return;
                    }
                    #endregion
                    try
                    {
                        _elevation = double.Parse(strs[9]);
                        string temp = strs[10];

                        _GGACorrection = double.Parse(strs[11]);
                        if (temp != "M")
                        {
                            _IsOK = false;
                        }
                        else
                        {

                        }
                    }
                    catch
                    {
                        _IsOK = false;
                    }
                    _IsOK = true;
                }
                else
                {
                    _IsOK = false;
                }
            }
            else if (strs[0].Contains("FPD") && strs.Length > 8)
            {
                if (strs[2].Length > 0 && strs[6].Length > 0 && strs[7].Length > 0 && strs[8].Length > 0)
                {
                    double weeksecond = (double.Parse(strs[2]) - 18) * 1000;//XW-G6615D2回传的周秒里面有18s的跳秒，减掉
                    double second = (int)weeksecond % 86400000;

                    DateTime sdate = DateTime.ParseExact("000000000", "HHmmssfff", System.Globalization.CultureInfo.InvariantCulture);
                    System.TimeSpan cpdate = new TimeSpan((long)(second * 10000));
                    _utctime = sdate.Add(cpdate);

                    _latitude = double.Parse(strs[6]);
                    _longitude = double.Parse(strs[7]);
                    _elevation = double.Parse(strs[8]) - _GGACorrection;
                    _IsOK = true;
                }
            }
            else if (strs[0].Contains("GPRMC") && strs.Length > 8)
            {

                if (strs[1].Length == 9 && strs[3].Length > 0 && strs[5].Length > 0 && strs[2] == "A")
                {
                    try
                    {
                        int hour = int.Parse(strs[1].Substring(0, 2));
                        int minute = int.Parse(strs[1].Substring(2, 2));
                        double second = double.Parse(strs[1].Substring(4, 5));

                        _utctime = new DateTime();
                        _utctime = _utctime.AddHours(hour);
                        _utctime = _utctime.AddMinutes(minute);
                        _utctime = _utctime.AddSeconds(second);
                    }
                    catch
                    {

                        _IsOK = false;
                        return;
                    }


                    if (!ConvertDegreesToDigital(strs[3], 2, out _latitude))
                    {
                        _IsOK = false;
                        return;
                    }

                    if (!ConvertDegreesToDigital(strs[5], 3, out _longitude))
                    {
                        _IsOK = false;
                        return;
                    }
                    #region cwb 20230816 
                    if (3.86 > _latitude || _latitude > 53.55)
                    {
                        _IsOK = false;
                        return;
                    }
                    if (73.66 > _longitude || _longitude > 135.05)
                    {
                        _IsOK = false;
                        return;
                    }
                    #endregion
                    try
                    {
                        _elevation = 0;
                    }
                    catch { }
                    if (_elevation == 0)
                    {

                    }
                    _IsOK = true;
                }
                else
                {
                    _IsOK = false;
                }
            }
            else
            {
                _IsOK = false;
            }
        }

        /// <summary>
        /// 经纬度 转 度
        /// </summary>
        /// <param name="inval"></param>
        /// <param name="sidx"></param>
        /// <param name="outval"></param>
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
            catch (Exception)
            {
            }
            return res;
        }

        /// <summary>
        /// GPS时间插值
        /// </summary>
        /// <param name="SInfo"></param>
        /// <param name="EInfo"></param>
        /// <param name="insertime"></param>
        public GPSInfo(GPSInfo SInfo, GPSInfo EInfo, DateTime insertime)
        {
            System.TimeSpan es_date = EInfo._utctime - SInfo._utctime;
            System.TimeSpan cs_date = insertime - SInfo._utctime;
            double k = cs_date.TotalMilliseconds / es_date.TotalMilliseconds;

            _utctime = insertime;
            _latitude = k * (EInfo._latitude - SInfo._latitude) + SInfo._latitude;
            _longitude = k * (EInfo._longitude - SInfo._longitude) + SInfo._longitude;
            _elevation = k * (EInfo._elevation - SInfo._elevation) + SInfo._elevation;
            _IsOK = true;
        }

        public override string ToString()
        {
            return string.Format("{0:HHmmssfff} {1:0.0000000} {2:0.0000000} {3:0.00}",
                _utctime, _longitude, _latitude, _elevation);
        }
    }

    /// <summary>
    /// GPS和桩号关联
    /// </summary>
    public class MapGPSMile
    {
        public GPSInfo _gpsinfo;

        /// <summary>
        /// 里程
        /// </summary>
        public float _dmi;

        /// <summary>
        /// 桩号
        /// </summary>
        public int _mile;

        public MapGPSMile() { }

        public override string ToString()
        {
            return string.Format("{0} {1} {2}", _gpsinfo, _dmi, _mile);
        }
    }

    public class DataIdx
    {
        public int _Idx;
        public int _LastIdx;
        public int _NextIdx;

        public DataIdx() { }
        public DataIdx(int ci, int li, int ni)
        {
            _Idx = ci;
            _LastIdx = li;
            _NextIdx = ni;
        }
    }

    public class ExcelGPS
    {
        public double _actualMile; // 新增：基于经纬度计算的累计距离

        /// <summary>
        /// utc时间
        /// </summary>
        public string _utctime;

        /// <summary>
        /// 纬度
        /// </summary>
        public string _latitude;

        /// <summary>
        /// 经度
        /// </summary>
        public string _longitude;

        /// <summary>
        /// 高程
        /// </summary>
        public string _elevation;

        /// <summary>
        /// 桩号
        /// </summary>
        public int _mile;
        public ExcelGPS()
        { }

        public ExcelGPS(string str)
        {
            string[] strs = str.Split(' ');
            _utctime = strs[0];
            _longitude = strs[1];
            _latitude = strs[2];
            _elevation = strs[3];
            _mile = int.Parse(strs[5]);

        }
        // 构造时计算实际里程（假设已有前一点的实际里程）
        public ExcelGPS(string data, double prevActualMile, ExcelGPS prevPoint)
        {
            string[] parts = data.Split(' ');
            _utctime = parts[0];
            _longitude = parts[1];
            _latitude = parts[2];
            _elevation = parts[3];
            _mile = int.Parse(parts[5]);
            // 计算实际空间距离（新增）
            if (prevPoint != null)
            {
                double dist = CalculateDistance(
                    double.Parse(prevPoint._longitude),
                    double.Parse(prevPoint._latitude),
                    double.Parse(this._longitude),
                    double.Parse(this._latitude)
                );
                _actualMile = prevActualMile + dist;
            }
            else
            {
                _actualMile = 0;
            }
        }
        private static double CalculateDistance(double lon1, double lat1, double lon2, double lat2)
        {
            const double R = 6371000; // 地球半径(米)
            var dLat = (lat2 - lat1) * Math.PI / 180;
            var dLon = (lon2 - lon1) * Math.PI / 180;
            var a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
                    Math.Cos(lat1 * Math.PI / 180) * Math.Cos(lat2 * Math.PI / 180) *
                    Math.Sin(dLon / 2) * Math.Sin(dLon / 2);
            return R * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
        }
    }
    /// <summary>
    /// 千寻设备 校准辅助类
    public class NeedSubOneSecTimeGpsHelp
    {
        private string[] m_allStr;
        private DateTime m_Otime;
        /// <summary>
        /// 年月日字符串
        /// </summary>
        private string m_YMD;

        public bool IsNewEquip { get; private set; } = false;


        public List<DateTime> AllDataTimes { get; private set; }
        public DateTime sTime { get; private set; } = default;

        public DateTime eTime { get; private set; } = default;
        /// <summary>
        /// 需要处理授时问题
        /// </summary>

        public bool NeedHandel { get; private set; }
        private NeedSubOneSecTimeGpsHelp()
        {

        }
        public NeedSubOneSecTimeGpsHelp(bool isNewEquip, string[] str, DateTime OTime, string ymd)
        {
            m_allStr = str;
            m_Otime = OTime;

            m_YMD = ymd;
            IsNewEquip = isNewEquip;
            NeedHandel = GetAllTmepGpsStr();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns>包含需要进行-1s操作的时间</returns>
        private bool GetAllTmepGpsStr()
        {

            List<DateTime> temp0711TimeList = new List<DateTime>();
            List<DateTime> temp0212TimeList = new List<DateTime>();
            AllDataTimes = new List<DateTime>();
            if (IsNewEquip)
            {
                //获取千寻授时时间
                string lastTimeStr = m_allStr.Last().Substring(1, m_allStr.Last().Length - 1);
                string[] lastTimeSplit = lastTimeStr.Split('-', '.');
                if (lastTimeSplit[0] == "0711")
                {
                    return true;
                }

                DateTime lastTime = DateTime.ParseExact(lastTimeSplit[1], "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture);
                //0212的情况

                //如果同步版时间授时在 0211之前 那么 trigger.txt所有都需要减去1s
                if (DateTime.Compare(lastTime, m_Otime) >= 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                //如果是老程序的  
                for (int i = 0; i < m_allStr.Length; i++)
                {
                    string tempStr = m_allStr[i];
                    string[] tempSp = tempStr.Split('$');
                    if (tempStr.StartsWith("b562"))
                    {
                        //获得该行语句的验证码   为18(0212)的不管 为17(0711)的将其时间与授时时间比较 如果

                        //三个时间点   开机<-1->同步版授时<-2->开始采集  在2时间段内如果从17->18 那么所有的时间也需要-1s
                        if (tempSp[0].Length <= 32)
                        {
                            continue;
                        }
                        string code = tempSp[0].Substring(28, 4);
                        if (code.Contains("0711"))
                        {
                            //授时后为17则该条对应的 trigger 需要-1s
                            //获取当前语句时间

                            if (tempSp.Length > 1)  //表示 b562语句后面就有时间
                            {
                                string nowTimeStr = m_YMD + tempSp[1].Split(',')[1];
                                DateTime now = DateTime.ParseExact(nowTimeStr, "yyyyMMddHHmmss.ff", System.Globalization.CultureInfo.InvariantCulture);
                                if (DateTime.Compare(now, m_Otime) == 1)
                                {
                                    temp0711TimeList.Add(now);
                                    AllDataTimes.Add(now);
                                }
                            }
                            else
                            {
                                if (m_allStr.Length > i + 1)
                                {
                                    string nowTimeStr = m_YMD + m_allStr[i + 1].Split(',')[1];
                                    DateTime now = DateTime.ParseExact(nowTimeStr, "yyyyMMddHHmmss.ff", System.Globalization.CultureInfo.InvariantCulture);
                                    //now 比 m_Otime 晚
                                    if (DateTime.Compare(now, m_Otime) == 1)
                                    {
                                        temp0711TimeList.Add(now);
                                        AllDataTimes.Add(now);
                                    }
                                }
                            }
                        }
                        else if (code.Contains("0212"))
                        {
                            //授时后为17则该条对应的 trigger 需要-1s
                            //获取当前语句时间

                            if (tempSp.Length > 1)  //表示 b562语句后面就有时间
                            {
                                string nowTimeStr = m_YMD + tempSp[1].Split(',')[1];
                                DateTime now = DateTime.ParseExact(nowTimeStr, "yyyyMMddHHmmss.ff", System.Globalization.CultureInfo.InvariantCulture);
                                if (DateTime.Compare(now, m_Otime) == 1)
                                {
                                    temp0212TimeList.Add(now);
                                    AllDataTimes.Add(now);
                                }
                            }
                            else
                            {
                                if (m_allStr.Length > i + 1)
                                {
                                    string nowTimeStr = m_YMD + m_allStr[i + 1].Split(',')[1];
                                    DateTime now = DateTime.ParseExact(nowTimeStr, "yyyyMMddHHmmss.ff", System.Globalization.CultureInfo.InvariantCulture);
                                    //now 比 m_Otime 晚
                                    if (DateTime.Compare(now, m_Otime) == 1)
                                    {
                                        temp0212TimeList.Add(now);
                                        AllDataTimes.Add(now);
                                    }
                                }
                            }
                        }

                        //
                    }
                }

                //如果0711
                if (temp0711TimeList.Count > 0)
                {
                    sTime = temp0711TimeList.First();
                    eTime = temp0711TimeList.Last();
                    if (temp0212TimeList.Count > 0)
                    {
                        eTime = temp0212TimeList.Last();
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }
}
