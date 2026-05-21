//#define 隐藏公司相关信息
using DevExpress.Accessibility;
using DevExpress.ClipboardSource.SpreadsheetML;
using DevExpress.Map.Native;
using DevExpress.Services.Implementation;
using DevExpress.Utils.Extensions;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Alerter;
using DevExpress.XtraBars.Docking;
using DevExpress.XtraCharts;
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Filtering.Templates;
using DevExpress.XtraLayout.Customization;
using DevExpress.XtraPrinting.Export.Pdf;
using Farmework.Other.enumTools;
using Framework.Log; 
using Framework.Office.Work;
using Framework.Other;
using LadybugAPI;
using MathNet.Numerics.Providers.LinearAlgebra;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word; 
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using Ookii.Dialogs.WinForms;
using OpenTK.Platform.Windows;
using OperateIniFile;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.Sec;
using RuralPavementDetect;
using Spire.Pdf;
using Spire.Xls;
using SqlSugar.Extensions;
using System;
using System.CodeDom;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Lifetime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using XRDataProcess.toolForms;
using static DevExpress.XtraEditors.Mask.MaskSettings;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
using static XRDataProcess.OutputXR;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSWord = Microsoft.Office.Interop.Word;


namespace XRDataProcess
{
    public partial class MainForm : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);

        [DllImport("hnDxfIO.dll", EntryPoint = "HighAccOut")]
        private static extern bool HighAccOut(string fpath, int diseaseNum, IntPtr diseasePtr);


        [DllImport("HnHighAccConvertPlane", EntryPoint = "initialParam")]
        private static extern void initialParam(IntPtr paramInfo);

        [DllImport("HnHighAccConvertPlane", EntryPoint = "convertBLHToProjection")]
        private static extern bool convertBLHToProjection(
            double dL,
            double dB,
            double dH,
            ref double dEast, ref double dNorth, ref double dHeight);

        private static MyLogger log = new MyLogger(typeof(MainForm));




        private List<SingleProject> _Projects;
        private SingleProject _CurProject;
        private bool _isMainLayoutSafeToSave = true;

        public static string chktxt_fpath = null;

        public static Random rdval;
        public static bool _IsSaveDisImg = false;

        public static bool _IsLinkShow = false;

        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public MainForm()
        {
            this.AutoScaleMode = AutoScaleMode.None;  // 必须！防止 WinForms 二次缩放
                                                      // 一行代码，解决灰色区 + 控件拖不动
            this.DpiChanged += (s, e) => ribbonControl1.Width = this.ClientSize.Width;
            InitializeComponent();
            if (File.Exists(_layoutpath))
            {
                // dockManager_main.RestoreLayoutFromXml(_layoutpath);
            }

            _Setting.ReadData();

            _RoadConfig.ReadData();

            string path = System.Windows.Forms.Application.StartupPath + _Setting.ICO;
            if (File.Exists(path))
            {
                this.Icon = new System.Drawing.Icon(path);
            }

            path = System.Windows.Forms.Application.StartupPath + _Setting.ICODX;
            if (File.Exists(path))
            {
                ribbonControl1.ApplicationIcon = new System.Drawing.Bitmap(path);
            }

            _Projects = new List<SingleProject>();
            _CurProject = null;
            rdval = new Random(System.DateTime.Now.Millisecond);

            RoadPavementPanel.IsDiseaseRemark = _Setting.IsCrackRemark;

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


            // 初始化 AlertControl
            alertControl = new AlertControl(this.components);
            alertControl.AutoHeight = true;
            alertControl.FormLocation = AlertFormLocation.BottomRight;  // 右上角
            alertControl.ShowCloseButton = true;
            alertControl.ShowPinButton = false;
            alertControl.AutoFormDelay = 300000;// 

            alertControl.FormClosing += (s, e) =>
            {
                // 关闭时记录已显示
                Properties.Settings.Default.HasShownUpdateToast = true;
                Properties.Settings.Default.Save();
            };
            alertControl.AlertClick += AlertControl_AlertClick;

        }

        private void AlertControl_AlertClick(object sender, AlertClickEventArgs e)
        {
            string text =    e.Info?.Text;
        }

        private void bar_ParmStyle_EditValueChanged(object sender, EventArgs e)
        {
            string cstr = this.bar_ParmStyle.EditValue.ToString();
            if (cstr == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
            {
                MessageBox.Show(string.Format("当前已经是【{0}】版", cstr));
            }
            else
            {
                DialogResult dr = MessageBox.Show(string.Format("确定改为【{0}】吗？", cstr), "修改版本", MessageBoxButtons.OKCancel);
                if (DialogResult.OK == dr)
                {
                    //切换版本 初始化
                    //_RoadConfig.RealHeight = 2.0;
                    //_RoadConfig.RealWidth = 3.75;
                    //_RoadConfig.DetectWidth = 3.75;

                    _RoadConfig.WriteData();
                    for (int i = 0; i < Framework.Other.MyGlobal.Global.g_ParmStyles.Length; ++i)
                    {
                        if (cstr == Framework.Other.MyGlobal.Global.g_ParmStyles[i])
                        {
                            _Setting.ParmStyle = (StandardParmType)i;
                            break;
                        }
                    }
                    _Setting.RQIJudgeType = 1;
                    _Setting.ExcelType = 0;
                    _Setting.SelectDrawDis = 0;
                    _Setting.WriteData();

                    MessageBox.Show("修改版本成功，即将重启软件！");
                    System.Windows.Forms.Application.Exit();
                    System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);
                }
            }
        }
        private AlertControl alertControl;
        
        
        private void ShowUpdateAlertIfFirstTime()
        {
#if DEBUG
            // 开发时强制重置（上线时删除这行）
            Properties.Settings.Default.Reset();
#endif

            if (Properties.Settings.Default.HasShownUpdateToast)
                return;

          string newInfoFile =   AppDomain.CurrentDomain.BaseDirectory + "修改日志\\newInfo.txt";
            string text =    File.ReadAllText(newInfoFile);

            AlertInfo info = new AlertInfo("欢迎使用最新版本！", text);
            alertControl.Show(this, info);
        } 

        private void MainForm_Load(object sender, EventArgs e)
        {
            ShowUpdateAlertIfFirstTime();
            RestoreMainLayoutSafely();
                // 方法 2.1：将 Icon 属性设置为 null
#if 隐藏公司相关信息
            this.Icon = null;
            ribbonControl1.ApplicationIcon = null;
#else
#endif


                barEditItem3.EditValue = _Setting.OutWordPasteDelay;

            barButtonItem39.Visibility = BarItemVisibility.Never;
            LoadUISetting();


            dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            dockPanel_main_Plist.Width = 200;
            if (this.Width > 250)
            {
                dockPanel_main_data.Width = this.Width - dockPanel_main_Plist.Width;
            }
            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.SelectDrawDis == 1)
            {
                barButtonItem31.Visibility = BarItemVisibility.Always;
            }
            else
            {

                barButtonItem31.Visibility = BarItemVisibility.Never;
            }

            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel || _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
            {
                //barButtonItem_IRM.Visibility = BarItemVisibility.Never;
                bar_gd.Visibility = BarItemVisibility.Always;
            }
            else
            {
                bar_gd.Visibility = BarItemVisibility.Never;
                // barButtonItem_IRM.Visibility = BarItemVisibility.Always;
            }
            if ((_Setting.ParmStyle == StandardParmType.CityRoad && (_Setting.ExcelType == 0 || _Setting.ExcelType == 6)))
            {
                barButtonItem1.Visibility = BarItemVisibility.Always;
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007 && _Setting.ExcelType == 1)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 1)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing && _Setting.ExcelType == 0)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 0)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 4)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else
            {
                ribbonPageGroup6.Visible = false;
            }

            if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
            {
                barButtonItem18.Visibility = BarItemVisibility.Always;
            }
            else
            {
                barButtonItem18.Visibility = BarItemVisibility.Never;
            }

            if (_Setting.ExcelType == 2 || _Setting.ExcelType == 3)
            {
                ribbonPageGroup3.Visible = true;
            }
            else
            {
                ribbonPageGroup3.Visible = false;
            }
            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007 || _Setting.ParmStyle == StandardParmType.DegreeRoad2018)
            {
                if (_Setting.SelectDrawDis == 0)
                {
                    barButtonItem_snbkDIs.Visibility = BarItemVisibility.Always;
                }
            }
          
            {
                barCheckItem1.Checked = _Setting.showGpsInfoToPicture;
            }
            foreach (string str in Framework.Other.MyGlobal.Global.g_ParmStyles)
            {
                ((DevExpress.XtraEditors.Repository.RepositoryItemComboBox)this.bar_ParmStyle.Edit).Items.Add(str);
            }
            string standard = Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle];
            
            switch (_Setting.ParmStyle)
            {
                case StandardParmType.DegreeRoad2007:
                    break;
                case StandardParmType.CityRoad:
                    standard = "CJJ 36-2016 城镇道路养护技术规范";
                    break;
                case StandardParmType.RuralRoadBeijing:
                    break;
                case StandardParmType.DegreeRoad2018:
                    standard = "JTG 5210-2018 公路技术状况评定标准";
                    break;
                case StandardParmType.DegreeRoad2001:
                    break;
                case StandardParmType.CityRoadShanghai:
                    break;
                case StandardParmType.RuralRoadLiaoning:
                    break;
                case StandardParmType.RuralRoadGuangxi:
                    break;
                case StandardParmType.RuralRoadChongqing:
                    break;
                case StandardParmType.RuralRoadHunan:
                    standard = "（湖南地标）农村公路技术状况评定规范DB43T 3087-2024";
                    break;
                case StandardParmType.RuralRoadlowLevel:
                    standard = "《公路技术状况评定标准》JTG-5211-2024";
                    break;
                default:
                    break;
            }
            this.Text = string.Format("{0} 【{1}】", this.Text, standard);
            this.bar_ParmStyle.Edit.NullText = Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle];
             

        }

     
        /// <summary>
        /// 获取用户专属的布局文件完整路径（%LocalAppData%）
        /// </summary>
        private string GetUserLayoutPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);               // 确保目录存在
            return Path.Combine(appFolder, _layoutFileName);
        }
        
        
        /// <summary>
         /// 首次运行时把默认布局复制到用户目录
         /// </summary>
        private void EnsureDefaultLayoutCopied()
        {
            string userPath = GetUserLayoutPath();
            string defaultPath = GetDefaultLayoutPath();

            if (!File.Exists(userPath) && File.Exists(defaultPath))
            {
                try
                {
                    File.Copy(defaultPath, userPath);
                }
                catch { /* 复制失败也不影响主流程 */ }
            }
        }
        private const string _layoutpath = "MainLayout.xml";
        private const string _layoutpathdefault = "MainDefaultLayout.xml";


        private const string _layoutFileName = "MainLayout.xml";        // 运行时布局
        private const string _layoutDefaultFileName = "MainDefaultLayout.xml"; // 安装目录的默认布局
        /// <summary>
        /// 获取安装目录下的默认布局路径（只读）
        /// </summary>
        private string GetDefaultLayoutPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _layoutDefaultFileName);
        }

        private void RestoreMainLayoutSafely()
        {
            string userLayoutPath = GetUserLayoutPath();
            string defaultLayoutPath = GetDefaultLayoutPath();

            if (File.Exists(userLayoutPath))
            {
                if (TryRestoreLayout(userLayoutPath, "用户主界面布局"))
                {
                    _isMainLayoutSafeToSave = true;
                    return;
                }

                BackupBadLayoutFile(userLayoutPath);
                if (File.Exists(defaultLayoutPath) && TryRestoreLayout(defaultLayoutPath, "默认主界面布局"))
                {
                    _isMainLayoutSafeToSave = true;
                    return;
                }

                _isMainLayoutSafeToSave = false;
                return;
            }

            if (File.Exists(defaultLayoutPath))
            {
                TryRestoreLayout(defaultLayoutPath, "默认主界面布局");
            }

            _isMainLayoutSafeToSave = true;
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

        private void SaveMainLayoutSafely()
        {
            if (!_isMainLayoutSafeToSave)
            {
                System.Diagnostics.Debug.WriteLine("跳过主界面布局保存: 本次启动未能恢复有效布局。");
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

                string folder = Path.GetDirectoryName(layoutPath);
                string fileName = Path.GetFileNameWithoutExtension(layoutPath);
                string extension = Path.GetExtension(layoutPath);
                string backupPath = Path.Combine(folder, fileName + ".bad" + extension);
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

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveMainLayoutSafely();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("保存布局失败: " + ex.Message);
            }
        }
        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                // 1. 项目相关保存
                if (_CurProject != null)
                {
                    _CurProject.SaveCurDisease();
                    _CurProject.SingleProject_FormClosed(null, e);
                }

                // 2. 皮肤设置保存
                _Setting.SkinName = this.defaultLookAndFeel1.LookAndFeel.SkinName;
                _Setting.WriteData();

                // 【删除】所有关于 dockManager 的 Save/Restore 代码
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MainForm_FormClosed 错误: " + ex.Message);
            }



            //try
            //{
            //    // 1. 项目相关保存（保持原有逻辑）
            //    if (_CurProject != null)
            //    {
            //        _CurProject.SaveCurDisease();
            //        _CurProject.SingleProject_FormClosed(null, e);
            //    }

            //    // 2. 皮肤设置保存
            //    _Setting.SkinName = this.defaultLookAndFeel1.LookAndFeel.SkinName;
            //    _Setting.WriteData();

            //    // 3. 布局保存（关键改动）
            //    // 3.1 确保默认布局已复制（仅第一次需要）
            //    EnsureDefaultLayoutCopied();

            //    // 3.2 读取布局前先恢复（如果用户目录有文件则使用它）
            //    string userLayoutPath = GetUserLayoutPath();
            //    if (File.Exists(userLayoutPath))
            //    {
            //        try { dockManager_main.RestoreLayoutFromXml(userLayoutPath); } catch { }
            //    }

            //    // 3.3 保存当前布局到用户目录（一定有写权限）
            //    dockManager_main.SaveLayoutToXml(userLayoutPath);
            //}
            //catch (Exception ex)
            //{
            //    // 防止任何异常导致程序崩溃，可记录日志
            //    System.Diagnostics.Debug.WriteLine("MainForm_FormClosed 错误: " + ex.Message);
            //    // 如果你用了 log4net：
            //    // log.Error("MainForm_FormClosed 异常", ex);
            //}

          
        }
        SortedList<string, SingleProject> sotedList = new SortedList<string, SingleProject>();
        private void LoadUISetting()
        {
            this.defaultLookAndFeel1.LookAndFeel.SkinName = _Setting.SkinName;
        }
        /// <summary>
        /// 上海浦公 根据工程名称 进行排序
        /// 特定需求
        /// </summary>
        private void sortFormProName(Dictionary<double, SingleProject> shpgTemp, List<SingleProject> shpgTemp1, DirectoryInfo dir, SingleProject proj)
        {
            sotedList.Add(dir.Name, proj);
            //if (int.TryParse(dir.Name.Split('-').FirstOrDefault(), out num))
            //{
            //    if (int.TryParse(dir.Name.Split('_').FirstOrDefault().Split('-').LastOrDefault(), out num1))
            //    {
            //        if (dir.Name.Contains("上行") && !shpgTemp.Keys.Contains(num * 100000 + num1 - 0.5))
            //        {
            //            shpgTemp.Add(num * 100000 + num1 - 0.5, proj);

            //        }
            //        else
            //        {
            //            if (!shpgTemp.Keys.Contains(num * 100000 + num1))
            //            {
            //                shpgTemp.Add(num * 100000 + num1, proj);
            //            }

            //        }
            //    }
            //    else if (int.TryParse(dir.Name.Split('&').FirstOrDefault().Split('-').LastOrDefault(), out num1))
            //    {
            //        if (dir.Name.Contains("上行") && !shpgTemp.Keys.Contains(num * 100000 + num1 + 0.5))
            //        {
            //            shpgTemp.Add(num * 100000 + num1 + 0.5, proj);

            //        }
            //        else
            //        {
            //            if (!shpgTemp.Keys.Contains(num * 100000 + num1 + 1))
            //            {
            //                shpgTemp.Add(num * 100000 + num1 + 1, proj);
            //            }

            //        }
            //    }
            //    else
            //    {
            //        shpgTemp1.Add(proj);
            //    }
            //}
            //else
            //{
            //    shpgTemp1.Add(proj);
            //}
        }
        SortedList<string, SingleProject> tempSortedPG = new SortedList<string, SingleProject>();
        private void barButtonItem_load_ItemClick(object sender, ItemClickEventArgs e)
        {
            
            treeView_main.Enabled = false;

            VistaFolderBrowserDialog fd;
            if (Directory.Exists(_Setting.DefaultPath))
            {
                fd = new VistaFolderBrowserDialog
                {
                    Description = "选择文件夹",

                    SelectedPath = _Setting.DefaultPath,
                    ShowNewFolderButton = true
                };
            }
            else
            {
                fd = new VistaFolderBrowserDialog
                {
                    Description = "选择文件夹",

                    
                    ShowNewFolderButton = true
                };
            } 
            Console.WriteLine(_Setting.DefaultPath); 
            if (fd.ShowDialog(this) != DialogResult.OK)
            {// 恢复窗口状态

                return;
            }

            // 恢复窗口状态

            List<DirectoryInfo> projects = new List<DirectoryInfo>();
            tempSortedPG.Clear();
            if (fd.SelectedPath != string.Empty)
            {
                string selectPath = fd.SelectedPath;
                //  string selectPath = Path.GetDirectoryName(fd.FileName);
                if (selectPath.Substring(selectPath.Length - 1) == "\\")
                {
                    selectPath = selectPath.Remove(selectPath.Length - 1);
                }

                chktxt_fpath = selectPath;

                _Setting.DefaultPath = selectPath;
                _Setting.WriteData();

                dockPanel_main_data.Controls.Clear();
                treeView_main.Nodes.Clear();
                _Projects.Clear();

                if (_CurProject != null) _CurProject.Dispose();
                _CurProject = null;
                projects = GetAllProjectPath(selectPath);
                foreach (DirectoryInfo dir in projects)
                {
                    handelVillagePara(dir.FullName);


                    var daqFiles = dir.GetFiles("*.daq");
                    if (daqFiles.Length > 0)
                    {
                        _Setting.isGDIriCalculate = true;
                    }
                    else
                    {
                        _Setting.isGDIriCalculate = false;
                    }
                    _RoadConfig.ReadData();
                    try
                    {
                        SingleProject proj = new SingleProject(dir);
                        _Projects.Add(proj);
                    }
                    catch (Exception ex)
                    {

                        throw new Exception($"{dir}工程读取失败\n{ex.Message}\n{ex.StackTrace}");
                    }
                   
                }

                if (!_Setting.hasCamsetting)
                {
                  //  MessageBox.Show("在工程下未检测到任何裁剪矫正文件\\CamSetting.ini,请注意人工设置【软件设置\\其他】下拍摄宽度!","提示信息", MessageBoxButtons.RetryCancel);
                }

                if (_Setting.ParmStyle != StandardParmType.CityRoad || _Setting.ParmStyle != StandardParmType.CityRoadShanghai)
                {
                    //先按照上下行排一下顺序，下行在前上行在后，便于将上下行数据输出到同一个报表里面
                    _Projects.Sort(delegate (SingleProject x, SingleProject y) { return x._ProjectInfo._Direction.CompareTo(y._ProjectInfo._Direction); });
                }

                _Projects.Sort((a, b) => StrCmpLogicalW(a._DataDir.Name, b._DataDir.Name));
                foreach (SingleProject tproject in _Projects)
                {
                    TreeNode node = new TreeNode() { Text = tproject._DataDir.Name, Tag = tproject._DataDir };
                    treeView_main.Nodes.Add(node);
                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
                } 
                #region 获取本地人工删除病害目录 
                //先判断非C盘可否使用  如果都不行使用C盘

                {
                    bool canOutput = false;
                    string outHumanDeleteDiseasePath = "RoadData\\DiseaseNegativeSample";
                    DriveInfo[] drives = DriveInfo.GetDrives();
                    foreach (DriveInfo drive in drives)
                    {
                        if (drive.DriveType == DriveType.Fixed)
                        {

                            string driveLetter = "C:";
                            bool dDriveNotC = !drive.Name.StartsWith(driveLetter, StringComparison.OrdinalIgnoreCase);
                            if (dDriveNotC)
                            {
                                long requiredFreeSpace = 10L * 1024 * 1024 * 1024; //10GB in bytes 
                                if (drive.IsReady)
                                {
                                    long freeSpace = drive.AvailableFreeSpace;

                                    if (freeSpace > requiredFreeSpace)
                                    {
                                        outHumanDeleteDiseasePath = drive.Name + outHumanDeleteDiseasePath;
                                        canOutput = true;
                                        break;
                                    }
                                    else
                                    {
                                        canOutput = false;
                                    }
                                }
                                else
                                {
                                    canOutput = false;
                                }

                            }
                            else
                            {
                                canOutput = false;
                            }
                        }
                    }
                    if (canOutput)
                    {
                        _Setting.outHumanDeleteDiseasePath = outHumanDeleteDiseasePath;
                        _Setting.outHumanDeleteDisease = true;
                        if (!Directory.Exists(_Setting.outHumanDeleteDiseasePath))
                        {
                            Directory.CreateDirectory(_Setting.outHumanDeleteDiseasePath);

                        }
                    }
                    else
                    {
                        string driveLetter = "C:";
                        DriveInfo[] allDrives = DriveInfo.GetDrives();
                        bool dDriveExists = Array.Exists(allDrives, drive => drive.Name.StartsWith(driveLetter + ":", StringComparison.OrdinalIgnoreCase));

                        if (dDriveExists)
                        {
                            long requiredFreeSpace = 10L * 1024 * 1024 * 1024; //10GB in bytes 
                            DriveInfo drive = new DriveInfo(driveLetter);
                            if (drive.IsReady)
                            {
                                long freeSpace = drive.AvailableFreeSpace;

                                if (freeSpace > requiredFreeSpace)
                                {
                                    outHumanDeleteDiseasePath = drive.Name + outHumanDeleteDiseasePath;
                                    canOutput = true;

                                }
                                else
                                {
                                    canOutput = false;
                                }
                            }
                            else
                            {
                                canOutput = false;
                            }
                        }
                        else
                        {
                            canOutput = false;
                        }
                    }
                    if (canOutput)
                    {
                        _Setting.outHumanDeleteDiseasePath = outHumanDeleteDiseasePath;
                        _Setting.outHumanDeleteDisease = true;
                        if (!Directory.Exists(_Setting.outHumanDeleteDiseasePath))
                        {
                            Directory.CreateDirectory(_Setting.outHumanDeleteDiseasePath);

                        }
                    }
                    else
                    {
                        _Setting.outHumanDeleteDisease = false;
                        _Setting.outHumanDeleteDiseasePath = "";
                    }
                    _Setting.WriteData();
                }
                #endregion 
            }
            treeView_main.Enabled = true; 

        }
        private static void GetFileListByPath(ref List<SingleProject> list)
        {
            list.Sort((a, b) => StrCmpLogicalW(b._DataDir.FullName, a._DataDir.FullName));
        }

        private void handelVillagePara(string path)
        {
            if (!Directory.Exists(path))
            {
                return;
            }
            DirectoryInfo dir = new DirectoryInfo(path);
            #region 低等级农村路参数初始化
             
            if (File.Exists(dir.FullName + "\\CamSetting.ini") )
            {
                readCamSetting(dir.FullName); 
            }
            else
            {
                _Setting.hasCamsetting = false;
            }
            _Setting.WriteData();
            _RoadConfig.WriteData();
            #endregion
        }

        private void GetVillagePara(string path, out double jzWidth)
        {
            jzWidth = 0;
            DirectoryInfo dir = new DirectoryInfo(path);

           
            if (File.Exists(dir.FullName + "\\CamSetting.ini"))
            {
                IniFiles inifile = new IniFiles(Path.Combine(path, "CamSetting.ini"));
                //0 矫正  1裁剪
                int isJZorCJInt = inifile.ReadInteger("RoadConfig", "jzorcj", 0);
                if (isJZorCJInt == -1)
                {
                    jzWidth = 0;
                    return;
                }
                //当用户只裁剪 即 bool值为false的时候需要   进行矩阵转换 
                double roadRealWidth = inifile.ReadDouble("RoadConfig", "width", 0);
                double roadRealHight = inifile.ReadDouble("RoadConfig", "height", 0);
                double roadWidScale = inifile.ReadDouble("RoadConfig", "widthScale", 0);
                double roadHigScale = inifile.ReadDouble("RoadConfig", "hightScale", 0);
                double oriWidth = inifile.ReadDouble("RoadConfig", "OriImgWidth", 0);
                double cutWidth = inifile.ReadDouble("RoadConfig", "qxCutImgWidth", 0);
                jzWidth = inifile.ReadDouble("RoadConfig", "jz_with", 0);
                if (roadRealHight == 0 || roadRealWidth == 0)
                {
                    _Setting.hasCamsetting = false;
                    jzWidth = 0;
                    return;
                }
            }
        }

        /// <summary>
        /// 查找工程下是否存在camsetting配置文件
        /// </summary>
        /// <returns></returns>
        private bool readCamSetting(string path)
        {
            IniFiles inifile = new IniFiles(Path.Combine(path, "CamSetting.ini"));
            //0 矫正  1原图裁剪
            int isJZorCJInt = inifile.ReadInteger("RoadConfig", "jzorcj", 0);
            if (isJZorCJInt == -1)
            {
                return true;
            }
            bool isJZorCJ = isJZorCJInt == 0 ? true : false;
            //当用户只裁剪 即 bool值为false的时候需要   进行矩阵转换
            VillageHandleCoord s = VillageHandleCoord.getInstance(isJZorCJ, path);

            double roadRealWidth = inifile.ReadDouble("RoadConfig", "width", 0);
            double roadRealHight = inifile.ReadDouble("RoadConfig", "height", 0);
            double roadWidScale = inifile.ReadDouble("RoadConfig", "widthScale", 0);
            double roadHigScale = inifile.ReadDouble("RoadConfig", "hightScale", 0);
            double oriWidth = inifile.ReadDouble("RoadConfig", "OriImgWidth", 0);
            double cutWidth = inifile.ReadDouble("RoadConfig", "qxCutImgWidth", 0);
            double cutWidthReal = inifile.ReadDouble("RoadConfig", "jz_with", 0);
            if (roadRealHight == 0 || roadRealWidth == 0)
            {
                _Setting.hasCamsetting = false;

                return false;
            }
            else
            {
                //道路宽度
                int width = 0;
                int height = 0;

                string imgDir = string.Format("{0}\\RoadImg\\Camera0\\Image_0000", path);
                var timgname = Directory.GetFiles(imgDir, "*.jpg").FirstOrDefault();

                if (File.Exists(timgname))
                {
                    FileInfo file = new FileInfo(timgname);
                    if (file.Length == 0)
                    {
                        throw new Exception("图片" + timgname + "大小为0Kb,请检查原始数据问题!");
                    }

                    using (FileStream fs = new FileStream(timgname, FileMode.Open, FileAccess.Read))
                    {
                        System.Drawing.Image _image = System.Drawing.Image.FromStream(fs);
                        width = _image.Width;
                        height = _image.Height;
                        _RoadConfig.ImageWidth = width;
                        _image.Dispose();
                        _image = null;
                    }

                }
                //注意这里  小数位数
                if (cutWidthReal != 0)
                {
                    roadRealWidth = cutWidthReal;
                }
                else
                {
                    roadRealWidth = Math.Round((width * roadRealWidth / cutWidth), 2);
                }

                if (_Setting.SelectDrawDis == 1)
                {
                    //对真实宽度和真实高度和计算宽度进行赋值
                    //图像分辨率在需要展示图片的时候再一次进行计算
                    _RoadConfig.RealWidth = roadRealWidth;
                    _RoadConfig.RealHeight = roadRealHight;
                    _RoadConfig.DetectWidth = roadRealWidth;
                    _RoadConfig.WidthScale = roadWidScale;
                    _RoadConfig.HeightScale = roadHigScale;
                    //非农村路非模块化
                    _Setting.hasCamsetting = true;
                }
                else
                {
                    //如果原图宽度 和 camsetting里面 OriImgWidth 一致 说明 图像没有经过采集
                    // 例如:高等级公路的图片 用低等级农村路处理 
                    if (width == oriWidth)
                    {
                        _Setting.hasCamsetting = false;
                    }
                    else
                    {
                        //对真实宽度和真实高度和计算宽度进行赋值
                        //图像分辨率在需要展示图片的时候再一次进行计算
                        _RoadConfig.RealWidth = roadRealWidth;
                        _RoadConfig.RealHeight = roadRealHight;
                        _RoadConfig.DetectWidth = roadRealWidth;
                        //非农村路非模块化
                        _Setting.hasCamsetting = true;
                    }

                }
            }

            return true;

        }

        /// <summary>
        /// 获取所有文件
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private List<DirectoryInfo> GetAllProjectPath(string path) 
        {
            List<DirectoryInfo> projects = new List<DirectoryInfo>();
            if (Directory.Exists(path))
            {
                DirectoryInfo dir = null;
                DirectoryInfo[] sdirs = null;
                try
                {
                    dir = new DirectoryInfo(path);
                    sdirs = dir.GetDirectories();
                }
                catch (System.Exception)
                {

                    dir = null;
                    sdirs = null;
                }
                if (sdirs != null)
                {
                    foreach (DirectoryInfo d in sdirs)
                    {
                        if (File.Exists(d.FullName + "\\ProjectInfo.txt"))
                        {
                            projects.Add(d);

                        }

                        else
                        {
                            projects.AddRange(GetAllProjectPath(d.FullName).ToArray());
                        }
                    }
                }
            }
            if (File.Exists(path + "\\ProjectInfo.txt"))
            {
                projects.Add(new DirectoryInfo(path));
            }
            return projects;
        }
        private TreeNode lastSelectedNode = null;
        private void treeView_main_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView_main.SelectedNode != null
            && treeView_main.SelectedNode.Index < _Projects.Count)
            {
                try
                {
                    // 重置上一个选中节点的颜色
                    if (lastSelectedNode != null)
                    {
                        lastSelectedNode.BackColor = treeView_main.BackColor;
                        lastSelectedNode.ForeColor = treeView_main.ForeColor;
                    }

                    // 设置当前选中节点的颜色
                    if (e.Node != null)
                    {
                        e.Node.BackColor = SystemColors.Highlight;
                        e.Node.ForeColor = SystemColors.HighlightText;
                        lastSelectedNode = e.Node;
                    }

                    dockPanel_main_data.Controls.Clear();
                    if (_CurProject!=null)
                    {
                        _CurProject.SaveCurDisease();

                    }
                    _CurProject = _Projects[treeView_main.SelectedNode.Index];
                    _CurProject.TopLevel = false;
                    _CurProject.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                    dockPanel_main_data.Controls.Add(_CurProject);
                    _CurProject.Dock = DockStyle.Fill;
                    _CurProject.Height = this.dockPanel_main_data.Height;
                    var dock = dockPanel_main_data.Dock;
                    var vis = dockPanel_main_data.Visibility;
                    _CurProject.Show();

                }
                catch (System.Exception)
                {

                }

            }
            try
            {
                //if (File.Exists(_layoutpathdefault))
                //{
                //    dockManager_main.RestoreLayoutFromXml(_layoutpathdefault);
                //}
                _CurProject.RestoreSavedLayout();
            }
            catch (Exception)
            {
                // 处理异常
            }
        }

        private void dockPanel_main_Plist_Collapsed(object sender, DevExpress.XtraBars.Docking.DockPanelEventArgs e)
        {
            dockPanel_main_data.Width = this.Width;
        }

        private void ribbonControl1_ApplicationButtonClick(object sender, EventArgs e)
        {
#if 隐藏公司相关信息

#else
 string info = _Setting.CompanyInfo;
            info = info.Replace("\\n", Environment.NewLine);
            MessageBox.Show(info, "关于");
#endif

        }

        /// <summary>
        /// 计算IRI、Rut和MTD的数值
        /// 注意需要添加进度条，和进度提示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem_IRM_ItemClick(object sender, ItemClickEventArgs e)
        {

            if (_Projects.Count < 1)
            {
                MessageBox.Show("没有待处理的工程！");
            }
            else
            {
                SelectIRM sltIRM = new SelectIRM(0);
                sltIRM.ShowDialog();
                if (sltIRM.IsYes())
                {
                    this.Cursor = Cursors.WaitCursor;
                    if (JudgeFreeSpace())
                    {
                        WinProcessBar winbar = new WinProcessBar(_Projects);
                        StartIRMThread(winbar);
                        winbar.ShowDialog();
                    }
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private Thread ThreadIRM;
        private void StartIRMThread(WinProcessBar winbar)
        {
            ThreadIRM = new Thread(IRMThreadMethod) { IsBackground = true };
            ThreadIRM.Start(winbar);
        }
        private void IRMThreadMethod(object prj)
        {
            WinProcessBar winbar = (WinProcessBar)prj;
            ComputeIRM(winbar);
        }
        private void ComputeIRM(WinProcessBar winbar)
        {
            winbar.SetMainMax(_Projects.Count);
            foreach (SingleProject proj in _Projects)
            {
                winbar.TextInfoAdd("正在处理：" + proj._DataDir.Name);
                proj.ComputeIRM(winbar);
                winbar.TextInfoAdd("处理完成：" + proj._DataDir.Name);
                winbar.AddMainVal(1);
            }
            MessageBox.Show("生成IRM完成!");
        }

        //判断磁盘剩余空间，返回true-磁盘剩余空间足够，false-磁盘剩余空间不够，剩余空间不够弹窗提示
        private bool JudgeFreeSpace()
        {
            if (_Projects.Count < 1)
                return true;

            long freespace = GetHardDiskSpace(_Projects[0]._DataDir.Root.Name);
            long needspace = 0;
            foreach (SingleProject proj in _Projects)
            {
                GetDirectoryLength(proj._DataDir.FullName + "\\camera0", ref needspace);
                GetDirectoryLength(proj._DataDir.FullName + "\\camera1", ref needspace);
            }
            needspace = needspace / (1024 * 1024);
            if (needspace + 100 < freespace)
            {
                return true;
            }
            else
            {
                MessageBox.Show(string.Format("警告：【{0}】磁盘剩余空间【{1}】不足！所需空间大小【{2}】", _Projects[0]._DataDir.Root.Name, freespace, needspace));
                return false;
            }
        }
        public static void GetDirectoryLength(string dirPath, ref long len, string pattern = "*.*")
        {
            if (!Directory.Exists(dirPath))
            {
                return;
            }
            DirectoryInfo di = new DirectoryInfo(dirPath);
            foreach (FileInfo fi in di.GetFiles(pattern))
            {
                len += fi.Length;
            }
            DirectoryInfo[] dis = di.GetDirectories();
            if (dis.Length > 0)
            {
                for (int i = 0; i < dis.Length; i++)
                {
                    GetDirectoryLength(dis[i].FullName, ref len);
                }
            }
        }

        public static long GetHardDiskSpace(string str_HardDiskName)
        {
            long totalSize = 0;
            //str_HardDiskName = str_HardDiskName + ":\\";
            System.IO.DriveInfo[] drives = System.IO.DriveInfo.GetDrives();
            foreach (System.IO.DriveInfo drive in drives)
            {
                if (drive.Name == str_HardDiskName)
                {
                    totalSize = drive.TotalFreeSpace / (1024 * 1024);
                }
            }
            return totalSize;
        }


     
        /// <summary>
        /// 批量生成报表
        /// 注意需要添加进度条，和进度提示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem_xls_ItemClick(object sender, ItemClickEventArgs e)
        {
          // AutoTest();
            导出路面报表区间 exceldis = new 导出路面报表区间(_Projects.Count);
            exceldis.ShowDialog();
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 6)
            {
                设置单元编号 ZYSetUnitNum = new 设置单元编号();
                ZYSetUnitNum.ShowDialog();
                if (!设置单元编号.fg) return;
            }
            //需要进行分段出表

            if (exceldis.NeedSub)
            {
                _Setting.needSub = true;
                string subStr = string.Join(",", exceldis.SubData.ToArray());
                _Setting.subData = subStr;
                //写到配置文件
            }
            else
            {
                _Setting.needSub = false;
                string subStr = "";
                _Setting.subData = subStr;
            }


            int[][] _ExcelDisVal;//存放10，100，1000 等自定义的区间
            if (exceldis._IsExcel)
            {
                _ExcelDisVal = new int[_Setting.LenExcelNum][];
                for (int i = 0; i < _Setting.LenExcelNum; i++)
                {
                    string[] strs = _Setting.LenExcel[i].Split(',');
                    _ExcelDisVal[i] = new int[strs.Length];
                    for (int j = 0; j < strs.Length; ++j)
                        _ExcelDisVal[i][j] = int.Parse(strs[j]);
                }
                FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    switch (_Setting.ParmStyle)
                    {
                        case StandardParmType.DegreeRoad2007: MyExcelDegree2007.LoadXlsParm(); break;
                        case StandardParmType.CityRoad: MyExcelCity.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadBeijing: MyExcelBJDegree.LoadXlsParm(); break;
                        case StandardParmType.DegreeRoad2018:
                            if (_Setting.SelectDrawDis == 0)
                            {
                                MyExcelDegree2018.LoadXlsParm();
                            }
                            else
                            {
                                MyExcelDegreeSmall2018.LoadXlsParm();
                            }
                            break;
                        case StandardParmType.DegreeRoad2001: MyExcelDegree2001.LoadXlsParm(); break;
                        case StandardParmType.CityRoadShanghai: MyExcelCitySH2013.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadLiaoning: MyExcelLNDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadGuangxi: MyExcelGXDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadChongqing: MyExcelCQDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadlowLevel:
                            if (_Setting.SelectDrawDis == 0)
                            {
                                MyExcelVillageDegree.LoadXlsParm();
                            }
                            else
                            {
                                MyExcelVillageDegreeSmall.LoadXlsParm();
                            }
                            break;
                        case StandardParmType.RuralRoadHunan:
                            if (_Setting.SelectDrawDis == 1)
                            {
                                MyExcelHNDegreeSmall.LoadXlsParm_new();
                            }
                            else
                            {
                                MyExcelHNDegree.LoadXlsParm_new();
                            }
                            break;
                        default: break;
                    }

                    MSExcel.Application excelApp = new MSExcel.Application()
                    {
                        Visible = true,
                        DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                        AlertBeforeOverwriting = false
                    };

                    string outdirpath = null;

                    if (_Setting.needSub)
                    {
                        MessageBox.Show($"当前分段区间{_Setting.subData}");
                        string[] subNames = _Setting.subData.Split(',');
                        for (int i = 0; i < subNames.Length; i += 2)
                        {
                            foreach (SingleProject proj in _Projects)
                            {
                                //需要进行分段输出

                                outdirpath = fd.SelectedPath;

                                if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                                {

                                    _Setting.nowSubIndexStr = subNames[i] + "," + subNames[i + 1];
                                    string smile = int.Parse(subNames[i]).ToString("K0000+000");
                                    string emile = int.Parse(subNames[i + 1]).ToString("K0000+000");


                                    string tt = "\\" + proj._DataDir.Name;
                                    tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                                    outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);


                                    if (!Directory.Exists(outdirpath))
                                    {
                                        Directory.CreateDirectory(outdirpath);
                                    }
                                }
                                else
                                {
                                    _Setting.nowSubIndexStr = subNames[i] + "," + subNames[i + 1];

                                }

                                proj.GenerateExcel(excelApp, outdirpath, _ExcelDisVal, _Setting.IsExcel);
                            }
                        }
                    }
                    else
                    {
                        foreach (SingleProject proj in _Projects)
                        {


                            outdirpath = fd.SelectedPath;
                            if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                            {
                                string smile = proj._ProjectInfo._StartMile.ToString("K0000+000");
                                string emile = proj._ProjectInfo._EndMile.ToString("K0000+000");

                                string tt = "\\" + proj._DataDir.Name;
                                var index = tt.LastIndexOf('_') - 9;
                                if (index <= 0)
                                {
                                    log.Error($"{proj._DataDir.Name}工程数据文件夹名称错误");
                                }

                                tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                                outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);


                                if (!Directory.Exists(outdirpath))
                                {
                                    Directory.CreateDirectory(outdirpath);
                                }
                            }

                            proj.GenerateExcel(excelApp, outdirpath, _ExcelDisVal, _Setting.IsExcel);
                        }
                    }
                    try
                    {
                        excelApp.Quit();

                    }
                    catch
                    {

                    }
                    if (_Setting.needSub)
                        MessageBox.Show("导出分段报表完成！\n注意!为保证其他功能正常,使用分段出表后请重新导入工程!");
                    else
                        MessageBox.Show("导出报表完成！");
                }
            }
            _Setting.WriteData();
        }
      

        public void AutoTest()
        {
            int recordDrawDis = _Setting.SelectDrawDis;
            var recordParm = _Setting.ParmStyle;
            var recordWidth = _RoadConfig.DetectWidth;
            _Setting.needSub = false;
            string ncDataUp = "D:\\统一测试数据\\二维软件\\C078422823_庙麸线_上行_官渡口镇_湖北省_恩施土家族苗族自治州_巴东县_20230424_161550";
            string ncDataDown = "D:\\统一测试数据\\二维软件\\C135430623_水赤线_下行_1_湖南省_岳阳市_华容县_20211014_105108";
            string djData = "D:\\统一测试数据\\二维软件\\XC27211221_红德线_上行_2_辽宁省_铁岭市_铁岭经济技术开发区_20251027_153626";
            string czData = "D:\\统一测试数据\\二维软件\\1513_建设一路_上行_上行_湖北省_武汉市_青山区_20251116_091135";
            
            {
                SingleProject proj = new SingleProject(ncDataUp);
                handelVillagePara(ncDataUp);
                proj.VerifyCalculationResults(StandardParmType.RuralRoadlowLevel, 0, ncDataUp);

            }

            {
                SingleProject proj = new SingleProject(ncDataUp);
                handelVillagePara(ncDataUp);
                proj.VerifyCalculationResults(StandardParmType.RuralRoadlowLevel, 1, ncDataUp);
            }
            {
                SingleProject proj = new SingleProject(ncDataDown);
                handelVillagePara(ncDataDown);
                proj.VerifyCalculationResults(StandardParmType.RuralRoadlowLevel, 0, ncDataDown);

            }

            {
                SingleProject proj = new SingleProject(ncDataDown);
                handelVillagePara(ncDataDown);
                proj.VerifyCalculationResults(StandardParmType.RuralRoadlowLevel, 1, ncDataDown);
            }

            {
                SingleProject proj = new SingleProject(djData);
                proj.VerifyCalculationResults(StandardParmType.DegreeRoad2018, 0, djData);
            }

            {
                SingleProject proj = new SingleProject(djData);
                proj.VerifyCalculationResults(StandardParmType.DegreeRoad2018, 1, djData);
            }

            {
                SingleProject proj = new SingleProject(czData);
                proj.VerifyCalculationResults(StandardParmType.CityRoad, 0, czData);

            }
            _Setting.SelectDrawDis = recordDrawDis;
            _Setting.ParmStyle = recordParm;
            _RoadConfig.DetectWidth = recordWidth;

            _Setting.WriteData();
            _RoadConfig.WriteData();
            Environment.Exit(0); // 明确退出，防止卡死

        }


        private void MainForm_SizeChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    dockManager_main.BeginUpdate(); // 暂停布局更新
            //    if (File.Exists(_layoutpathdefault))
            //    {
            //        dockManager_main.RestoreLayoutFromXml(_layoutpathdefault);
            //    }
            //    _CurProject.RestoreDefaultLayout();
            //}
            //catch (Exception)
            //{
            //    // 处理异常
            //}
            //finally
            //{
            //    dockManager_main.EndUpdate(); // 恢复布局更新
            //}
        }
        private int _lastWidth = 0;
        private int _lastHeight = 0;
        protected  override void OnResizeEnd(EventArgs e)
        {


            int currentWidth = this.Width;
            int currentHeight = this.Height;

            // 检查尺寸变化是否超过阈值
            if (Math.Abs(currentWidth - _lastWidth) > 50 || Math.Abs(currentHeight - _lastHeight) > 50)
            {
                _lastWidth = currentWidth;
                _lastHeight = currentHeight;

                try
                {
                    if (_CurProject != null)
                    {
                        _CurProject.Resize(this.dockPanel_main_data.Height);
                    }
                    ribbonControl1.Width = this.ClientSize.Width;
                    dockManager_main.ForceInitialize();
                }
                catch (Exception)
                {
                    // 处理异常
                }
            }
        }

        private void barButtonItem_help_ItemClick(object sender, ItemClickEventArgs e)
        {
            string helpfile = "";
            DialogResult result = MessageBox.Show("有html格式与.docx格式文档供查询，请问是否查询.docx格式文档？", "选择窗口", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.No)
            {
                helpfile = System.Windows.Forms.Application.StartupPath + @"\数据处理软件操作手册 V1.0.html";
                Help.ShowHelp(this, helpfile);
            }
            if (result == DialogResult.Yes)
            {
                helpfile = System.Windows.Forms.Application.StartupPath + "\\软件说明\\12 产品内业软件使用手册V1.1-2024.docx";
                System.Diagnostics.Process.Start(helpfile);
            }
            else
            {
                return;
            }
        }

        #region 导出报告相关
        private bool _IsSetDocRoad = false;//是否导出报告相关参数
        public List<string> _ExcelPathList = new List<string>();

        /// <summary>
        /// 技术明细
        /// </summary>
        public List<string> _ExcelJSMXList = new List<string>();

        /// <summary>
        /// 病害统计
        /// </summary>
        public List<string> _ExcelBHTJList = new List<string>();

        /// <summary>
        /// IRI报表
        /// </summary>
        public List<string> _ExcelIRIList = new List<string>();

        /// <summary>
        /// DR报表
        /// </summary>
        public List<string> _ExcelDRList = new List<string>();
        /// <summary>
        /// 空间定位表格
        /// </summary>
        public List<string> _ExcelGpss = new List<string>();
        /// <summary>
        /// 沿线设施及路基损害汇总表
        /// </summary>
        public List<string> _LJSHAndLJSumList = new List<string>();
        /// <summary>
        /// 沥青路面病害统计表
        /// </summary>
        public List<string> _DiseaseLQ = new List<string>();
        /// <summary>
        /// 水泥病害统计表
        /// </summary>
        public List<string> _DiseaseSN = new List<string>();
        /// <summary>
        /// 病害流水表表格
        /// </summary>
        public List<string> _DiseaseAll = new List<string>();
        private string _ExcelListFilePath;
        private string _RoadPName;//路段名
        private string _RoadLName;//路线名
        private string[] _xlstype = { "IRI", "MTD", "PCI", "病害统计", "PQI", "Rut" };
        private string[] _sidetypeL = { "左一幅", "左二幅", "左三幅", "左四幅", "左五幅", "左六幅", "左七幅", "左八幅" };
        private string[] _sidetypeR = { "右一幅", "右二幅", "右三幅", "右四幅", "右五幅", "右六幅", "右七幅", "右八幅" };
        private bool[] _sidetypeLf = { false, false, false, false, false, false, false, false };
        private bool[] _sidetypeRf = { false, false, false, false, false, false, false, false };
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            路段配置 myset = new 路段配置();
            myset.ShowDialog();
            _IsSetDocRoad = myset._IsSet;
        }

        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            _ExcelPathList.Clear();
            _ExcelBHTJList.Clear();
            _ExcelJSMXList.Clear();
            _ExcelIRIList.Clear();
            _ExcelDRList.Clear();
            _DiseaseAll.Clear();
            _DiseaseSN.Clear();
            _DiseaseLQ.Clear();
            _LJSHAndLJSumList.Clear();
            _ExcelGpss.Clear();
            //城镇单独出表
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 0)
            {
                if (!_IsSetDocRoad)
                {
                    MessageBox.Show("请先进行路段配置并检查！");
                }
                else
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    _RoadPName = inisetting.ReadString("Road", "RoadPName", "123").Replace("\0", "");
                    _RoadLName = inisetting.ReadString("Road", "RoadLName", "123").Replace("\0", "");
                    for (int i = 0; i < 8; ++i)
                    {
                        _sidetypeLf[i] = inisetting.ReadBool("Road", "RoadLine0" + i.ToString(), false);
                        _sidetypeRf[i] = inisetting.ReadBool("Road", "RoadLine1" + i.ToString(), false);
                    }

                    FolderBrowserDialog fd = new FolderBrowserDialog();
                    fd.Description = "请选择路段报表文件夹：";
                    fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                    fd.ShowDialog();
                    if (fd.SelectedPath != string.Empty)
                    {
                        if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                        {
                            fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                        }

                        inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                        treeView_main.Nodes.Clear();
                        _Projects.Clear();
                        _CurProject = null;
                        dockPanel_main_data.Controls.Clear();
                        _ExcelPathList.Clear();

                        //文件夹里面的文件夹里面的报表文件
                        DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                        DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                        foreach (DirectoryInfo tdir in dirInfoArray)
                        {
                            FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                            foreach (FileInfo tfile in tfileInfoArray)
                            {
                                if (tfile.FullName.Contains(_RoadPName) && tfile.FullName.Contains(_RoadLName))
                                    _ExcelPathList.Add(tfile.FullName);
                            }
                        }

                        foreach (string fname in _ExcelPathList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }

                        dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                        dockPanel_main_data.Width = this.Width;
                    }
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 6) //模板1
            {
                if (!_IsSetDocRoad)
                {
                    MessageBox.Show("请先进行路段配置并检查！");
                }
                else
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    _RoadPName = inisetting.ReadString("Road", "RoadPName", "123").Replace("\0", "");
                    _RoadLName = inisetting.ReadString("Road", "RoadLName", "123").Replace("\0", "");
                    for (int i = 0; i < 8; ++i)
                    {
                        _sidetypeLf[i] = inisetting.ReadBool("Road", "RoadLine0" + i.ToString(), false);
                        _sidetypeRf[i] = inisetting.ReadBool("Road", "RoadLine1" + i.ToString(), false);
                    }

                    FolderBrowserDialog fd = new FolderBrowserDialog();
                    fd.Description = "请选择路段报表文件夹：";
                    fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                    fd.ShowDialog();
                    if (fd.SelectedPath != string.Empty)
                    {
                        if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                        {
                            fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                        }

                        inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                        treeView_main.Nodes.Clear();
                        _Projects.Clear();
                        _CurProject = null;
                        dockPanel_main_data.Controls.Clear();
                        _ExcelPathList.Clear();

                        DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                        //文件夹里面的文件夹里面的报表文件
                        DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                        foreach (DirectoryInfo tdir in dirInfoArray)
                        {
                            FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                            foreach (FileInfo tfile in tfileInfoArray)
                            {
                                if (tfile.FullName.Contains(_RoadPName) && tfile.FullName.Contains(_RoadLName))
                                {
                                    _ExcelPathList.Add(tfile.FullName);
                                }
                            }
                        }
                        foreach (string fname in _ExcelPathList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }

                        dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                        dockPanel_main_data.Width = this.Width;
                    }
                }
            }
            // 07 综合
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007 && _Setting.ExcelType == 1)
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains("1000m.xlsx") && !tfile.FullName.Contains("IRIMTD"))
                            _ExcelPathList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains("1000m.xlsx") && !tfile.FullName.Contains("IRIMTD"))
                                _ExcelPathList.Add(tfile.FullName);
                        }
                    }
                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            //18标准 综合出表的报告  低等级农村路
            else if ((_Setting.ParmStyle == StandardParmType.DegreeRoad2018
                || _Setting.ParmStyle == StandardParmType.RuralRoadGuangxi ||
                _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel) && _Setting.ExcelType == 1)
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains("1000m.xlsx") && !tfile.FullName.Contains("IRIMTD"))
                            _ExcelPathList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains("1000m.xlsx") && !tfile.FullName.Contains("IRIMTD"))
                                _ExcelPathList.Add(tfile.FullName);
                        }
                    }

                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            //18标准 综合出表的报告 

            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 20)
            {
                //上海惠普
                DialogResult result = MessageBox.Show("请选择包含Excel文件名包含_病害统计_1000m.xlsx文件", "统计报表提示");
                if (result != DialogResult.OK)
                {
                    return;
                }
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains("_病害统计_1000m.xlsx"))
                            _ExcelPathList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains("_病害统计_1000m.xlsx"))
                                _ExcelPathList.Add(tfile.FullName);
                        }
                    }

                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.FilterIndex = 1;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _ExcelListFilePath = openFileDialog.FileName;
                }
            }

            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018
                && _Setting.ExcelType == 0)
            {
                统计单元长度 unitlendlg = new 统计单元长度();
                unitlendlg.ShowDialog();
                if (!unitlendlg._IsOK)
                {
                    return;
                }

                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelBHTJList.Clear();
                    _ExcelJSMXList.Clear();
                    _ExcelIRIList.Clear();
                    _ExcelDRList.Clear();

                    int unitlen = 统计单元长度._DisUnitLen;

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                            _ExcelBHTJList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                            _ExcelJSMXList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                            _ExcelIRIList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                            _ExcelDRList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                _ExcelBHTJList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                _ExcelJSMXList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                _ExcelIRIList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                _ExcelDRList.Add(tfile.FullName);
                        }
                    }

                    //报表列表
                    foreach (string fname in _ExcelBHTJList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelJSMXList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelIRIList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelDRList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            #region 农村路统计报表合并
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 3)
            {
                MulitExcelFrom selectForm = new MulitExcelFrom();

                selectForm.StartPosition = FormStartPosition.CenterParent;
                DialogResult result0 = selectForm.ShowDialog(this);
                if (result0 != DialogResult.OK)
                {
                    return;
                }

                统计单元长度 unitlendlg = new 统计单元长度();
                unitlendlg.label1.Visible = false;
                unitlendlg.comboBox1.Visible = false;
                unitlendlg.ShowDialog();
                if (!unitlendlg._IsOK)
                {
                    return;
                }

                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath == string.Empty)
                {
                    return;
                }
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }
                inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                treeView_main.Nodes.Clear();
                _Projects.Clear();
                _CurProject = null;
                dockPanel_main_data.Controls.Clear();
                int unitlen = 统计单元长度._IndexUnitLen;
                // 文件夹里面的报表文件
                DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");

                switch (_Setting.multiExcelMergeType)
                {
                    case 0:
                        #region 多车道统计
                        {   //统计报表 弹出选择界面

                            foreach (FileInfo tfile in fileInfoArray)
                            {
                                if (tfile.FullName.Contains(string.Format("_PQI_{0}m", unitlen)))
                                    _ExcelPathList.Add(tfile.FullName);
                            }
                            //文件夹里面的文件夹里面的报表文件
                            DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                            foreach (DirectoryInfo tdir in dirInfoArray)
                            {
                                FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                                foreach (FileInfo tfile in tfileInfoArray)
                                {
                                    if (tfile.FullName.Contains(string.Format("_PQI_{0}m", unitlen)))
                                        _ExcelPathList.Add(tfile.FullName);
                                }
                            }
                        }
                        #endregion
                        break;
                    case 1:
                        #region 咸宁
                        {
                            //文件夹里面的文件夹里面的报表文件
                            foreach (FileInfo tfile in fileInfoArray)
                            {

                                if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                    _ExcelJSMXList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                    _ExcelIRIList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                    _ExcelDRList.Add(tfile.FullName);
                            }
                            DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                            foreach (DirectoryInfo tdir in dirInfoArray)
                            {

                                FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                                if (tfileInfoArray.Length < 22)
                                {

                                }
                                foreach (FileInfo tfile in tfileInfoArray)
                                {

                                    if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                        _ExcelJSMXList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                        _ExcelIRIList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                        _ExcelDRList.Add(tfile.FullName);
                                }
                            }
                        }
                        #endregion
                        break;
                    case 2:
                        #region 孝感统计
                        {
                            foreach (FileInfo tfile in fileInfoArray)
                            {

                                if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                    _ExcelJSMXList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                    _ExcelIRIList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                    _ExcelDRList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_路基破损及沿线设施病害_10米")))
                                    _LJSHAndLJSumList.Add(tfile.FullName);

                            }

                            //文件夹里面的文件夹里面的报表文件
                            DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                            foreach (DirectoryInfo tdir in dirInfoArray)
                            {
                                FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                                if (tfileInfoArray.Length < 22)
                                {
                                    MessageBox.Show(tdir.FullName + "\n工程下表格数量低于22请检查,是否有缺损");
                                }
                                foreach (FileInfo tfile in tfileInfoArray)
                                {

                                    if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                        _ExcelJSMXList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                        _ExcelIRIList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                        _ExcelDRList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_路基破损及沿线设施病害_10米")))
                                        _LJSHAndLJSumList.Add(tfile.FullName);
                                }
                            }
                        }
                        #endregion
                        break;
                    case 3:
                        #region 南昌农村路
                        {
                            //文件夹里面的文件夹里面的报表文件
                            foreach (FileInfo tfile in fileInfoArray)
                            {

                                if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                    _ExcelJSMXList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                    _ExcelBHTJList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                    _ExcelDRList.Add(tfile.FullName);
                            }
                            DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                            foreach (DirectoryInfo tdir in dirInfoArray)
                            {

                                FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                                if (tfileInfoArray.Length < 22)
                                {

                                }
                                foreach (FileInfo tfile in tfileInfoArray)
                                {

                                    if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                        _ExcelJSMXList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                        _ExcelBHTJList.Add(tfile.FullName);
                                    else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                        _ExcelDRList.Add(tfile.FullName);
                                }
                            }
                        }
                        #endregion
                        break;
                    default:
                        break;
                }

                //报表列表

                foreach (string fname in _ExcelPathList)
                {
                    TreeNode node = new TreeNode() { Text = fname };
                    treeView_main.Nodes.Add(node);
                }
                foreach (string fname in _ExcelJSMXList)
                {
                    TreeNode node = new TreeNode() { Text = fname };
                    treeView_main.Nodes.Add(node);
                }
                foreach (string fname in _ExcelIRIList)
                {
                    TreeNode node = new TreeNode() { Text = fname };
                    treeView_main.Nodes.Add(node);
                }
                foreach (string fname in _ExcelDRList)
                {
                    TreeNode node = new TreeNode() { Text = fname };
                    treeView_main.Nodes.Add(node);
                }
                foreach (string fname in _ExcelBHTJList)
                {
                    TreeNode node = new TreeNode() { Text = fname };
                    treeView_main.Nodes.Add(node);
                }

                dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                dockPanel_main_Plist.Width = this.Width;


            }
            #endregion
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 15
                || _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 7
                )
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    if (!Directory.Exists(fd.SelectedPath))
                    {
                        return;
                    }
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");

                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.Name.Contains("合肥路况") && !tfile.Name.Contains("~"))
                            _ExcelPathList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");

                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.Name.Contains("合肥路况") && !tfile.Name.Contains("~"))
                                _ExcelPathList.Add(tfile.FullName);
                        }
                    }
                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && _Setting.ExcelType == 1)
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains("裂缝、接缝料损坏分公里明细表_1000m"))
                            _ExcelPathList.Add(tfile.FullName);

                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains("裂缝、接缝料损坏分公里明细表_1000m"))
                                _ExcelPathList.Add(tfile.FullName);

                        }
                    }
                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            //湖南定制
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 4)
            {
                统计单元长度 unitlendlg = new 统计单元长度();
                unitlendlg.label1.Visible = false;
                unitlendlg.comboBox1.Visible = false;
                unitlendlg.ShowDialog();
                if (!unitlendlg._IsOK)
                {
                    return;
                }

                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelBHTJList.Clear();
                    _ExcelJSMXList.Clear();
                    _ExcelIRIList.Clear();
                    _ExcelDRList.Clear();

                    int unitlen = 统计单元长度._IndexUnitLen;

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains(string.Format("_病害流水表表格_{0}m.xlsx", unitlen)))
                            _DiseaseAll.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_空间定位数据_{0}m.xlsx", unitlen)))
                            _ExcelGpss.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_沥青路面损坏_{0}m.xlsx", unitlen)))
                            _DiseaseLQ.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_水泥路面损坏_{0}m.xlsx", unitlen)))
                            _DiseaseSN.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {

                            if (tfile.FullName.Contains(string.Format("_病害流水表表格_{0}m.xlsx", unitlen)))
                                _DiseaseAll.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_空间定位数据_{0}m.xlsx", unitlen)))
                                _ExcelGpss.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_沥青路面损坏_{0}m.xlsx", unitlen)))
                                _DiseaseLQ.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_水泥路面损坏_{0}m.xlsx", unitlen)))
                                _DiseaseSN.Add(tfile.FullName);
                        }
                    }

                    //报表列表

                    foreach (string fname in _DiseaseAll)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelGpss)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _DiseaseLQ)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _DiseaseSN)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 0)
            {
                统计单元长度 unitlendlg = new 统计单元长度();
                unitlendlg.ShowDialog();
                if (!unitlendlg._IsOK)
                {
                    return;
                }
                treeView_main.Nodes.Clear();
                _Projects.Clear();
                _CurProject = null;
                dockPanel_main_data.Controls.Clear();
                _ExcelBHTJList.Clear();
                _ExcelJSMXList.Clear();
                _ExcelIRIList.Clear();
                _ExcelDRList.Clear();
                //统计报表 弹出选择界面 
                try
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    FolderBrowserDialog fd = new FolderBrowserDialog();
                    fd.Description = "请选择路段报表文件夹：";
                    fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                    fd.ShowDialog();
                    if (fd.SelectedPath != string.Empty)
                    {
                        if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                        {
                            fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                        }

                        inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                        treeView_main.Nodes.Clear();
                        _Projects.Clear();
                        _CurProject = null;
                        dockPanel_main_data.Controls.Clear();
                        _ExcelPathList.Clear();
                        int unitlen = 统计单元长度._DisUnitLen;
                        // 文件夹里面的报表文件
                        DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                        FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in fileInfoArray)
                        {
                            if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                _ExcelBHTJList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                _ExcelJSMXList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                _ExcelIRIList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                _ExcelDRList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains("_PQI_1000m"))
                                _ExcelPathList.Add(tfile.FullName);
                        }

                        //文件夹里面的文件夹里面的报表文件
                        DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                        foreach (DirectoryInfo tdir in dirInfoArray)
                        {
                            FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                            foreach (FileInfo tfile in tfileInfoArray)
                            {
                                if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                    _ExcelBHTJList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                    _ExcelJSMXList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                    _ExcelIRIList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                    _ExcelDRList.Add(tfile.FullName);
                                else if (tfile.FullName.Contains(string.Format("_PQI_{0}m.xlsx", unitlen)))
                                    _ExcelPathList.Add(tfile.FullName);
                            }
                        }

                        //报表列表
                        foreach (string fname in _ExcelPathList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }
                        //报表列表
                        foreach (string fname in _ExcelBHTJList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }
                        foreach (string fname in _ExcelJSMXList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }
                        foreach (string fname in _ExcelIRIList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }
                        foreach (string fname in _ExcelDRList)
                        {
                            TreeNode node = new TreeNode() { Text = fname };
                            treeView_main.Nodes.Add(node);
                        }

                        dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                        dockPanel_main_data.Width = this.Width;
                    }
                }
                catch (Exception ex)
                {

                    throw ex;
                }

            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing && _Setting.ExcelType == 0)
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains("_PQI_1000m"))
                            _ExcelPathList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains("_PQI_1000m"))
                                _ExcelPathList.Add(tfile.FullName);
                        }
                    }

                    //报表列表
                    foreach (string fname in _ExcelPathList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && _Setting.ExcelType == 0)
            {
                IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择路段报表文件夹：";
                fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                    treeView_main.Nodes.Clear();
                    _Projects.Clear();
                    _CurProject = null;
                    dockPanel_main_data.Controls.Clear();
                    _ExcelPathList.Clear();
                    _ExcelBHTJList.Clear();
                    _ExcelJSMXList.Clear();
                    _ExcelIRIList.Clear();
                    _ExcelDRList.Clear();

                    int unitlen = 1000;

                    // 文件夹里面的报表文件
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in fileInfoArray)
                    {
                        if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                            _ExcelBHTJList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                            _ExcelJSMXList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                            _ExcelIRIList.Add(tfile.FullName);
                        else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                            _ExcelDRList.Add(tfile.FullName);
                    }

                    //文件夹里面的文件夹里面的报表文件
                    DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                    foreach (DirectoryInfo tdir in dirInfoArray)
                    {
                        FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                        foreach (FileInfo tfile in tfileInfoArray)
                        {
                            if (tfile.FullName.Contains(string.Format("_病害统计_{0}m.xlsx", unitlen)))
                                _ExcelBHTJList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_技术状况评定明细表_{0}m.xlsx", unitlen)))
                                _ExcelJSMXList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_IRI_{0}m.xlsx", unitlen)))
                                _ExcelIRIList.Add(tfile.FullName);
                            else if (tfile.FullName.Contains(string.Format("_PCI_{0}m.xlsx", unitlen)))
                                _ExcelDRList.Add(tfile.FullName);
                        }
                    }

                    //报表列表
                    foreach (string fname in _ExcelBHTJList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelJSMXList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelIRIList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }
                    foreach (string fname in _ExcelDRList)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                }
            }
        }

        private List<FileInfo> GetAllFilePath(string dirpath, string type)
        {
            List<FileInfo> projects = new List<FileInfo>();
            if (Directory.Exists(dirpath))
            {
                DirectoryInfo dir = new DirectoryInfo(dirpath);
                DirectoryInfo[] sdirs = dir.GetDirectories();
                FileInfo[] files = dir.GetFiles(type);
                projects.AddRange(files);

                foreach (DirectoryInfo d in sdirs)
                {
                    projects.AddRange(GetAllFilePath(d.FullName, type));
                }
            }
            return projects;
        }

        private void barButtonItem_ExportDoc_ItemClick(object sender, ItemClickEventArgs e)
        {
            /*{
                MSWord.Application wordApp = new MSWord.ApplicationClass() { Visible = true };
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                MyWord.OutputDocTest(wordApp, excelApp, "{0}\\报告模板\\城镇道路\\{1}.docx");
            }
             */
            // 城镇单独
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 0)
            {
                if (CheckExcelType())
                {
                    MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyWordCity.OutputDoc(wordApp, excelApp, _ExcelPathList);

                    excelApp.Quit();
                    wordApp.Quit();

                    MessageBox.Show("导出报告完成");
                }
            }
            // 城镇附件模板
            else if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 6)
            {
                if (CheckExcelType())
                {
                    MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyWordCity.OutputModel1Doc(wordApp, excelApp, _ExcelPathList);

                    excelApp.Quit();
                    wordApp.Quit();

                    MessageBox.Show("导出报告完成");
                }
            }
            // 07 标准
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007 && _Setting.ExcelType == 1)
            {
                MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                foreach (string fpath in _ExcelPathList)
                {
                    if (File.Exists(fpath.Replace("_1000m.xlsx", "_100m.xlsx")))
                    {
                        MyWord.OutputDoc(wordApp, excelApp, fpath);
                    }
                    else
                    {
                        MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少100m报表!");
                    }
                }

                excelApp.Quit();
                wordApp.Quit();

                MessageBox.Show("导出报告完成");
            }
            // 18 标准
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 1)
            {
                输出报告类型 wordBox = new 输出报告类型();
                wordBox.ShowDialog();
                if (!wordBox._IsOK)
                    return;

                MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                if (wordBox._WordType == 0)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_100m.xlsx")))
                        {
                            MyWord._18OutputDoc(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少100m报表!");
                        }
                    }
                }
                else if (wordBox._WordType == 1)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_10m.xlsx")))
                        {
                            MyWord._18OutputDoc_10_1000(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少10m报表!");
                        }
                    }
                }

                excelApp.Quit();
                wordApp.Quit();

                MessageBox.Show("导出报告完成");
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 1)
            {
                输出报告类型 wordBox = new 输出报告类型();
                wordBox.ShowDialog();
                if (!wordBox._IsOK)
                    return;

                MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                if (wordBox._WordType == 0)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_100m.xlsx")))
                        {
                            MyWord._OutputDoc_low(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少100m报表!");
                        }
                    }
                }
                else if (wordBox._WordType == 1)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_10m.xlsx")))
                        {
                            MyWord._18OutputDoc_low_10_1000(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少10m报表!");
                        }
                    }
                }

                excelApp.Quit();
                wordApp.Quit();

                MessageBox.Show("导出报告完成");
            }
            // 18 标准
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadGuangxi && _Setting.ExcelType == 1)
            {
                输出报告类型 wordBox = new 输出报告类型();
                wordBox.ShowDialog();
                if (!wordBox._IsOK)
                    return;

                MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                if (wordBox._WordType == 0)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_100m.xlsx")))
                        {
                            MyWordGX.GuangXiOutputDoc(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少100m报表!");
                        }
                    }
                }
                else if (wordBox._WordType == 1)
                {
                    foreach (string fpath in _ExcelPathList)
                    {
                        if (File.Exists(fpath.Replace("_1000m.xlsx", "_10m.xlsx")))
                        {
                            MyWordGX.GuangXiOutputDoc_10_1000(wordApp, excelApp, fpath);
                        }
                        else
                        {
                            MessageBox.Show(fpath.Substring(fpath.LastIndexOf('\\') + 1) + " 缺少10m报表!");
                        }
                    }
                }

                excelApp.Quit();
                wordApp.Quit();

                MessageBox.Show("导出报告完成");
            }


            else if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                报告设置 myCfgDlg = new 报告设置();
                myCfgDlg.ShowDialog();

                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                        MyWordCity.OutputMode7Doc(wordApp, excelApp, _ExcelListFilePath, treport.m_roadpartlist);

                        excelApp.Quit();
                        wordApp.Quit();
                        MessageBox.Show("导出报告完成");
                    }
                }
            }
        }

        //检查所导入的报表数量和类型有没有缺漏的
        private bool CheckExcelType()
        {
            bool res = true;
            string fname10 = null;

            for (int i = 0; i < _sidetypeLf.Length; ++i)
            {
                if (!_sidetypeLf[i])
                {
                    continue;
                }
                for (int j = 0; j < _xlstype.Length; ++j)
                {
                    bool flag = false;
                    foreach (string str in _ExcelPathList)
                    {
                        fname10 = str;
                        if (str.Contains(_sidetypeL[i]) && str.Contains(_xlstype[j]))
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        res = false;
                        MessageBox.Show(string.Format("导入的报表中缺少：\r\n【{0} {1} {2} {3}】\r\n的相关【100米】报表", _RoadLName, _RoadPName, _sidetypeL[i], _xlstype[j]));
                    }
                    fname10 = fname10.Replace("\\100米报表\\", "\\10米报表\\").Replace("_100m.xlsx", "_10m.xlsx");
                    fname10 = fname10.Replace("\\100米报表\\", "\\10米报表\\").Replace("_100m.xls", "_10m.xls");
                    if (!File.Exists(fname10))
                    {
                        MessageBox.Show(string.Format("导入的报表中缺少：\r\n【{0} {1} {2} {3}】\r\n的相关【10米】报表", _RoadLName, _RoadPName, _sidetypeL[i], _xlstype[j]));
                    }
                }
            }

            for (int i = 0; i < _sidetypeRf.Length; ++i)
            {
                if (!_sidetypeRf[i])
                {
                    continue;
                }
                for (int j = 0; j < _xlstype.Length; ++j)
                {
                    bool flag = false;
                    foreach (string str in _ExcelPathList)
                    {
                        if (str.Contains(_sidetypeR[i]) && str.Contains(_xlstype[j]))
                        {
                            flag = true;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        res = false;
                        MessageBox.Show(string.Format("导入的报表中缺少：\r\n【{0} {1} {2} {3}】\r\n的相关报表", _RoadLName, _RoadPName, _sidetypeR[i], _xlstype[j]));
                    }
                }
            }

            return res;
        }
        #endregion
        private void barButtonItem3_ItemClick_0(object sender, ItemClickEventArgs e)
        { 
            

            string defaultPath = GetDefaultLayoutPath();
            if (File.Exists(defaultPath))
            {
                try { dockManager_main.RestoreLayoutFromXml(defaultPath); }
                catch { /* 损坏的布局直接忽略 */ }
            }

            foreach (SingleProject tpro in _Projects)
            {
                tpro.RestoreDefaultLayout();
            }
            // 关键：加载布局后，强制刷新！
            this.AutoScaleMode = AutoScaleMode.None;
            ribbonControl1.Width = this.ClientSize.Width;
            dockManager_main.ForceInitialize(); // 强制重绘 

        }

        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            string defaultPath = GetDefaultLayoutPath();
            if (File.Exists(defaultPath))
            {
                try { dockManager_main.RestoreLayoutFromXml(defaultPath); }
                catch { }
            }

            foreach (SingleProject tpro in _Projects)
                tpro.RestoreDefaultLayout();

            // 终极修复：让主面板填满剩余空间
            var mainPanel = dockManager_main.Panels[1]; // 工程数据（Item2）
            var leftPanel = dockManager_main.Panels[0]; // 工程列表（Item1）

            int totalWidth = this.ClientSize.Width;
            int leftWidth = leftPanel.Width; // 200
            mainPanel.Width = totalWidth - leftWidth; // 强制填满！

            ribbonControl1.Width = totalWidth;
            dockManager_main.ForceInitialize();
        }


        private void barButtonItem_cfg_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2001)
            {
                SelectShow2001 tshow = new SelectShow2001(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis);
                tshow.ShowDialog();
            }
            else
            {
                设置病害勾选类型 tshow = new 设置病害勾选类型(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis);
                tshow.ShowDialog();
            }
        }

        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            ChoseExcel tbox = new ChoseExcel();
            tbox.ShowDialog();
            if (tbox._IsOK)
            {
                MSExcel.Application excelApp = new MSExcel.Application()
                {
                    Visible = true,
                    DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                    AlertBeforeOverwriting = false
                };
                MyWord.OutputZNExcel(excelApp, tbox._leftpath, tbox._rightpath, tbox._destpath);
                excelApp.Quit();
                MessageBox.Show("合并两车道报表完成！");
            }
        }
        /// <summary>
        /// gps桩号匹配
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
        {

            if (_CurProject == null)
            {
                MessageBox.Show("请打开任一工程!");
                return;
            }
            string highGpsFilePath = _CurProject._DataDir.FullName + "/GPSModel/gps.txt";
            if (File.Exists(highGpsFilePath))
            {
                DialogResult result = MessageBox.Show("二三维设备采集选择【是】,模块化设备采集选择【否】", "警告", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    _Setting.equipType = 1;
                }
                else
                {
                    _Setting.equipType = 0;
                }
                // result = MessageBox.Show("高精度定位模块计算是否时间统一减去一秒？", "警告", MessageBoxButtons.YesNo);
                //if (result == DialogResult.Yes)
                //{
                //    _Setting.allSub1s = true;
                //}
                //else
                //{
                //    _Setting.allSub1s = false;
                //}
            }
            this.Cursor = Cursors.WaitCursor;

            foreach (SingleProject proj in _Projects)
            {
                proj.MappingGPS2Mile();
            }


            MessageBox.Show("处理完成！");

            this.Cursor = Cursors.Default;
        }

        private void barButtonItem6_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ExcelType == 3)
            {
                ChoseExcel tbox = new ChoseExcel();
                tbox.ShowDialog();
                if (tbox._IsOK)
                {
                    MSExcel.Application excelApp = new MSExcel.Application()
                    {
                        Visible = true,
                        DisplayAlerts = false,
                        AlertBeforeOverwriting = false
                    };
                    MyWord.OutputZJGTExcel(excelApp, tbox._leftpath, tbox._destpath);
                    excelApp.Quit();
                    MessageBox.Show("合并两车道报表完成！");
                }
            }
        }

        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)  
        {
            if (_Projects.Count < 1)
            {
                MessageBox.Show("没有待处理的工程！");
            }
            else
            {
                this.Cursor = Cursors.WaitCursor;
                WinPanoProcessBar winbar = new WinPanoProcessBar(_Projects);
                StartPanoThread(winbar);
                winbar.ShowDialog();
                this.Cursor = Cursors.Default;
            }
        }

        private Thread ThreadPano;
        private void StartPanoThread(WinPanoProcessBar winbar)
        {
            ThreadPano = new Thread(PanoThreadMethod) { IsBackground = true };
            ThreadPano.Start(winbar);
        }
        private void PanoThreadMethod(object prj)
        {
            WinPanoProcessBar winbar = (WinPanoProcessBar)prj;
            StitchPanoImg(winbar);
        }
        private void StitchPanoImg(WinPanoProcessBar winbar)
        {
            winbar.SetMainMax(_Projects.Count);
            foreach (SingleProject proj in _Projects)
            {
                if (IsContainChinese(proj._DataDir.FullName))
                {
                    winbar.TextInfoAdd("工程路径中包含中文！跳过：" + proj._DataDir.FullName);
                }
                else
                {
                    winbar.TextInfoAdd("正在处理：" + proj._DataDir.Name);
                    proj.StitchPanoImg(winbar);
                    winbar.TextInfoAdd("处理完成：" + proj._DataDir.Name);
                }
                winbar.AddMainVal(1);
            }
            MessageBox.Show("全景图像拼接完成!");
        }
        private bool IsContainChinese(string input)
        {
            string strRex = @"[\u4e00-\u9fa5]";
            return System.Text.RegularExpressions.Regex.IsMatch(input, strRex);
        }
        //清除结果
        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
        {
            SelectIRM sltIRM = new SelectIRM(1);
            sltIRM.ShowDialog();
            if (sltIRM.IsYes())
            {
                this.Cursor = Cursors.WaitCursor;
                foreach (SingleProject proj in _Projects)
                {
                    proj.CleanIRMVal();
                }
                this.Cursor = Cursors.Default;
                MessageBox.Show("清除IRM中间计算结果成功！请重新计算IRM");
            }
        }

        private void btn_reportset_ItemClick(object sender, ItemClickEventArgs e)
        {
            软件设置 ReportSet = new 软件设置();


            ReportSet.ShowDialog();
        }

        private void ribbonControl1_Click(object sender, EventArgs e)
        {
            ribbonPageGroup6.Visible = true;
            barButtonItem38.Visibility = BarItemVisibility.Never;
            barButtonItem1.Visibility = BarItemVisibility.Always;
            if ((_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 0))
            {
                barButtonItem1.Visibility = BarItemVisibility.Always;
            }
            else if ((_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 6))
            {
                barButtonItem1.Visibility = BarItemVisibility.Always;
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007 && _Setting.ExcelType == 1)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 1)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadGuangxi && _Setting.ExcelType == 1)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing && _Setting.ExcelType == 0)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;
            }

            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel
                && (
                _Setting.ExcelType == 0 ||
                _Setting.ExcelType == 3 || _Setting.ExcelType == 6 ||
                _Setting.ExcelType == 4 ||
                _Setting.ExcelType == 7 || _Setting.ExcelType == 19))
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;

            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && (_Setting.ExcelType == 0 || _Setting.ExcelType == 1))
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;

            }
            else
            {
                ribbonPageGroup6.Visible = false;
            }

            if (_Setting.ExcelType == 3)
            {
                ribbonPageGroup3.Visible = true;
            }
            else
            {
                ribbonPageGroup3.Visible = false;
            }

            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                barButtonItem_ExportDoc.Caption = "生成报告主体";
                ribbonPageGroup6.Visible = true;
                barButtonItem38.Visibility = BarItemVisibility.Never;
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Never;

                barButtonItem13.Visibility = BarItemVisibility.Always;

                barButtonItem15.Visibility = BarItemVisibility.Always;
                barButtonItem16.Visibility = BarItemVisibility.Always;
                barButtonItem17.Visibility = BarItemVisibility.Always;
                barButtonItem19.Visibility = BarItemVisibility.Always;
                barButtonItem20.Visibility = BarItemVisibility.Always;
                barEditItem3.Visibility = BarItemVisibility.Always;
                barEditItem4.Visibility = BarItemVisibility.Always;
                barButtonItem22.Visibility = BarItemVisibility.Always;
            }
            else if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 0)
            {
                barButtonItem_ExportDoc.Caption = "生成报告";
                barButtonItem38.Visibility = BarItemVisibility.Never;
                barButtonItem1.Visibility = BarItemVisibility.Always;
                barButtonItem2.Visibility = BarItemVisibility.Always;

                barButtonItem13.Visibility = BarItemVisibility.Never;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never; //2022.2.18 修改
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
            }
            else if ((_Setting.ParmStyle == StandardParmType.DegreeRoad2007
                || _Setting.ParmStyle == StandardParmType.DegreeRoad2018
                || _Setting.ParmStyle == StandardParmType.RuralRoadGuangxi
                || _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                && _Setting.ExcelType == 1)
            {
                barButtonItem_ExportDoc.Caption = "生成报告";
                ribbonPageGroup6.Visible = true;
                barButtonItem13.Visibility = BarItemVisibility.Never;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never;
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
                barButtonItem_ExportDoc.Visibility = BarItemVisibility.Always;
            }
            else if ((_Setting.ParmStyle == StandardParmType.RuralRoadChongqing || _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel) && _Setting.ExcelType == 0)
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never;
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
                barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && (_Setting.ExcelType == 3 || _Setting.ExcelType == 4 || _Setting.ExcelType == 6 || _Setting.ExcelType == 7))
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never;
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
                barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && (_Setting.ExcelType == 1 || _Setting.ExcelType == 0))
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Always;
                barButtonItem13.Visibility = BarItemVisibility.Always;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never;
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
                barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;
            }
            else
            {
                barButtonItem1.Visibility = BarItemVisibility.Never;
                barButtonItem2.Visibility = BarItemVisibility.Never;
                barButtonItem13.Visibility = BarItemVisibility.Never;
                barButtonItem15.Visibility = BarItemVisibility.Never;
                barButtonItem16.Visibility = BarItemVisibility.Never;
                barButtonItem17.Visibility = BarItemVisibility.Never;
                barButtonItem19.Visibility = BarItemVisibility.Never;
                barButtonItem20.Visibility = BarItemVisibility.Never;
                barEditItem3.Visibility = BarItemVisibility.Never;
                barEditItem4.Visibility = BarItemVisibility.Never;
                barButtonItem22.Visibility = BarItemVisibility.Never;
                barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;
            }

            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018)
            {
                if (_Setting.ExcelType == 0)
                {
                    ribbonPageGroup6.Visible = true;
                    ribbonPageGroup3.Visible = false;
                    barButtonItem38.Visibility = BarItemVisibility.Never;
                    barButtonItem1.Visibility = BarItemVisibility.Never;
                    barButtonItem2.Visibility = BarItemVisibility.Always;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem13.Visibility = BarItemVisibility.Always;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;

                    barButtonItem25.Visibility = BarItemVisibility.Always;
                }
                else if (_Setting.ExcelType == 1)
                {
                    barButtonItem_ExportDoc.Caption = "生成报告";
                    barButtonItem13.Visibility = BarItemVisibility.Never;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barEditItem3.Visibility = BarItemVisibility.Never;
                    barEditItem4.Visibility = BarItemVisibility.Never;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Always;

                }

                else if (_Setting.ExcelType == 14)
                {
                    ribbonPageGroup6.Visible = true;
                    ribbonPageGroup3.Visible = false;
                    barButtonItem38.Visibility = BarItemVisibility.Always;
                    barButtonItem1.Visibility = BarItemVisibility.Never;
                    barButtonItem2.Visibility = BarItemVisibility.Never;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem13.Visibility = BarItemVisibility.Never;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;
                    barButtonItem25.Visibility = BarItemVisibility.Always;
                    // barEditItem4.Visibility = BarItemVisibility.Always;
                    barEditItem3.Visibility = BarItemVisibility.Always;
                }
                else if (_Setting.ExcelType == 19)
                {
                    ribbonPageGroup6.Visible = true;
                    ribbonPageGroup3.Visible = false;
                    barButtonItem38.Visibility = BarItemVisibility.Never;
                    barButtonItem1.Visibility = BarItemVisibility.Never;
                    barButtonItem2.Visibility = BarItemVisibility.Always;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem13.Visibility = BarItemVisibility.Always;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;

                    barButtonItem25.Visibility = BarItemVisibility.Never;
                }
                else if (_Setting.ExcelType == 15)
                {
                    ribbonPageGroup6.Visible = true;
                    ribbonPageGroup3.Visible = false;
                    barButtonItem38.Visibility = BarItemVisibility.Never;
                    barButtonItem1.Visibility = BarItemVisibility.Never;
                    barButtonItem2.Visibility = BarItemVisibility.Always;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem13.Visibility = BarItemVisibility.Always;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;

                    barButtonItem25.Visibility = BarItemVisibility.Never;
                }
                else
                {

                    ribbonPageGroup6.Visible = true;
                    ribbonPageGroup3.Visible = false;
                    barButtonItem38.Visibility = BarItemVisibility.Never;
                    barButtonItem1.Visibility = BarItemVisibility.Never;
                    barButtonItem2.Visibility = BarItemVisibility.Never;
                    barButtonItem22.Visibility = BarItemVisibility.Never;
                    barButtonItem13.Visibility = BarItemVisibility.Never;
                    barButtonItem15.Visibility = BarItemVisibility.Never;
                    barButtonItem16.Visibility = BarItemVisibility.Never;
                    barButtonItem17.Visibility = BarItemVisibility.Never;
                    barButtonItem19.Visibility = BarItemVisibility.Never;
                    barButtonItem20.Visibility = BarItemVisibility.Never;
                    barButtonItem_ExportDoc.Visibility = BarItemVisibility.Never;

                    barButtonItem25.Visibility = BarItemVisibility.Always;
                }
            }
            else
            {
                barButtonItem25.Visibility = BarItemVisibility.Never;
            }
            //新加

            if ((_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 15) || (
               _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 7)
               )
            {
                barButtonItem46.Visibility = BarItemVisibility.Always;
                barEditItem3.Visibility = BarItemVisibility.Always;
            }
            else
            {
                barButtonItem46.Visibility = BarItemVisibility.Never;
            }
        }

        private void barButtonItem_snbkDIs_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Projects.Count <= 0)
            {
                MessageBox.Show("请先导入工程！");
                return;
            }

            if (_Setting.SelectDrawDis == 1)
            {
                MessageBox.Show("不支持自动识别病害模式，\r\n请在软件设置里，将病害勾画方式设置为人工调查");
                return;
            }
            SnbkSetForm snform = new SnbkSetForm();
            snform.ShowDialog();
            if (!snform.falg) return;

            if (SnbkSetForm.bknum >= 0)
            {
                FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    if (_Setting.ParmStyle == StandardParmType.DegreeRoad2007)
                    {
                        MyExcelDegree2007.LoadXlsParm();
                    }

                    else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018)
                    {
                        if (_Setting.SelectDrawDis == 0)
                        {
                            MyExcelDegree2018.LoadXlsParm();
                        }
                        else
                        {
                            MyExcelDegreeSmall2018.LoadXlsParm();
                        }
                    }
                    else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                    {
                        if (_Setting.SelectDrawDis == 0)
                        {
                            MyExcelVillageDegree.LoadXlsParm();
                        }
                        else
                        {
                            MyExcelVillageDegreeSmall.LoadXlsParm();
                        }
                    }
                    else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2001)
                    {
                        MyExcelDegree2001.LoadXlsParm();
                    }


                    MSExcel.Application excelApp = new MSExcel.Application()
                    {
                        Visible = true,
                        DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                        AlertBeforeOverwriting = false
                    };

                    string outdirpath = null;
                    foreach (SingleProject proj in _Projects)
                    {
                        outdirpath = fd.SelectedPath;
                        if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                        {
                            string smile = proj._ProjectInfo._StartMile.ToString("K0000+000");
                            string emile = proj._ProjectInfo._EndMile.ToString("K0000+000");

                            string tt = "\\" + proj._DataDir.Name;
                            tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                            outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);

                            if (!Directory.Exists(outdirpath))
                            {
                                Directory.CreateDirectory(outdirpath);
                            }
                        }

                        proj.BkExcel(excelApp, outdirpath);
                    }

                    excelApp.Quit();
                    MessageBox.Show("导出报表完成！");
                }
            }
        }

        private void barButtonItem_street_xls_ItemClick(object sender, ItemClickEventArgs e)
        {
            导出景观报表区间 exceldis = new 导出景观报表区间();
            exceldis.ShowDialog();
            if (exceldis.NeedSub)
            {
                _Setting.needSub = true;
                string subStr = string.Join(",", exceldis.SubData.ToArray());
                _Setting.subData = subStr;
                //写到配置文件
            }
            else
            {
                _Setting.needSub = false;
                string subStr = "";
                _Setting.subData = subStr;
            }


            int[][] _ExcelDisVal;//存放10，100，1000 等自定义的区间
            if (exceldis._IsExcel)
            {
                _ExcelDisVal = new int[_Setting.StreetLenExcelNum][];
                for (int i = 0; i < _Setting.StreetLenExcelNum; i++)
                {
                    string[] strs = _Setting.StreetLenExcel[i].Split(',');
                    _ExcelDisVal[i] = new int[strs.Length];
                    for (int j = 0; j < strs.Length; ++j)
                        _ExcelDisVal[i][j] = int.Parse(strs[j]);
                }

                FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    switch (_Setting.ParmStyle)
                    {
                        case StandardParmType.DegreeRoad2007: MyExcelDegree2007.LoadXlsParm(); break;
                        case StandardParmType.CityRoad: MyExcelCity.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadBeijing: MyExcelBJDegree.LoadXlsParm(); break;
                        case StandardParmType.DegreeRoad2018:
                            if (_Setting.SelectDrawDis == 0)
                            {
                                MyExcelDegree2018.LoadXlsParm();
                            }
                            else
                            {
                                MyExcelDegreeSmall2018.LoadXlsParm();
                            }
                            break;
                        case StandardParmType.DegreeRoad2001: MyExcelDegree2001.LoadXlsParm(); break;
                        case StandardParmType.CityRoadShanghai: MyExcelCitySH2013.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadLiaoning: MyExcelLNDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadGuangxi: MyExcelGXDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadChongqing: MyExcelCQDegree.LoadXlsParm(); break;
                        case StandardParmType.RuralRoadlowLevel:

                            if (_Setting.SelectDrawDis == 0)
                            {
                                MyExcelVillageDegree.LoadXlsParm();
                            }
                            {

                                MyExcelVillageDegreeSmall.LoadXlsParm();
                            }
                            break;
                        case StandardParmType.RuralRoadHunan:
                            if (_Setting.SelectDrawDis == 0)
                            {
                                MyExcelHNDegree.LoadXlsParm_new();
                            }
                            {

                                MyExcelHNDegreeSmall.LoadXlsParm_new();
                            }
                           break;
                        default: break;
                    }

                    MSExcel.Application excelApp = new MSExcel.Application()
                    {
                        Visible = true,
                        DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                        AlertBeforeOverwriting = false
                    };

                    string outdirpath = null;

                    if (_Setting.needSub)
                    {
                        MessageBox.Show($"当前分段区间{_Setting.subData}");
                        string[] subNames = _Setting.subData.Split(',');
                        for (int i = 0; i < subNames.Length; i += 2)
                        {
                            foreach (SingleProject proj in _Projects)
                            {
                                //需要进行分段输出

                                outdirpath = fd.SelectedPath;

                                if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                                {

                                    _Setting.nowSubIndexStr = subNames[i] + "," + subNames[i + 1];
                                    string smile = int.Parse(subNames[i]).ToString("K0000+000");
                                    string emile = int.Parse(subNames[i + 1]).ToString("K0000+000");


                                    string tt = "\\" + proj._DataDir.Name;
                                    tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                                    outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);


                                    if (!Directory.Exists(outdirpath))
                                    {
                                        Directory.CreateDirectory(outdirpath);
                                    }
                                }

                                if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018
                               || _Setting.ParmStyle == StandardParmType.RuralRoadGuangxi
                               || _Setting.ParmStyle == StandardParmType.RuralRoadChongqing
                               || _Setting.ParmStyle == StandardParmType.RuralRoadHunan || _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                                {
                                    proj.GenerateExcel_Street(excelApp, outdirpath, _ExcelDisVal, _Setting.StreetIsExcel);
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (SingleProject proj in _Projects)
                        {
                            outdirpath = fd.SelectedPath;

                            if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                            {
                                string smile = proj._ProjectInfo._StartMile.ToString("K0000+000");
                                string emile = proj._ProjectInfo._EndMile.ToString("K0000+000");

                                string tt = "\\" + proj._DataDir.Name;
                                tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                                outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);

                                if (!Directory.Exists(outdirpath))
                                {
                                    Directory.CreateDirectory(outdirpath);
                                }
                            }

                            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018
                                || _Setting.ParmStyle == StandardParmType.RuralRoadGuangxi
                                || _Setting.ParmStyle == StandardParmType.RuralRoadChongqing
                                || _Setting.ParmStyle == StandardParmType.RuralRoadHunan || _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                            {
                                proj.GenerateExcel_Street(excelApp, outdirpath, _ExcelDisVal, _Setting.StreetIsExcel);
                            }
                        }
                    } 
                    excelApp.Quit();
                    MessageBox.Show("导出报表完成！");
                }
            }
        }

        private void barEditItem1_EditValueChanged(object sender, EventArgs e)
        {
            _IsLinkShow = Convert.ToBoolean(barEditItem1.EditValue);
        }

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            string fpath;
            if (_Projects.Count > 0)
            {
                fpath = _Projects[0]._DataDir.Parent.FullName + "\\数据检查结果.txt";
                if (File.Exists(fpath))
                {
                    File.Delete(fpath);
                }
            }
            else
            {
                return;
            }
            List<string> errorProjectInfo = new List<string>();
            List<string> errorProjectCameraSettingInfo = new List<string>();
            List<string> errorProjectBinFileExistInfo = new List<string>();
            foreach (SingleProject proj in _Projects)
            {
                proj.CheckOriDataComplete(ref errorProjectInfo);
            }
            #region 检查cameraSetting.ini文件裁剪是否一致 

            double firstJzWidth = 0;
            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
            {
                //检查cameraSetting.ini文件裁剪是否一致 
                for (int i = 0; i < _Projects.Count; i++)
                {
                    var project = _Projects[i];
                    double jzWidth = 0;
                    DirectoryInfo dir = new DirectoryInfo(project._DataDir.FullName);
                    GetVillagePara(dir.FullName, out jzWidth);
                    if (i == 0)
                    {
                        firstJzWidth = jzWidth;
                    }
                    else
                    {
                        if (jzWidth != firstJzWidth)
                        {
                            string info = $"检测到[{project._DataDir.Name}]工程裁剪宽度为{jzWidth}，与其他工程不一致。";
                            errorProjectCameraSettingInfo.Add(info);

                        }
                    }
                }
                //检测.bin文件是否存在
                for (int i = 0; i < _Projects.Count; i++)
                {
                    var project = _Projects[i];
                    string binFilePath = Path.Combine(project._DataDir.FullName, "u_d.bin");
                    if (!File.Exists(binFilePath))
                    {
                        string info = $"检测到低等级农村公路工程[{project._DataDir.Name}]不存在u_d.bin文件。";
                        errorProjectBinFileExistInfo.Add(info);

                    }
                }
            }

            if (errorProjectInfo.Count > 0)
            {
                MessageBox.Show($"检测到{errorProjectInfo.Count}条错误信息，具体错误信息请查看输出数据检查结果.txt。\n位置在{fpath}");
            }
            if (errorProjectCameraSettingInfo.Count > 0)
            {
                using (StreamWriter sw = new StreamWriter(fpath, true))
                {
                    foreach (var errorMsg in errorProjectCameraSettingInfo)
                    {
                        sw.WriteLine(errorMsg);
                    }
                }
                MessageBox.Show("检测到农村路导入工程存在裁剪幅宽不一致，具体工程信息请查看输出数据检查结果.txt。");
            }

            if (errorProjectBinFileExistInfo.Count > 0)
            {
                using (StreamWriter sw = new StreamWriter(fpath, true))
                {
                    foreach (var errorMsg in errorProjectBinFileExistInfo)
                    {
                        sw.WriteLine(errorMsg);
                    }
                }
                MessageBox.Show("检测到农村路导入工程不存在u_d.bin文件，具体工程信息请查看输出数据检查结果.txt。");
            }

            #endregion

            this.Cursor = Cursors.Default;
            MessageBox.Show("数据检查结束！");
        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Projects.Count < 1)
                return;

            GetDiseaseFiles getDiseaseFiles = new GetDiseaseFiles();
            getDiseaseFiles.openPath = _Projects[0]._DataDir.Parent.FullName;
            List<string> dirs = new List<string>();
            foreach (SingleProject proj in _Projects)
            {
                dirs.Add(proj._DataDir.FullName);
            }
            getDiseaseFiles.srcPaths = dirs;
            getDiseaseFiles.ShowDialog();
        }

        private void barEditItem2_EditValueChanged(object sender, EventArgs e)
        {
            _IsSaveDisImg = Convert.ToBoolean(barEditItem2.EditValue);
            if (_IsSaveDisImg)
            {
                MessageBox.Show("进入保存路面病害图片模式，播放/翻页路面图片时，仅显示有病害图片，不可进行病害拉框操作！", "提示");
            }
            else
            {
                MessageBox.Show("退出保存路面病害图片模式，可正常进行病害拉框操作！", "提示");
            }
        }

        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                        MyWordCity.OutputMode7DocXls(excelApp, _ExcelListFilePath, treport.m_roadpartlist);
                        excelApp.Quit();
                        MessageBox.Show("生成统计报表完成！");
                    }
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && _Setting.ExcelType == 1)
            {
                if (_ExcelPathList.Count > 0)
                {

                    if (_ExcelPathList.Count > 0)
                    {
                        IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                        string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                        MyExcelHNDegree.OutputAllRoadStatistics_hn(excelApp, fpath, _ExcelPathList);
                        excelApp.Quit();
                        MessageBox.Show("生成统计报表完成！");
                    }
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 0)
            {
                if (_ExcelBHTJList.Count > 0 && _ExcelJSMXList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                    if (_Setting.SelectDrawDis == 0)
                    {
                        MyExcelDegree2018.LoadXlsParm();
                        MyExcelDegree2018.OutputAreaStatistics(excelApp, _ExcelBHTJList, _ExcelJSMXList, fpath);
                        MyExcelDegree2018.OutputAllRoadStatistics(excelApp, fpath, _ExcelBHTJList, _ExcelJSMXList, _ExcelDRList, _ExcelIRIList);
                    }
                    else if (_Setting.SelectDrawDis == 1)
                    {
                        MyExcelDegreeSmall2018.LoadXlsParm();
                        MyExcelDegreeSmall2018.OutputAreaStatistics(excelApp, _ExcelBHTJList, _ExcelJSMXList, fpath);
                    }


                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 20)
            {
                //上海惠普
                MyExcelDegree2018.LoadXlsParm();
                if (_ExcelPathList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                    if (_Setting.SelectDrawDis == 0)
                    {
                        MyExcelDegree2018.LoadXlsParm();

                        MyExcelDegree2018.OutputAllRoadDisease_HP(excelApp, fpath, _ExcelPathList);
                    }



                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing && _Setting.ExcelType == 0)
            {
                if (_ExcelPathList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                    MyExcelCQDegree.OutputAllRoadStatistics(excelApp, fpath, _ExcelPathList);

                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 0)
            {
                if (_ExcelPathList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                    //DialogResult result =  MessageBox.Show("请选择是否为大框版本?", "选择窗口", MessageBoxButtons.YesNo);
                    // if (result == DialogResult.Yes)
                    {
                        MyExcelVillageDegree.LoadXlsParm();
                        MyExcelVillageDegree.OutputAllRoadStatistics(excelApp, fpath, _ExcelBHTJList, _ExcelJSMXList, _ExcelDRList, _ExcelIRIList);
                    }

                    excelApp.Quit();
                    excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyExcelVillageDegree.OutputAllRoadStatistics(excelApp, fpath, _ExcelPathList);
                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }

            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 3)
            {

                switch (_Setting.multiExcelMergeType)
                {
                    case 0:
                        #region 多车道统计

                        if (_ExcelPathList.Count > 0)
                        {
                            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                            string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                            MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                            MyExcelVillageDegree.OutputAllRoadStatistics(excelApp, fpath, _ExcelPathList);
                            excelApp.Quit();
                            MessageBox.Show("生成统计报表完成！");
                        }
                        #endregion
                        break;
                    case 1:
                        #region 咸宁
                        if (_ExcelJSMXList.Count > 0)
                        {
                            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                            string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                            MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                            MyExcelVillageDegree.LoadXlsParm();
                            MyExcelVillageDegree.OutputAllRoadStatistics_XN(excelApp, fpath, _ExcelJSMXList, _ExcelDRList, _ExcelIRIList);
                            excelApp.Quit();
                            MessageBox.Show("生成统计报表完成！");
                        }
                        #endregion
                        break;
                    case 2:
                        #region 孝感统计
                        if (_ExcelJSMXList.Count > 0)
                        {
                            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                            string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                            MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                            MyExcelVillageDegree.LoadXlsParm();
                            MyExcelVillageDegree.OutputAllRoadStatistics_XG(excelApp, fpath, _ExcelJSMXList, _ExcelDRList, _ExcelIRIList, _LJSHAndLJSumList);
                            excelApp.Quit();
                            MessageBox.Show("生成统计报表完成！");
                        }
                        #endregion
                        break;
                    case 3:
                        #region 南昌农村路
                        if (_ExcelJSMXList.Count > 0)
                        {
                            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                            string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                            MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                            MyExcelVillageDegree.LoadXlsParm();
                            MyExcelVillageDegree.outMulitExcel_NC(excelApp, fpath, _ExcelJSMXList);

                            excelApp.Quit();
                            excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                            MyExcelVillageDegree.LoadXlsParm();
                            MyExcelVillageDegree.outMulitExcel_NC1(excelApp, fpath, _ExcelBHTJList);

                            excelApp.Quit();

                            excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                            MyExcelVillageDegree.LoadXlsParm();
                            MyExcelVillageDegree.outMulitExcel_NC2(excelApp, fpath, _ExcelBHTJList);

                            excelApp.Quit();
                            MessageBox.Show("生成统计报表完成！");
                        }
                        #endregion
                        break;
                    default:
                        break;
                }
            }
            else if (
                _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 7)
            {
                if (_ExcelPathList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyExcelVillageDegreeSmall.LoadXlsParm();
                    MyExcelVillageDegreeSmall.OutputAllRoadStatistics_Hefei(excelApp, fpath, _ExcelPathList);
                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.ExcelType == 15)
            {
                if (_ExcelPathList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyExcelVillageDegreeSmall.LoadXlsParm();
                    MyExcelVillageDegreeSmall.OutputAllRoadStatistics_2018Hefei(excelApp, fpath, _ExcelPathList);
                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            //湖南定制
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 4)
            {
                if (_ExcelGpss.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MyExcelVillageDegree.LoadXlsParm();
                    MyExcelVillageDegree.OutputAllRoadStatistics_hunan(excelApp, fpath, _ExcelGpss, _DiseaseLQ, _DiseaseSN, _DiseaseAll);
                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else if (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && _Setting.ExcelType == 0)
            {
                if (_ExcelBHTJList.Count > 0 && _ExcelJSMXList.Count > 0)
                {
                    IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
                    string fpath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                    MyExcelHNDegree.LoadXlsParm_new();
                    MyExcelHNDegree.OutputAllRoadStatistics(excelApp, fpath, _ExcelBHTJList, _ExcelJSMXList, _ExcelDRList, _ExcelIRIList);
                    excelApp.Quit();
                    MessageBox.Show("生成统计报表完成！");
                }
            }
            else
            {
                MessageBox.Show("请注意车道统计功能未进入任何分支，\n检查模块类型及软件设置内出表选项");
                log.Warn("请注意车道统计功能未进入任何分支，检查模块类型及软件设置内出表选项");
            }
        }

        private void barButtonItem15_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                        MyWordCity.OutputMode7DocAppendix(wordApp, excelApp, _ExcelListFilePath, treport.m_roadpartlist);

                        excelApp.Quit();
                        wordApp.Quit();
                        MessageBox.Show("导出报告附录完成！");
                    }
                }
            }
        }

        private void barButtonItem16_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                SHPG_ExportWord_LoadData();
                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MyWordCity.OutputMode7DocMerge(wordApp, _ExcelListFilePath, treport.m_roadpartlist, treport);
                        wordApp.Quit();
                        MessageBox.Show("合并报告完成！");
                    }
                }
            }

        }

        private void barButtonItem17_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && _Setting.ExcelType == 7)

            {

            }


            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                报告设置 myCfgDlg = new 报告设置();
                myCfgDlg.ShowDialog();

                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ProjectProject != null)
                {
                    ProjectProjectClass tproject = new ProjectProjectClass();
                    tproject = ProjectProjectClass.DeepCopyByBinary(_ProjectProject);

                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                        MyWordCity.OutputMode7DocXls(excelApp, _ExcelListFilePath, treport.m_roadpartlist);
                        MyWordCity.OutputMode7DocAppendix(wordApp, excelApp, _ExcelListFilePath, treport.m_roadpartlist);
                        MyWordCity.OutputMode7Doc(wordApp, excelApp, _ExcelListFilePath, treport.m_roadpartlist);
                        MyWordCity.OutputMode7DocHeader(wordApp, excelApp, _ExcelListFilePath, tproject, treport);
                        MyWordCity.OutputMode7DocSummary(wordApp, excelApp, _ExcelListFilePath);
                        MyWordCity.OutputMode7DocMerge(wordApp, _ExcelListFilePath, treport.m_roadpartlist, treport);

                        excelApp.Quit();
                        wordApp.Quit();
                        MessageBox.Show("导出报告完成！");
                    }
                }
            }
        }

        private void barEditItem3_EditValueChanged(object sender, EventArgs e)
        {
            _Setting.OutWordPasteDelay = Convert.ToInt32(barEditItem3.EditValue);
            GlobalWord.wd_sleep_us = Convert.ToInt32(barEditItem3.EditValue);
        }

        private void barEditItem4_EditValueChanged(object sender, EventArgs e)
        {
            GlobalWord.wd_sleep_us2 = Convert.ToInt32(barEditItem4.EditValue);
        }

        private void barButtonItem18_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Projects.Count <= 0)
            {
                return;
            }
            this.Cursor = Cursors.WaitCursor;

            // 将所有工程分成上下行
            List<SingleProject> UpProject = new List<SingleProject>(); //上行工程集合
            List<SingleProject> DownProject = new List<SingleProject>(); //上行工程集合
            foreach (SingleProject proj in _Projects)
            {
                if (proj._ProjectInfo._Direction > 0)
                {
                    UpProject.Add(proj);
                }
                else
                {
                    DownProject.Add(proj);
                }
            }

            //上下行配对
            foreach (SingleProject projdown in DownProject)
            {
                bool IsFindUpProject = false;
                foreach (SingleProject projup in UpProject)
                {
                    // 同条路的上下行道路编号一样
                    if (projdown._ProjectInfo._RoadCode == projup._ProjectInfo._RoadCode)
                    {
                        // 同条路的上下行车道编号一样
                        if (projdown._ProjectInfo._RoadNum == projup._ProjectInfo._RoadNum)
                        {
                            // 将下行的单元分段和上行的对齐
                            RoadMatchUnit(projup, projdown);
                            IsFindUpProject = true;
                            break;
                        }
                    }
                }
                if (!IsFindUpProject)
                {
                    MessageBox.Show("没有找到\r\n" + projdown._DataDir.FullName + "\r\n匹配上行车道工程！\r\n请点击确定继续！");
                }
            }
            this.Cursor = Cursors.Default;
            MessageBox.Show("上下行路口桩号对齐调整完成！");
        }

        // 将上下行的桩号相匹配起来
        private void RoadMatchUnit(SingleProject projup, SingleProject projdown)
        {
            string oldmarkpathup = projup._DataDir.FullName + "\\RoadStatuMarkInfo.txt";
            string oldmarkpathdown = projdown._DataDir.FullName + "\\RoadStatuMarkInfo.txt";

            // 上行的打标文件不存在
            if (!File.Exists(oldmarkpathup))
            {
                if (!File.Exists(oldmarkpathdown))
                {
                    // 如果上下行的打标文件都不存在，就直接返回
                    return;
                }
                else
                {
                    // 如果上行的打标文件不存在，下行的打标文件存在就将下行打标文件中的路段单元去掉
                    // 先将原来的文件另存为备份
                    string[] markstrs = File.ReadAllLines(oldmarkpathdown);
                    string oldmarkpathdownnew = projdown._DataDir.FullName + "\\RoadStatuMarkInfo.xrbak.txt";
                    if (!File.Exists(oldmarkpathdownnew))
                    {
                        File.Move(oldmarkpathdown, oldmarkpathdownnew);
                    }

                    // 先将打标文件中的路段单元去掉
                    List<string> markstrslistnew = new List<string>();
                    foreach (string str in markstrs)
                    {
                        if (!(str.Contains("路面单元") && (str.Contains("进路口") || str.Contains("出路口"))))
                        {
                            markstrslistnew.Add(str);
                        }
                    }

                    // 将更新后的打标记录写入文件
                    if (markstrslistnew.Count > 0)
                    {
                        File.WriteAllLines(oldmarkpathdown, markstrslistnew.ToArray(), Encoding.UTF8);
                    }
                    else
                    {
                        File.Delete(oldmarkpathdown);
                    }
                }
            }
            else //上行的打标文件存在
            {
                if (!File.Exists(oldmarkpathdown))
                {
                    // 如果下行的打标文件不存在，将上行的路段单元信息写入，将上行的进出路口调换写入下行
                    string[] markstrs = File.ReadAllLines(oldmarkpathup);
                    List<string> markstrslistnew = new List<string>();
                    foreach (string str in markstrs)
                    {
                        if (str.Contains("路面单元"))
                        {
                            if (str.Contains("进路口"))
                            {
                                string[] strs = str.Split(' ');
                                int mile = projup._ProjectInfo.Dmi2Mile(int.Parse(strs[2]));
                                int dmi = projdown._ProjectInfo.Mile2Dmi(mile);
                                string newstr = string.Format("{0} {1} {2} 路面单元:出路口", mile, mile, dmi);
                                markstrslistnew.Add(newstr);
                            }
                            if (str.Contains("出路口"))
                            {
                                string[] strs = str.Split(' ');
                                int mile = projup._ProjectInfo.Dmi2Mile(int.Parse(strs[2]));
                                int dmi = projdown._ProjectInfo.Mile2Dmi(mile);
                                string newstr = string.Format("{0} {1} {2} 路面单元:进路口", mile, mile, dmi);
                                markstrslistnew.Add(newstr);
                            }
                        }
                    }


                    // 将更新后的打标记录写入文件
                    if (markstrslistnew.Count > 0)
                    {
                        File.WriteAllLines(oldmarkpathdown, markstrslistnew.ToArray(), Encoding.UTF8);
                    }
                }
                else
                {
                    // 如果下行的打标文件存在，去掉原来下行打标文件中的路段单元信息，将上行的路段单元信息写入 
                    // 先将原来的文件另存为备份
                    string[] markstrs = File.ReadAllLines(oldmarkpathdown);
                    string oldmarkpathdownnew = projdown._DataDir.FullName + "\\RoadStatuMarkInfo.xrbak.txt";
                    if (!File.Exists(oldmarkpathdownnew))
                    {
                        File.Move(oldmarkpathdown, oldmarkpathdownnew);
                    }

                    // 先将原来下行的打标文件中的路段单元去掉
                    List<string> markstrslistnew = new List<string>();
                    foreach (string str in markstrs)
                    {
                        if (!(str.Contains("路面单元") && (str.Contains("进路口") || str.Contains("出路口"))))
                        {
                            markstrslistnew.Add(str);
                        }
                    }

                    // 将上行的进出路口调换写入下行
                    markstrs = File.ReadAllLines(oldmarkpathup);
                    foreach (string str in markstrs)
                    {
                        if (str.Contains("路面单元"))
                        {
                            if (str.Contains("进路口"))
                            {
                                string[] strs = str.Split(' ');
                                int mile = projup._ProjectInfo.Dmi2Mile(int.Parse(strs[2]));
                                int dmi = projdown._ProjectInfo.Mile2Dmi(mile);
                                string newstr = string.Format("{0} {1} {2} 路面单元:出路口", mile, mile, dmi);
                                markstrslistnew.Add(newstr);
                            }
                            if (str.Contains("出路口"))
                            {
                                string[] strs = str.Split(' ');
                                int mile = projup._ProjectInfo.Dmi2Mile(int.Parse(strs[2]));
                                int dmi = projdown._ProjectInfo.Mile2Dmi(mile);
                                string newstr = string.Format("{0} {1} {2} 路面单元:进路口", mile, mile, dmi);
                                markstrslistnew.Add(newstr);
                            }
                        }
                    }

                    // 将更新后的打标记录写入文件
                    if (markstrslistnew.Count > 0)
                    {
                        File.WriteAllLines(oldmarkpathdown, markstrslistnew.ToArray(), Encoding.UTF8);
                    }
                    else
                    {
                        File.Delete(oldmarkpathdown);
                    }
                }
            }
        }

        private void barButtonItem19_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                报告设置 myCfgDlg = new 报告设置();
                myCfgDlg.ShowDialog();

                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                        MyWordCity.OutputMode7DocSummary(wordApp, excelApp, _ExcelListFilePath);

                        excelApp.Quit();
                        wordApp.Quit();
                        MessageBox.Show("导出报告完成");
                    }
                }
            }
        }

        private void barButtonItem20_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }

                SHPG_ExportWord_LoadData();
                if (_ProjectProject != null && _ReportProject != null)
                {
                    ProjectProjectClass tproject = new ProjectProjectClass();
                    tproject = ProjectProjectClass.DeepCopyByBinary(_ProjectProject);

                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);

                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        {
                            if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                            {
                                treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                                --i2;
                            }
                        }
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查报表是否已导出！");
                    }
                    else
                    {
                        MSWord.Application wordApp = new MSWord.Application() { Visible = true };
                        MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                        MyWordCity.OutputMode7DocHeader(wordApp, excelApp, _ExcelListFilePath, tproject, treport);

                        excelApp.Quit();
                        wordApp.Quit();
                        MessageBox.Show("生成报告头完成！");
                    }
                }
            }
        }

        /// <summary>
        /// 上海浦公要导出的报告 所属的项目
        /// </summary>
        ProjectProjectClass _ProjectProject = null;

        /// <summary>
        /// 上海浦公要导出的报告
        /// </summary>
        ReportProjectClass _ReportProject = null;

        /// <summary>
        /// 选中要导出的检测报告ID
        /// </summary>
        public string m_checked_reportID = null;

        public string[] _LaneLayoutStr = { "零", "一", "两", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六" };

        /// <summary>
        /// 选中要导出的项目ID
        /// </summary>
        public string m_checked_prjID = null;
        private void SHPG_ExportWord_LoadData()
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择路段报表文件夹：";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                if (fd.SelectedPath != string.Empty)
                {
                    List<ProjectProjectClass> srcProjectinfoList = new List<ProjectProjectClass>();

                    bool chkres = false;
                    string roadinfosheet = fd.SelectedPath + "\\1基础信息.xlsx";
                    string paraminfosheet = fd.SelectedPath + "\\2标准及参数.xlsx";
                    string reportinfosheet = fd.SelectedPath + "\\3项目.xlsx";
                    string wcdatainfofolder = fd.SelectedPath + "\\原始弯沉数据";
                    string datalistfolder = fd.SelectedPath + "\\原始工程数据";
                    string xlsxfolder = fd.SelectedPath + "\\车道报表数据";
                    string mapimgfolder = fd.SelectedPath + "\\地理位置示意图";

                    if (!File.Exists(roadinfosheet))
                        MessageBox.Show("没找到文件：" + roadinfosheet + "，请检查！");
                    else if (!File.Exists(paraminfosheet))
                        MessageBox.Show("没找到文件：" + paraminfosheet + "，请检查！");
                    else if (!File.Exists(reportinfosheet))
                        MessageBox.Show("没找到文件：" + reportinfosheet + "，请检查！");
                    else if (!Directory.Exists(datalistfolder))
                        MessageBox.Show("没找到文件夹：" + datalistfolder + "，请检查！");
                    else if (!Directory.Exists(mapimgfolder))
                        MessageBox.Show("没找到文件夹：" + mapimgfolder + "，请检查！");
                    else
                    {
                        this.Cursor = Cursors.WaitCursor;
                        chkres = CheckData(roadinfosheet, paraminfosheet, reportinfosheet, wcdatainfofolder,
                            datalistfolder, xlsxfolder, mapimgfolder, ref srcProjectinfoList);
                        this.Cursor = Cursors.Default;
                    }

                    // 导入的数据完整性和合法性检查都没有问题，开始生成各个车道工程的报表文件
                    if (chkres)
                    {
                        MessageBox.Show("导入清单文件成功！");
                    }
                    else
                    {
                        _ReportProject = null;
                        MessageBox.Show("导入清单文件失败！请和下一个弹窗信息数据进行核对！");
                        return;
                    }

                    选择导出报告 chkbox = new 选择导出报告(srcProjectinfoList);
                    chkbox.ShowDialog();

                    if (!chkres)
                    {
                        chkbox.m_ischek = false;
                        _ReportProject = null;
                    }

                    if (chkbox.m_ischek)
                    {
                        _ExcelListFilePath = xlsxfolder;
                        m_checked_reportID = chkbox.m_reportID;
                        m_checked_prjID = chkbox.m_prjID;
                        MessageBox.Show("选中了检测报告ID=" + m_checked_prjID + "！");
                    }
                    else
                    {
                        _ReportProject = null;
                        m_checked_reportID = null;
                        m_checked_prjID = null;
                        MessageBox.Show("没有选择要导出的检测报告！");
                    }

                    foreach (ProjectProjectClass tprj in srcProjectinfoList)
                    {
                        if (tprj.m_project.m_id == m_checked_prjID)
                        {
                            _ProjectProject = new ProjectProjectClass();
                            _ProjectProject = ProjectProjectClass.DeepCopyByBinary(tprj);
                            foreach (ReportProjectClass treport in tprj.m_reportlist)
                            {
                                if (treport.m_report.m_id == m_checked_reportID)
                                {
                                    _ReportProject = new ReportProjectClass();
                                    _ReportProject = ReportProjectClass.DeepCopyByBinary(treport);
                                }
                            }
                        }
                    }
                }
                else
                {
                    _ReportProject = null;
                }
            }
            else
            {
                _ReportProject = null;
            }
        }

        private void ReadWcData(FileInfo wcFile, ref WcDataClass data)
        {
            MSExcel.Application excelApp = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            int userow = 0;
            object[,] srcobj = null;
            excelApp = new MSExcel.Application() { Visible = false, DisplayAlerts = false, AlertBeforeOverwriting = false };
            excelApp.DisplayAlerts = false; // 禁用所有提示
            srcbook = excelApp.Workbooks.Open(wcFile.FullName, Type.Missing,
              true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
              Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            int generation = 0;

            try
            {
                srcsheet = srcbook.Sheets["总体信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                string msg = $"【{wcFile.Name}】中不存在【总体信息】表单，请检查！";
                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);

            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A1:L" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            try
            {
                data.traffic = srcobj[4, 2].ToString();

                DateTime dateTime = DateTime.FromOADate(double.Parse(srcobj[1, 5].ToString()));

                string dateString = dateTime.ToString("yyyy/MM/dd");
                data.time = dateString;
            }
            catch (Exception)
            {
                string msg = $"【{wcFile.Name}】【总体信息】表单中 指定位置无法获得交通量等级数据，请检查！";
                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);
            }

            try
            {
                srcsheet = srcbook.Sheets["单元信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                string msg = $"【{wcFile.Name}】中不存在【单元信息】表单，请检查！";
                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);
            }
            int errorRow = 0;
            try
            {
                userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                srcrange = srcsheet.get_Range("A3:T" + userow.ToString());
                srcobj = (object[,])srcrange.Value2;
                data.unitDatas.Columns.Add("sMile", typeof(int));
                data.unitDatas.Columns.Add("eMile", typeof(int));
                data.unitDatas.Columns.Add("保证率系数", typeof(double));
                data.unitDatas.Columns.Add("基层类型", typeof(string));
                userow -= 1;
                for (int i = 1; i < userow; i++)
                {
                    errorRow = i;
                    DataRow row = data.unitDatas.NewRow();
                    row[0] = int.Parse(srcobj[i, 2].ToString());
                    row[1] = int.Parse(srcobj[i, 3].ToString());
                    row[2] = double.Parse(srcobj[i, 8].ToString());
                    row[3] = srcobj[i, 9].ToString();
                    data.unitDatas.Rows.Add(row);
                }

            }
            catch (Exception ex)
            {
                string msg = $"【{wcFile.Name}】中【单元信息】表单，第{errorRow}行数据解析错误 请检查。\n 详细错误信息为 {ex.Message}";
                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);
            }

            try
            {
                srcsheet = srcbook.Sheets["弯沉"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                string msg = $"【{wcFile.Name}】中不存在【弯沉】表单，请检查！";
                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);
            }
            int direction = 1;
            try
            {
                userow = GlobalExcel.judegeusedrow(srcsheet, 1, 1);
                srcrange = srcsheet.get_Range("A2:AD" + userow.ToString());
                srcobj = (object[,])srcrange.Value2;

                data.wcDatas.Columns.Add("序号", typeof(int));
                data.wcDatas.Columns.Add("方向", typeof(string));
                data.wcDatas.Columns.Add("车道编号", typeof(int));
                data.wcDatas.Columns.Add("Mile", typeof(int));
                data.wcDatas.Columns.Add("弯沉值", typeof(double));


                for (int i = 1; i < userow; i++)
                {
                    errorRow = i;
                    DataRow row = data.wcDatas.NewRow();
                    row[0] = int.Parse(srcobj[i, 1].ToString());
                    row[1] = srcobj[i, 2].ToString();
                    if (srcobj[i, 2].ToString() == "下行")
                    {
                        //direction = -1;
                    }
                    row[2] = int.Parse(srcobj[i, 3].ToString());
                    row[3] = int.Parse(srcobj[i, 4].ToString());
                    row[4] = double.Parse(srcobj[i, 30].ToString());
                    data.wcDatas.Rows.Add(row);
                }

            }
            catch (Exception ex)
            {
                string msg = $"【{wcFile.Name}】中【弯沉】表单，第{errorRow}行数据解析错误 请检查。\n 详细错误信息为 {ex.Message}";

                MessageBox.Show(msg);
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                throw new Exception(msg);
            }

            srcbook.Save();
            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            //计算需要输出得表格数据
            //根据单元信息得到 结果数据行数
            data.wcResultDatas.Columns.Add("sMile", typeof(int));
            data.wcResultDatas.Columns.Add("eMile", typeof(int));
            data.wcResultDatas.Columns.Add("交通量等级", typeof(string));
            data.wcResultDatas.Columns.Add("基层类型", typeof(string));
            data.wcResultDatas.Columns.Add("弯沉值", typeof(double));
            data.wcResultDatas.Columns.Add("评价等级", typeof(string));

            double tempBzlxs = 1;

            Dictionary<DataRow, List<double>> wcSplitValues = new Dictionary<DataRow, List<double>>();

            for (int i = 0; i < data.unitDatas.Rows.Count; i++)
            {
                List<double> values = new List<double>();
                var curRow = data.unitDatas.Rows[i];
                DataRow row = data.wcResultDatas.NewRow();
                int smile = curRow["sMile"].ObjToInt();
                int emile = curRow["eMile"].ObjToInt();
                row["sMile"] = smile;
                row["eMile"] = emile;
                row["基层类型"] = curRow["基层类型"];
                row["交通量等级"] = data.traffic;
                //计算弯沉值

                double bzlxs = double.Parse(curRow["保证率系数"].ToString());
                tempBzlxs = bzlxs;

                for (int t = 0; t < data.wcDatas.Rows.Count; t++)
                {
                    var ccurRow = data.wcDatas.Rows[t];
                    int csmile = ccurRow["Mile"].ObjToInt();
                    double value = double.Parse(ccurRow["弯沉值"].ToString());

                    if (direction * smile <= csmile * direction && csmile < direction * emile)
                    {
                        values.Add(value);
                    }
                    if (t == data.wcDatas.Rows.Count-1)
                    {
                        if (csmile == emile)
                        {
                            values.Add(value);
                        }
                    }

                }
                wcSplitValues.Add(curRow, values);
                if (values.Count==0)
                {
                    row["弯沉值"] = double.NaN;
                    row["评价等级"] = "/";
                    // MessageBox.Show("【" + wcFile.Name + "】中【单元信息】表单，第" + i + "行数据，没有找到对应弯沉数据，请检查！");
                }
                else if (values.Count==1)
                {
                    row["弯沉值"] = values[0];

                    //求评价等级
                    row["评价等级"] = shhpGetJudgeStr(data.traffic, row["基层类型"].ToString(), values[0]);
                }
                else
                {
                    double sdValue = CalculateStandardDeviation(values);
                    double result = values.Average() + bzlxs * sdValue;
                    row["弯沉值"] = result;

                    //求评价等级
                    row["评价等级"] = shhpGetJudgeStr(data.traffic, row["基层类型"].ToString(), result);

                }

                data.wcResultDatas.Rows.Add(row);


            }
            //求全车道弯沉值
            List<double> wcAllTemps = new List<double>();
            for (int t = 0; t < data.wcDatas.Rows.Count; t++)
            {
                var ccurRow = data.wcDatas.Rows[t];
                int csmile = ccurRow["Mile"].ObjToInt();
                double value = double.Parse(ccurRow["弯沉值"].ToString());
                wcAllTemps.Add(value);


            }

            //获得最长的基层类型
            string maxLengthJclx = "";
            Dictionary<string, int> jclxLenList = new Dictionary<string, int>();
            for (int i = 0; i < data.unitDatas.Rows.Count; i++)
            {
                var curRow = data.unitDatas.Rows[i];
                DataRow row = data.wcResultDatas.NewRow();
                int smile = curRow["sMile"].ObjToInt();
                int emile = curRow["eMile"].ObjToInt();
                row["sMile"] = smile;
                row["eMile"] = emile;
                int len = Math.Abs(smile - emile);
                string jclxTemp = curRow["基层类型"].ToString();
                if (jclxLenList.Keys.Contains(jclxTemp))
                {
                    jclxLenList[jclxTemp] += len;
                }
                else
                {
                    jclxLenList.Add(jclxTemp, len);
                }
            }
            // 将字典转换为KeyValuePair列表
            List<KeyValuePair<string, int>> list = jclxLenList.ToList();

            // 使用LINQ根据值进行排序
            list.Sort((pair1, pair2) => pair2.Value.CompareTo(pair1.Value));

            maxLengthJclx = list.First().Key;

            if (list.Count > 1)
            {
                data.WcLjlx = "/";
            }
            else
            {
                data.WcLjlx = maxLengthJclx;
            }

            //求每个区间的弯沉值   key 道路长度  value 对应区间弯沉值
            //Dictionary<int, double> wcDataValues = new Dictionary<int, double>();
            List<(int, double)> wcDataValues = new List<(int, double)>();
            foreach (var item in wcSplitValues)
            {
                double sdValueTemp = CalculateStandardDeviation(item.Value);
                double bzlxsTemp = double.Parse(item.Key["保证率系数"].ToString());
                if (item.Value.Count == 1)
                {
                    int smile = item.Key["sMile"].ObjToInt();
                    int emile = item.Key["eMile"].ObjToInt();
                    int len = Math.Abs(smile - emile);
                    wcDataValues.Add((len, item.Value[0]));
                }
                else if (item.Value.Count>0)
                {
                    double resultTemp0 = item.Value.Average() + tempBzlxs * sdValueTemp;
                    int smile = item.Key["sMile"].ObjToInt();
                    int emile = item.Key["eMile"].ObjToInt();
                    int len = Math.Abs(smile - emile);
                    wcDataValues.Add((len, resultTemp0));
                } 
                else
                {

                }
            }
            //加权平均获得整个路段弯沉值
            double roadLength = 0;
            double sumTempValues = 0;
            foreach (var item in wcDataValues)
            {
                roadLength += item.Item1;
                sumTempValues += (item.Item1 * item.Item2);
            }
            data.WcValue = sumTempValues / roadLength;
            data.WcJudge = shhpGetJudgeStr(data.traffic, maxLengthJclx, data.WcValue);
            data.wcLength = roadLength;
        }

        // 创建评价表
        private static Dictionary<string, Dictionary<string, List<EvaluationRange>>> evaluationTable
             = new Dictionary<string, Dictionary<string, List<EvaluationRange>>>();

        /// <summary>
        /// 上海惠普获得评价等级
        /// </summary>
        /// <param name="jtldj">交通量等级</param>
        /// <param name="jclx">基层类型</param>
        /// <param name="cxValue">弯沉值</param>
        /// <returns></returns>
        private static string shhpGetJudgeStr(string jtldj, string jclx, double cxValue)
        {
          
            initHpJudegeDictionary();
            // 检查表中是否存在指定的基层类型和交通量等级
            if (evaluationTable.ContainsKey(jclx) && evaluationTable[jclx].ContainsKey(jtldj))
            {
               // string evaluation = "";
                // 遍历对应的评价区间
                foreach (var range in evaluationTable[jclx][jtldj])
                {
                    // 检查值是否在当前区间内
                    if (cxValue >= range.MinValue && cxValue <= range.MaxValue)
                    {
                        return range.Evaluation; // 返回评价结果
                    }
                }
            }
            else
            {
                throw new Exception("请检查沉陷文件\n【基层类型(粒料及沥青稳定|半刚性)】【交通量等级(很轻|轻|中|重|特重)】\n是否填写正确!");
            }

            return "未找到";


        }
        // 定义评价区间类
        class EvaluationRange
        {
            public double MinValue { get; set; }
            public double MaxValue { get; set; }
            public string Evaluation { get; set; }
        }

        private static void initHpJudegeDictionary()
        {
            if (evaluationTable.Count == 0)
            {
                // 添加基层类型和交通量等级的组合及其对应的评价区间
                evaluationTable["粒料及沥青稳定"] = new Dictionary<string, List<EvaluationRange>>
                {
                     {
                     "很轻", new List<EvaluationRange>
                         {
                             new EvaluationRange { MinValue = -10000, MaxValue = 98, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 98, MaxValue = 126, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 126, MaxValue = 10000, Evaluation = "不足" },

                         }
                     },
                    {
                    "轻", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -1000, MaxValue = 77, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 77, MaxValue = 98, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 98, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "中", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 60, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 60, MaxValue = 81, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 81, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                      {
                    "重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 46, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 46, MaxValue = 67, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 67, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "特重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 35, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 35, MaxValue = 56, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 56, MaxValue = 10000, Evaluation = "不足" },
                         }
                    }

                 };

                // 添加基层类型和交通量等级的组合及其对应的评价区间
                evaluationTable["半刚性"] = new Dictionary<string, List<EvaluationRange>>
                {
                     {
                     "很轻", new List<EvaluationRange>
                         {
                             new EvaluationRange { MinValue = -10000, MaxValue = 77, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 77, MaxValue = 98, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 98, MaxValue = 10000, Evaluation = "不足" },

                         }
                     },
                    {
                    "轻", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -1000, MaxValue = 56, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 56, MaxValue = 77, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 77, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "中", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 42, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 42, MaxValue =59, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 59, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                      {
                    "重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 31, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 31, MaxValue = 46, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 46, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "特重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 21, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 21, MaxValue = 35, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 35, MaxValue = 10000, Evaluation = "不足" },
                         }
                    }

                 };
            }
        }
        /// <summary>
        /// 求标准差  Excel的=STDEV方法
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private double CalculateStandardDeviation(List<double> arrData)
        {
            double xSum = 0;
            double xAvg = 0;
            double sSum = 0;
            double tmpStDev = 0;
            int arrNum = arrData.Count;
            for (int i = 0; i < arrNum; i++)
            {
                xSum += arrData[i];
            }
            xAvg = xSum / arrNum;
            for (int j = 0; j < arrNum; j++)
            {
                sSum += ((arrData[j] - xAvg) * (arrData[j] - xAvg));
            }
            tmpStDev = Math.Sqrt(sSum / (arrNum - 1));
            return tmpStDev;
        }

        private bool CheckData(string roadinfosheet, string paraminfosheet, string reportinfosheet, string wcdatainfoFolder,
            string datalistfolder, string xlsxfolder, string mapimgfolder, ref List<ProjectProjectClass> projectinfolist)
        {
            bool res = true;

            MSExcel.Application excelApp = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            object[,] srcobj = null;

            int tidx = 0;
            int userow = 0;
            Dictionary<string, RoadInfoClass> roadinfos = new Dictionary<string, RoadInfoClass>();
            Dictionary<string, RoadPartInfoClass> roadpartinfos = new Dictionary<string, RoadPartInfoClass>();
            Dictionary<string, LaneInfoClass> laneinfos = new Dictionary<string, LaneInfoClass>();
            Dictionary<string, ProjectInfoClass> projectinfos = new Dictionary<string, ProjectInfoClass>();
            Dictionary<string, ReportInfoClass> reportinfos = new Dictionary<string, ReportInfoClass>();
            Dictionary<string, ReoprtRoadPartInfoClass> reportroadpartinfos = new Dictionary<string, ReoprtRoadPartInfoClass>();
            Dictionary<string, TestingStardardClass> standardinfos = new Dictionary<string, TestingStardardClass>();


            List<IndexInfoClass> indexinfolist = new List<IndexInfoClass>();
            List<TestingPersonClass> personlist = new List<TestingPersonClass>();

            excelApp = new MSExcel.Application() { Visible = false, DisplayAlerts = false, AlertBeforeOverwriting = false };
            int generation = 0;

            // 读取标准及参数            
            srcbook = null;
            srcsheet = null;
            srcbook = excelApp.Workbooks.Open(paraminfosheet, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            try
            {
                srcsheet = srcbook.Sheets["检测标准"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【2标准及参数.xlsx】中不存在【检测标准】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:AH" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                TestingStardardClass tinfo = new TestingStardardClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_code = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_function = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_type = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_industry = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_dependency = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_remarks = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;

                try
                {
                    standardinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【2标准及参数.xlsx】的【检测标准】表单中存在多个相同的检测标准ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }
            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);

            //1、读取所有的道路信息数据
            srcbook = null;
            srcsheet = null;
            srcbook = excelApp.Workbooks.Open(roadinfosheet, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            try
            {
                srcsheet = srcbook.Sheets["道路信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【1基础信息.xlsx】中不存在【道路信息】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:AF" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                RoadInfoClass tinfo = new RoadInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_province = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_city = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_district = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_town = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_village = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_code = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_properity = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_grade = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_length = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_width = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadway_area = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_sidewalk_area = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadtype = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadstartlocation = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadendlocation = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadsartmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_roadendmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_buildyear = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_buildunit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_designunit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_constructionunit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_controlunit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_managementunit_province = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_managementunit_city = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_managementunit_district = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_managementunit_department = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_maintenance_center = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_maintenance_section = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_maintenance_unit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_project_department = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;

                try
                {
                    roadinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【1基础信息.xlsx】的【道路信息】表单中存在多个相同的道路信息ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }

            //2、读取所有的路段信息数据
            try
            {
                srcsheet = srcbook.Sheets["路段信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【1基础信息.xlsx】中不存在【路段信息】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 2, 2);
            srcrange = srcsheet.get_Range("A2:P" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                RoadPartInfoClass tinfo = new RoadPartInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                string troadid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_roadinfo = RoadInfoClass.DeepCopyByBinary(roadinfos[troadid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【路段信息】表单中路段ID=" + tinfo.m_id + "的路段所对应的道路信息ID=" + troadid + "在【道路信息】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
                tidx = 8;
                tinfo.m_startlocation = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_endlocation = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_startmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_endmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_part_grade = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_length = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_width = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_area = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_type = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;

                try
                {
                    roadpartinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【1基础信息.xlsx】的【路段信息】表单中存在多个相同的路段信息ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }

            //3、读取所有的车道数据
            try
            {
                srcsheet = srcbook.Sheets["车道信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【1基础信息.xlsx】中不存在【车道信息】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:M" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                LaneInfoClass tinfo = new LaneInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                string troadid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_roadpartinfo = RoadPartInfoClass.DeepCopyByBinary(roadpartinfos[troadid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【车道信息】表单中车道ID=" + tinfo.m_id + "的车道所对应的路段信息ID=" + troadid + "在【路段信息】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
                tidx = 6;
                tinfo.m_roadfunctiontype = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_direction = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_lanenum = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_width = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_pavementtype = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_startmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_endmile = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_carwaytype = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                if (tinfo.m_width == string.Empty)
                {
                    res = false;
                    MessageBox.Show("【车道信息】表单中车道ID=" + tinfo.m_id + "的车道所对应车道宽度数值为空，请检查！");
                }
                try
                {
                    laneinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【1基础信息.xlsx】的【车道信息】表单中存在多个相同的车道信息ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }
            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);

            srcbook = null;
            srcsheet = null;
            srcbook = excelApp.Workbooks.Open(reportinfosheet, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //4、读取 项目 数据
            try
            {
                srcsheet = srcbook.Sheets["检测参数"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【3项目.xlsx】中不存在【检测参数】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:G" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;

            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                IndexInfoClass tinfo = new IndexInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_projectid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_standardid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_index = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_pavementtype = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_tesing = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                indexinfolist.Add(tinfo);
            }
            try
            {
                srcsheet = srcbook.Sheets["检测人员"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【3项目.xlsx】中不存在【检测人员】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 2, 2);
            srcrange = srcsheet.get_Range("B2:F" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                TestingPersonClass tinfo = new TestingPersonClass();
                tinfo.m_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_CertificateNo = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_title = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_post = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_duty = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                personlist.Add(tinfo);
            }

            try
            {
                srcsheet = srcbook.Sheets["项目信息"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【3项目.xlsx】中不存在【项目信息】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:L" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                ProjectInfoClass tinfo = new ProjectInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_project_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_entrust_client = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_entrust_serial = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_contract_num = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_entrust_date = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                if (tinfo.m_entrust_date.Length < 8 && tinfo.m_entrust_date.Length > 0)
                {
                    System.DateTime dt = new System.DateTime(1900, 1, 1);
                    dt = dt.AddDays(Convert.ToInt64(tinfo.m_entrust_date) - 2);
                    tinfo.m_entrust_date = dt.ToString("yyyyMMdd");
                }

                tinfo.m_testing_unit = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_project_dutyperson = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_testing_start_date = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                if (tinfo.m_testing_start_date.Length < 8 && tinfo.m_testing_start_date.Length > 0)
                {
                    System.DateTime dt = new System.DateTime(1900, 1, 1);
                    dt = dt.AddDays(Convert.ToInt64(tinfo.m_testing_start_date) - 2);
                    tinfo.m_testing_start_date = dt.ToString("yyyyMMdd");
                }

                tinfo.m_testing_end_date = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                if (tinfo.m_testing_end_date.Length < 8 && tinfo.m_testing_end_date.Length > 0)
                {
                    System.DateTime dt = new System.DateTime(1900, 1, 1);
                    dt = dt.AddDays(Convert.ToInt64(tinfo.m_testing_end_date) - 2);
                    tinfo.m_testing_end_date = dt.ToString("yyyyMMdd");
                }

                string tstandardid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_testing_standard = TestingStardardClass.DeepCopyByBinary(standardinfos[tstandardid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【项目信息】表单中项目ID=" + tinfo.m_id + "所对应的检测标准ID=" + tstandardid + "在【2标准及参数.xlsx】的【检测标准】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }

                tinfo.m_date = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                if (tinfo.m_date.Length < 8 && tinfo.m_date.Length > 0)
                {
                    System.DateTime dt = new System.DateTime(1900, 1, 1);
                    dt = dt.AddDays(Convert.ToInt64(tinfo.m_date) - 2);
                    tinfo.m_date = dt.ToString("yyyyMMdd");
                }

                foreach (IndexInfoClass tindx in indexinfolist)
                {
                    if (tindx.m_projectid == tinfo.m_id)
                    {
                        tinfo.m_indexlist.Add(tindx);
                    }
                }

                try
                {
                    projectinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【项目信息】表单中存在多个相同的项目信息ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }

            //5、读取 报告 数据
            try
            {
                srcsheet = srcbook.Sheets["检测报告"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【3项目.xlsx】中不存在【检测报告】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:E" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                ReportInfoClass tinfo = new ReportInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                string troadid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_projectinfo = ProjectInfoClass.DeepCopyByBinary(projectinfos[troadid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【检测报告】表单中检测报告ID=" + tinfo.m_id + "所对应的项目ID=" + troadid + "在【项目信息】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
                tinfo.m_report_num = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_report_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                tinfo.m_project_name = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    reportinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【检测报告】表单中存在多个相同的报告信息ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }

            //6、读取 报告检测路段 中的 项目总体信息
            try
            {
                srcsheet = srcbook.Sheets["报告检测路段"] as MSExcel.Worksheet;
            }
            catch (Exception)
            {
                MessageBox.Show("【3项目.xlsx】中不存在【报告检测路段】表单，请检查！");
                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                return false;
            }
            userow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
            srcrange = srcsheet.get_Range("A2:E" + userow.ToString());
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < userow; ++i)
            {
                tidx = 1;
                ReoprtRoadPartInfoClass tinfo = new ReoprtRoadPartInfoClass();
                tinfo.m_id = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                string troadid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_reoprt = ReportInfoClass.DeepCopyByBinary(reportinfos[troadid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【报告检测路段】表单中报告检测路段ID=" + tinfo.m_id + "所对应的检测报告ID=" + troadid + "在【检测报告】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }

                tidx = 5;
                troadid = srcobj[i, tidx] == null ? string.Empty : Convert.ToString(srcobj[i, tidx]); ++tidx;
                try
                {
                    tinfo.m_roadpart = RoadPartInfoClass.DeepCopyByBinary(roadpartinfos[troadid]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【报告检测路段】表单中报告检测路段ID=" + tinfo.m_id + "所对应的路段信息ID=" + troadid + "在【1基础信息.xlsx】的【路段信息】表单中没有找到！\r\n请检查导入的数据是否完整！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }

                try
                {
                    reportroadpartinfos.Add(tinfo.m_id, tinfo);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("【3项目.xlsx】的【报告检测路段】表单中存在多个相同的报告检测路段ID=" + tinfo.m_id + "请检查！");
                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    excelApp.Quit();
                    return false;
                }
            }

            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            excelApp.Quit();

            //7、将以上信息整理成链表：项目-报告-路段-车道
            //整理 路段-车道 链表
            Dictionary<string, RoadPartProjectClass> roadpartinfolist = new Dictionary<string, RoadPartProjectClass>();
            foreach (KeyValuePair<string, RoadPartInfoClass> kvp_roadpartinfo in roadpartinfos)
            {
                RoadPartProjectClass tlist = new RoadPartProjectClass();
                tlist.m_roadpart = RoadPartInfoClass.DeepCopyByBinary(kvp_roadpartinfo.Value);
                foreach (KeyValuePair<string, LaneInfoClass> kvp_laneinfo in laneinfos)
                {

                    if (kvp_laneinfo.Value.m_roadpartinfo.m_id == kvp_roadpartinfo.Value.m_id)
                    {


                        LaneProjectClass tlane = new LaneProjectClass();
                        tlane.m_lane = LaneProjectClass.DeepCopyByBinary(kvp_laneinfo.Value);

                        tlane.m_projectdatapathlist.Clear();
                        tlist.m_lanelist.Add(tlane);
                    }

                }
                roadpartinfolist.Add(kvp_roadpartinfo.Value.m_id, tlist);
            }


            // 整理 项目-报告 链表
            // 先清空
            projectinfolist.Clear();
            foreach (KeyValuePair<string, ProjectInfoClass> kvp_projectinfo in projectinfos)
            {
                ProjectProjectClass tproject = new ProjectProjectClass();
                tproject.m_project = ProjectInfoClass.DeepCopyByBinary(kvp_projectinfo.Value);
                tproject.m_reportlist = new List<ReportProjectClass>();

                foreach (KeyValuePair<string, ReportInfoClass> kvp_reportinfo in reportinfos)
                {
                    if (kvp_reportinfo.Value.m_projectinfo.m_id == kvp_projectinfo.Value.m_id)
                    {
                        ReportProjectClass treportinfo = new ReportProjectClass();
                        treportinfo.m_personList.AddRange(personlist);
                        treportinfo.m_report = ReportInfoClass.DeepCopyByBinary(kvp_reportinfo.Value);
                        treportinfo.m_roadpartlist = new List<RoadPartProjectClass>();

                        foreach (KeyValuePair<string, ReoprtRoadPartInfoClass> kvp_reportroadpartinfo in reportroadpartinfos)
                        {
                            if (kvp_reportroadpartinfo.Value.m_reoprt.m_id == kvp_reportinfo.Value.m_id)
                            {
                                RoadPartProjectClass troadpart = new RoadPartProjectClass();
                                troadpart = roadpartinfolist[kvp_reportroadpartinfo.Value.m_roadpart.m_id];
                                treportinfo.m_roadpartlist.Add(troadpart);
                            }
                        }

                        tproject.m_reportlist.Add(treportinfo);
                    }
                }
                projectinfolist.Add(tproject);
            }

            //5、检查 报告生成路段清单.xlsx 的 车道清单 sheet页中所列出的所有车道原始工程数据是否都完整
            DirectoryInfo srcdir = new DirectoryInfo(datalistfolder);
            DirectoryInfo[] prjdirs = srcdir.GetDirectories();
            foreach (DirectoryInfo tdir in prjdirs)
            {
                bool isfinddir = false;
                for (int i1 = 0; i1 < projectinfolist.Count; ++i1)
                {
                    for (int i2 = 0; i2 < projectinfolist[i1].m_reportlist.Count; ++i2)
                    {
                        for (int i3 = 0; i3 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist.Count; ++i3)
                        {
                            string lanelayout = null;
                            bool updir = false;
                            bool downdir = false;

                            for (int i4 = 0; i4 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist.Count; ++i4)
                            {
                                projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report
                                    = ReportInfoClass.DeepCopyByBinary(projectinfolist[i1].m_reportlist[i2].m_report);

                                LaneInfoClass tlane = projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_lane;
                                if (tlane.m_direction.Contains("上行")) updir = true;
                                if (tlane.m_direction.Contains("下行")) downdir = true;
                                //else downdir = true;

                                if (tdir.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_code)
                                       && tdir.Name.Contains("_" + tlane.m_direction + "_")
                                       && tdir.Name.Contains("_" + tlane.m_lanenum + "_"))
                                {
                                    if (tdir.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_name))
                                    {
                                        if (tdir.Name.Contains(tlane.m_roadpartinfo.m_startlocation))
                                        {
                                            if (tdir.Name.Contains(tlane.m_roadpartinfo.m_endlocation))
                                            {
                                                isfinddir = true;
                                                projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_projectdatapathlist.Add(tdir.FullName);

                                            }
                                        }
                                    }


                                }
                                // if (isfinddir) break;
                            }
                            // 车道布置的信息
                            if (updir & downdir) lanelayout = "双向";
                            else lanelayout = "单向";
                            lanelayout = lanelayout + _LaneLayoutStr[projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist.Count] + "车道";
                            projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_LaneLayout = lanelayout;

                            if (isfinddir) break;
                        }
                        if (isfinddir) break;
                    }
                    if (isfinddir) break;
                }
                if (!isfinddir)
                {
                    res = false;
                    MessageBox.Show("文件夹【" + tdir.Name + "】\r\n没找到相关联的车道信息，请检查文件夹名是否合法及车道信息是否完整！");
                }
            }

            //5、检查 报告生成路段清单.xlsx 的 车道清单 sheet页中所列出的所有车道原始工程数据是否都完整 
            if (Directory.Exists(wcdatainfoFolder))
            {
                DirectoryInfo wcdir = new DirectoryInfo(wcdatainfoFolder);
                var temp = wcdir.GetFiles("*.xlsx", SearchOption.AllDirectories).Where(
                   t => !t.Name.Contains("~")
                   );
                foreach (FileInfo wc in temp)
                {
                    bool isfinddir = false;
                    for (int i1 = 0; i1 < projectinfolist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < projectinfolist[i1].m_reportlist.Count; ++i2)
                        {
                            for (int i3 = 0; i3 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist.Count; ++i3)
                            {
                                for (int i4 = 0; i4 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist.Count; ++i4)
                                {
                                    /*  projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report
                                          = ReportInfoClass.DeepCopyByBinary(projectinfolist[i1].m_reportlist[i2].m_report);
                                    */
                                    LaneInfoClass tlane = projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_lane;

                                    //else downdir = true;

                                    if (wc.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_code)
                                           && wc.Name.Contains("_" + tlane.m_direction + "_")
                                           && wc.Name.Contains("_" + tlane.m_lanenum))
                                    {
                                        if (wc.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_name))
                                        {
                                            if (wc.Name.Contains(tlane.m_roadpartinfo.m_startlocation))
                                            {
                                                if (wc.Name.Contains(tlane.m_roadpartinfo.m_endlocation))
                                                {
                                                    isfinddir = true;
                                                    if (projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_projectdatapathlist.Count > 0)
                                                    {

                                                        projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_projectwcDataFilePath.Add(wc);

                                                        //弯沉信息解析
                                                        WcDataClass wcData = new WcDataClass();
                                                        try
                                                        {
                                                            ReadWcData(wc, ref wcData);


                                                            projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_wcDataClasses.Add(wcData);
                                                        }
                                                        catch (Exception ex)
                                                        {

                                                            throw ex;
                                                        }
                                                    }



                                                }
                                            }
                                        }


                                    }
                                    // if (isfinddir) break;
                                }
                                if (isfinddir) break;
                            }
                            if (isfinddir) break;
                        }
                        if (isfinddir) break;
                    }
                    if (!isfinddir)
                    {
                        res = false;
                        MessageBox.Show("文件夹【" + wc.Name + "】\r\n没找到相关联的弯沉数据，请检查文件夹名是否合法及车道信息是否完整！");
                    }
                }
            }



            // 将地理位置示意图 和 路段信息关联起来
            if (Directory.Exists(mapimgfolder))
            {
                DirectoryInfo mapimgdir = new DirectoryInfo(mapimgfolder);
                FileInfo[] mapimgfiles = mapimgdir.GetFiles("*.jpg");
                foreach (FileInfo tfile in mapimgfiles)
                {
                    bool isfindfile = false;
                    for (int i1 = 0; i1 < projectinfolist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < projectinfolist[i1].m_reportlist.Count; ++i2)
                        {
                            for (int i3 = 0; i3 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist.Count; ++i3)
                            {
                                RoadPartProjectClass troadpart = projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3];
                                if (tfile.Name.Contains(troadpart.m_roadpart.m_roadinfo.m_code)
                                       && tfile.Name.Contains(troadpart.m_roadpart.m_roadinfo.m_name)
                                       && tfile.Name.Contains(troadpart.m_roadpart.m_startlocation)
                                       && tfile.Name.Contains(troadpart.m_roadpart.m_endlocation))
                                {
                                    isfindfile = true;
                                    projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_MapImg = tfile.FullName;
                                }
                                if (isfindfile) break;
                            }
                            if (isfindfile) break;
                        }
                        if (isfindfile) break;
                    }
                }
            }

            //6、将xlsx报表和车道信息关联起来
            if (Directory.Exists(xlsxfolder))
            {
                DirectoryInfo xlsxdir = new DirectoryInfo(xlsxfolder);
                FileInfo[] xlsxfiles = xlsxdir.GetFiles("*.xlsx");
                foreach (FileInfo tfile in xlsxfiles)
                {
                    bool isfindfile = false;
                    for (int i1 = 0; i1 < projectinfolist.Count; ++i1)
                    {
                        for (int i2 = 0; i2 < projectinfolist[i1].m_reportlist.Count; ++i2)
                        {
                            for (int i3 = 0; i3 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist.Count; ++i3)
                            {
                                for (int i4 = 0; i4 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist.Count; ++i4)
                                {
                                    projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report
                                        = ReportInfoClass.DeepCopyByBinary(projectinfolist[i1].m_reportlist[i2].m_report);

                                    LaneInfoClass tlane = projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_lane;
                                    if (tfile.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_code)
                                           && tfile.Name.Contains(tlane.m_roadpartinfo.m_roadinfo.m_name)
                                           && tfile.Name.Contains(tlane.m_roadpartinfo.m_startlocation)
                                           && tfile.Name.Contains(tlane.m_roadpartinfo.m_endlocation)
                                           && tfile.Name.Contains("_" + tlane.m_direction + "_")
                                           && tfile.Name.Contains("_" + tlane.m_lanenum + "_"))
                                    {
                                        isfindfile = true;
                                        projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_xlsxpath = tfile.FullName;
                                    }
                                    if (isfindfile) break;
                                }
                                if (isfindfile) break;
                            }
                            if (isfindfile) break;
                        }
                        if (isfindfile) break;
                    }
                    if (!isfindfile)
                    {
                        res = false;
                        MessageBox.Show("报表【" + tfile.Name + "】\r\n没找到相关联的车道信息，请检查文件夹名是否合法及车道信息是否完整！");
                    }
                }
            }

            // 得到报告的 起始 和 终止 日期
            for (int i1 = 0; i1 < projectinfolist.Count; ++i1)
            {
                for (int i2 = 0; i2 < projectinfolist[i1].m_reportlist.Count; ++i2)
                {
                    for (int i3 = 0; i3 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist.Count; ++i3)
                    {
                        for (int i4 = 0; i4 < projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist.Count; ++i4)
                        {
                            string startstr = "99990101";
                            string endstr = "19000101";
                            foreach (string tpath in projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_projectdatapathlist)
                            {
                                DirectoryInfo tdirinfo = new DirectoryInfo(tpath);
                                SingleProject proj = new SingleProject(tdirinfo);
                                string tstr = proj._ProjectInfo._DataDate;
                                if (string.Compare(startstr, tstr) > 0)
                                {
                                    startstr = tstr;
                                }
                                if (string.Compare(endstr, tstr) < 0)
                                {
                                    endstr = tstr;
                                }
                            }
                            projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report
                                = ReportInfoClass.DeepCopyByBinary(projectinfolist[i1].m_reportlist[i2].m_report);
                            projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report.m_report_start_date = startstr;
                            projectinfolist[i1].m_reportlist[i2].m_roadpartlist[i3].m_lanelist[i4].m_report.m_report_end_date = endstr;
                        }
                    }
                }
            }

            return res;
        }

        private void barButtonItem22_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.CityRoad && _Setting.ExcelType == 7)
            {
                if (MyExcelCity._PQIGrade == null)
                {
                    MyExcelCity.LoadXlsParm();
                }
                _ReportProject = null;
                SHPG_ExportWord_LoadData();

                if (_ReportProject != null)
                {
                    ReportProjectClass treport = new ReportProjectClass();
                    treport = ReportProjectClass.DeepCopyByBinary(_ReportProject);
                    for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
                    {
                        //for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                        //{
                        //    if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath == null)
                        //    {
                        //        treport.m_roadpartlist[i1].m_lanelist.RemoveAt(i2);
                        //        --i2;
                        //    }
                        //}
                        if (treport.m_roadpartlist[i1].m_lanelist.Count == 0)
                        {
                            treport.m_roadpartlist.RemoveAt(i1);
                            --i1;
                        }
                    }
                    if (treport.m_roadpartlist.Count == 0)
                    {
                        MessageBox.Show("没有可导出的路段，请检查路段信息是否正确完整！");
                    }
                    else
                    {
                        MSExcel.Application excelApp = new MSExcel.Application()
                        {
                            Visible = true,
                            DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                            AlertBeforeOverwriting = false
                        };

                        bool flag = true;
                        foreach (RoadPartProjectClass troadpart in treport.m_roadpartlist)
                        {
                            foreach (LaneProjectClass tlane in troadpart.m_lanelist)
                            {
                                if (tlane.m_projectdatapathlist.Count > 0)
                                {
                                    foreach (string tpath in tlane.m_projectdatapathlist)
                                    {
                                        DirectoryInfo tdirinfo = new DirectoryInfo(tpath);
                                        SingleProject proj = new SingleProject(tdirinfo);
                                        string xlspath = tdirinfo.Parent.Parent.FullName + "\\车道报表数据";

                                        if (flag)
                                        {
                                            flag = false;
                                            if (File.Exists(tdirinfo.Parent.Parent.FullName + "\\路面材质不一致记录.txt"))
                                            {
                                                File.Delete(tdirinfo.Parent.Parent.FullName + "\\路面材质不一致记录.txt");
                                            }
                                        }

                                        if (!Directory.Exists(xlspath))
                                        {
                                            Directory.CreateDirectory(xlspath);
                                        }
                                        proj.GenerateExcel(excelApp, xlspath, null, null, tlane);
                                    }
                                }
                            }
                        }

                        int generation = System.GC.GetGeneration(excelApp);
                        System.GC.Collect(generation);//垃圾回收
                        System.GC.WaitForPendingFinalizers();
                        excelApp.Quit();
                        MessageBox.Show("导出车道报表完成！");
                    }
                }
            }
        }

        private void barButtonItem21_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_CurProject != null)
            {
                MessageBox.Show("请重新导入工程，不要点击左侧工程列表，直接进行病害转换！");
                return;
            }

            病害类型转换 FormTran = new 病害类型转换(_Projects);
            FormTran.ShowDialog();
        }

        /// <summary>
        /// 是否输出空的技术状况评定明细表
        /// </summary>
        public static bool _IsOutputEmptyExcel = false;
        private void barEditItem5_EditValueChanged(object sender, EventArgs e)
        {
            _IsOutputEmptyExcel = Convert.ToBoolean(barEditItem5.EditValue);
        }

        private void barButtonItem24_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Projects.Count < 1)
            {
                MessageBox.Show("工程列表为空！");
                return;
            }

            FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }

                MSExcel.Application excelApp = new MSExcel.Application()
                {
                    Visible = true,
                    DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                    AlertBeforeOverwriting = false
                };

                string srcxls = string.Format(@"{0}\报表模板\工程台账汇总.xlsx", System.Windows.Forms.Application.StartupPath);
                string Destxls = string.Format(@"{0}\多车道统计.xlsx", fd.SelectedPath);

                MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _Workbook.SaveAs(Destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                int cnt = 0;
                MSExcel.Worksheet destsheet = _Workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                object[,] obj = new object[_Projects.Count, 21];
                foreach (SingleProject tproject in _Projects)
                {
                    obj[cnt, 0] = tproject._ProjectInfo._Province;
                    obj[cnt, 1] = tproject._ProjectInfo._City;
                    obj[cnt, 2] = tproject._ProjectInfo._District;
                    obj[cnt, 3] = tproject._ProjectInfo._RoadCode;
                    obj[cnt, 4] = tproject._ProjectInfo._RoadName;
                    obj[cnt, 5] = tproject._ProjectInfo._Direction > 0 ? "上行" : "下行";
                    obj[cnt, 6] = tproject._ProjectInfo._RoadNum;
                    obj[cnt, 7] = tproject._ProjectInfo._StartMile;
                    obj[cnt, 8] = tproject._ProjectInfo._EndMile;
                    obj[cnt, 9] = string.Format("=ABS(H{0}-I{0})*0.001", cnt + 2);
                    //obj[cnt, 9] = tproject._ProjectInfo._EndDmi * 0.001;
                    obj[cnt, 10] = tproject._ProjectInfo._RoadGrade;
                    obj[cnt, 11] = GlobalExcel._RoadTypeStr[tproject._ProjectInfo._RoadType];
                    obj[cnt, 12] = tproject._ProjectInfo._DataDate;
                    obj[cnt, 13] = tproject._ProjectInfo._DataTime;
                    obj[cnt, 14] = tproject._ProjectInfo._DataPerson;
                    obj[cnt, 15] = tproject._ProjectInfo._DataWeather;
                    obj[cnt, 16] = tproject._DataDir.FullName;

                    string markstr = null;
                    string filename = tproject._DataDir.FullName + "\\RoadStatuMarkInfo.txt";
                    if (File.Exists(filename))
                    {
                        string[] infos = File.ReadAllLines(filename);
                        int strlen = infos.Length;
                        for (int ti = 0; ti < strlen; ++ti)
                        {
                            markstr = markstr + infos[ti];
                            if (ti < strlen - 1)
                            {
                                markstr = markstr + "\r\n";
                            }
                        }
                    }
                    obj[cnt, 20] = markstr;

                    ++cnt;
                }

                MSExcel.Range destrange = destsheet.get_Range(string.Format("A2:U{0}", cnt + 1));
                destrange.Value2 = obj;
                GlobalExcel.SetBorderLine(destrange, 63);

                _Workbook.Save();
                _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();

                MessageBox.Show("生成工程台账完成！");
            }
        }

        private void barButtonItem25_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018)
            {
                if (_Setting.SelectDrawDis == 0)
                {
                    int len = _ExcelJSMXList.Count;
                    if (_ExcelBHTJList.Count != len
                        || _ExcelDRList.Count != len
                        || _ExcelIRIList.Count != len)
                    {
                        MessageBox.Show("导入的报表中，IRI报表数量 和 病害统计报表数量 和 技术状况评定明细表 和 PCI报表数量 不一致，请检查！");
                        return;
                    }

                    MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                    MSExcel.Workbook drworkbook = null;
                    MSExcel.Workbook iriworkbook = null;
                    MSExcel.Workbook mqiworkbook = null;
                    MSExcel.Workbook disworkbook = null;

                    FileInfo fileinfo = new FileInfo(_ExcelDRList[0]);

                    string newmqi_dpath = string.Format("{0}\\{1}-{2}-MQI-新路网", fileinfo.Directory.Parent.Parent.FullName, _Setting.DistrictCode, _Setting.DetectYear);
                    string newdis_dpath = string.Format("{0}\\{1}-{2}-病害-新路网", fileinfo.Directory.Parent.Parent.FullName, _Setting.DistrictCode, _Setting.DetectYear);
                    string newidx_dpath = string.Format("{0}\\{1}-{2}-指标-新路网", fileinfo.Directory.Parent.Parent.FullName, _Setting.DistrictCode, _Setting.DetectYear);

                    if (!Directory.Exists(newmqi_dpath))
                    {
                        Directory.CreateDirectory(newmqi_dpath);
                    }
                    if (!Directory.Exists(newdis_dpath))
                    {
                        Directory.CreateDirectory(newdis_dpath);
                    }
                    if (!Directory.Exists(newidx_dpath))
                    {
                        Directory.CreateDirectory(newidx_dpath);
                    }

                    for (int i = 0; i < len; ++i)
                    {
                        MSExcel.Worksheet iriworksheet = null;
                        MSExcel.Worksheet drworksheet = null;
                        MSExcel.Worksheet mqiworksheet = null;
                        MSExcel.Worksheet lqworksheet = null;
                        MSExcel.Worksheet snworksheet = null;
                        MSExcel.Worksheet infoworksheet = null;

                        drworkbook = excelApp.Workbooks.Open(_ExcelDRList[i], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        iriworkbook = excelApp.Workbooks.Open(_ExcelIRIList[i], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        mqiworkbook = excelApp.Workbooks.Open(_ExcelJSMXList[i], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        disworkbook = excelApp.Workbooks.Open(_ExcelBHTJList[i], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        infoworksheet = iriworkbook.Sheets["工程信息"] as MSExcel.Worksheet;
                        iriworksheet = iriworkbook.Sheets["Sheet1"] as MSExcel.Worksheet;

                        drworksheet = drworkbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                        mqiworksheet = mqiworkbook.Sheets["Sheet1"] as MSExcel.Worksheet;

                        try
                        {
                            lqworksheet = disworkbook.Sheets["沥青病害汇总表"] as MSExcel.Worksheet;
                        }
                        catch (System.Exception) { }

                        try
                        {
                            snworksheet = disworkbook.Sheets["水泥病害汇总表"] as MSExcel.Worksheet;
                        }
                        catch (System.Exception) { }

                        MSExcel.Range srcrange = infoworksheet.get_Range("A2:B17");
                        object[,] infoobj = (object[,])srcrange.Value2;
                        string roadcode = infoobj[4, 2].ToString().Substring(0, 4);
                        string roaddirection = infoobj[7, 2].ToString();
                        int direction = 0;
                        if (roaddirection == "上行")
                        {
                            roaddirection = "1-上行";
                            direction = 1;
                        }
                        else if (roaddirection == "下行")
                        {
                            roaddirection = "2-下行";
                            direction = -1;
                        }
                        string roadnum = null;
                        if (infoobj[9, 2] == null)
                        {
                            roadnum = "1车道";
                        }
                        else
                        {
                            roadnum = infoobj[9, 2].ToString().Replace("车道", "") + "车道";
                        }
                        string datadate = infoobj[10, 2].ToString();

                        string srcxls = null;

                        //路面病害
                        if (lqworksheet != null)
                        {
                            int lqdisuserow = GlobalExcel.judegeusedrow(lqworksheet, 2, 5);
                            MSExcel.Range disrange = lqworksheet.get_Range(string.Format("A5:AB{0}", lqdisuserow));
                            lqdisuserow = lqdisuserow - 4;
                            object[,] lqdisobj = (object[,])disrange.Value2;
                            object[,] lqdisobj_dest = new object[lqdisuserow, 27];
                            for (int kk = 1, destrow = 0; kk < lqdisuserow; ++kk)
                            {
                                if (lqdisobj[kk, 1] != null || lqdisobj[kk, 1].ToString() != "小计" || lqdisobj[kk, 1].ToString() != "总计")
                                {
                                    int colidx = 0;
                                    lqdisobj_dest[destrow, colidx++] = destrow + 1;
                                    lqdisobj_dest[destrow, colidx++] = string.Format("{0:K0+000}-{1:K0+000}", lqdisobj[kk, 1], lqdisobj[kk, 2]);
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 27];

                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 15];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 16];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 11];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 12];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 9];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 10];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 19];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 20];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 21];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 22];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 13];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 14];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 17];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 18];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 23];
                                    lqdisobj_dest[destrow, colidx++] = Convert.ToDouble(lqdisobj[kk, 24]) + Convert.ToDouble(lqdisobj[kk, 25]);
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 4];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 5];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 6];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 7];
                                    lqdisobj_dest[destrow, colidx++] = lqdisobj[kk, 8];

                                    lqdisobj_dest[destrow, colidx++] = Math.Abs(Convert.ToInt32(lqdisobj[kk, 1]) - Convert.ToInt32(lqdisobj[kk, 2]));
                                    lqdisobj_dest[destrow, colidx++] = _Setting.DistrictCode;
                                    lqdisobj_dest[destrow, colidx++] = infoobj[8, 2];

                                    ++destrow;
                                }
                            }

                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\003-平台病害模板（沥青）.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_dislq = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_dislq.SaveAs(string.Format("{0}\\{1}-病害（沥青）-{2}", newdis_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_dislq = destworkbook_dislq.Sheets["Sheet1"] as MSExcel.Worksheet;
                            disrange = destworksheet_dislq.get_Range(string.Format("A8:AA{0}", lqdisuserow + 7));
                            disrange.Value2 = lqdisobj_dest;

                            destworksheet_dislq.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_dislq.Cells[1, 4] = roadcode;
                            destworksheet_dislq.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_dislq.Cells[2, 2] = "沥青路面";
                            destworksheet_dislq.Cells[2, 4] = roaddirection;
                            destworksheet_dislq.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_dislq.Cells[3, 2] = roadnum;
                            destworksheet_dislq.Cells[3, 4] = datadate;

                            destworkbook_dislq.Save();
                            destworkbook_dislq.Close(Type.Missing, Type.Missing, Type.Missing);
                        }

                        if (snworksheet != null)
                        {
                            int sndisuserow = GlobalExcel.judegeusedrow(snworksheet, 2, 5);
                            MSExcel.Range disrange = snworksheet.get_Range(string.Format("A5:AA{0}", sndisuserow));
                            sndisuserow = sndisuserow - 4;
                            object[,] sndisobj = (object[,])disrange.Value2;
                            object[,] sndisobj_dest = new object[sndisuserow, 26];
                            for (int kk = 1, destrow = 0; kk < sndisuserow; ++kk)
                            {
                                if (sndisobj[kk, 1] != null || sndisobj[kk, 1].ToString() != "小计" || sndisobj[kk, 1].ToString() != "总计")
                                {
                                    int colidx = 0;
                                    sndisobj_dest[destrow, colidx++] = destrow + 1;
                                    sndisobj_dest[destrow, colidx++] = string.Format("{0:K0+000}-{1:K0+000}", sndisobj[kk, 1], sndisobj[kk, 2]);
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 26];

                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 4];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 5];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 6];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 7];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 8];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 9];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 10];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 11];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 12];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 13];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 14];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 15];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 16];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 17];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 18];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 19];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 20];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 21];
                                    sndisobj_dest[destrow, colidx++] = sndisobj[kk, 22];
                                    sndisobj_dest[destrow, colidx++] = Convert.ToDouble(sndisobj[kk, 23]) + Convert.ToDouble(sndisobj[kk, 24]);

                                    sndisobj_dest[destrow, colidx++] = Math.Abs(Convert.ToInt32(sndisobj[kk, 1]) - Convert.ToInt32(sndisobj[kk, 2]));
                                    sndisobj_dest[destrow, colidx++] = _Setting.DistrictCode;
                                    sndisobj_dest[destrow, colidx++] = infoobj[8, 2];

                                    ++destrow;
                                }
                            }

                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\003-平台病害模板（水泥）.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_dissn = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_dissn.SaveAs(string.Format("{0}\\{1}-病害（水泥）-{2}", newdis_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_dissn = destworkbook_dissn.Sheets["Sheet1"] as MSExcel.Worksheet;
                            disrange = destworksheet_dissn.get_Range(string.Format("A8:Z{0}", sndisuserow + 7));
                            disrange.Value2 = sndisobj_dest;

                            destworksheet_dissn.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_dissn.Cells[1, 4] = roadcode;
                            destworksheet_dissn.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_dissn.Cells[2, 2] = "水泥路面";
                            destworksheet_dissn.Cells[2, 4] = roaddirection;
                            destworksheet_dissn.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_dissn.Cells[3, 2] = roadnum;
                            destworksheet_dissn.Cells[3, 4] = datadate;

                            destworkbook_dissn.Save();
                            destworkbook_dissn.Close(Type.Missing, Type.Missing, Type.Missing);
                        }

                        //pqi
                        int druserow = GlobalExcel.judegeusedrow(drworksheet, 1, 3);
                        druserow = druserow - 2;

                        int iriuserow = GlobalExcel.judegeusedrow(iriworksheet, 1, 4);
                        iriuserow = iriuserow - 3;

                        int pqiuserow = GlobalExcel.judegeusedrow(mqiworksheet, 15, 5);
                        pqiuserow = pqiuserow - 4;

                        MSExcel.Range irirange = iriworksheet.get_Range(string.Format("A4:K{0}", iriuserow + 3));
                        object[,] obj_iri = (object[,])irirange.Value2;

                        MSExcel.Range drrange = drworksheet.get_Range(string.Format("A3:I{0}", druserow + 2));
                        object[,] obj_dr = (object[,])drrange.Value2;

                        MSExcel.Range mqirange = mqiworksheet.get_Range(string.Format("A5:O{0}", pqiuserow + 4));
                        object[,] obj_mqi = (object[,])mqirange.Value2;

                        if (druserow != iriuserow || druserow != pqiuserow || iriuserow != pqiuserow)
                        {
                            MessageBox.Show("PCI/RQI/MQI报表的数据行数不一致，请检查！");
                        }

                        int destuserow = Math.Min(Math.Min(druserow, iriuserow), pqiuserow);
                        int obj_pqi_destlq_rowcnt = 0;
                        int obj_pqi_destsn_rowcnt = 0;

                        object[,] obj_pqi_destlq = new object[pqiuserow, 13];
                        object[,] obj_pqi_destsn = new object[pqiuserow, 13];

                        object[,] obj_mqi_destlq = new object[pqiuserow, 11];
                        object[,] obj_mqi_destsn = new object[pqiuserow, 11];

                        int tmp = 0;
                        if (direction > 0)
                        {
                            for (int kk = 0; kk < destuserow; ++kk)
                            {
                                tmp = kk + 1;
                                if (obj_mqi[tmp, 15].ToString() == "沥青")
                                {
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 0] = tmp;
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 3] = obj_dr[tmp, 4];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 4] = obj_iri[tmp, 6];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 6] = obj_mqi[tmp, 6];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 7] = obj_mqi[tmp, 7];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 9] = obj_mqi[tmp, 5];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 10] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destlq_rowcnt + 6);
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 11] = _Setting.DistrictCode;
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 12] = infoobj[8, 2];

                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 0] = tmp;
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 3] = obj_mqi[tmp, 5];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 4] = obj_mqi[tmp, 4];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 5] = obj_mqi[tmp, 14];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 6] = obj_mqi[tmp, 13];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 7] = obj_mqi[tmp, 3];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 8] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destlq_rowcnt + 6);
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 9] = _Setting.DistrictCode;
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 10] = infoobj[8, 2];

                                    ++obj_pqi_destlq_rowcnt;
                                }
                                else if (obj_mqi[tmp, 15].ToString() == "水泥")
                                {
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 0] = tmp;
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 3] = obj_dr[tmp, 4];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 4] = obj_iri[tmp, 6];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 6] = obj_mqi[tmp, 6];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 7] = obj_mqi[tmp, 7];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 9] = obj_mqi[tmp, 5];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 10] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destsn_rowcnt + 6);
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 11] = _Setting.DistrictCode;
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 12] = infoobj[8, 2];

                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 0] = tmp;
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 3] = obj_mqi[tmp, 5];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 4] = obj_mqi[tmp, 4];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 5] = obj_mqi[tmp, 14];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 6] = obj_mqi[tmp, 13];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 7] = obj_mqi[tmp, 3];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 8] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destsn_rowcnt + 6);
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 9] = _Setting.DistrictCode;
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 10] = infoobj[8, 2];

                                    ++obj_pqi_destsn_rowcnt;
                                }
                            }
                        }
                        else
                        {
                            for (int kk = 0; kk < destuserow; ++kk)
                            {
                                tmp = destuserow - kk;
                                if (obj_iri[tmp, 9].ToString() == "沥青")
                                {
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 0] = kk + 1;
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 3] = obj_dr[tmp, 4];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 4] = obj_iri[tmp, 6];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 6] = obj_mqi[tmp, 6];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 7] = obj_mqi[tmp, 7];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 9] = obj_mqi[tmp, 5];
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 10] = string.Format("=ABS(B{0}-C{0})", kk + 6);
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 11] = _Setting.DistrictCode;
                                    obj_pqi_destlq[obj_pqi_destlq_rowcnt, 12] = infoobj[8, 2];

                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 0] = tmp;
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 3] = obj_mqi[tmp, 5];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 4] = obj_mqi[tmp, 4];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 5] = obj_mqi[tmp, 14];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 6] = obj_mqi[tmp, 13];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 7] = obj_mqi[tmp, 3];
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 8] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destlq_rowcnt + 6);
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 9] = _Setting.DistrictCode;
                                    obj_mqi_destlq[obj_pqi_destlq_rowcnt, 10] = infoobj[8, 2];

                                    ++obj_pqi_destlq_rowcnt;
                                }
                                else if (obj_iri[tmp, 9].ToString() == "水泥")
                                {
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 0] = kk + 1;
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 3] = obj_dr[tmp, 4];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 4] = obj_iri[tmp, 6];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 6] = obj_mqi[tmp, 6];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 7] = obj_mqi[tmp, 7];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 9] = obj_mqi[tmp, 5];
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 10] = string.Format("=ABS(B{0}-C{0})", kk + 6);
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 11] = _Setting.DistrictCode;
                                    obj_pqi_destsn[obj_pqi_destsn_rowcnt, 12] = infoobj[8, 2];

                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 0] = tmp;
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 1] = obj_iri[tmp, 1];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 2] = obj_iri[tmp, 2];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 3] = obj_mqi[tmp, 5];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 4] = obj_mqi[tmp, 4];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 5] = obj_mqi[tmp, 14];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 6] = obj_mqi[tmp, 13];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 7] = obj_mqi[tmp, 3];
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 8] = string.Format("=ABS(B{0}-C{0})", obj_pqi_destsn_rowcnt + 6);
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 9] = _Setting.DistrictCode;
                                    obj_mqi_destsn[obj_pqi_destsn_rowcnt, 10] = infoobj[8, 2];

                                    ++obj_pqi_destsn_rowcnt;
                                }
                            }
                        }

                        if (obj_pqi_destlq_rowcnt > 0)
                        {
                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\004-平台技术指标模板.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_pqilq = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_pqilq.SaveAs(string.Format("{0}\\{1}-七项指标（沥青）-{2}", newidx_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_pqilq = destworkbook_pqilq.Sheets["Sheet1"] as MSExcel.Worksheet;
                            MSExcel.Range pqirange_destlq = destworksheet_pqilq.get_Range(string.Format("A6:M{0}", destuserow + 5));
                            pqirange_destlq.Value2 = obj_pqi_destlq;
                            destworksheet_pqilq.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_pqilq.Cells[1, 4] = roadcode;
                            destworksheet_pqilq.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_pqilq.Cells[2, 2] = "沥青路面";
                            destworksheet_pqilq.Cells[2, 4] = roaddirection;
                            destworksheet_pqilq.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_pqilq.Cells[3, 2] = roadnum;
                            destworksheet_pqilq.Cells[3, 4] = datadate;
                            destworkbook_pqilq.Save();
                            destworkbook_pqilq.Close(Type.Missing, Type.Missing, Type.Missing);

                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\005-平台MQI模板.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_mqilq = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_mqilq.SaveAs(string.Format("{0}\\{1}-{2}-MQI（沥青）", newmqi_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_mqilq = destworkbook_mqilq.Sheets["Sheet1"] as MSExcel.Worksheet;
                            MSExcel.Range mqirange_destlq = destworksheet_mqilq.get_Range(string.Format("A6:K{0}", destuserow + 5));
                            mqirange_destlq.Value2 = obj_mqi_destlq;
                            destworksheet_mqilq.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_mqilq.Cells[1, 4] = roadcode;
                            destworksheet_mqilq.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_mqilq.Cells[2, 2] = "沥青路面";
                            destworksheet_mqilq.Cells[2, 4] = roaddirection;
                            destworksheet_mqilq.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_mqilq.Cells[3, 2] = roadnum;
                            destworksheet_mqilq.Cells[3, 4] = datadate;
                            destworkbook_mqilq.Save();
                            destworkbook_mqilq.Close(Type.Missing, Type.Missing, Type.Missing);
                        }

                        if (obj_pqi_destsn_rowcnt > 0)
                        {
                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\004-平台技术指标模板.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_pqisn = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_pqisn.SaveAs(string.Format("{0}\\{1}-七项指标（水泥）-{2}", newidx_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_pqisn = destworkbook_pqisn.Sheets["Sheet1"] as MSExcel.Worksheet;
                            MSExcel.Range pqirange_destsn = destworksheet_pqisn.get_Range(string.Format("A6:M{0}", destuserow + 5));
                            pqirange_destsn.Value2 = obj_pqi_destsn;
                            destworksheet_pqisn.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_pqisn.Cells[1, 4] = roadcode;
                            destworksheet_pqisn.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_pqisn.Cells[2, 2] = "水泥路面";
                            destworksheet_pqisn.Cells[2, 4] = roaddirection;
                            destworksheet_pqisn.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_pqisn.Cells[3, 2] = roadnum;
                            destworksheet_pqisn.Cells[3, 4] = datadate;
                            destworkbook_pqisn.Save();
                            destworkbook_pqisn.Close(Type.Missing, Type.Missing, Type.Missing);

                            srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2018\奥路通\005-平台MQI模板.xlsx", System.Windows.Forms.Application.StartupPath);
                            MSExcel.Workbook destworkbook_mqisn = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            destworkbook_mqisn.SaveAs(string.Format("{0}\\{1}-{2}-MQI（水泥）", newmqi_dpath, roadcode, _Setting.DetectYear),
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            MSExcel.Worksheet destworksheet_mqisn = destworkbook_mqisn.Sheets["Sheet1"] as MSExcel.Worksheet;
                            MSExcel.Range mqirange_destsn = destworksheet_mqisn.get_Range(string.Format("A6:K{0}", destuserow + 5));
                            mqirange_destsn.Value2 = obj_mqi_destsn;
                            destworksheet_mqisn.Cells[1, 2] = 统计单元长度._DisUnitLen;
                            destworksheet_mqisn.Cells[1, 4] = roadcode;
                            destworksheet_mqisn.Cells[1, 6] = _Setting.DetectYear;
                            destworksheet_mqisn.Cells[2, 2] = "水泥路面";
                            destworksheet_mqisn.Cells[2, 4] = roaddirection;
                            destworksheet_mqisn.Cells[2, 6] = _Setting.DetectNum;
                            destworksheet_mqisn.Cells[3, 2] = roadnum;
                            destworksheet_mqisn.Cells[3, 4] = datadate;
                            destworkbook_mqisn.Save();
                            destworkbook_mqisn.Close(Type.Missing, Type.Missing, Type.Missing);
                        }

                        drworkbook.Close();
                        iriworkbook.Close();
                        mqiworkbook.Close();
                        disworkbook.Close();
                    }

                    excelApp.Quit();
                    MessageBox.Show("报表转换完成！");

                    int generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                }
            }
        }

        private void barButtonItem26_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.Cursor = Cursors.WaitCursor;
                string fName = openFileDialog.FileName;
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                MSExcel.Workbook workbook = excelApp.Workbooks.Open(fName, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MSExcel.Worksheet excelsheet = workbook.Sheets["Sheet1"] as MSExcel.Worksheet;

                int userow = GlobalExcel.judegeusedrow(excelsheet, 1);
                MSExcel.Range srcrange = excelsheet.get_Range(string.Format("A2:T{0}", userow));
                object[,] srcobj = (object[,])srcrange.Value2;
                workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();

                for (int i = 1; i < userow; ++i)
                {
                    if (srcobj[i, 17] != null && srcobj[i, 8] != null && srcobj[i, 9] != null)
                    {
                        string fpath = string.Format("{0}\\ProjectInfo.txt", srcobj[i, 17]);
                        if (File.Exists(fpath))
                        {
                            bool ischange = false;
                            string[] datastrs = File.ReadAllLines(fpath, Encoding.UTF8);
                            int datastrlen = datastrs.Length;
                            for (int j = 0; j < datastrlen; ++j)
                            {
                                if (datastrs[j].Contains("工程起点桩号："))
                                {
                                    if (srcobj[i, 18] != null)
                                    {
                                        int oldmile = 0;
                                        int newmile = 0;

                                        try
                                        {
                                            oldmile = Convert.ToInt32(srcobj[i, 8].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的原始起点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 8]));
                                            continue;
                                        }

                                        try
                                        {
                                            newmile = Convert.ToInt32(srcobj[i, 18].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的新起点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 18]));
                                            continue;
                                        }

                                        if (oldmile != newmile)
                                        {
                                            datastrs[j] = string.Format("工程起点桩号：K{0:0000}+{1:000}", newmile / 1000, newmile % 1000);
                                            ischange = true;
                                        }
                                    }
                                }
                                else if (datastrs[j].Contains("工程终点道路标识桩号："))
                                {
                                    if (srcobj[i, 19] != null)
                                    {
                                        int oldmile = 0;
                                        int newmile = 0;
                                        try
                                        {
                                            oldmile = Convert.ToInt32(srcobj[i, 9].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的原始终点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 9]));
                                            continue;
                                        }

                                        try
                                        {
                                            newmile = Convert.ToInt32(srcobj[i, 19].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的新终点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 19]));
                                            continue;
                                        }

                                        if (oldmile != newmile)
                                        {
                                            datastrs[j] = string.Format("工程终点道路标识桩号：K{0:0000}+{1:000}", newmile / 1000, newmile % 1000);
                                            ischange = true;
                                        }
                                    }
                                }
                                else if (datastrs[j].Contains("工程终点道路实际桩号："))
                                {
                                    if (srcobj[i, 19] != null)
                                    {
                                        int oldmile = 0;
                                        int newmile = 0;
                                        try
                                        {
                                            oldmile = Convert.ToInt32(srcobj[i, 9].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的原始终点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 9]));
                                            continue;
                                        }

                                        try
                                        {
                                            newmile = Convert.ToInt32(srcobj[i, 19].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的新终点桩号出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 19]));
                                            continue;
                                        }
                                        if (oldmile != newmile)
                                        {
                                            datastrs[j] = string.Format("工程终点道路实际桩号：K{0:0000}+{1:000}", newmile / 1000, newmile % 1000);
                                            ischange = true;
                                        }
                                    }
                                }
                                else if (datastrs[j].Contains("工程总里程数："))
                                {
                                    if (srcobj[i, 20] != null)
                                    {
                                        string[] ttstrs = datastrs[j].Split('：');
                                        int olddmi = 0;
                                        int newdmi = 0;

                                        try
                                        {
                                            Convert.ToInt32(ttstrs[1].Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的原始里程出错：{2}", srcobj[i, 1], srcobj[i, 17], datastrs[j]));
                                            continue;
                                        }

                                        try
                                        {
                                            newdmi = Convert.ToInt32(srcobj[i, 20].ToString().Replace("K", "").Replace("+", "").Replace("k", ""));
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(string.Format("读取{0}，{1}的新里程出错：{2}", srcobj[i, 1], srcobj[i, 17], srcobj[i, 20]));
                                            continue;
                                        }

                                        if (olddmi != newdmi)
                                        {
                                            datastrs[j] = string.Format("工程总里程数：K{0:0000}+{1:000}", newdmi / 1000, newdmi % 1000);
                                            ischange = true;
                                        }
                                    }
                                }
                            }
                            if (ischange)
                            {
                                File.WriteAllLines(fpath, datastrs, Encoding.UTF8);
                            }
                        }
                    }
                }

                this.Cursor = Cursors.Default;
                MessageBox.Show("批量调整工程起止点桩号完成，请重新导入工程数！");
            }
        }

        private void barButtonItem27_ItemClick(object sender, ItemClickEventArgs e)
        {
            车道报表选择 chksheet = new 车道报表选择();
            chksheet.ShowDialog();

            if (chksheet._IsOK)
            {
                MyExcelMerge.OutputMerge(chksheet._MergeType, chksheet._ExcelType, chksheet._MergeIdxInfo,
                    chksheet._UpXlsFiles, chksheet._DownXlsFiles,
                    chksheet._UpXlsFilesKM, chksheet._DownXlsFilesKM,
                    chksheet._OutputPath);
                MessageBox.Show("合并多车道报表完成！");
            } 
        }

        private void barButtonItem28_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择导出简易工程：";
            fd.SelectedPath = _Setting.DefaultPath;

            List<DirectoryInfo> projects = new List<DirectoryInfo>();
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }

                foreach (SingleProject proj in _Projects)
                {
                    proj.CreateSlimProj(fd.SelectedPath);
                }
                MessageBox.Show("导出简易工程成功");
            }
        }

        private void barButtonItem29_ItemClick(object sender, ItemClickEventArgs e)
        {
            if ((_Setting.ParmStyle == StandardParmType.DegreeRoad2018 || _Setting.ParmStyle == StandardParmType.CityRoad)
                &&_Setting.SelectDrawDis== 0)
            {

            }
            else
            {
                MessageBox.Show("该功能仅支持【等级公路2018】【城镇道路】大框病害导出。");
                return;
            }
                dxf导出车道报表选择 unitlendlg = new dxf导出车道报表选择();
            unitlendlg.ShowDialog();
            if (!unitlendlg._IsOK)
            {
                return;
            }

            if (unitlendlg._IsProvinceRoad)
            {
                if (unitlendlg._UpXlsFileList.Count() > 0)
                {
                    OutputXR.OutputDxfByExcel_ProvinceRoad(unitlendlg._UpXlsFileList, 1, unitlendlg._savePath, unitlendlg._BeginMile, unitlendlg._EndMile);

                }
                if (unitlendlg._DownXlsFileList.Count() > 0)
                {
                    OutputXR.OutputDxfByExcel_ProvinceRoad(unitlendlg._DownXlsFileList, -1, unitlendlg._savePath, unitlendlg._BeginMile, unitlendlg._EndMile);

                }
                MessageBox.Show("导出dxf成功");
                return;
            }
            if (unitlendlg._UpXlsFileList.Count() > 0)
            {
                OutputXR.OutputDxfByExcel(unitlendlg._UpXlsFileList, 1, unitlendlg._savePath, unitlendlg._BeginMile, unitlendlg._EndMile, _Setting.ParmStyle);
            }

            if (unitlendlg._DownXlsFileList.Count() > 0)
            {
                OutputXR.OutputDxfByExcel(unitlendlg._DownXlsFileList, -1, unitlendlg._savePath, unitlendlg._BeginMile, unitlendlg._EndMile, _Setting.ParmStyle);
            }
            MessageBox.Show("导出dxf成功");

        }

        private void dockPanel_main_data_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 惯导平整度计算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem30_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.outDaqAccelerate)
            {
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择惯导数据文件夹：";
                var result = fd.ShowDialog();

                if (result == DialogResult.OK && fd.SelectedPath != string.Empty)
                {
                    DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                    var files = dir.GetFiles("*.daq", SearchOption.AllDirectories);
                    List<string> daqFiles = new List<string>();
                    foreach (var file in files)
                    {
                        daqFiles.Add(file.FullName);
                    }
                    foreach (string fname in daqFiles)
                    {
                        TreeNode node = new TreeNode() { Text = fname };
                        treeView_main.Nodes.Add(node);
                    }

                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;


                    
                    SingleProject.computeDaqToAcce(daqFiles);

                }

            }
            else
            {


                if (_Projects.Count < 1)
                {
                    MessageBox.Show("没有待处理的工程！");
                }
                else
                {
                    MessageBox.Show("激光平整度与加速度计算平整度一体化设备，请先计算激光平整度最后计算惯导平整度！");

                    this.Cursor = Cursors.WaitCursor;
                    if (JudgeFreeSpace())
                    {
                        WinGDProcessBar iriFrom = new WinGDProcessBar(_Projects);
                        lowStartIrmThread(iriFrom);
                        iriFrom.ShowDialog();
                    }
                    this.Cursor = Cursors.Default;
                }

            }


        }
        private void lowStartIrmThread(WinGDProcessBar form)
        {
            ThreadIRM = new Thread(lowIrmThreadMethod) { IsBackground = true };
            ThreadIRM.Start(form);

        }
        private void lowIrmThreadMethod(object obj)
        {
            WinGDProcessBar form1 = (WinGDProcessBar)obj;
            lowComputeIRM(form1);

        }
        private void lowComputeIRM(WinGDProcessBar obj)
        {
            obj.SetMainMax(_Projects.Count);
            bool resutl = true;
            foreach (SingleProject proj in _Projects)
            {
                obj.TextInfoAdd("正在处理：" + proj._DataDir.Name);
                bool resutl1 = proj.lowComputeIRI(obj);
                if (!resutl1)
                {
                    resutl = false;
                }
                if (resutl1)
                {
                    obj.TextInfoAdd("处理完成：" + proj._DataDir.Name);
                    obj.AddMainVal(1);
                }
                else
                {
                    obj.TextInfoAdd("处理失败：" + proj._DataDir.Name);
                }

            }

            if (resutl)
            {
                MessageBox.Show("生成IRI完成！");
            }
            else
            {
                MessageBox.Show("生成IRI存在问题请查看界面日志！");
            }



        }

        private void barButtonItem31_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
            fd.ShowDialog();
            //此处取消分段出表
            _Setting.needSub = false;
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }
                MyExcelDegreeSmall2018.LoadXlsParm();
                MSExcel.Application excelApp = new MSExcel.Application()
                {
                    Visible = true,
                    DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                    AlertBeforeOverwriting = false
                };

                string outdirpath = null;
                foreach (SingleProject proj in _Projects)
                {

                    outdirpath = fd.SelectedPath;
                    if (_Setting.Is_Multfolder != 0)  //等于0 导出到同一个文件夹
                    {
                        string smile = proj._ProjectInfo._StartMile.ToString("K0000+000");
                        string emile = proj._ProjectInfo._EndMile.ToString("K0000+000");

                        string tt = "\\" + proj._DataDir.Name;
                        tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                        outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);


                        if (!Directory.Exists(outdirpath))
                        {
                            Directory.CreateDirectory(outdirpath);
                        }
                    }

                    proj.GenerateExcel(excelApp, outdirpath, null, _Setting.IsExcel);
                }

                excelApp.Quit();
                MessageBox.Show("导出报表完成！");
            }



        }

        private List<FileInfo> disTxts = null;
        //private List<FileInfo> getAllDisTxt(DirectoryInfo dir)
        //{
        //    foreach (var item in dir.GetDirectories(""))
        //    {

        //    }
        //}
        /// <summary>
        /// 删除所有病害  
        /// 实现原理：删除图片目录下所有txt文本
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem32_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                var tipForm = MessageBox.Show("确定要删除路面病害？", "一键清空所有路面病害", MessageBoxButtons.OKCancel);
                if (tipForm == DialogResult.Cancel)
                {
                    return;
                }
                else
                {
                    foreach (SingleProject pro in _Projects)
                    {

                        DirectoryInfo pathInfo = pro._DataDir;


                        List<string> dics = new List<string> {
                    Path.Combine(pathInfo.FullName, "RoadImg", "Camera0"),
                    Path.Combine(pathInfo.FullName, "RoadImg", "Camera1"),
                    //Path.Combine(pathInfo.FullName, "StreetImg", "Camera0"),
                    //Path.Combine(pathInfo.FullName, "StreetImg", "Camera1")
                    };

                        foreach (var ImgPath in dics)
                        {
                            if (Directory.Exists(ImgPath))
                            {
                                DirectoryInfo dic = new DirectoryInfo(ImgPath);

                                foreach (var item in dic.GetDirectories())
                                {
                                    FileInfo[] disTxts = item.GetFiles("*.txt");
                                    foreach (var txt in disTxts)
                                    {
                                        //病害
                                        File.Delete(txt.FullName);
                                    }
                                }

                            }
                        }
                        MessageBox.Show("所有道路及景观病害清除完毕！");
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }
        /// <summary>
        /// 导出所有人工病害
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem34_ItemClick(object sender, ItemClickEventArgs e)
        {



        }
        private void outHumanDis(string path, string outPath, string basePath, string suffix)
        {
            if (File.Exists(path))
            {

                //读取大框文本信息
                string[] mes = File.ReadAllLines(path);
                if (mes.Length > 1)
                {
                    Directory.CreateDirectory(outPath);
                    Dictionary<string, List<string>> dicMes = new Dictionary<string, List<string>>();
                    //dicMes --> key:去除后缀后的文件地址+去除后缀后的输出路径  values:大框病害信息
                    for (int i = 0; i < mes.Length; i += 2)
                    {
                        //图片路径.jpg 
                        string picPathName = Path.Combine(basePath, mes[i]).Split('.').First() +
                            $"|{Path.Combine(outPath, mes[i]).Split('.').First()}";
                        string sourceTxt = basePath + mes[i].Split('.').First() + suffix;
                        string txtMegs = null;
                        bool isExist = false;

                        //去掉后缀后加入
                        if (dicMes.Keys.Contains(picPathName))
                        {
                            //dicMes[picPathName].Add(mes[i + 1]);
                        }
                        else
                        {
                            if (File.Exists(sourceTxt))
                            {
                                isExist = true;
                                txtMegs = File.ReadAllText(sourceTxt);
                            }
                            else
                            {
                                isExist = false;
                                if (mes[i + 1].Contains("clear"))
                                {
                                    dicMes.Add(picPathName, new List<string> { "clear" });
                                }


                            }
                            if (isExist)
                            {
                                dicMes.Add(picPathName, new List<string> { txtMegs });
                            }
                        }
                        //可能不存在图片

                    }
                    //写入文件
                    foreach (var mesOne in dicMes)
                    {


                        string sourcePicPath = basePath + mesOne.Key.Split('|')[0] + ".jpg";
                        //string sourceTxtPath = mesOne.Key.Split('+')[0] + ".jpg.txt";
                        string picOutPath = outPath + mesOne.Key.Split('|')[1] + ".jpg";
                        // string txtOutPath = mesOne.Key.Split('+')[1] + ".jpg.txt";



                        string txtOutPath = outPath + mesOne.Key.Split('|')[1] + suffix;
                        //先处理文本

                        string fatherPath = new DirectoryInfo(txtOutPath).Parent.FullName;
                        if (!Directory.Exists(fatherPath))
                        {
                            Directory.CreateDirectory(fatherPath);
                        }
                        if (File.Exists(txtOutPath))
                        {
                            File.Delete(txtOutPath);

                        }

                        if (!mesOne.Value.First().Contains("clear"))
                        {
                            using (StreamWriter sw = new StreamWriter(txtOutPath, true))
                            {
                                foreach (var item in mesOne.Value)
                                {
                                    sw.WriteLine(item);
                                }

                            }
                        }
                        //判断图片是否存在
                        if (!File.Exists(sourcePicPath))
                        {
                            continue;
                        }
                        //目标图像已经存在
                        if (File.Exists(picOutPath))
                        {
                            File.Delete(picOutPath);
                        }
                        File.Copy(sourcePicPath, picOutPath);
                    }
                }
            }
        }

        private void findAllXlsxFile(HashSet<string> paths, string path, string suffix)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            foreach (var dic in di.GetFiles(suffix))
            {
                paths.Add(dic.FullName);
            }
            foreach (var dic in di.GetDirectories())
            {
                foreach (var item in dic.GetFiles(suffix))
                {
                    paths.Add(item.FullName);
                }
                findAllXlsxFile(paths, dic.FullName, suffix);
            }
        }
        public bool showTip = true;
        private void XlsxToCsv(Spire.Xls.Workbook workbook, string sourceFile, string outFile)
        {
            workbook.LoadFromFile(sourceFile);
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
            if (!File.Exists(outFile))
            {
                sheet.SaveToFile(outFile, ",", Encoding.UTF8);
                File.Delete(sourceFile);
            }

        }


        private readonly object lockMy = new object();
        [STAThread]
        private void barButtonItem38_ItemClick(object sender, ItemClickEventArgs e)
        {
            string sourcePath = null;
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择报告资源文件夹：";
            fd.ShowDialog();

            if (string.IsNullOrEmpty(fd.SelectedPath))
            {

                return;

            }
            if (!Directory.Exists(fd.SelectedPath))
            {
                return;
            }
            sourcePath = fd.SelectedPath;
            //  sourcePath = @"D:\四川公路院\data";
            DirectoryInfo di = new DirectoryInfo(sourcePath);
            var filePaths = di.GetFiles("*.xlsx");
            int proValue = 0;
            int toProcess = filePaths.Length;
            if (toProcess == 0)
            {
                MessageBox.Show("excel数据源不存在请检查！");
                return;
            }
            using (ManualResetEvent manualResetEvent = new ManualResetEvent(false))//线程安全)
            {


                ProcessOperator p = new ProcessOperator();

                System.Action t1 = () =>
                {


                    int temp = 0;
                    int ProCount = filePaths.Length;
                    for (int d = 0; d < ProCount; d++)
                    {

                        int t = d;
                        var pathCopy = filePaths[t]; 
                        var thread = new Thread(() =>
                        {
                            if (!pathCopy.FullName.Contains("~$"))
                            {
                                MyWordSzechwanDQ sc = new MyWordSzechwanDQ(pathCopy.FullName);
                                sc.readModuleTxt();

                                if (sc.ReadExcel(pathCopy.FullName, t))
                                {
                                    sc.getTextInfoData();
                                    lock (lockMy)
                                    {

                                        proValue += (100 / ProCount) * 15 / 100;
                                        p._backgroundWorker.ReportProgress(proValue);
                                        sc.WriteWord(p);
                                        proValue += (100 / ProCount) * 80 / 100;
                                        p._backgroundWorker.ReportProgress(proValue);
                                    }

                                    sc.Disposed();
                                }
                                else
                                {
                                    log.Error("没有找到excel文件");
                                }
                                if (Interlocked.Decrement(ref toProcess) == 0)
                                    manualResetEvent.Set();
                            }
                        });
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start();
                        if (t > 50 * temp + 50)
                        {
                            temp++;
                            manualResetEvent.WaitOne();
                        }
                    }
                    manualResetEvent.WaitOne();


                    p._backgroundWorker.ReportProgress(100);
                    p.BackgroundWorkerCompleted += P_BackgroundWorkerCompleted1;
                    MessageBox.Show("所有word导出已经完成，请您查看!");

                };

                p.BackgroundWork = t1;
                p.BackgroundWorkerCompleted += P_BackgroundWorkerCompleted;
                p.Start();




            }
        }

        private void P_BackgroundWorkerCompleted1(object sender, EventArgs e)
        {

        }

        private void P_BackgroundWorkerCompleted(object sender, EventArgs e)
        {

        }

        private void barButtonItem39_ItemClick(object sender, ItemClickEventArgs e)
        {
            //遍历病害文本
            foreach (SingleProject pro in _Projects)
            {
                string baseTxtPath = pro._DataDir.FullName + "\\RoadImg\\Camera0";
                string outBasePath = pro._DataDir.FullName + "\\DisImg\\Camera0";
                if (Directory.Exists(outBasePath))
                {
                    Directory.Delete(outBasePath, true);
                }

                if (Directory.Exists(baseTxtPath))
                {
                    DirectoryInfo di = new DirectoryInfo(baseTxtPath);
                    var dirs = di.GetDirectories("Image*");
                    foreach (DirectoryInfo dir in dirs)
                    {
                        //寻找各自文本
                        string dirName = dir.Name;

                    }
                }
            }
        }

        private void barButtonItem40_ItemClick(object sender, ItemClickEventArgs e)
        {
            gjModelSelectForm form = new gjModelSelectForm();
            if (form.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            if (_Setting.gjStandardNew == hnEnumTools.CityModelItem.等级公路5210与农村路5211标准模板导出_2025年)
            {
                //通用性模板导出_2024 自2024年12月后的新增国检制作均通过 通用性模板进行定义输出
                string cityCode = string.Empty;
                bool isStart = false;
                //加载出表配置信息
                switch (_Setting.ParmStyle)
                {
                    case StandardParmType.DegreeRoad2018:
                        if (_Setting.SelectDrawDis == 1)
                        {
                            MyExcelDegreeSmall2018.LoadXlsParm();
                        }
                        else
                        {
                            MyExcelDegree2018.LoadXlsParm();

                        }
                        break;
                    case StandardParmType.RuralRoadlowLevel:
                        if (_Setting.SelectDrawDis == 0)
                        {
                            MyExcelVillageDegree.LoadXlsParm();
                        }
                        else
                        {   
                            MyExcelVillageDegreeSmall.LoadXlsParm();
                        }
                        break;
                    default:
                        break;
                }

                foreach (var project in _Projects)
                {
                    string proName = "";

                    bool ok = CreateConventSource_Universality(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2025\通用性模板导出.txt",
              System.Windows.Forms.Application.StartupPath);
                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);

                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018: 
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel: 
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("所有通用性文件已输出，请使用《国检转换软件》选择具体【国检标准】\n《国检转换软件》将进行【文件格式转换】【图片规范输出】等操作！");
                }
            }
        

            #region 其他国检模板_2024年12月17日前
            else if (_Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.交通部2024规范)
            {
                //国检转换 交通部2024
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\交通部2024规范表头文本.txt",
              System.Windows.Forms.Application.StartupPath);

                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {

                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();

                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }

            }
            else if (_Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.河南省单位一农村路定制
               )
            {
                if (_Setting.ParmStyle != StandardParmType.RuralRoadlowLevel)
                {
                    MessageBox.Show("仅支持低等级农村公路模块操作！");
                    return;
                }

                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\河南省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);

                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();

                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelVillageDegreeSmall.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelVillageDegree.LoadXlsParm();
                                }
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".csv");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }
            }
            else if (_Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.湖南省单位一定制
                 || _Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.农养国省道路况检测数据提交格式_2026年)
            {
                if (_Setting.ParmStyle != StandardParmType.DegreeRoad2018)
                {
                    MessageBox.Show("仅支持等级公路2018模块操作！");
                    return;
                }

                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\湖南省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);
                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();
                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew, ".csv");
                                break;
                            //case StandardParmType.RuralRoadlowLevel:
                            //    if (_Setting.SelectDrawDis == 1)
                            //    {
                            //        MyExcelVillageDegreeSmall.LoadXlsParm();
                            //    }
                            //    else
                            //    {
                            //        MyExcelVillageDegree.LoadXlsParm();
                            //    }
                            //    CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".csv");
                            //    break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }
            }


            
            else if (_Setting.gjStandardNew == hnEnumTools.CityModelItem.重庆市单位一定制)
            {
                //等级公路2018大框 + 农村路大框
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\重庆市单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);
                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:

                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();

                                }

                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;

                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }

            }
            else if (_Setting.gjStandardNew == hnEnumTools.CityModelItem.甘肃省单位一定制)
            {

                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\甘肃省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);

                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();

                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelVillageDegreeSmall.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelVillageDegree.LoadXlsParm();
                                }
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }
            }
            else if (_Setting.gjStandardNew == hnEnumTools.CityModelItem.河北省单位一定制)
            {
                //等级公路2018大框 + 农村路大框
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\河北省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);
                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:


                                MyExcelDegree2018.LoadXlsParm();

                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;

                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }
            }
            else if (_Setting.gjStandardNew ==  hnEnumTools.CityModelItem.安徽省单位一定制)
            {
                //国检转换 交通部2024
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\安徽省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);

                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();
                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }

            }
            else if (_Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.河北省单位二定制)
            {
                //等级公路2018大框 + 农村路大框
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\河北省单位二定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);
                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:


                                MyExcelDegree2018.LoadXlsParm();
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;

                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }
            }
            else if (_Setting.gjStandardNew == Farmework.Other.enumTools.hnEnumTools.CityModelItem.广东省单位一定制)
            {
                //国检转换 交通部2024
                string cityCode = string.Empty;
                bool isStart = false;
                foreach (var project in _Projects)
                {
                    string proName = "";
                    bool ok = CreateConventSource(project, out proName, ref cityCode, _Setting.gjStandardNew);
                    isStart = ok;
                    if (ok)
                    {
                        string modelTxtPath = string.Format(@"{0}\报表模板\国检转换2024\广东省单位一定制表头文本.txt",
              System.Windows.Forms.Application.StartupPath);

                        List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
                        GJTitles.SetPara(txtFiles);
                        switch (_Setting.ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                if (_Setting.SelectDrawDis == 1)
                                {
                                    MyExcelDegreeSmall2018.LoadXlsParm();
                                }
                                else
                                {
                                    MyExcelDegree2018.LoadXlsParm();
                                }
                                Create2018ExcelSource2024(project, proName, _Setting.gjStandardNew);
                                break;
                            case StandardParmType.RuralRoadlowLevel:
                                MyExcelVillageDegree.LoadXlsParm();
                                CreateVillageExcelSource2024(project, proName, _Setting.gjStandardNew, ".txt");
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    { break; }
                }
                if (isStart)
                {
                    MessageBox.Show("转换中间文件生成完毕请查看！");
                }

            }
            #endregion
            #region 2023年国检转换定制
            ////河南 甘肃
            //else
            //if (_Setting.gjStandard == 1)
            //{
            //    string cityCode = string.Empty;
            //    bool isStart = false;
            //    foreach (var project in _Projects)
            //    {
            //        string proName = "";
            // bool ok = CreateConventSource(project, out proName, ref cityCode,);
            //        isStart = ok;
            //        if (ok)
            //        {
            //            switch (_Setting.ParmStyle)
            //            {
            //                case StandardParmType.DegreeRoad2018:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelDegreeSmall2018.LoadXlsParm();
            //                    }
            //                    else
            //                    { MyExcelDegree2018.LoadXlsParm(); }
            //Create2018ExcelSource(project, proName);
            //                    break;
            //                case StandardParmType.RuralRoadlowLevel:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelVillageDegreeSmall.LoadXlsParm();
            //                    }
            //                    else
            //                    {
            //                        MyExcelVillageDegree.LoadXlsParm();
            //                    }
            //                    CreateVillageExcelSource(project, proName);
            //                    break;
            //                default:
            //                    break;
            //            }
            //        }
            //        else
            //        {
            //            break;
            //        }
            //    }
            //    if (isStart)
            //    {
            //        MessageBox.Show("转换中间文件生成完毕请查看！");
            //    }
            //}

            ////江西
            //else if (_Setting.gjStandard == 2)
            //{
            //    string cityCode = string.Empty;
            //    bool isStart = false;
            //    foreach (var project in _Projects)
            //    {
            //        string proName = "";
            //        bool ok = CreateConventSource(project, out proName, ref cityCode);
            //        isStart = ok;
            //        if (ok)
            //        {
            //            switch (_Setting.ParmStyle)
            //            {
            //                case StandardParmType.DegreeRoad2018:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelDegreeSmall2018.LoadXlsParm();
            //                    }
            //                    else
            //                    { MyExcelDegree2018.LoadXlsParm(); }
            //                    Create2018ExcelSource_JiangXi(project, proName);
            //                    break;
            //                case StandardParmType.RuralRoadlowLevel:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelVillageDegreeSmall.LoadXlsParm();
            //                    }
            //                    else
            //                    {
            //                        MyExcelVillageDegree.LoadXlsParm();
            //                    }
            //                    CreateVillageExcelSource_JiangXi(project, proName);
            //                    break;
            //                default:
            //                    break;
            //            }
            //        }
            //        else
            //        {
            //            break;
            //        }
            //    }
            //    if (isStart)
            //    {
            //        MessageBox.Show("转换中间文件生成完毕请查看！");
            //    }
            //}
            //else if (_Setting.gjStandard == 3)
            //{
            //    string modelTxtPath = string.Format(@"{0}\报表模板\低等级农村公路\国检转换2023\合肥表头文本.txt",
            //  System.Windows.Forms.Application.StartupPath);
            //    List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
            //    GJTitles.SetPara(txtFiles);
            //    string cityCode = string.Empty;
            //    bool isStart = false;
            //    foreach (var project in _Projects)
            //    {
            //        string proName = "";
            //        bool ok = CreateConventSource(project, out proName, ref cityCode);
            //        isStart = ok;
            //        if (ok)
            //        {
            //            switch (_Setting.ParmStyle)
            //            {
            //                case StandardParmType.DegreeRoad2018:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelDegreeSmall2018.LoadXlsParm();
            //                    }
            //                    else
            //                    { MyExcelDegree2018.LoadXlsParm(); }
            //                    Create2018ExcelSource_HeBei(project, proName);
            //                    break;
            //                case StandardParmType.RuralRoadlowLevel:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelVillageDegreeSmall.LoadXlsParm();
            //                    }
            //                    else
            //                    {
            //                        MyExcelVillageDegree.LoadXlsParm();
            //                    }
            //                    CreateVillageExcelSource_HeBei(project, proName);
            //                    break;
            //                default:
            //                    break;
            //            }
            //        }
            //        else
            //        {
            //            break;
            //        }
            //    }
            //    if (isStart)
            //    {
            //        MessageBox.Show("转换中间文件生成完毕请查看！");
            //    }
            //}
            ////安徽
            //else if (_Setting.gjStandard == 4)
            //{
            //    string cityCode = string.Empty;
            //    bool isStart = false;
            //    foreach (var project in _Projects)
            //    {
            //        string proName = "";
            // bool ok = CreateConventSource2023(project, out proName, ref cityCode);
            //        isStart = ok;
            //        if (ok)
            //        {
            //            string modelTxtPath = string.Format(@"{0}\报表模板\低等级农村公路\国检转换2023\表头文本.txt",
            //  System.Windows.Forms.Application.StartupPath);

            //            List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
            //            GJTitles.SetPara(txtFiles);
            //            switch (_Setting.ParmStyle)
            //            {
            //                case StandardParmType.DegreeRoad2018:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    { MyExcelDegreeSmall2018.LoadXlsParm(); }
            //                    else
            //                    { MyExcelDegree2018.LoadXlsParm(); }
            //                    Create2018ExcelSource2023(project, proName);
            //                    break;
            //                case StandardParmType.RuralRoadlowLevel:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelVillageDegreeSmall.LoadXlsParm();
            //                    }
            //                    else
            //                    {
            //                        MyExcelVillageDegree.LoadXlsParm();
            //                    }
            //   CreateVillageExcelSource2023(project, proName);
            //                    break;
            //                default:
            //                    break;
            //            }
            //        }
            //        else
            //        { break; }
            //    }
            //    if (isStart)
            //    {
            //        MessageBox.Show("转换中间文件生成完毕请查看！");
            //    }
            //}
            ////辽宁省
            //else if (_Setting.gjStandard == 5)
            //{
            //    string cityCode = string.Empty;
            //    bool isStart = false;
            //    foreach (var project in _Projects)
            //    {
            //        string proName = "";
            // bool ok = CreateConventSource2023(project, out proName, ref cityCode);
            //        isStart = ok;
            //        if (ok)
            //        {
            //            string modelTxtPath = string.Format(@"{0}\报表模板\低等级农村公路\国检转换\辽宁2024表头文件.txt",
            //             System.Windows.Forms.Application.StartupPath);

            //            List<string> txtFiles = File.ReadAllLines(modelTxtPath).ToList();
            //            GJTitles.SetPara(txtFiles);
            //            switch (_Setting.ParmStyle)
            //            {
            //                case StandardParmType.DegreeRoad2018:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    { MyExcelDegreeSmall2018.LoadXlsParm(); }
            //                    else
            //                    { MyExcelDegree2018.LoadXlsParm(); }
            //  Create2018ExcelSource2023(project, proName);
            //                    break;
            //                case StandardParmType.RuralRoadlowLevel:
            //                    if (_Setting.SelectDrawDis == 1)
            //                    {
            //                        MyExcelVillageDegreeSmall.LoadXlsParm();
            //                    }
            //                    else
            //                    {
            //                        MyExcelVillageDegree.LoadXlsParm();
            //                    }
            //                    CreateVillageExcelSource2023(project, proName);
            //                    break;
            //                default:
            //                    break;
            //            }
            //            //辽宁要求改变文本编码为GBK
            //            string outDirectoryPath = project._DataDir.FullName + "\\ConverSource\\" + proName + "\\";
            //            // 遍历文件夹中的所有文件
            //            foreach (string file in Directory.GetFiles(outDirectoryPath, "*.txt", SearchOption.AllDirectories))
            //            {
            //                try
            //                {
            //                    // 读取文件内容
            //                    string content = File.ReadAllText(file, Encoding.UTF8);

            //                    // 将内容转换为GBK编码
            //                    byte[] gbkBytes = Encoding.Convert(Encoding.UTF8, Encoding.GetEncoding("GBK"), Encoding.UTF8.GetBytes(content));
            //                    string gbkContent = Encoding.GetEncoding("GBK").GetString(gbkBytes);

            //                    // 将转换后的内容写回原文件
            //                    File.WriteAllText(file, gbkContent, Encoding.GetEncoding("GBK"));
            //                    //Console.WriteLine($"文件 {file} 已转换为GBK编码。");
            //                }
            //                catch (Exception ex)
            //                {
            //                    string msg = $"处理文件 {file} 时发生错误: {ex.Message}";
            //                    throw new Exception(msg);
            //                }
            //            }
            //        }
            //        else
            //        { break; }


            //    }
            //    if (isStart)
            //    {
            //        MessageBox.Show("转换中间文件生成完毕请查看,后续请使用国检转换软件[农村路标准处理]！");
            //    }
            //}
            #endregion
        }



        private void CreateVillageExcelSource2024(SingleProject pro, string ProName, hnEnumTools.CityModelItem standard, string suff)
        {
            
            switch (standard)
            {
                case hnEnumTools.CityModelItem.等级公路5210与农村路5211标准模板导出_2025年:
                    pro.Convent_Village2024_AnHui( ConverDic["RIFile"].FullName, ConverDic["LBIFile"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);
                    break;
                case hnEnumTools.CityModelItem.交通部2024规范:
                    pro.Convent_Village2024( ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);
                    break;
                case hnEnumTools.CityModelItem.河南省单位一农村路定制:
                    pro.Convent_Village2024_HeNan( ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);
                    break;
                case hnEnumTools.CityModelItem.重庆市单位一定制:

                    pro.Convent_Village2024_ChongQing(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, Encoding.Default);
                    break;
                case hnEnumTools.CityModelItem.河北省单位一定制:
                case hnEnumTools.CityModelItem.河北省单位二定制:
                    pro.Convent_Village2024_HeBei(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ConverDic["RD"].FullName, ConverDic["PB"].FullName,
               ConverDic["TEXTFile"].FullName, ConverDic["RDFile"].FullName, ProName, Encoding.Default);
                    break;
                case hnEnumTools.CityModelItem.甘肃省单位一定制:
                    pro.Convent_Village2024_GanSu( ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);
                    break;
                case hnEnumTools.CityModelItem.安徽省单位一定制:

                    pro.Convent_Village2024_AnHui(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);
                    break;
                case hnEnumTools.CityModelItem.广东省单位一定制:
                    pro.Convent_Village2024_GuangDong(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, standard, suff);

                    break;
                default:
                    break;
            }
        }

        private void Create2018ExcelSource2024(SingleProject pro, string ProName, hnEnumTools.CityModelItem standard, string suff = ".txt")
        {
            
            switch (standard)
            {
                case hnEnumTools.CityModelItem.等级公路5210与农村路5211标准模板导出_2025年:
                    pro.Convent_Standard(ConverDic["RIFile"].FullName, ConverDic["LBIFile"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ConverDic["RD"].FullName, ConverDic["PB"].FullName,
                    ConverDic["MPD"].FullName, ConverDic["TTFile"].FullName, ConverDic["RDFile"].FullName, ProName);
                    break;
                case  hnEnumTools.CityModelItem.交通部2024规范:
                case  hnEnumTools.CityModelItem.安徽省单位一定制:
                case  hnEnumTools.CityModelItem.广东省单位一定制:
                case  hnEnumTools.CityModelItem.甘肃省单位一定制:
                    pro.Convent_2024(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, Encoding.UTF8);
                    break;
                case  hnEnumTools.CityModelItem.湖南省单位一定制:
                case hnEnumTools.CityModelItem.农养国省道路况检测数据提交格式_2026年:
                    //bump  gps=lbi   iri   RDFile=RDFile
                    pro.Convent_2024_HuNan( ConverDic["GPS"].FullName, ConverDic["IRI"].FullName, ConverDic["RDFile"].FullName, ConverDic["BUMP"].FullName, ProName, suff);

                    break;

                //case enumTools.hnEnumTools.CityModelItem.重庆市单位一定制:
                //    pro.Convent_2024_ChongQing( ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ConverDic["RD"].FullName, ConverDic["PB"].FullName,
                //     ConverDic["MPD"].FullName, ConverDic["TEXTFile"].FullName, ConverDic["RDFile"].FullName, ProName);

                case  hnEnumTools.CityModelItem.重庆市单位一定制:
                    {
                        if (_Setting.SelectDrawDis == 0)
                        {
                            pro.Convent_2024_ChongQing(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, Encoding.Default);

                        }
                        else
                        {

                            pro.Convent_2024_ChongQing(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ProName, Encoding.Default);

                        }

                    }
                    break;
                case  hnEnumTools.CityModelItem.河北省单位一定制:
                case  hnEnumTools.CityModelItem.河北省单位二定制:
                    pro.Convent_2024_HeBei(ConverDic["RIFile"].FullName, ConverDic["LBI"].FullName, ConverDic["DR"].FullName, ConverDic["IRI"].FullName, ConverDic["RD"].FullName, ConverDic["PB"].FullName,
                    ConverDic["MPD"].FullName, ConverDic["TEXTFile"].FullName, ConverDic["RDFile"].FullName, ProName, Encoding.Default); 
                    break;
              
                default:
                    break;
            }

           

        }
        Dictionary<string, DirectoryInfo> ConverDic = new Dictionary<string, DirectoryInfo>();

        private Dictionary<string, string> huNanCodeTranslateDic;
        private bool CreateConventSource(SingleProject pro, out string RoadName, ref string cityCode, hnEnumTools.CityModelItem standard)
        {

            ConverDic.Clear();
            string dirc = pro._ProjectInfo._Direction > 0 ? "A" : "B";
            //检查是否具有县级行政代码
            if (standard ==  hnEnumTools.CityModelItem.湖南省单位一定制)
            {
                huNanCodeTranslateDic = new Dictionary<string, string>
                {
                    { "01","C"},
                    { "02","D"},
                    { "03","E"},
                    { "05","F"},
                    { "09","G"},
                    { "11","H"},
                    { "12","I"},
                    { "13","J"},
                    { "14","K"},
                    { "15","L"},
                    { "16","M"},
                    { "21","N"},
                    { "22","P"},
                    { "31","Q"},
                    { "51","R"}
                };
                string tempName = pro._ProjectInfo._RoadCode;
                if (tempName.StartsWith("S") || tempName.StartsWith("X"))
                {
                    tempName = tempName.Substring(0, 4);
                }
                if (tempName.Length == 2)
                {//路线号为1 "A1"
                    tempName = tempName[0] + "A0" + tempName[1];
                }
                if (tempName.Length == 3)
                {
                    //路线号为 2 "A22"
                    tempName = tempName[0] + "A" + tempName[1] + tempName[2];
                }
                if (tempName.Length == 4)
                {
                    //"G252"
                }
                if (tempName.Length > 3)
                {
                    //路线号超过四位  "a2511"
                    string codeTemp = tempName[tempName.Length - 2].ToString() + tempName[tempName.Length - 1].ToString();
                    if (huNanCodeTranslateDic.Keys.Contains(codeTemp))
                    {
                        tempName = tempName.Substring(0, tempName.Length - 2) + huNanCodeTranslateDic[codeTemp];
                    }
                    else
                    {

                    }
                }

                RoadName = tempName + dirc;



            }
            else
            {
                if (string.IsNullOrWhiteSpace(cityCode))
                {
                    if (string.IsNullOrEmpty(pro._ProjectInfo._RoadCode))
                    {
                        MessageBox.Show("请检查" + pro._ProjectInfo._RoadName + "道路未设置道路编号!");
                        RoadName = "";
                        return false;
                    }
                    ConventGetRoadNumberForm form = new ConventGetRoadNumberForm(pro._ProjectInfo._RoadCode);
                    form.ShowDialog();
                    if (form.Ok)
                    {
                        RoadName = $"{form.Name}{dirc}"; //路线编码+县级行政代码+方向
                        string tmppath = pro._DataDir.FullName + @"\ProjectInfo.txt";
                        pro._ProjectInfo._CityCode = form.CityNum;
                        cityCode = form.CityNum;

                    }
                    else
                    {
                        RoadName = "";
                        MessageBox.Show("转换过程必须根据格式赋予项目名称!");
                        return false;
                    }



                }
                else
                {
                    string temp = pro._ProjectInfo._RoadCode.Substring(0, 4);
                    RoadName = $"{temp}{cityCode}{dirc}"; //路线编码+县级行政代码+方向
                }
            }
            string timeYeals = pro._ProjectInfo._DataDate;
            string timeHours = pro._ProjectInfo._DataTime;
            string son2PathName = RoadName + "_" + timeHours;
            string outPath = pro._DataDir.FullName + "\\" + "ConverSource\\" + RoadName;
            if (Directory.Exists(pro._DataDir.FullName + "\\" + "ConverSource\\"))
            {
                Directory.Delete(pro._DataDir.FullName + "\\" + "ConverSource\\", true);
            }
            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }
            List<string> sonPaths = new List<string>();
            if (standard == hnEnumTools.CityModelItem.交通部2024规范)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));
            }
            else if (standard == hnEnumTools.CityModelItem.河南省单位一农村路定制)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals, RoadName + "_" + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "景观图像", timeYeals, RoadName + "_" + timeHours));
            }
            else if (standard == hnEnumTools.CityModelItem.湖南省单位一定制
                || standard == hnEnumTools.CityModelItem.农养国省道路况检测数据提交格式_2026年)
            {
                string strRoadNum = RoadName;           //道路编号
                sonPaths.Add(string.Join("\\", outPath, "BUMP", timeYeals, strRoadNum));
                sonPaths.Add(string.Join("\\", outPath, "GPS"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals, strRoadNum, "0"));
                sonPaths.Add(string.Join("\\", outPath, "IRI", timeYeals, strRoadNum.Substring(0, strRoadNum.Length - 1)));
                sonPaths.Add(string.Join("\\", outPath, "RDFile", timeYeals, strRoadNum));
                sonPaths.Add(string.Join("\\", outPath, "前方图像", timeYeals, strRoadNum));
            }
            else if (standard == hnEnumTools.CityModelItem.重庆市单位一定制
                )
            {

                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));

                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));

            }
            else if (standard == hnEnumTools.CityModelItem.河北省单位一定制)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));

                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));

                sonPaths.Add(string.Join("\\", outPath, "RD"));
                sonPaths.Add(string.Join("\\", outPath, "RDFile"));

                sonPaths.Add(string.Join("\\", outPath, "PB"));
                sonPaths.Add(string.Join("\\", outPath, "MPD"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));

                sonPaths.Add(string.Join("\\", outPath, "TEXTFile"));
            }
            else if (standard == hnEnumTools.CityModelItem.河北省单位二定制)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));

                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));

                sonPaths.Add(string.Join("\\", outPath, "RD"));
                sonPaths.Add(string.Join("\\", outPath, "RDFile"));

                sonPaths.Add(string.Join("\\", outPath, "PB"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "MPD"));

                sonPaths.Add(string.Join("\\", outPath, "TEXTFile"));
            }
            else if (standard == hnEnumTools.CityModelItem.甘肃省单位一定制)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));
            }
            else if (standard == hnEnumTools.CityModelItem.安徽省单位一定制 ||
                standard == hnEnumTools.CityModelItem.广东省单位一定制)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBI"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours));
            }
            foreach (var path in sonPaths)
            {
                DirectoryInfo dir = null;
                if (!Directory.Exists(path))
                    dir = Directory.CreateDirectory(path);
                else
                    dir = new DirectoryInfo(path);
                ConverDic.Add(path.Replace(outPath, " ").Split('\\')[1], dir);
            }
            return true;
        }

        /// <summary>
        /// 创建统一模板目录 写入配置参数
        /// </summary>
        /// <param name="pro"></param>
        /// <param name="RoadName"></param>
        /// <param name="cityCode"></param>
        /// <param name="standard"></param>
        /// <returns></returns>
        private bool CreateConventSource_Universality(SingleProject pro, out string RoadName, ref string cityCode,  hnEnumTools.CityModelItem standard)
        {

            ConverDic.Clear();
            string dirc = pro._ProjectInfo._Direction > 0 ? "A" : "B";
            //检查是否具有县级行政代码 

            if (string.IsNullOrWhiteSpace(cityCode))
            {
                if (string.IsNullOrEmpty(pro._ProjectInfo._RoadCode))
                {
                    MessageBox.Show("请检查" + pro._ProjectInfo._RoadName + "道路未设置道路编号!");
                    RoadName = "";
                    return false;
                }
                ConventGetRoadNumberForm form = new ConventGetRoadNumberForm(pro._ProjectInfo._RoadCode);
                form.ShowDialog();
                if (form.Ok)
                {
                    RoadName = $"{form.Name}{dirc}"; //路线编码+县级行政代码+方向
                    string tmppath = pro._DataDir.FullName + @"\ProjectInfo.txt";
                    pro._ProjectInfo._CityCode = form.CityNum;
                    cityCode = form.CityNum;

                }
                else
                {
                    RoadName = "";
                    MessageBox.Show("转换过程必须根据格式赋予项目名称!");
                    return false;
                }
            }
            else
            {
                string temp = pro._ProjectInfo._RoadCode.Substring(0, 4);
                RoadName = $"{temp}{cityCode}{dirc}"; //路线编码+县级行政代码+方向
            }
            string timeYeals = pro._ProjectInfo._DataDate;
            string timeHours = pro._ProjectInfo._DataTime;
            string son2PathName = RoadName + "_" + timeHours;
            string outPath = pro._DataDir.FullName + "\\" + "ConverSource\\";
            if (Directory.Exists(outPath))
            {
                Directory.Delete(outPath, true);
            }
            string configFile = outPath + "ConventConfig.txt";
            outPath += RoadName;
            Directory.CreateDirectory(outPath);
            List<string> sonPaths = new List<string>();

            if (standard == hnEnumTools.CityModelItem.等级公路5210与农村路5211标准模板导出_2025年)
            {
                sonPaths.Add(string.Join("\\", outPath, "DR"));
                sonPaths.Add(string.Join("\\", outPath, "Images", timeYeals + timeHours));
                sonPaths.Add(string.Join("\\", outPath, "IRI"));
                sonPaths.Add(string.Join("\\", outPath, "LBIFile"));
                sonPaths.Add(string.Join("\\", outPath, "RIFile"));
                sonPaths.Add(string.Join("\\", outPath, "ViewImages", timeYeals + timeHours)); 
                if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018)
                {
                    sonPaths.Add(string.Join("\\", outPath, "LFile"));
                    sonPaths.Add(string.Join("\\", outPath, "MPD"));
                    sonPaths.Add(string.Join("\\", outPath, "PB"));
                    sonPaths.Add(string.Join("\\", outPath, "RD"));
                    sonPaths.Add(string.Join("\\", outPath, "RDFile"));
                    sonPaths.Add(string.Join("\\", outPath, "SFC"));
                    sonPaths.Add(string.Join("\\", outPath, "SFCOV"));
                    sonPaths.Add(string.Join("\\", outPath, "SSR"));
                    sonPaths.Add(string.Join("\\", outPath, "TTFile"));
                }
               
            }
            //写入工程配置信息
            List<string> configData = new List<string>()
            {
                $"drawType:{_Setting.SelectDrawDis}",
                $"stardand:{_Setting.ParmStyle.ToString()}"

            };
          //  File.WriteAllLines(configFile, configData);

            foreach (var path in sonPaths)
            {
                DirectoryInfo dir = null;
                if (!Directory.Exists(path))
                    dir = Directory.CreateDirectory(path);
                else
                    dir = new DirectoryInfo(path);
                ConverDic.Add(path.Replace(outPath, " ").Split('\\')[1], dir);
            }
            return true;
        }
        private void barButtonItem40_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");

            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择路段报表文件夹：";
            fd.SelectedPath = inisetting.ReadString("Road", "ExcelPath", @"C:\").Replace("\0", "");
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }

                inisetting.WriteString("Road", "ExcelPath", fd.SelectedPath);
                treeView_main.Nodes.Clear();
                _Projects.Clear();
                _CurProject = null;
                dockPanel_main_data.Controls.Clear();
                _ExcelPathList.Clear();

                // 文件夹里面的报表文件
                DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                FileInfo[] fileInfoArray = dir.GetFiles("*.xlsx");
                foreach (FileInfo tfile in fileInfoArray)
                {
                    _ExcelPathList.Add(tfile.FullName);
                }

                //文件夹里面的文件夹里面的报表文件
                DirectoryInfo[] dirInfoArray = dir.GetDirectories();
                foreach (DirectoryInfo tdir in dirInfoArray)
                {
                    FileInfo[] tfileInfoArray = tdir.GetFiles("*.xlsx");
                    foreach (FileInfo tfile in tfileInfoArray)
                    {
                        _ExcelPathList.Add(tfile.FullName);
                    }
                }

            }
            //未完

        }

        private void barButtonItem46_ItemClick(object sender, ItemClickEventArgs e)
        {
            

            OpenFileDialog fd = new OpenFileDialog() { 
            
                 Filter = "数据报表|*.xlsx",
                 Title = "选择数据报表"
            };


            //VistaFolderBrowserDialog fd = new VistaFolderBrowserDialog
            //{
            //    Description = "选择报告数据来源报表",

            //    ShowNewFolderButton = true
            //};


            string inputFilePath = "";
            if (fd.ShowDialog() !=  DialogResult.OK)
            {
                return;
            }

            if (fd.FileName != string.Empty)
            {
                inputFilePath = fd.FileName;
                if (!inputFilePath.Contains(".xlsx"))
                {
                    return;
                }
            }
            else
                return;
            string inputFileTxtPath = "";
            if (MessageBox.Show("是否要选则该报表对应的文本文件?", "提示窗口", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                OpenFileDialog fd0 = new OpenFileDialog()
                {
                 
                    Title = "选择报告数据对应的文本文件",

                };

                if (fd0.ShowDialog() !=  DialogResult.OK)
                {
                    return;
                }

                if (fd0.FileName != string.Empty)
                {
                    inputFileTxtPath = fd0.FileName;
                    if (!inputFileTxtPath.Contains(".txt"))
                    {
                        return;
                    }
                }
                else
                    return;

            }
            bool outDiseasePic = false;
            if (MessageBox.Show("报告中是否要出病害个数图，\n请确定所需报表在同级目录下?", "提示窗口", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                outDiseasePic = true;
            }
            bool outPartEight = false;
            if (MessageBox.Show("是否需要导出附件八各路分项数据?", "提示窗口", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                outPartEight = true;
            }
            MyWordHeFei hefei = new MyWordHeFei(inputFilePath, inputFileTxtPath, outDiseasePic, outPartEight);

            if (hefei.outWord())
            {

                MessageBox.Show("导出已完成");
            }
            else
            {

            }
        }

        private void barButtonItem37_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                List<string> ss = new List<string>();
                foreach (var item in System.Diagnostics.Process.GetProcesses())
                {
                    ss.Add(item.ProcessName);
                }
                foreach (var item in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                {

                    item.Kill();
                    item.WaitForExit();
                }
                foreach (var item in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
                {

                    item.Kill();
                    item.WaitForExit();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private void barButtonItem36_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            HashSet<string> xlsxPaths = new HashSet<string>();

            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择xlsx文件路径：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (showTip)
                {
                    DialogResult dr = MessageBox.Show("请注意该路径下所有xlsx文件都将转换为csv，是否继续?", "提示窗口", MessageBoxButtons.OKCancel);
                    if (dr == DialogResult.OK)
                    {
                        showTip = false;
                        findAllXlsxFile(xlsxPaths, fd.SelectedPath, "*.xlsx");
                        Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                        foreach (var fileFullName in xlsxPaths)
                        {
                            string outFileFullName = fileFullName.Split('.').First() + ".csv";
                            XlsxToCsv(workbook, fileFullName, outFileFullName);
                        }
                        MessageBox.Show("转换完成！");
                    }
                    else
                    {
                        showTip = true;
                    }
                }
                else
                {
                    findAllXlsxFile(xlsxPaths, fd.SelectedPath, "*.xlsx");
                    Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                    foreach (var fileFullName in xlsxPaths)
                    {
                        string outFileFullName = fileFullName.Split('.').First() + ".csv";
                        XlsxToCsv(workbook, fileFullName, outFileFullName);
                    }
                    MessageBox.Show("转换完成！");

                }
            }
        }


        private void barButtonItem50_ItemClick(object sender, ItemClickEventArgs e)
        {


            FolderBrowserDialog fb = new FolderBrowserDialog()
            {
                Description = "选择输出文件夹",

            };
            var result = fb.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }
            string outBasePath = fb.SelectedPath + "\\二维负样本";
            if (!Directory.Exists(outBasePath))
            {
                Directory.CreateDirectory(outBasePath);
            }
            foreach (var item in _Projects)
            {
                string deleteFilePath = string.Format("{0}\\RoadImg\\Camera0\\deleteDisPath.txt", item._DataDir.FullName);
                string[] disLines = File.ReadAllLines(deleteFilePath);

                File.Copy(deleteFilePath, outBasePath + "\\deleteDisPath.txt", true);
                foreach (var dis in disLines)
                {
                    string picParaPath = dis.Split('|').Last();
                    var temp = picParaPath.Split('\\');
                    string picPath = string.Format("{0}\\{1}", outBasePath, temp[1]);
                    if (!Directory.Exists(picPath))
                    {
                        Directory.CreateDirectory(picPath);
                    }
                    picPath = string.Format("{0}\\RoadImg\\Camera0{1}", item._DataDir.FullName, picParaPath);

                    if (File.Exists(picPath))
                    {
                        string outPicPath = outBasePath + picParaPath;
                        File.Copy(picPath, outPicPath, true);

                    }
                }

            }
            MessageBox.Show("负样本输出完成请检查！");

        }

        private void barButtonItem47_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                int toProcess = _Projects.Count;
                if (toProcess <= 0)
                {
                    return;
                }
                using (ManualResetEvent manualResetEvent = new ManualResetEvent(false))//线程安全)
                {
                    foreach (SingleProject pro in _Projects)
                    {
                        ThreadPool.QueueUserWorkItem(x =>
                        {
                            var basePath = Path.Combine(pro._DataDir.FullName, "RoadImg", "Camera0");
                            string outPath = Path.Combine(pro._DataDir.FullName, "RoadImg", "Camera0", "HumanDis");
                            string bigPath = Path.Combine(basePath, "HumanBigDisMessage.txt");
                            string smallPath = Path.Combine(basePath, "HumanSmallDisMessage.txt");
                            this.outHumanDis(bigPath, outPath, basePath, ".jpg.txt");
                            this.outHumanDis(smallPath, outPath, basePath, ".jpg_PartClass.txt");
                            //信号量设置为true  waitone会直接通过
                            if (Interlocked.Decrement(ref toProcess) == 0)
                                manualResetEvent.Set();
                            //manualResetEvent.Reset();
                        });
                    }
                    manualResetEvent.WaitOne();
                    MessageBox.Show("所有人工病害输出完毕!");
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void barButtonItem51_ItemClick(object sender, ItemClickEventArgs e)
        {
            VistaFolderBrowserDialog fd;
             if (Directory.Exists(_Setting.DefaultPath))
            {
                fd = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",

                    SelectedPath = _Setting.DefaultPath,
                    ShowNewFolderButton = true
                };
            }
            else
            {
                fd = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",


                    ShowNewFolderButton = true
                };
            }
            var fdResult =  fd.ShowDialog();
            if (fdResult != DialogResult.OK)
            {
                return;
            }
            if (!Directory.Exists(fd.SelectedPath))
            {
                return;
            }

            MyTipForm tipForm = new MyTipForm();
            tipForm.StartPosition = FormStartPosition.CenterParent;
            DialogResult result = tipForm.ShowDialog();

            if (result == DialogResult.OK)
            {
                string filter = tipForm.FilterNameStr;

                bool insert = tipForm.NeedInstallRoadCode;
                if (!string.IsNullOrEmpty(filter))
                {
                    string selectPath = fd.SelectedPath  ;
                    FileInfo[] excelFile = new DirectoryInfo(selectPath).GetFiles("*" + filter + "*" + ".xlsx", SearchOption.AllDirectories);
                    treeView_main.Width = 800;
                    treeView_main.Nodes.Clear();
                    foreach (FileInfo fname in excelFile)
                    {
                        TreeNode node = new TreeNode() { Text = fname.Name };
                        treeView_main.Nodes.Add(node);
                    }
                    dockPanel_main_data.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
                    dockPanel_main_data.Width = this.Width;
                    //开始合并
                    if (excelFile.Length <= 0)
                    {
                        return;
                    }

                    //创建模板文件

                    string modelExcelPath = selectPath + $"\\{filter}_合并结果.xlsx";
                    if (File.Exists(modelExcelPath))
                    {
                        File.Delete(modelExcelPath);
                    }
                    File.Copy(excelFile[0].FullName, modelExcelPath, true);
                    //先打开模板文件删除多余的数据
                    #region  出表
                    MSExcel.Application excelApp = new MSExcel.Application()
                    {
                        Visible = true,
                        DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                        AlertBeforeOverwriting = false
                    };
                    MSExcel.Workbook _Workbook = excelApp.Workbooks.Open(modelExcelPath, Type.Missing,
                      false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //  MSExcel.Worksheet _Worksheet = _Workbook.Sheets[1] as MSExcel.Worksheet;
                    MSExcel.Worksheet _Worksheet = _Workbook.Sheets[tipForm.TargetSheetName] as MSExcel.Worksheet;

                    // 删除非tipForm.TargetSheetName的sheet页面
                    for (int i = _Workbook.Sheets.Count; i >= 1; i--)
                    {
                        MSExcel.Worksheet sheet = _Workbook.Sheets[i] as MSExcel.Worksheet;
                        if (sheet.Name != tipForm.TargetSheetName)
                        {
                            sheet.Delete();
                        }
                    }



                    // 将所有表的数据累加到 _Worksheet中 
                    int rowIndexToPaste = _Worksheet.Cells.SpecialCells(MSExcel.XlCellType.xlCellTypeLastCell).Row + 1; // Next row to paste data
                    int startDataRowIndex = tipForm.StartRowIndex;
                    if (insert)
                    {
                        string roadCode = Path.GetFileNameWithoutExtension(excelFile[0].FullName).Split('_').FirstOrDefault();


                        //在第一列插入道路编号
                        MSExcel.Range firstColumnRange = _Worksheet.Range["A" + startDataRowIndex , "A" + (rowIndexToPaste-1)];
                        firstColumnRange.Insert(MSExcel.XlInsertShiftDirection.xlShiftToRight);

                        firstColumnRange = _Worksheet.Range["A" + (startDataRowIndex-1), "A" + (startDataRowIndex - 1)]; 
                        firstColumnRange.Insert(MSExcel.XlInsertShiftDirection.xlShiftToRight);
                        _Worksheet.Cells[startDataRowIndex - 1, 1] = "道路编号";
                        for (int row = startDataRowIndex; row < rowIndexToPaste; row++)
                        {
                            _Worksheet.Cells[row, 1] = roadCode;
                        } 
                    }

                    for (int t = 1; t < excelFile.Length; ++t)
                    {
                        FileInfo file = excelFile[t];
                        MSExcel.Workbook sourceWorkbook = excelApp.Workbooks.Open(file.FullName, Type.Missing,
                            false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                       

                        MSExcel.Worksheet sourceWorksheet = sourceWorkbook.Sheets[tipForm.TargetSheetName] as MSExcel.Worksheet;

                        // Determine the last used row in the source sheet
                        int lastUsedRow = sourceWorksheet.Cells.SpecialCells(MSExcel.XlCellType.xlCellTypeLastCell).Row;

                        // Copy data from source worksheet to destination worksheet
                        string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                        string lastColumn;

                        if (sourceWorksheet.UsedRange.Columns.Count <= alphabet.Length)
                        {
                            lastColumn = alphabet[sourceWorksheet.UsedRange.Columns.Count - 1].ToString();
                        }
                        else
                        {
                            int firstDigit = (sourceWorksheet.UsedRange.Columns.Count - 1) / alphabet.Length - 1;
                            int secondDigit = (sourceWorksheet.UsedRange.Columns.Count - 1) % alphabet.Length;
                            lastColumn = alphabet[firstDigit].ToString() + alphabet[secondDigit].ToString();
                        }

                        MSExcel.Range sourceRange = sourceWorksheet.Range["A" + startDataRowIndex, lastColumn + lastUsedRow.ToString()];
                        MSExcel.Range destinationRange = _Worksheet.Range["A" + rowIndexToPaste.ToString()];
                        sourceRange.Copy(destinationRange);
                        if (insert)
                        { 
                            string roadCode = Path.GetFileNameWithoutExtension(file.Name).Split('_').FirstOrDefault();
                            //在第一列插入道路编号
                            MSExcel.Range firstColumnRange = _Worksheet.Range["A" + (rowIndexToPaste), "A" + (rowIndexToPaste + lastUsedRow - startDataRowIndex)];
                            firstColumnRange.Insert(MSExcel.XlInsertShiftDirection.xlShiftToRight);
                            for (int row = rowIndexToPaste; row < rowIndexToPaste + lastUsedRow - startDataRowIndex + 1; row++)
                            {
                                _Worksheet.Cells[row, 1] = roadCode;
                            }


                        }
                        lastUsedRow -= (startDataRowIndex - 1);

                        rowIndexToPaste += lastUsedRow; // Move the destination row pointer to the next available row 
                        sourceWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    }
                    _Workbook.Save();
                    _Workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    int generation = System.GC.GetGeneration(excelApp);
                    System.GC.Collect(generation);//垃圾回收
                    System.GC.WaitForPendingFinalizers();
                    try
                    {
                        excelApp.Quit();
                    }
                    catch
                    {

                    }
                    #endregion
                    MessageBox.Show("合并完成!");
                }
                else
                {
                    MessageBox.Show("过滤字段为空!");
                }
            }
        }

        private void barCheckItem1_CheckedChanged(object sender, ItemClickEventArgs e)
        {
            if (barCheckItem1.Checked)
            {

                _Setting.showGpsInfoToPicture = true;
            }
            else
            {
                _Setting.showGpsInfoToPicture = false;
            }
            if (_CurProject == null)
            {
                if (barCheckItem1.Checked)
                {
                    MessageBox.Show("请先打开任一工程！");
                    barCheckItem1.Checked = false;
                    _Setting.showGpsInfoToPicture = false;
                    return;
                }
                else
                {
                    _Setting.showGpsInfoToPicture = false;
                    return;
                }

            }
            string highGpsFilePath = _CurProject._DataDir.FullName + "/GPSModel/gps.txt";
            if (File.Exists(highGpsFilePath))
            {

            }
            else
            {
                if (barCheckItem1.Checked)
                {
                    barCheckItem1.Checked = false;
                    _Setting.showGpsInfoToPicture = false;
                    MessageBox.Show("仅支持搭载高精度定位模块的设备!");
                }

            }
            _Setting.WriteData();

        }

        private void barButtonItem52_ItemClick(object sender, ItemClickEventArgs e)
        {

            if (Directory.Exists(_Setting.DefaultPath))
            {
                Process.Start("explorer.exe", _Setting.DefaultPath);
            }

        }

        private List<Disease> getPorjectAllDiseases(string projectpath, int direction)
        {
            List<Disease> disease = new List<Disease>();
            string[] ImgMilestr = null;

            if (File.Exists(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt"))
            {
                ImgMilestr = File.ReadAllLines(projectpath + "\\RoadImg\\Camera0\\Road2Mile.txt");
                int temp = 0;
                foreach (string infostr in ImgMilestr)
                {
                    string[] s = infostr.Split(' ');

                    string disfile = string.Format("{0}\\RoadImg\\Camera0{1}.txt", projectpath, s[1]);
                    temp = s[1].LastIndexOf('\\');
                    string tname = s[1].Substring(temp + 1);
                    string tpath = "\\RoadImg\\Camera0" + s[1].Substring(0, temp);
                    int imgmile = (int)Math.Round(Convert.ToDouble(s[0]));

                    if (File.Exists(disfile))
                    {
                        string[] dises = File.ReadAllLines(disfile);
                        foreach (string dis in dises)
                        {
                            try
                            {
                                Disease tdis = new Disease(dis, imgmile);
                                if (tdis.Area > 0)
                                {
                                    tdis.imgname = tname;
                                    tdis.imgpath = tpath;
                                }

                                if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
                                {
                                    tdis.m_mile += direction;
                                }
                                if (tdis.m_mile >= 0)
                                {
                                    disease.Add(tdis);

                                }
                            }
                            catch (Exception)
                            {

                                throw;
                            }

                        }
                    }
                }
                if (direction > 0)
                {
                    disease.Sort(delegate (Disease x, Disease y) { return x.m_mile.CompareTo(y.m_mile); });
                }
                else
                {
                    disease.Sort(delegate (Disease x, Disease y) { return y.m_mile.CompareTo(x.m_mile); });
                }

            }
            return disease;
        }

        public List<ValueTuple<double, double, double>> MidPoint_HighGps(int prj_Index)
        {
            List<ValueTuple<double, double, double>> highGPS_Result = new List<ValueTuple<double, double, double>>();
            SingleProject curPro = _Projects.ElementAt(prj_Index);

            List<HighAccuracyDisease> highAccuracyDiseases = new List<HighAccuracyDisease>();

            List<Disease> diseases = getPorjectAllDiseases(curPro._DataDir.FullName, curPro._ProjectInfo._Direction);
            var prjinfo = curPro._ProjectInfo;
            string outHighGpstxtPath = prjinfo._PrjPath + "/HighGps2Mile.txt";
            if (!File.Exists(outHighGpstxtPath))
            {
                MessageBox.Show("未找到HighGps2Mile.txt！");
                return highGPS_Result;
            }

            List<string> highGpsTxts = File.ReadAllLines(outHighGpstxtPath).ToList();
            List<(double, GPSInfo)> highGpss = new List<(double, GPSInfo)>();
            foreach (var line in highGpsTxts)
            {
                string[] strings = line.Split(',');
                GPSInfo gpsInfo = new GPSInfo();
                gpsInfo._longitude = double.Parse(strings[0]);
                gpsInfo._latitude = double.Parse(strings[1]);
                gpsInfo._elevation = double.Parse(strings[2]);
                highGpss.Add((double.Parse(strings[3]), gpsInfo));
            }
            if (highGpsTxts.Count > 0)
            {
                HighAccuracyPositioning.UpdateAllImg(prjinfo._PrjPath + "\\RoadImg\\Camera0");
                foreach (Disease tdis in diseases)
                {
                    HighAccuracyDisease dis = new HighAccuracyDisease
                    {
                        DiseaseName = tdis.RoadDisType
                    };
                    if (string.IsNullOrEmpty(dis.DiseaseName))
                    {
                        continue;
                    }

                    // 疑问： 中心点坐标不一定为整形，但未找到传入double的方法
                    //double half_x = tdis.rect.X + (double)tdis.rect.Width / 2;
                    //double half_y = tdis.rect.Y + (double)tdis.rect.Height / 2;
                    int half_x = tdis.rect.X + tdis.rect.Width / 2;
                    int half_y = tdis.rect.Y + tdis.rect.Height / 2;
                    //var point = points[index];
                    double dDiseaseLon = 0, dDiseaseLat = 0, dDiseaseH = 0; //当前像素
                    HighAccuracyPositioning.getHighAccPosition(_Setting.gpsformat, highGpss, _Setting.equipType, prjinfo._PrjPath, tdis.m_mile, half_x, half_y, _RoadConfig.ImageWidth
                        , _RoadConfig.ImageHeight, prjinfo._Direction, _RoadConfig.RealWidth, _RoadConfig.RealHeight,
                            ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH);
                    highGPS_Result.Add((dDiseaseLon, dDiseaseLat, dDiseaseH));

                }
            }
            return highGPS_Result;

        }

        private void barButtonItem53_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Setting.SelectDrawDis == 1)
            {
                MessageBox.Show("仅支持大框模式导出!");
                return;
            }
            if (!_Setting.showGpsInfoToPicture)
            {
                MessageBox.Show("请选中软件上方【显示定位】选项后,重新点击GPS桩号匹配后重试\n仅支持具备高精度定位模块的工程");
                return;
            }


            VistaFolderBrowserDialog fd = new VistaFolderBrowserDialog()
            {
          
                Description = "选择文件夹"
            };

            if (fd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string outPath;
            HighAccuracySettingForm form = new HighAccuracySettingForm();


            form.StartPosition = FormStartPosition.CenterParent;
            if (form.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            //获得窗口参数

            POS_CONVERT_INFO info = form.getUserSelectInfo();
            List<string> debugText = new List<string>();
            for (int i = 0; i < _Projects.Count; i++)
            {
                SingleProject curPro = _Projects.ElementAt(i);

                outPath = fd.SelectedPath + "\\" + curPro._DataDir.Name + "_高精度定位病害图.dxf";
                List<HighAccuracyDisease> highAccuracyDiseases = new List<HighAccuracyDisease>();

                List<Disease> diseases = getPorjectAllDiseases(curPro._DataDir.FullName, curPro._ProjectInfo._Direction);
                var prjinfo = curPro._ProjectInfo;


                int ptrSize = 0, diseaseCount = 0;
                IntPtr diseaseInfoPtr = default;
                IntPtr ptr = default;
                string outHighGpstxtPath = prjinfo._PrjPath + "/HighGps2Mile.txt";
                if (File.Exists(outHighGpstxtPath))
                {
                    List<string> highGpsTxts = File.ReadAllLines(outHighGpstxtPath).ToList();
                    List<(double, GPSInfo)> highGpss = new List<(double, GPSInfo)>();
                    foreach (var line in highGpsTxts)
                    {
                        string[] strings = line.Split(',');
                        GPSInfo gpsInfo = new GPSInfo();
                        gpsInfo._longitude = double.Parse(strings[0]);
                        gpsInfo._latitude = double.Parse(strings[1]);
                        gpsInfo._elevation = double.Parse(strings[2]);
                        highGpss.Add((double.Parse(strings[3]), gpsInfo));
                    }
                    if (highGpsTxts.Count > 0)
                    {
                        HighAccuracyPositioning.UpdateAllImg(prjinfo._PrjPath + "\\RoadImg\\Camera0");
                        for (int dd = 0; dd < diseases.Count; dd++)
                        {

                            Disease tdis = diseases[dd];
                            HighAccuracyDisease dis = new HighAccuracyDisease();
                            dis.DiseaseName = tdis.RoadDisType;
                            if (string.IsNullOrEmpty(dis.DiseaseName))
                            {
                                continue;
                            }
                            List<System.Drawing.Point> points = new List<System.Drawing.Point>()
                                { 
                                //四个点 
                                 new System.Drawing.Point (tdis.rect.X, tdis.rect.Y),
                                  new System.Drawing.Point(tdis.rect.X + tdis.rect.Width, tdis.rect.Y),
                                  new System.Drawing.Point(tdis.rect.X + tdis.rect.Width, tdis.rect.Y + tdis.rect.Height),
                                 new System.Drawing.Point(tdis.rect.X, tdis.rect.Y + tdis.rect.Height)
                                };

                            for (int index = 0; index < points.Count; index++)
                            {
                                var point = points[index];
                                double dDiseaseLon = 0, dDiseaseLat = 0, dDiseaseH = 0; //当前像素
                                HighAccuracyPositioning.getHighAccPosition(true, highGpss, _Setting.equipType, prjinfo._PrjPath, tdis.m_mile, point.X, point.Y, _RoadConfig.ImageWidth
                                    , _RoadConfig.ImageHeight, prjinfo._Direction, _RoadConfig.RealWidth, _RoadConfig.RealHeight,
                                      ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH);

                                ptrSize = Marshal.SizeOf(typeof(POS_CONVERT_INFO));

                                // 分配内存来保存疾病信息结构体的指针数组
                                diseaseCount = 1;
                                diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
                                // 获取指向指针数组的指针
                                ptr = diseaseInfoPtr;
                                // 将每个疾病信息结构体转换为指针，并存储到指针数组中
                                Marshal.StructureToPtr(info, ptr, false);
                                initialParam(diseaseInfoPtr);

                                double dEast = 0;
                                double dNorth = 0;
                                double dHeight = 0;

                                //测试代码
                                //double dL = 111.5693797675;
                                //double dB = 39.9577920059;
                                //double dH = 1041.949463;
                                //// 验证84;
                                //double dEast84 = 497383.1583;
                                //double dNorth84 = 4425608.6447;
                                //double dH84 = 1041.9495;

                                //// 使用七参数验证2000;
                                //double dEast2000 = 497383.2083;
                                //double dNorth2000 = 4425609.7709;
                                //double dH2000 = 1063.2585;
                                //convertBLHToProjection(dL, dB , dH , ref dEast,ref dNorth, ref dHeight);


                                convertBLHToProjection(dDiseaseLon, dDiseaseLat, dDiseaseH, ref dEast, ref dNorth, ref dHeight);

                                // debugText.Add(string.Format("{0},{1},{2},{3},{4}", tdis.m_mile, dis.DiseaseName, dEast, dNorth, dHeight));
                                if (dEast < 0 || dNorth < 0)
                                {

                                }
                                // 释放分配的内存
                                Marshal.FreeHGlobal(diseaseInfoPtr);

                                switch (index)
                                {
                                    case 0:
                                        dis.HighAccureacyGpsP0 = new HighAccureacyGps
                                         (dEast, dNorth, dHeight);

                                        break;
                                    case 1:
                                        dis.HighAccureacyGpsP1 = new HighAccureacyGps
                                         (dEast, dNorth, dHeight);
                                        break;
                                    case 2:
                                        dis.HighAccureacyGpsP2 = new HighAccureacyGps
                                        (dEast, dNorth, dHeight);
                                        break;
                                    case 3:
                                        dis.HighAccureacyGpsP3 = new HighAccureacyGps
                                         (dEast, dNorth, dHeight);
                                        break;

                                    default:
                                        break;
                                }


                            }

                            highAccuracyDiseases.Add(dis);

                        }

                        // File.WriteAllLines("D:\\DiseaseMsg.txt", debugText);
                        diseaseCount = highAccuracyDiseases.Count;

                        ptrSize = Marshal.SizeOf(typeof(HighAccuracyDisease));
                        // 分配内存来保存疾病信息结构体的指针数组
                        diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
                        // 获取指向指针数组的指针
                        ptr = diseaseInfoPtr;
                        // 将每个疾病信息结构体转换为指针，并存储到指针数组中
                        for (int m = 0; m < diseaseCount; m++)
                        {
                            Marshal.StructureToPtr(highAccuracyDiseases[m], ptr, false);
                            ptr = (IntPtr)(long)ptr + ptrSize;
                        }

                        if (HighAccOut(outPath, diseaseCount, diseaseInfoPtr))
                        {

                        }
                        else
                        {
                            MessageBox.Show("工程文件输出错误，请检查!\n" + outPath);
                        }


                        // 释放分配的内存
                        Marshal.FreeHGlobal(diseaseInfoPtr);
                    }
                }

            }

            MessageBox.Show("输出完毕,请检查！");

        }

        private void barButtonItem54_ItemClick(object sender, ItemClickEventArgs e)
        {
            VistaFolderBrowserDialog fd = new VistaFolderBrowserDialog()
            {
             
                Description = "选择文件夹"
            };

            if (fd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            // 遍历文件夹中的所有文件
            foreach (string file in Directory.GetFiles(fd.SelectedPath, "*.txt", SearchOption.AllDirectories))
            {
                try
                {
                    // 读取文件内容
                    string content = File.ReadAllText(file, Encoding.UTF8);

                    // 将内容转换为GBK编码
                    byte[] gbkBytes = Encoding.Convert(Encoding.UTF8, Encoding.GetEncoding("GBK"), Encoding.UTF8.GetBytes(content));
                    string gbkContent = Encoding.GetEncoding("GBK").GetString(gbkBytes);

                    // 将转换后的内容写回原文件
                    File.WriteAllText(file, gbkContent, Encoding.GetEncoding("GBK"));
                    //Console.WriteLine($"文件 {file} 已转换为GBK编码。");
                }
                catch (Exception ex)
                {
                    string msg = $"处理文件 {file} 时发生错误: {ex.Message}";
                    throw new Exception(msg);
                }
            }
            MessageBox.Show("所有文件转换完成。");
        }

        private void barButtonItem55_ItemClick(object sender, ItemClickEventArgs e)
        {
            MessageBox.Show("即将启动资产表数据库生成软件，请将生成或修改后的propertyInfo.db复制到【二维内业数据处理软件】运行目录下，以供进行资产表操作。");

            string helpfile = System.Windows.Forms.Application.StartupPath + "\\ExcelDataToSqllite\\ExcelDataToSqllite.exe";
            System.Diagnostics.Process.Start(helpfile);
        }

        private void dockPanel_main_data_Click_1(object sender, EventArgs e)
        {

        }

        private void barButtonItem56_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog() { Description = "请选择报表放置位置：" };
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }
                string outdirpath = null;
                {
                    foreach (SingleProject proj in _Projects)
                    {


                        outdirpath = fd.SelectedPath;
                        {
                            string smile = proj._ProjectInfo._StartMile.ToString("K0000+000");
                            string emile = proj._ProjectInfo._EndMile.ToString("K0000+000");

                            string tt = "\\" + proj._DataDir.Name;
                            var index = tt.LastIndexOf('_') - 9;
                            if (index <= 0)
                            {
                                log.Error($"{proj._DataDir.Name}工程数据文件夹名称错误");
                            }

                            tt = tt.Remove(tt.LastIndexOf('_') - 9);//  例如减去_20190212  9个字符
                            outdirpath = string.Format("{0}{1}({2}~{3})", fd.SelectedPath, tt, smile, emile);


                            if (!Directory.Exists(outdirpath))
                            {
                                Directory.CreateDirectory(outdirpath);
                            }
                            string outFilePath = outdirpath + $"//{proj._DataDir.Name}_景观自定义信息表.xlsx";
                            //获取数据
                            string streetPath = proj._DataDir.FullName + "\\StreetImg";
                            DirectoryInfo dir = new DirectoryInfo(streetPath);
                            FileInfo[] files = dir.GetFiles("*_UserSign.txt", SearchOption.AllDirectories);
                            List<string> dataInfos = new List<string>();
                            foreach (var item in files)
                            {
                                dataInfos.AddRange(File.ReadAllLines(item.FullName));
                            }

                            // 创建工作簿
                            IWorkbook workbook;
                            workbook = new XSSFWorkbook(); // 创建 .xlsx 工作簿
                                                           // 创建工作表
                            ISheet sheet = workbook.CreateSheet("Sheet1");
                            // 创建单元格样式（带边框）
                            ICellStyle cellStyle = workbook.CreateCellStyle();
                            cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin; // 上边框
                            cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin; // 下边框
                            cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin; // 左边框
                            cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin; // 右边框

                            //创建表头
                            int colIndex = 0;
                            IRow row0 = sheet.CreateRow(0);
                            row0.CreateCell(colIndex++).SetCellValue("路线编码");
                            row0.CreateCell(colIndex++).SetCellValue("检测方向");
                            row0.CreateCell(colIndex++).SetCellValue("检测车道");
                            row0.CreateCell(colIndex++).SetCellValue("标记点桩号(km)");
                            row0.CreateCell(colIndex++).SetCellValue("标记点名称");
                            row0.CreateCell(colIndex++).SetCellValue("备注");
                            row0.CreateCell(colIndex++).SetCellValue("经度");
                            row0.CreateCell(colIndex++).SetCellValue("纬度");
                            row0.CreateCell(colIndex++).SetCellValue("高程");
                            // 为表头单元格设置样式
                            for (int i = 0; i < 9; i++)
                            {
                                row0.GetCell(i).CellStyle = cellStyle;
                            }
                            // 将 List<string> 写入 Excel
                           
                            string[] gpsinfostrs;
                            ExcelGPS[] GPSInfos = null;
                            if (File.Exists(proj._DataDir.FullName + "\\GPS2Mile.txt"))
                            {
                                gpsinfostrs = File.ReadAllLines(proj._DataDir.FullName + "\\GPS2Mile.txt");
                                GPSInfos = new ExcelGPS[gpsinfostrs.Length];
                                for (int i = 0; i < gpsinfostrs.Length; ++i)
                                {
                                    GPSInfos[i] = new ExcelGPS(gpsinfostrs[i]);
                                }
                            }
                            else
                            {
                                MessageBox.Show($"工程{proj._DataDir.Name}获取经纬度信息失败！");
                                continue;
                            }
                            for (int i = 1; i < dataInfos.Count + 1; i++)
                            {
                                colIndex = 0;
                                int curMile = 0;

                                IRow row = sheet.CreateRow(i); // 创建行
                                string[] info = dataInfos[i-1].Split(' ',',');
                                if (info.Length>=7)
                                {
                                    //对于带有坐标的信息直接跳过
                                    continue;
                                }
                                if (info.Length < 3)
                                { 
                                    continue;
                                }
                                bool newDis = info.Length == 4 ? true : false;
                                 
                                int qian = int.Parse(info[0].Substring(1, info[0].Length - 1).Split('+')[0]);
                                int bai = int.Parse(info[0].Substring(1, info[0].Length - 1).Split('+')[1]);
                                curMile = qian * 1000 + bai;
                                ExcelGPS closest = GPSInfos
                                                    .OrderBy(gps => Math.Abs(gps._mile - curMile)).First();
                                // 设置每一列的值
                                SetCellValue(row, colIndex++, proj._ProjectInfo._RoadCode, cellStyle);
                                SetCellValue(row, colIndex++, proj._ProjectInfo._Direction == 1 ? "上行" : "下行", cellStyle);
                                SetCellValue(row, colIndex++, proj._ProjectInfo._RoadNum, cellStyle);
                                if (newDis)
                                {
                                    SetCellValue(row, colIndex++, info[0], cellStyle);
                                    SetCellValue(row, colIndex++, info[1], cellStyle);
                                    SetCellValue(row, colIndex++, info[3], cellStyle);
                                }
                                else
                                {
                                    SetCellValue(row, colIndex++, info[0], cellStyle);
                                    SetCellValue(row, colIndex++, info[1], cellStyle);
                                    SetCellValue(row, colIndex++, info[2], cellStyle);
                                }
                                   
                                SetCellValue(row, colIndex++, closest._longitude , cellStyle);
                                SetCellValue(row, colIndex++, closest._latitude , cellStyle);
                                SetCellValue(row, colIndex++, closest._elevation , cellStyle); 

                            }

                            // 保存文件
                            using (FileStream fs = new FileStream(outFilePath, FileMode.Create, FileAccess.Write))
                            {
                                workbook.Write(fs);
                            }

                        }


                    }
                }
                MessageBox.Show("导出报表完成！");
            }
        }
        // 辅助方法：设置单元格值
        private static void SetCellValue(IRow row, int columnIndex, object value, ICellStyle cellStyle)
        {
            ICell cell = row.CreateCell(columnIndex); // 创建单元格
            if (value is string)
            {
                cell.SetCellValue((string)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)value);
            }
            else if (value is double)
            {
                cell.SetCellValue((double)value);
            }
            // 其他类型...
            cell.CellStyle = cellStyle; // 设置样式
        }

        private void barButtonItem58_ItemClick(object sender, ItemClickEventArgs e)
        {
            string helpfile = System.Windows.Forms.Application.StartupPath + "\\HNRoadFormatConverter.exe";
            System.Diagnostics.Process.Start(helpfile);
        }

        private void barButtonItem58_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            // 1. 在方法内部直接动态创建一个 DevExpress 弹窗
            using (XtraForm inputForm = new XtraForm())
            {
                inputForm.Text = "修改天地图密钥";
                inputForm.Size = new Size(500, 250);
                inputForm.StartPosition = FormStartPosition.CenterParent;
                inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                inputForm.MaximizeBox = false;
                inputForm.MinimizeBox = false;

                LabelControl lbl = new LabelControl() { Text = "请输入新的天地图浏览器端开发密钥 (TK):", Location = new System.Drawing.Point(20, 20) };
                TextEdit txtToken = new TextEdit() { Location = new System.Drawing.Point(20, 50), Width = 320 };
                SimpleButton btnOk = new SimpleButton() { Text = "确定", Location = new System.Drawing.Point(170, 85), DialogResult = DialogResult.OK };
                SimpleButton btnCancel = new SimpleButton() { Text = "取消", Location = new System.Drawing.Point(260, 85), DialogResult = DialogResult.Cancel };

                inputForm.Controls.Add(lbl);
                inputForm.Controls.Add(txtToken);
                inputForm.Controls.Add(btnOk);
                inputForm.Controls.Add(btnCancel);
                inputForm.AcceptButton = btnOk;
                inputForm.CancelButton = btnCancel;

                // 2. 显示弹窗并等待用户操作
                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    string newToken = txtToken.Text.Trim();

                    // 3. 校验输入
                    if (string.IsNullOrWhiteSpace(newToken))
                    {
                        XtraMessageBox.Show("密钥不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    try
                    {
                        // 4. 获取 HTML 文件路径
                        string mapHtmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Map", "tianditu", "tianditu.html");

                        if (!File.Exists(mapHtmlPath))
                        {
                            XtraMessageBox.Show("未找到地图文件：" + mapHtmlPath, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // 5. 读取并正则替换
                        string htmlTxts = File.ReadAllText(mapHtmlPath);
                        string pattern = @"tk=[a-zA-Z0-9]+";
                        string replacement = "tk=" + newToken;

                        if (Regex.IsMatch(htmlTxts, pattern))
                        {
                            string updatedHtml = Regex.Replace(htmlTxts, pattern, replacement);

                            // 6. 写回文件
                            File.WriteAllText(mapHtmlPath, updatedHtml);

                            XtraMessageBox.Show("地图密钥更新成功！请刷新或重新加载地图界面。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            // TODO: 如果你想让地图自动刷新，可以在这里调用你浏览器控件的 Reload 方法
                            // 例如: webView.Reload(); 
                        }
                        else
                        {
                            XtraMessageBox.Show("在 HTML 中未找到密钥配置标识（tk=），请检查 HTML 代码。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        XtraMessageBox.Show("更新失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void barButtonItem59_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                int toProcess = _Projects.Count;
                if (toProcess <= 0)
                {
                    return;
                }
                foreach (SingleProject pro in _Projects)
                {
                    var basePath = Path.Combine(pro._DataDir.FullName, "RoadImg", "Camera0");
                    string outPath = Path.Combine(pro._DataDir.FullName, "RoadImg", "Camera0", "AllRoadBigDis.txt");
                    if (_Setting.SelectDrawDis == 1)
                        outPath = Path.Combine(pro._DataDir.FullName, "RoadImg", "Camera0", "AllRoadSmallDis.txt");
                    if (File.Exists(outPath))
                    {
                        File.Delete(outPath);
                    }
                    DirectoryInfo disDir = new DirectoryInfo(basePath);
                    if (_Setting.SelectDrawDis == 0)
                    {
                        
                        disDir.GetFiles("*.jpg.txt", SearchOption.AllDirectories).ToList().ForEach(f =>
                        {
                            string fileName = f.Directory.FullName + "\\" + f.Name.Replace(".txt","");
                            var lines = File.ReadAllLines(f.FullName);
                            List<string> outLines = new List<string>();
                            for (int i = 0; i < lines.Length; i++)
                            {
                                if (string.IsNullOrWhiteSpace( lines[i]))
                                {
                                    continue;
                                }  
                                outLines.Add(fileName);
                                outLines.Add(lines[i]);
                            }
                            File.AppendAllLines(outPath, outLines);


                        });
                    }
                    else
                    {
                        disDir.GetFiles("*.jpg_PartClass.txt", SearchOption.AllDirectories).ToList().ForEach(f =>
                        {
                            
                                string fileName = f.Directory.FullName + "\\" + f.Name.Replace("_PartClass.txt", "");
                                var lines = File.ReadAllLines(f.FullName);
                                List<string> outLines = new List<string>();
                                for (int i = 0; i < lines.Length; i++)
                                {
                                    if (string.IsNullOrWhiteSpace(lines[i]))
                                    {
                                        continue;
                                    }
                                    string line = fileName + "\n" + lines[i];

                                    outLines.Add(fileName);
                                    outLines.Add(lines[i]);
                                }
                                File.AppendAllLines(outPath, outLines); 

                        });
                    }
                    disDir.GetFiles();


                    
                } 
                string text = "/RoadImg"+ "/Camera0"+ "/AllRoadBigDis.txt";
                if (_Setting.SelectDrawDis ==1 )
                {
                      text = "/RoadImg" + "/Camera0" + "/AllRoadSmallDis.txt";

                }
                MessageBox.Show("所有路面病害输出完毕!保存地址: " + text);
            }
            catch (Exception)
            {

                throw;
            }
        }
    }

}
