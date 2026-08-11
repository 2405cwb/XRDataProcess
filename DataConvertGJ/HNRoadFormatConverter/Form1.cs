
using DevExpress.Portable.Input;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraTreeList.Nodes;
using Farmework.Other;
using Farmework.Other.enumTools;
using HNRoadFormatConverter.Commons;
using HNRoadFormatConverter.Entitys;
using HNRoadFormatConverter.Exporters;
using HNRoadFormatConverter.MyConfig;
using HNRoadFormatConverter.MyEntitys;
using HNRoadFormatConverter.toolForms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Ookii.Dialogs.WinForms; 
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
namespace HNRoadFormatConverter
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public string _Suf = ".jpg"; // 文件名后缀
        private static Config _config = ConfigManager.GetConfig();
        //   private static LogHelper _log = new LogHelper(typeof(Form1));
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        private static extern int StrCmpLogicalW(string psz1, string psz2);
        private string Chktxt_Path { get; set; }


        private bool dataFrom2D = true;
        public Form1(string source)
        {
            InitializeComponent();
            AddUI();
            SetDefaultUI();
            // HandelMap();
            _Projects = new List<ProjectInfo>();

            setFromType(source);


            Bar1process = new Progress<int>(value =>
            {
                if (value == -1)
                {
                    return;
                }
                else
                {
                    progressBar1.BeginInvoke(() =>
                    {
                        progressBar1.Value = value;
                        double value1 = (value / progressBar1.Maximum) * 100;
                        label2.Text = value1.ToString() + "%";
                    });
                }
            });

        }

        private static Progress<int> Bar1process;

        private void setFromType(string source)
        {
            this.Text = $"主窗口 - 来源：{source ?? "未选择"}";
            if (source.Contains("二三维内业数据处理软件"))
            {
                dataFrom2D = false;
            }
            else
            {
                dataFrom2D = true;

            }
            //if (dataFrom2D)
            //{
            //    ribbonPage1.Groups.RemoveAt(1);
            //}
            //else
            //{
            //    ribbonPage1.Groups.RemoveAt(0);
            //}

            // 方法1：用两个页（最清晰）
            ribbonPageGroup1.Enabled = dataFrom2D;
            ribbonPageGroup9.Enabled = !dataFrom2D;


        }
        //

        string userSettingsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "夕睿光电",
                "国检转换软件",
                "Settings"
            );
       string defaultLayoutPath =  Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RibbonSettings.xml");

        private void Form1_Load(object sender, EventArgs e)
        {
            string ribbonPath = Path.Combine(userSettingsPath, "RibbonSettings.xml");

         

            if (File.Exists(ribbonPath))
            {
                try
                {
                    ribbonControl1.RestoreLayoutFromXml(ribbonPath);
                }// ribbonControl1 是你的 Ribbon 实例名 }
                catch { /* 损坏的布局直接忽略 */ }
            }
            else
            {
                string defaultPath = defaultLayoutPath;
                if (File.Exists(defaultPath))
                {
                    try { ribbonControl1.RestoreLayoutFromXml(defaultPath); }
                    catch { /* 损坏的布局直接忽略 */ }
                }
            }


            string files = ribbonControl1.AutoSaveLayoutToXmlPath;
            barButtonItem21.Visibility = BarItemVisibility.Always;

            barButtonItem15.Caption = "交通部数据";
            //  barLoad.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            barButtonItem3.Visibility = DevExpress.XtraBars.BarItemVisibility.Never;
            barLoad.Visibility = DevExpress.XtraBars.BarItemVisibility.Always;
            barEditItem1.EditValue = _config.NowModel;
            barEditItem1.EditValueChanged += BarEditItem1_EditValueChanged;

            // 创建 RepositoryItemComboBox 并设置为 BarEditItem 的编辑控件
            RepositoryItemComboBox repositoryItemComboBox = barEditItem1.Edit as RepositoryItemComboBox;
            repositoryItemComboBox.Items.Clear();
            // 向 RepositoryItemComboBox 添加子项
            repositoryItemComboBox.Items.Add("等级公路5210与农村路5211标准模板导出_2025年");
            repositoryItemComboBox.Items.Add("农养国省道路况检测数据提交格式_2026年");
            repositoryItemComboBox.Items.Add("河南省单位一农村路定制");
            repositoryItemComboBox.Items.Add("河北省单位定制");
            repositoryItemComboBox.Items.Add("辽宁省2025单位定制");
            //repositoryItemComboBox.Items.Add("湖南省单位一定制");
            //repositoryItemComboBox.Items.Add("重庆市单位一定制");
            //repositoryItemComboBox.Items.Add("甘肃省单位一定制");
            //repositoryItemComboBox.Items.Add("河北省单位一定制");
            //repositoryItemComboBox.Items.Add("河北省单位二定制");
            //repositoryItemComboBox.Items.Add("江苏省单位一定制");
            //repositoryItemComboBox.Items.Add("安徽省单位一定制");
            //repositoryItemComboBox.Items.Add("广东省单位一定制");
            // ===== 关键：禁用 DevExpress 自动保存到安装目录 =====
            ribbonControl1.AutoSaveLayoutToXml = false;  // 这一行必须加！
                                                           // 或者更彻底（推荐两者都加）：
          
        }
        private void HandelMap()
        {
            string str_uil = AppDomain.CurrentDomain.BaseDirectory + "Map\\BaiduMap.html";
            Uri uri = new Uri(str_uil);
            //myWebBrowser1.webBrowser1.Url = uri;
            // myWebBrowser1.webBrowser1.ObjectForScripting = true;
        }
        private void AddUI()
        {
            var skinCon = new DevExpress.XtraBars.SkinRibbonGalleryBarItem();
            ribbonPageGroup5.ItemLinks.Add(skinCon);
            SkinHelper.InitSkinGallery(skinCon);
        }
        private void SetDefaultUI()
        {
            this.defaultLookAndFeel1.LookAndFeel.SkinName = _config.DefaultSkin;


        }
        private List<ProjectInfo> _Projects;
        private void barLoad_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            getProject(false);

        }
        /// <summary>
        /// 获取所有文件
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private List<DirectoryInfo> GetAllProjectPath(string path)
        {
            List<DirectoryInfo> projects = new List<DirectoryInfo>();
            _Projects.Clear();
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


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                ribbonControl1.SaveLayoutToXml(Path.Combine(userSettingsPath, "RibbonSettings.xml"));
            }
            catch { /* 保存失败忽略，避免权限异常崩溃程序 */ }
            _config.DefaultSkin = this.defaultLookAndFeel1.LookAndFeel.SkinName;
            ConfigManager.SaveConfig();
        }

        private void ribbonStatusBar1_Click(object sender, EventArgs e)
        {

        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {


        }

        private void btn_check_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }


        private CityModelItem currentStandard = CityModelItem.等级公路5210与农村路5211标准模板导出_2025年;
        private async void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            handelCsvFiles.Clear();
            string targetBasePath = "";
            if (_Projects.Count <= 0)
            {
                MessageBox.Show("您还没有导入任何合格工程");
                return;
            }
            if (_Projects.Count > 1)
            {
                if ((CityModelItem)Enum.Parse(typeof(CityModelItem), barEditItem1.EditValue.ToString()) == CityModelItem.湖南省单位一定制)
                {
                    MessageBox.Show("请注意目前同样【道路编号及上下行仅桩号不同的路段】，不可同时处理导入到同一个路径下，否则会发生图片文件被覆盖的问题\n此类工程需要单个导出到各自独立的文件夹防止出现覆盖问题");

                }
            }
            VistaFolderBrowserDialog dlg;
            if (Directory.Exists(_config.UserPath))
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",

                    SelectedPath = _config.UserPath,
                    ShowNewFolderButton = true
                };
            }
            else
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",


                    ShowNewFolderButton = true
                };
            }


            if (dlg.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            if (string.IsNullOrEmpty(dlg.SelectedPath))
            {
                return;
            }
            targetBasePath = dlg.SelectedPath;

            currentStandard = (CityModelItem)Enum.Parse(typeof(CityModelItem), barEditItem1.EditValue.ToString());
            int ImgCnt = 0; // 总待处理图片数量

            // 将索引目录按5000大小 分区后，存入缓冲队列
            List<ValueTuple<DirectoryInfo, List<List<PicAndMile>>>> streePic_CacheQueue = new();
            List<ValueTuple<DirectoryInfo, List<List<PicAndMile>>>> roadImage_CacheQueue = new();
            List<National2026PictureExportTask> national2026PictureTasks = new();
            
             
                
            for (int proIndex = 0; proIndex < _Projects.Count; proIndex++)
            {
                ProjectInfo pro = _Projects[proIndex];

                string outPath;
                if (currentStandard == CityModelItem.江苏省单位一定制)
                {
                    outPath = targetBasePath + $"\\{pro._City + pro._District}";
                }
                else if (currentStandard == CityModelItem.农养国省道路况检测数据提交格式_2026年)
                {
                    outPath = National2026ExportService.BuildExportDataPath(pro, targetBasePath);
                }
                else
                {
                    outPath = targetBasePath + "\\结果数据";
                }

                if (currentStandard == CityModelItem.农养国省道路况检测数据提交格式_2026年)
                {
                    _Suf = ".jpg";
                    Directory.CreateDirectory(outPath);
                    if (!National2026ExportService.ExportMetricFiles(pro, outPath, handelCsvFiles, out string errorMessage))
                    {
                        MessageBox.Show(errorMessage);
                        continue;
                    }
                    National2026ExportService.ExportProjectInfoCsv(pro, outPath, handelCsvFiles);

                    if (chebtn.Checked)
                    {
                        if (!string.IsNullOrWhiteSpace(pro.RoadPicPath))
                        {
                            National2026PictureExportTask roadTask =
                                National2026ExportService.CreatePictureTask(pro, outPath, true);
                            ImgCnt += roadTask.Count;
                            national2026PictureTasks.Add(roadTask);
                        }

                        if (!string.IsNullOrWhiteSpace(pro.StreetPicPath))
                        {
                            National2026PictureExportTask frontTask =
                                National2026ExportService.CreatePictureTask(pro, outPath, false);
                            ImgCnt += frontTask.Count;
                            national2026PictureTasks.Add(frontTask);
                        }
                    }

                    continue;
                }

                string outPathResult = outPath + "\\" + pro.ConvertProName;
                string sourcePath = pro.ConvertPath.FullName;
                FileHelpter.CopyFileAndDir(sourcePath, outPathResult); // 第二次调用时，路径sourcePath不更新，仍会使用旧项目的文件。如果删除则找不到文件
                DirectoryInfo info = new DirectoryInfo(outPathResult);

                DirectoryInfo tempPath;
                //景观图像存储路径
                DirectoryInfo streePic = null;
                DirectoryInfo roadImage = null;

                void DeleteEmptyFolder(string path)
                {
                    foreach (var d in Directory.GetDirectories(path))
                    {
                        try { Directory.Delete(d); }
                        catch (Exception ex)
                        {
                        }
                    }

                }
                // 处理不同的格式要求
                switch (currentStandard)
                {
                    case CityModelItem.河南省单位一农村路定制:
                        if (chebtn.Checked)
                        {
                            streePic = info.GetDirectories("景观图像").First().GetDirectories(pro._DataDate).First().GetDirectories(pro.ConvertProName + "_" + pro._DataTime, SearchOption.AllDirectories).First();
                            roadImage = info.GetDirectories("Images").First().GetDirectories(pro._DataDate).First().GetDirectories(pro.ConvertProName + "_" + pro._DataTime, SearchOption.AllDirectories).First();
                        }
                        else
                        {
                            //删除影像文件夹
                            string temp = outPathResult + "\\Images";
                            string temp1 = outPathResult + "\\景观图像";
                            Directory.Delete(temp, true);
                            Directory.Delete(temp1, true);
                        }
                        break;
                    case CityModelItem.湖南省单位一定制:
                    case CityModelItem.农养国省道路况检测数据提交格式_2026年:
                        if (chebtn.Checked)
                        {
                            tempPath = info.GetDirectories("前方图像").First();
                            streePic = tempPath.GetDirectories().First().GetDirectories().First();
                            roadImage = info.GetDirectories("Images").First().GetDirectories("0", SearchOption.AllDirectories).First();
                        }
                        else
                        {
                            //删除影像文件夹
                            string temp = outPathResult + "\\Images";
                            string temp1 = outPathResult + "\\前方图像";
                            Directory.Delete(temp, true);
                            Directory.Delete(temp1, true);
                        }
                        break;
                    case CityModelItem.江苏省单位一定制:
                        if (chebtn.Checked)
                        {
                            _Suf = ".jpeg";

                            streePic = info.GetDirectories("ViewImages").First().GetDirectories(pro._DataDate + pro._DataTime, SearchOption.AllDirectories).First();
                            roadImage = info.GetDirectories("Images").First().GetDirectories(pro._DataDate + pro._DataTime, SearchOption.AllDirectories).First();
                        }
                        else
                        {
                            //删除影像文件夹
                            string temp = outPathResult + "\\Images";
                            string temp1 = outPathResult + "\\ViewImage";
                            Directory.Delete(temp, true);
                            Directory.Delete(temp1, true);
                        }
                        break;
                    default:
                        if (chebtn.Checked)
                        {
                            streePic = info.GetDirectories("ViewImages").First().GetDirectories(pro._DataDate + pro._DataTime, SearchOption.AllDirectories).First();
                            roadImage = info.GetDirectories("Images").First().GetDirectories(pro._DataDate + pro._DataTime, SearchOption.AllDirectories).First();
                        }
                        else
                        {
                            //删除影像文件夹
                            string temp = outPathResult + "\\Images";
                            string temp1 = outPathResult + "\\ViewImages";
                            Directory.Delete(temp, true);
                            Directory.Delete(temp1, true);
                        }
                        break;
                }
                DeleteEmptyFolder(info.FullName);
                if (chebtn.Checked)
                {
                    if (!string.IsNullOrWhiteSpace(pro.RoadPicPath))
                    { 
                        var (cnt, roadImageIdx_5000_split) = ConvertPci(pro, roadImage, true);
                        ImgCnt += cnt;
                        roadImage_CacheQueue.Add((roadImage, roadImageIdx_5000_split));
                    }
                    else
                    {
                        Console.WriteLine((proIndex, roadImage));
                    }
                    if (!string.IsNullOrWhiteSpace(pro.StreetPicPath))
                    {
                        
                        var (cnt, streePicIdx_5000_split) = ConvertPci(pro, streePic, false);
                        ImgCnt += cnt;
                        streePic_CacheQueue.Add((streePic, streePicIdx_5000_split));
                    }
                    else
                    {
                        Console.WriteLine((proIndex, streePic));
                    }
                } 
                FileInfo[] indexFiles = info.GetFiles("fileindex.txt", SearchOption.AllDirectories);
                foreach (FileInfo file in indexFiles)
                {
                    var strings = File.ReadAllLines(file.FullName).ToList();
                    for (int i = strings.Count - 1; i >= 0; i--)
                    {
                        string temp = strings[i];
                        if (string.IsNullOrEmpty(temp))
                        {
                            strings.RemoveAt(i);
                        }
                    }
                    SaveFile(file.FullName, strings.ToArray(), false);
                }
                DirectoryInfo[] hasDirs = info.GetDirectories("Text", SearchOption.AllDirectories);
                ConvertToCity(info, pro);
                var allFiles = info.GetFiles("*", SearchOption.AllDirectories);
                foreach (var file in allFiles)
                {
                    handelCsvFiles.Add(file.FullName);
                }

            }

            // 进度条 最大值设置为图片总数
            this.progressBar1.Value = 0;
            this.progressBar1.Minimum = 0;
            label2.Text = "0%";
            this.progressBar1.Maximum = Math.Max(ImgCnt, 1);
            this.Process = new Progress<int>(


                Value =>

                {
                    if (ImgCnt <= 0)
                    {
                        return;
                    }

                    progressBar1.Value += Value;
                    int value = progressBar1.Value * 100 / ImgCnt;
                    label2.Text = value.ToString() + "%";
                }

                );

            foreach (National2026PictureExportTask task in national2026PictureTasks)
            {
                await Task.Run(() => National2026ExportService.ExportPictures(task, this.Process, _Suf));
            }

            //输出图片
            for (int proIndex = 0; proIndex < _Projects.Count; proIndex++)
            {
                if (proIndex < streePic_CacheQueue.Count)
                {
                    var (streePic, streePic_spList) = streePic_CacheQueue[proIndex];
                    await MovePicutre(streePic, streePic_spList, currentStandard);
                }
                if (proIndex < roadImage_CacheQueue.Count)
                {
                    var (roadImage, roadImage_spList) = roadImage_CacheQueue[proIndex];
                    await MovePicutre(roadImage, roadImage_spList, currentStandard);
                }

            }

            //所有文件合并
            MessageBox.Show("所有项目处理完毕请检查!");
        }



        private static List<string> handelCsvFiles = new List<string>(); //已经处理过的文件
        /// <summary>
        /// 转换结果文件符合各地标准
        /// </summary>
        /// <param name="info"></param>
        private void ConvertToCity(DirectoryInfo info, ProjectInfo pro)
        {
            //c111-IRI-0.1-20220711
            //名称不包含时刻

            FileInfo[] lbiFiles = info.GetFiles("*LBI*.csv", SearchOption.AllDirectories)
      .Union(info.GetFiles("*LBI*.txt", SearchOption.AllDirectories))
     .ToArray();


            FileInfo[] textFiles = info.GetFiles("*TT*.csv", SearchOption.AllDirectories)
     .Union(info.GetFiles("*TT*.txt", SearchOption.AllDirectories))
    .ToArray();
            //只有等级公路有车辙
            FileInfo[] rdFiles = info.GetFiles("*RD*.csv", SearchOption.AllDirectories);
            FileInfo[] mpdFiles = info.GetFiles("*MPD*.csv", SearchOption.AllDirectories);

            FileInfo[] drFiles = info.GetFiles("*DR*", SearchOption.AllDirectories);

            //c111-IRI-0.1-2022071-131313
            //时刻由-间隔
            FileInfo[] lpFiles = info.GetFiles("*LP*.csv", SearchOption.AllDirectories).Union(info.GetFiles("*LP*.txt", SearchOption.AllDirectories))
     .ToArray(); ;
            FileInfo[] pbFiles = info.GetFiles("*PB*.csv", SearchOption.AllDirectories);
            FileInfo[] vbiFiles = info.GetFiles("*VBI*.csv", SearchOption.AllDirectories);
            FileInfo[] acceleFiles = info.GetFiles("*加速度*.csv", SearchOption.AllDirectories);
            //C6666-DR-0.6-10.6-20220716-标准沥青.CSV       等级公路病害
            //C8888-DR-0.8-10.8-20220718-指南沥青.CSV       低等级农村路病害
            FileInfo[] lqFiles = info.GetFiles("*沥*", SearchOption.AllDirectories);
            FileInfo[] snFiles = info.GetFiles("*水*", SearchOption.AllDirectories);
            FileInfo[] ssFiles = info.GetFiles("*砂*", SearchOption.AllDirectories);

            FileInfo[] iriFiles = info.GetFiles("*IRI*.csv", SearchOption.AllDirectories).Union(info.GetFiles("*IRI*.txt", SearchOption.AllDirectories))
     .ToArray(); ;
            switch (currentStandard)
            {
                case CityModelItem.等级公路5210与农村路5211标准模板导出_2025年:
                    break;
                case CityModelItem.河北省单位定制:
                    //lbi加有效列
                    // 为每个IRI文件添加有效性列
                    foreach (FileInfo file in lbiFiles)
                    {
                        // 读取原始文件内容
                        var lines = File.ReadAllLines(file.FullName);

                        // 处理第一行（表头）
                        if (lines.Length > 0)
                        {

                            lines[0] = lines[0] + ",有效性";
                            lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                        }

                        // 处理数据行
                        for (int i = 1; i < lines.Length; i++)
                        {
                            lines[i] = lines[i] + ",A";
                        }

                        // 写回文件
                        File.WriteAllLines(file.FullName, lines);
                    }

                    foreach (var file in drFiles)
                    {
                        // 读取原始文件内容
                        var lines = File.ReadAllLines(file.FullName);
                        // 处理第一行（表头）
                        if (lines.Length > 0)
                        {


                            lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                        }
                        // 写回文件
                        File.WriteAllLines(file.FullName, lines);
                    }
                    foreach (var file in lpFiles)
                    {
                        // 读取原始文件内容
                        var lines = File.ReadAllLines(file.FullName);
                        // 处理第一行（表头）
                        if (lines.Length > 0)
                        {


                            lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                        }
                        // 写回文件
                        File.WriteAllLines(file.FullName, lines);
                    }
                    break;

                case CityModelItem.农养国省道路况检测数据提交格式_2026年:
                    { 
                    }
                    break;
                case CityModelItem.辽宁省2025单位定制:
                    {
                        if (textFiles.Count() == 0)
                        {
                            //为低等级农村路
                            break;
                        }
                        //处理平整度文件 
                        foreach (FileInfo file in iriFiles)
                        {
                            // 读取原始文件内容
                            var lines = File.ReadAllLines(file.FullName);

                            lines = deleteColumn(lines, new int[] { 1, 2 }, ",");

                            // 处理第一行（表头）
                            if (lines.Length > 0)
                            {
                                lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                            }
                            // 写回文件
                            File.WriteAllLines(file.FullName, lines);
                        }
                        //修改平整度
                        //修改平整度文件夹名称 LBIFile->LBI
                        string lbiDirPath = Path.Combine(info.FullName, "LBIFile");
                        Microsoft.VisualBasic.FileIO.FileSystem.RenameDirectory(lbiDirPath, "LBI");
                        lbiFiles = info.GetFiles("*LBI*.csv", SearchOption.AllDirectories).Union(info.GetFiles("*LBI*.txt", SearchOption.AllDirectories)).ToArray();
                        foreach (FileInfo file in lbiFiles)
                        {
                            // 读取原始文件内容
                            var lines = File.ReadAllLines(file.FullName);

                            // 处理第一行（表头）
                            if (lines.Length > 0)
                            {

                                lines[0] = lines[0] + ",有效性";
                                lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                            }

                            // 处理数据行
                            for (int i = 1; i < lines.Length; i++)
                            {
                                lines[i] = lines[i] + ",A";
                            }

                            // 写回文件
                            File.WriteAllLines(file.FullName, lines);
                        }
                        //处理LP文件
                        foreach (var file in lpFiles)
                        {
                            // 读取原始文件内容
                            var lines = File.ReadAllLines(file.FullName);

                            // 处理第一行（表头）
                            if (lines.Length > 0)
                            {
                                lines[0] = lines[0].Replace("起点桩号(km)", "桩号(km)");
                            }
                            File.WriteAllLines(file.FullName, lines);
                        }
                        //修改磨耗
                        //修改磨耗文件夹名称 TTFile ->TEXTFile
                        string ttDirPath = Path.Combine(info.FullName, "TTFile");
                        Microsoft.VisualBasic.FileIO.FileSystem.RenameDirectory(ttDirPath, "TEXTFile");
                        textFiles = info.GetFiles("*TT*.csv", SearchOption.AllDirectories)
          .Union(info.GetFiles("*TT*.txt", SearchOption.AllDirectories))
         .ToArray(); 
                    }
                    break;

                case CityModelItem.交通部2024规范:
                    break;
                case CityModelItem.河南省单位一农村路定制:
                    break;
                case CityModelItem.湖南省单位一定制:
                    break;
                case CityModelItem.重庆市单位一定制:
                    break;
                case CityModelItem.甘肃省单位一定制:
                    break;
                case CityModelItem.河北省单位一定制:
                    break;
                case CityModelItem.河北省单位二定制:
                    break;
                case CityModelItem.江苏省单位一定制:
                    break;
                case CityModelItem.安徽省单位一定制:
                    break;
                case CityModelItem.广东省单位一定制:
                    break;
                default:
                    break;
            }

        }
        //河北规范, 移动目标文件
        private void changeDirNameMoveFile()
        {

        }
        /// <summary>
        /// 删除指定列数据（高效版本，使用HashSet）
        /// </summary>
        private string[] deleteColumn(string[] texts, int[] indexes, string splitString)
        {
            if (texts == null || texts.Length == 0)
                return new string[0];

            // 使用HashSet提高查找效率
            var indexesToRemove = new HashSet<int>(indexes ?? new int[0]);

            return texts.Select(line =>
            {
                string[] columns = line.Split(new string[] { splitString }, StringSplitOptions.None);

                // 如果没有要删除的索引，直接返回原行
                if (indexesToRemove.Count == 0)
                    return line;

                // 过滤掉所有在HashSet中的索引列
                var newColumns = columns.Where((col, idx) => !indexesToRemove.Contains(idx)).ToArray();
                return string.Join(splitString, newColumns);
            }).ToArray();
        }

        public static void DeleteFolders(string path)
        {
            try
            {
                // 检查路径是否存在
                if (!Directory.Exists(path))
                {
                    Console.WriteLine($"路径 {path} 不存在。");
                    return;
                }

                // 获取该路径下所有文件夹
                string[] folders = Directory.GetDirectories(path);

                foreach (string folder in folders)
                {

                    // 递归删除子文件夹
                    DeleteFolders(folder);

                    // 删除当前文件夹
                    Directory.Delete(folder, true);
                    Console.WriteLine($"已删除文件夹 {folder}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除文件夹时发生错误：{ex.Message}");
            }
        }
        public void WriteDataToExcel(string modelPath, DataTable dt, string outPath, int startIndex = -1)
        {
            // 读取模板文件
            using (FileStream fs = new FileStream(modelPath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0); // 获取第一个工作表
                int rowIndex = startIndex;
                // 找到第一个没有数据的行
                if (startIndex == -1)
                {
                    rowIndex = FindFirstEmptyRow(sheet);
                }


                // 写入数据
                foreach (DataRow row in dt.Rows)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        object value = row[i];
                        dataRow.CreateCell(i).SetCellValue(value.ToString());
                    }

                    rowIndex++;
                }

                // 保存文件
                using (FileStream output = new FileStream(outPath, FileMode.Create))
                {
                    workbook.Write(output);
                }
            }
        }

        private int FindFirstEmptyRow(ISheet sheet)
        {
            int rowIndex = 0;
            IRow row = sheet.GetRow(rowIndex);

            while (row != null && !HasEmptyCells(row))
            {
                rowIndex++;
                row = sheet.GetRow(rowIndex);
            }

            return rowIndex;
        }

        private bool HasEmptyCells(IRow row)
        {

            ICell cell = row.GetCell(0, MissingCellPolicy.RETURN_NULL_AND_BLANK);

            if (cell == null || cell.CellType == CellType.Blank)
            {
                return true;
            }
            else
            {
                return false;
            }


        }
        private string findModelExcelPath(string basePath, string fillter)
        {
            string path = "";
            DirectoryInfo directory = new DirectoryInfo(basePath);



            foreach (FileInfo file in directory.GetFiles("*.xlsx"))
            {
                if (file.Name.Contains(fillter))
                {
                    path = file.FullName;
                    break;
                }
            }
            if (string.IsNullOrEmpty(path))
            {
                throw new Exception("找不到模板文件" + fillter);
            }
            return path;
        }
        public void InsertColumnWithValue(string value, DataTable dataTable, int index = 0)
        {
            DataColumn newColumn = new DataColumn();
            newColumn.DefaultValue = value;

            dataTable.Columns.Add(newColumn);
            dataTable.Columns[dataTable.Columns.Count - 1].SetOrdinal(index);
        }
        public void SwapColumns(DataTable dataTable, int column1Index, int column2Index)
        {
            if (column1Index < 0 || column1Index >= dataTable.Columns.Count
                || column2Index < 0 || column2Index >= dataTable.Columns.Count)
            {
                throw new ArgumentException("Invalid column index.");
            }

            DataColumn column1 = dataTable.Columns[column1Index];
            DataColumn column2 = dataTable.Columns[column2Index];

            int column1Ordinal = column1.Ordinal;
            column1.SetOrdinal(column2.Ordinal);
            column2.SetOrdinal(column1Ordinal);
        }
        public void InsertColumnWithValue(DataColumn newColumn, DataTable dataTable, int index = 0)
        {
            dataTable.Columns.Add(newColumn);
            dataTable.Columns[dataTable.Columns.Count - 1].SetOrdinal(index);
        }
        private static void ChangeFileName3_HN(ProjectInfo pro, FileInfo[] lqFiles, string dName)
        {
            foreach (var file in lqFiles)
            {
                if (handelCsvFiles.Contains(file.FullName))
                {
                    continue;
                }
                string fullName = file.FullName;

                string fatherPath = Path.GetDirectoryName(fullName);


                string FileName = Path.GetFileNameWithoutExtension(file.FullName);
                string[] spS = FileName.Split('-');
                if (spS.Length > 0)
                {
                    string cityCode = spS[0];

                    string startMile = spS[2];
                    string endMile = spS[3];
                    var year = pro._DataDate;
                    string newName = string.Join("-", cityCode, "DR", startMile, endMile, year, dName) + ".csv";
                    string nowFullPath = Path.Combine(fatherPath, newName);

                    if (!File.Exists(nowFullPath))
                    {

                        Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(fullName, newName);

                    }
                }


            }

        }

        private static void ChangeFileName2_HN(ProjectInfo pro, FileInfo[] lpFiles)
        {
            foreach (var file in lpFiles)
            {
                string fullName = file.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                string fatherPath = Path.GetDirectoryName(fullName);

                string fileName = Path.GetFileNameWithoutExtension(file.FullName);
                int len = fileName.Length;

                string newName = fileName.Substring(0, len - 6) + "-" + pro._DataTime + ".csv";

                string nowFullPath = Path.Combine(fatherPath, newName);

                if (!File.Exists(nowFullPath))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(fullName, newName);

                }
            }
        }

        private static void ChangeFileName1_HN(FileInfo[] iriFiles)
        {
            foreach (var file in iriFiles)
            {

                string fullName = file.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                string fatherPath = Path.GetDirectoryName(fullName);

                string fileName = Path.GetFileNameWithoutExtension(file.FullName);
                int len = fileName.Length;
                string newName = fileName.Substring(0, len - 6) + ".csv";
                string nowFullPath = Path.Combine(fatherPath, newName);


                if (!File.Exists(nowFullPath))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(fullName, newName);

                }

            }
        }

        private List<string> stringToListString(string str)
        {
            List<string> ttemp = new List<string>();
            return ttemp = str.Split(',').ToList();

        }
        /// <summary>
        /// 移动 alldir下某个文件夹到某个文件夹下
        /// </summary>
        /// <param name="allDir"></param>
        /// <param name="sourceDir"></param>
        /// <param name="disDir"></param>
        private void MoveDirToDir(DirectoryInfo info, string sourceDir, string disDir)
        {
            DirectoryInfo iriDir = info.GetDirectories(sourceDir).FirstOrDefault();
            DirectoryInfo RIFileDir = info.GetDirectories(disDir).FirstOrDefault();
            DirectoryInfo directoryInfo = Directory.CreateDirectory(Path.Combine(RIFileDir.FullName, sourceDir));
            FileHelpter.CopyFileAndDir(iriDir.FullName, directoryInfo.FullName);
            FileInfo[] fileInfos = iriDir.GetFiles();
            Directory.Delete(iriDir.FullName, true);
        }
        private void ChangeConvert_henan(FileInfo[] iriFiles, string head, List<string> headTxt)
        {
            foreach (var item in iriFiles)
            {
                string fullName = item.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);

                DataTable dtOut = ChangeMile_henan(dt, headTxt);

                if (!string.IsNullOrEmpty(head))
                {
                    InsertFirstName(dtOut, head);
                }
                CsvHelper.WriteDataToCsv(dtOut, item.FullName);

            }
        }

        /// <summary>
        /// 断面高程 桩号为0.1m间隔  4位小数
        /// </summary>
        /// <param name="iriFiles"></param>
        /// <param name="head"></param>
        private void ChangeConvert_henan_iriH(FileInfo[] iriFiles, string head, List<string> headTxts)
        {
            foreach (var item in iriFiles)
            {
                string fullName = item.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);

                DataTable dtOut = ChangeMile_henan_iriH(dt, headTxts);
                if (!string.IsNullOrEmpty(head))
                {
                    InsertFirstName(dtOut, head);
                }
                CsvHelper.WriteDataToCsv(dtOut, item.FullName);

            }
        }
        /// <summary>
        ///当DataTable中有值时，是不允许修改列的DataType
        /// 修改数据表DataTable某一列的数据类型和记录值
        /// </summary>
        /// <param name="argDataTable">数据表DataTable</param>
        /// <returns>数据表DataTable</returns>
        private DataTable UpdateDataTable(DataTable argDataTable, int colNum)
        {
            DataTable dtResult = new DataTable();
            //克隆表结构
            dtResult = argDataTable.Clone();
            //修改数据列类型
            dtResult.Columns[colNum].DataType = typeof(string);
            //foreach (DataColumn col in dtResult.Columns)
            //{

            //    col.DataType = typeof(int);

            //}
            return dtResult;
        }

        private DataTable ChangeMile_henan(DataTable dt, List<string> heads)
        {
            DataTable dtOut = UpdateDataTable(dt, 0);
            //桩号为K0+000格式

            int colNum = dtOut.Columns.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dtOut.NewRow();
                var nowRow = dt.Rows[i];
                for (int t = 0; t < colNum; t++)
                {
                    var cell = nowRow[t];
                    if (t == 0)
                    {
                        string txt = cell.ToString();
                        double temp;
                        bool result = double.TryParse(txt, out temp);
                        if (result)
                        {
                            cell = (temp * 1000).ToString("K0+000");
                        }
                    }
                    if (heads != null && heads.Count == colNum)
                    {
                        if (i == 0)
                        {
                            cell = heads[t];
                        }
                    }
                    dr[t] = cell;
                }
                dtOut.Rows.Add(dr);
            }
            //foreach (DataRow drOne in dtOut.Rows)
            //{
            //    var cell = drOne[0];

            //    string txt = drOne[0].ToString();
            //    double temp;
            //    bool result = double.TryParse(txt, out temp);
            //    if (result)
            //    {
            //        cell = "ccc";
            //        //cell = (temp * 1000).ToString("K0+000");
            //    }
            //}
            return dtOut;
        }
        /// <summary>
        /// 修改表头
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="heads"></param>
        /// <returns></returns>
        private DataTable ChangeHead_gansu(DataTable dt, List<string> heads)
        {
            DataTable dtOut = UpdateDataTable(dt, 0);


            int colNum = dtOut.Columns.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dtOut.NewRow();
                var nowRow = dt.Rows[i];
                for (int t = 0; t < colNum; t++)
                {
                    var cell = nowRow[t];
                    //if (t == 0)
                    //{
                    //    string txt = cell.ToString();
                    //    double temp;
                    //    bool result = double.TryParse(txt, out temp);
                    //    if (result)
                    //    {
                    //        cell = (temp * 1000).ToString("K0+000");
                    //    }
                    //}
                    if (heads != null && heads.Count == colNum)
                    {
                        if (i == 0)
                        {
                            cell = heads[t];
                        }
                    }
                    dr[t] = cell;
                }
                dtOut.Rows.Add(dr);
            }
            //foreach (DataRow drOne in dtOut.Rows)
            //{
            //    var cell = drOne[0];

            //    string txt = drOne[0].ToString();
            //    double temp;
            //    bool result = double.TryParse(txt, out temp);
            //    if (result)
            //    {
            //        cell = "ccc";
            //        //cell = (temp * 1000).ToString("K0+000");
            //    }
            //}
            return dtOut;
        }
        private DataTable ChangeHead_gansu_iri(DataTable dt, List<string> heads)
        {
            DataTable dtOut = UpdateDataTable(dt, 0);


            int colNum = dtOut.Columns.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dtOut.NewRow();
                var nowRow = dt.Rows[i];
                for (int t = 0; t < colNum; t++)
                {
                    var cell = nowRow[t];
                    if (t == 1) //iri数据列
                    {
                        string txt = cell.ToString();
                        double temp;
                        bool result = double.TryParse(txt, out temp);
                        if (result)
                        {
                            cell = temp.ToString("f2");
                        }
                    }
                    if (heads != null && heads.Count == colNum)
                    {
                        if (i == 0)
                        {
                            cell = heads[t];
                        }
                    }
                    dr[t] = cell;
                }
                dtOut.Rows.Add(dr);
            }
            //foreach (DataRow drOne in dtOut.Rows)
            //{
            //    var cell = drOne[0];

            //    string txt = drOne[0].ToString();
            //    double temp;
            //    bool result = double.TryParse(txt, out temp);
            //    if (result)
            //    {
            //        cell = "ccc";
            //        //cell = (temp * 1000).ToString("K0+000");
            //    }
            //}
            return dtOut;
        }
        /// <summary>
        /// 修改表头文字
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private DataTable ChangeHead_henan(DataTable dt, List<string> heads)
        {
            DataTable dtOut = UpdateDataTable(dt, 0);


            DataRow dt0 = dt.Rows[0];

            int colNum = dtOut.Columns.Count;

            for (int i = 0; i < colNum; i++)
            {
                var cell = dt0[i];
                cell = heads[i];
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dtOut.NewRow();
                var nowRow = dt.Rows[i];
                for (int t = 0; t < colNum; t++)
                {
                    var cell = nowRow[t];
                    if (t == 0)
                    {
                        string txt = cell.ToString();
                        double temp;
                        bool result = double.TryParse(txt, out temp);
                        if (result)
                        {
                            cell = (temp * 1000).ToString("K0+000");
                        }
                    }

                    dr[t] = cell;
                    if (i == 0)
                    {
                        string txt = cell.ToString();
                        txt = txt.Replace("\"", "");
                        cell = txt;

                    }
                }
                dtOut.Rows.Add(dr);
            }

            return dtOut;
        }
        private DataTable ChangeMile_henan_iriH(DataTable dt, List<string> heads)
        {
            DataTable dtOut = UpdateDataTable(dt, 0);
            //桩号为K0+000格式

            int colNum = dtOut.Columns.Count;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dtOut.NewRow();
                var nowRow = dt.Rows[i];
                for (int t = 0; t < colNum; t++)
                {
                    var cell = nowRow[t];
                    if (t == 0)
                    {
                        string txt = cell.ToString();
                        double temp;
                        bool result = double.TryParse(txt, out temp);
                        if (result)
                        {
                            cell = (temp * 10000).ToString("K0+0000");
                        }
                    }
                    if (heads != null && heads.Count == colNum)
                    {
                        if (i == 0)
                        {
                            cell = heads[t];
                        }
                    }
                    dr[t] = cell;
                }
                dtOut.Rows.Add(dr);
            }
            //foreach (DataRow drOne in dtOut.Rows)
            //{
            //    var cell = drOne[0];

            //    string txt = drOne[0].ToString();
            //    double temp;
            //    bool result = double.TryParse(txt, out temp);
            //    if (result)
            //    {
            //        cell = "ccc";
            //        //cell = (temp * 1000).ToString("K0+000");
            //    }
            //}
            return dtOut;
        }

        private void InsertFirstName(DataTable dt, string name)
        {
            DataRow dr = dt.NewRow();
            dr[0] = name;
            dt.Rows.InsertAt(dr, 0);
        }
        private void ChangeConvert_gansu(FileInfo[] iriFiles, string head, List<string> headTxt)
        {
            foreach (var item in iriFiles)
            {
                string fullName = item.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);

                DataTable dtOut = ChangeHead_gansu(dt, headTxt);


                CsvHelper.WriteDataToCsv(dtOut, item.FullName);

            }
        }
        /// <summary>
        /// 总是报错  iri精度不足2位  再这里处理一下
        /// </summary>
        /// <param name="iriFiles"></param>
        /// <param name="head"></param>
        /// <param name="headTxt"></param>
        private void ChangeConvert_gansu_iri(FileInfo[] iriFiles, string head, List<string> headTxt)
        {
            foreach (var item in iriFiles)
            {
                string fullName = item.FullName;
                if (handelCsvFiles.Contains(fullName))
                {
                    continue;
                }
                DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);

                DataTable dtOut = ChangeHead_gansu_iri(dt, headTxt);


                CsvHelper.WriteDataToCsv(dtOut, item.FullName);

            }
        }
        /// <summary>
        /// 转换图片工程
        /// </summary>
        /// <param name="pro"></param>
        /// <param name="image"></param>
        /// <param name="isRoad">true 路面破损 false 景观</param>
        private ValueTuple<int, List<List<PicAndMile>>> ConvertPci(ProjectInfo pro, DirectoryInfo image, bool isRoad)
        {

            string indexText = image.FullName + "\\fileindex.txt";

            List<PicAndMile> _picAndMiles = pro.GetPicAndMiles(isRoad, currentStandard);
            int total_cnt = _picAndMiles.Count;
            string proName = pro.ConvertProName;
            string timeYear = pro.DateDay;

            WriteIndexTxt(pro, indexText, ref _picAndMiles, proName, timeYear, false, isRoad);
            List<List<PicAndMile>> spList = SplitPicAndMileList_5000(_picAndMiles);
            return (total_cnt, spList);

            //await MovePicutre(image, spList, currentStandard);


        }

        private async Task ConvertPciHuNan(ProjectInfo pro, DirectoryInfo image, bool isRoad, int indexPro)
        {

            string indexText = image.FullName + "\\0\\" + "fileindex.txt";
            if (!Directory.Exists(image.FullName + "\\0"))
            {
                Directory.CreateDirectory(image.FullName + "\\0");
            }

            if (File.Exists(indexText))
            {

                List<PicAndMile> _picAndMiles = pro.GetPicAndMilesHuNan(isRoad);
                string proName = pro.ConvertProName;
                string timeYear = pro.DateDay;
                WriteIndexTxtHuNan(indexText, _picAndMiles, proName, timeYear, true);
                List<List<PicAndMile>> spList = SplitPicAndMileList_5000HuNan(_picAndMiles);


                await MovePicutreHuNan(image, spList, indexPro);
            }
            else
            {
                File.Create(indexText);
                List<PicAndMile> _picAndMiles = pro.GetPicAndMilesHuNan(isRoad);
                string proName = pro.ConvertProName;
                string timeYear = pro.DateDay;
                WriteIndexTxtHuNan(indexText, _picAndMiles, proName, timeYear, false);
                List<List<PicAndMile>> spList = SplitPicAndMileList_5000HuNan(_picAndMiles);


                await MovePicutreHuNan(image, spList, indexPro);
            }


        }

        /// <summary>
        /// 复制图片 
        /// 多线程
        /// </summary>
        /// <param name="image"></param>
        /// <param name="spList"></param>
        private async Task MovePicutre(DirectoryInfo image, List<List<PicAndMile>> spList, CityModelItem stanard)
        {
            await Task.Run(() =>
                PictureExportService.ExportStandardBatches(image, spList, stanard, _Suf, this.Process)
            );
        }

        private async Task MovePicutreHuNan(DirectoryInfo image, List<List<PicAndMile>> spList, int indexPro)
        {
            await Task.Run(() => PictureExportService.ExportHunanBatches(image, spList));
        }
        public static int ProcessValue
        {
            set
            {

            }
        }
        /// <summary>
        /// 按照5000间隔对总图片结构进行切断
        /// </summary>
        /// <param name="_picAndMiles"></param>
        /// <returns></returns>
        private static List<List<PicAndMile>> SplitPicAndMileList_5000(List<PicAndMile> _picAndMiles)
        {
            return PictureExportService.SplitByBatchSize(_picAndMiles);
        }
        private static List<List<PicAndMile>> SplitPicAndMileList_5000HuNan(List<PicAndMile> _picAndMiles)
        {
            return PictureExportService.SplitByBatchSize(_picAndMiles);
        }
        /// <summary>
        /// 写 fileindex.txt
        /// </summary>
        /// <param name="fs"></param>
        /// <param name="_picAndMiles"></param>
        /// <param name="proName"></param>
        /// <param name="timeYear"></param>
        private void WriteIndexTxt(ProjectInfo pro, string path, ref List<PicAndMile> _picAndMiles, string proName, string timeYear, bool apped, bool isRoad)
        {
            using (StreamWriter sw = new StreamWriter(path, apped))
            {
                string allMsg = "";
                for (int i = 0; i < _picAndMiles.Count; i++)
                {
                    string dirName = (i / 5000 + 1).ToString("00");

                    string msg = "";
                    string picName = "";
                    PicAndMile updatedPicAndMile = _picAndMiles[i];
                    switch (currentStandard)
                    {
                        case CityModelItem.河南省单位一农村路定制:
                            if (isRoad)
                            {
                                picName = Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000");
                                msg = proName + "->" + Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000")
                            + "->" + dirName + "->" + picName + _Suf;
                                // 创建一个新的 PicAndMile 实例并更新 ResultPicName
                                updatedPicAndMile.updateResultPicName(picName);
                                _picAndMiles[i] = updatedPicAndMile; // 更新列表中的元素
                            }
                            else
                            {
                                picName = _picAndMiles[i].Mile.ToString("K0+000");
                                msg = proName + "->" + Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000")
                            + "->" + dirName + "->" + picName + _Suf;
                                // 创建一个新的 PicAndMile 实例并更新 ResultPicName
                                updatedPicAndMile.updateResultPicName(picName);
                                _picAndMiles[i] = updatedPicAndMile; // 更新列表中的元素
                            }

                            break;
                        case CityModelItem.湖南省单位一定制:
                            picName = proName + "-" + _picAndMiles[i].Mile.ToString("000+000") + "000-" + _picAndMiles[i].Mile.ToString("000+000") + "000";
                            dirName = (i / 5000).ToString("0");
                            msg = proName + "->" + _picAndMiles[i].Mile.ToString("000+000")
                        + "->" + dirName + "->" + picName + _Suf;
                            // 创建一个新的 PicAndMile 实例并更新 ResultPicName
                            updatedPicAndMile.updateResultPicName(picName);
                            _picAndMiles[i] = updatedPicAndMile; // 更新列表中的元素
                            break;
                        case CityModelItem.甘肃省单位一定制:
                            // msg = $"{dirName}\\{proName.Substring(0,proName.Length-1)}-{Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000")}-{timeYear}.jpg";
                            msg = proName + "->" + _picAndMiles[i].Mile.ToString("0+000")
                      + "->" + dirName + "->" + proName.Substring(0, proName.Length - 1) + "_" + Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000") + "_"
                      + timeYear + _Suf;
                            break;
                        case CityModelItem.江苏省单位一定制:
                            msg = $"{pro._City + pro._District}" + "->" + proName + "->" + _picAndMiles[i].Mile.ToString("0+000")
                       + "->" + dirName + "->" + timeYear + "_" + proName + "_"
                       + Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000") + _Suf;
                            break;
                        default:
                            msg = proName + "->" + _picAndMiles[i].Mile.ToString("000+000")
                      + "->" + dirName + "->" + timeYear + "_" + proName + "_"
                      + Math.Round(_picAndMiles[i].Mile * 0.001, 3).ToString("0.000") + _Suf;
                            break;
                    }
                    allMsg += msg + "\r\n";
                }
                sw.WriteLine(allMsg);
            }

        }
        private static void WriteIndexTxtHuNan(string fs, List<PicAndMile> _picAndMiles, string proName, string timeYear, bool append)
        {
            using (StreamWriter sw = new StreamWriter(fs, append))
            {
                string allMsg = "";
                for (int i = 0; i < _picAndMiles.Count; i++)
                {

                    // string dirName = (i / 5000 + 1).ToString();
                    string dirName = (i / 5000).ToString();
                    string mile = ConvertIntToFormattedString(_picAndMiles[i].Mile);
                    string picName = proName + "-" + mile + "-" + mile + ".jpg";
                    string msg = proName + "->" + mile
                        + "->" + dirName + "->" + picName;
                    allMsg += msg + "\r\n";
                }
                sw.WriteLine(allMsg);
            }
        }


        public static string ConvertIntToFormattedString(int value)
        {
            int thousands = value / 1000; // 获取千位以上的数字部分
            int hundreds = value % 1000; // 获取百位以下的数字部分

            string result = $"{thousands.ToString().PadLeft(3, '0')}+{hundreds.ToString().PadLeft(3, '0') + "000"}";
            return result;
        }

        List<FileInfo> DaqFiles = new List<FileInfo>();
        private Progress<int> Process;

        /// <summary>
        /// 自定义惯导平整度文件名称
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            this.treeList2.Nodes.Clear();
            DaqFiles.Clear();
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.Description = "请选择需要转换的工程";
            dlg.SelectedPath = _config.UserPath;
            dlg.ShowDialog();
            if (dlg.SelectedPath != String.Empty)
            {
                if (dlg.SelectedPath.Substring(dlg.SelectedPath.Length - 1) == "\\")
                {
                    dlg.SelectedPath = dlg.SelectedPath.Remove(dlg.SelectedPath.Length - 1);
                }
                this.Chktxt_Path = dlg.SelectedPath;
                _config.UserPath = dlg.SelectedPath;
                ConfigManager.SaveConfig();
                DirectoryInfo directory = new DirectoryInfo(dlg.SelectedPath);
                DaqFiles = directory.GetFiles("*.daq", SearchOption.AllDirectories).ToList();
                DaqFiles.Sort((a, b) => StrCmpLogicalW(a.FullName, b.FullName));
                foreach (var pro in DaqFiles)
                {
                    TreeListNode node = this.treeList2.AppendNode(null, null);
                    // node.SetValue("name",pro._DataDir.Name) ;
                    node.SetValue("data", pro.Name);
                }
            }
        }



        private void BarEditItem1_EditValueChanged(object sender, EventArgs e)
        {
            _config.NowModel = barEditItem1.EditValue.ToString();
            ConfigManager.SaveConfig();
        }

        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            foreach (var proj in _Projects)
            {
                var nowDir = proj._DataDir;
                DirectoryInfo[] ConvertDirs = nowDir.GetDirectories("ConverSource");
                foreach (var item in ConvertDirs)
                {
                    Directory.Delete(item.FullName, true);
                }

                FileInfo[] projectInfos = nowDir.GetFiles("ProjectInfo.txt");
                foreach (FileInfo file in projectInfos)
                {
                    var strings = File.ReadAllLines(file.FullName).ToList();
                    for (int i = strings.Count - 1; i >= 0; i--)
                    {
                        string temp = strings[i];
                        if (temp.Contains("县级行政区划代码") || temp.Contains("ConvertOk"))
                        {
                            strings.RemoveAt(i);
                        }
                    }
                    //ClearText(file.FullName);

                    SaveFile(file.FullName, strings.ToArray(), false);
                }

            }
            if (_Projects.Count > 0)
            {
                MessageBox.Show("清理完毕！");
            }
            else
            {
                MessageBox.Show("您还没有导入任何合格工程");
            }

        }

        private void ClearText(string textPath)
        {
            FileStream stream = File.Open(textPath, FileMode.OpenOrCreate, FileAccess.Write);
            stream.Seek(0, SeekOrigin.Begin);
            stream.SetLength(0);
            stream.Close();

        }
        private void SaveFile(string str, string text, bool saOrAp)
        {
            StreamWriter sw = new StreamWriter(str, saOrAp);
            sw.WriteLine(text);
            sw.Close();

        }
        private void SaveFile(string str, string[] text, bool saOrAp)
        {
            StreamWriter sw = new StreamWriter(str, saOrAp);
            foreach (string str2 in text)
            {
                sw.WriteLine(str2);
            }

            sw.Close();

        }

        private void barButtonItem6_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void barButtonItem7_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择河南2导出的结果文件夹";
            var result = fd.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }
            //获取上下行 
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);

            DirectoryInfo[] dis = dir.GetDirectories();
            for (int i = 0; i < dis.Length; i++)
            {
                string proName = "";
                proName = dis[i].Name;
                string basePath = dis[i].FullName;
                DirectoryInfo baseDir = new DirectoryInfo(basePath);
                try
                {


                    //获得上下行 
                    string dirction = "上行";

                    if (proName.EndsWith("B"))
                    {
                        dirction = "下行";
                    }

                    //获得路线编码 
                    string roadCode = proName.Substring(0, 10);


                    //获得检测日期
                    FileInfo[] files = dis[i].GetFiles("*.csv", SearchOption.AllDirectories);
                    string firstName = files.Where(t => t.Name.Contains("IRI")).First().Name;
                    string time = firstName.Split('-').Last().Split('-').First();
                    DateTime date;
                    try
                    {
                        date = DateTime.ParseExact(time, "yyyyMMdd", CultureInfo.InvariantCulture);

                    }
                    catch (Exception)
                    {
                        date = DateTime.ParseExact(time.Split('.').First(), "yyyyMMdd", CultureInfo.InvariantCulture);

                    }
                    time = date.ToString("yyyy/M/d");
                    #region 更新病害表
                    foreach (var item in files.Where(t => t.Name.Contains("DR")).ToList())
                    {
                        string fileName = item.Name;
                        string outName = basePath + "\\" + fileName;
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        if (string.IsNullOrEmpty(dt.Rows[0][1].ToString()))
                        {
                            dt.Rows[0].Delete();
                        }
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;
                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();

                        }
                        //    CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                    }
                    #endregion
                    #region 更新空间定位表
                    foreach (var item in files.Where(t => t.Name.Contains("LBI")).ToList())
                    {
                        string fileName = item.Name;
                        string outName = basePath + "\\" + fileName;
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        dt.Columns.RemoveAt(3);
                        dt.Columns.RemoveAt(3);
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        // 在第0列插入一列
                        dt.Columns.Add("检测方向", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测方向";
                        dt.Columns.Add("检测日期", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测日期";

                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;

                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();
                            dt.Rows[t][4] = dirction;
                            dt.Rows[t][5] = time;

                        }
                        //  CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                    }
                    #endregion


                    #region 更新平整度表
                    foreach (var item in files.Where(t => t.Name.Contains("IRI")).ToList())
                    {
                        string fileName = item.Name;
                        string outName = basePath + "\\" + fileName;
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        if (string.IsNullOrEmpty(dt.Rows[0][1].ToString()))
                        {
                            dt.Rows[0].Delete();
                        }
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        // 在第0列插入一列
                        dt.Columns.Add("检测方向", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测方向";
                        dt.Columns.Add("检测日期", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测日期";

                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;

                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();
                            dt.Rows[t][4] = dirction;
                            dt.Rows[t][5] = time;

                        }
                        //  CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                    }

                    //删除除了图片的所有文件夹
                    DirectoryInfo[] dirs = baseDir.GetDirectories();

                    foreach (var item in dirs)
                    {
                        if (item.Name.Contains("景观图像") || item.Name.Contains("Images"))
                        {

                        }
                        else
                        {
                            Directory.Delete(item.FullName, true);
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {

                    throw new Exception(ex.Message + "\n" + proName + "工程解析出错");
                }


            }
            MessageBox.Show("所有项目处理完毕请检查!");

        }


        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择河南2导出的结果文件夹";
            List<string> needDirPath = new List<string> { "LBI", "Images", "RIFile", "IRI", "景观图像", "识别结果" };
            var result = fd.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }
            //获取上下行 
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);

            DirectoryInfo[] dis = dir.GetDirectories();
            for (int i = 0; i < dis.Length; i++)
            {
                string proName = "";
                proName = dis[i].Name;
                string basePath = dis[i].FullName;
                DirectoryInfo baseDir = new DirectoryInfo(basePath);
                try
                {

                    //获得上下行 
                    string dirction = "上行";

                    if (proName.EndsWith("B"))
                    {
                        dirction = "下行";
                    }

                    //获得路线编码 
                    string roadCode = proName.Substring(0, 10);


                    //获得检测日期
                    FileInfo[] files = dis[i].GetFiles("*.csv", SearchOption.AllDirectories);



                    string firstName = Path.GetFileNameWithoutExtension(files.Where(t =>
                    {
                        if (t.Name.Contains("IRI"))
                        {
                            string[] s = Path.GetFileNameWithoutExtension(t.FullName).Split('-');
                            return s.Last().Length == 8;
                        }

                        return false;
                    }
                    ).First().FullName);
                    string time = firstName.Split('-').Last().Split('-').First();
                    DateTime date;
                    try
                    {
                        date = DateTime.ParseExact(time, "yyyyMMdd", CultureInfo.InvariantCulture);

                    }
                    catch (Exception)
                    {
                        date = DateTime.ParseExact(time.Split('.').First(), "yyyyMMdd", CultureInfo.InvariantCulture);

                    }
                    time = date.ToString("yyyy/M/d");
                    //路面文件夹
                    DirectoryInfo roadPicDir = baseDir.GetDirectories("Images").First();
                    DirectoryInfo streetPicDir = baseDir.GetDirectories("景观图像").First();

                    DirectoryInfo lbiDir = baseDir.GetDirectories("LBI").First();
                    DirectoryInfo iriDir = baseDir.GetDirectories("IRI").First();
                    DirectoryInfo disDir = baseDir.GetDirectories("识别结果").First();
                    FileInfo indexRoadFile = roadPicDir.GetFiles("fileindex.txt", SearchOption.AllDirectories).FirstOrDefault();
                    FileInfo indexStreetFile = streetPicDir.GetFiles("fileindex.txt", SearchOption.AllDirectories).FirstOrDefault();
                    FileInfo GPSFile = baseDir.GetFiles("GPS2Mile.txt").First();
                    //获得时分秒
                    string timeMille = roadPicDir.GetDirectories().First().GetDirectories().First().Name.Split('_').Last();
                    #region 处理图像
                    //处理路面图像
                    //修改 index文件
                    Dictionary<double, string> lbiAndGpsInfoDic = new Dictionary<double, string>();
                    List<string> gpsInfos = File.ReadAllLines(GPSFile.FullName).ToList();
                    foreach (var item in gpsInfos)
                    {
                        string[] strs = item.Split(' ');
                        try
                        {
                            if (!lbiAndGpsInfoDic.Keys.Contains(double.Parse(strs.Last()) / 1000))

                            {
                                lbiAndGpsInfoDic.Add(double.Parse(strs.Last()) / 1000, strs[1] + "," + strs[2]);

                            }

                        }
                        catch (Exception)
                        {
                            throw;

                        }

                    }
                    if (indexRoadFile != null)
                    {
                        List<string> indexFileTXT = File.ReadAllLines(indexRoadFile.FullName).ToList();


                        var pcis = roadPicDir.GetFiles("*.jpg", SearchOption.AllDirectories);
                        string oldGpsInfo = "";
                        foreach (var item in pcis)
                        {
                            string name = Path.GetFileNameWithoutExtension(item.FullName);
                            //桩号
                            double mile = double.Parse(name.Split('_').Last());
                            string gpsInfo = lbiAndGpsInfoDic.Last().Value;

                            if (dirction == "上行")
                            {
                                gpsInfo = lbiAndGpsInfoDic.First().Value;
                            }
                            oldGpsInfo = gpsInfo;

                            try
                            {
                                gpsInfo = lbiAndGpsInfoDic[mile];
                                oldGpsInfo = gpsInfo;
                            }
                            catch (Exception)
                            {

                                gpsInfo = oldGpsInfo;
                            }

                            string[] strs = gpsInfo.Split(',');
                            string mileStr = (mile * 1000).ToString("K000+000");

                            string newName = mileStr + "_" + strs[0] + "_" + strs[1] + ".jpg";

                            string newPath = Path.GetDirectoryName(item.FullName) + "\\" + newName;
                            File.Move(item.FullName, newPath);
                            for (int d = 0; d < indexFileTXT.Count; d++)
                            {
                                string line = indexFileTXT[d];
                                string[] temp2 = line.Split('>');
                                string temp22 = line.Split('>')[1].Replace("+", "");
                                double mileTxt = double.Parse(temp22.Substring(0, temp22.Length - 1));
                                if (mile * 1000 == mileTxt)
                                {
                                    string newLine = "";
                                    for (int ddd = 0; ddd < temp2.Length - 1; ddd++)
                                    {
                                        newLine += temp2[ddd] + ">";
                                    }
                                    newLine += newName;
                                    indexFileTXT[d] = newLine;
                                    break;
                                }
                            }
                        }
                        File.WriteAllLines(indexRoadFile.FullName, indexFileTXT);

                    }
                    if (indexStreetFile != null)
                    {
                        string oldGpsInfo = "";
                        List<string> indexFileTXT = File.ReadAllLines(indexStreetFile.FullName).ToList();
                        var pcis = streetPicDir.GetFiles("*.jpg", SearchOption.AllDirectories);
                        foreach (var item in pcis)
                        {
                            string name = Path.GetFileNameWithoutExtension(item.FullName);
                            //桩号
                            double mile = double.Parse(name.Split('_').Last());
                            string gpsInfo = lbiAndGpsInfoDic.Last().Value;
                            if (dirction == "上行")
                            {
                                gpsInfo = lbiAndGpsInfoDic.First().Value;
                            }
                            oldGpsInfo = gpsInfo;
                            try
                            {
                                gpsInfo = lbiAndGpsInfoDic[mile];
                                oldGpsInfo = gpsInfo;
                            }
                            catch (Exception)
                            {

                                gpsInfo = oldGpsInfo;
                            }

                            string[] strs = gpsInfo.Split(',');
                            string mileStr = (mile * 1000).ToString("K000+000");

                            string newName = mileStr + "_" + strs[0] + "_" + strs[1] + ".jpg";

                            string newPath = Path.GetDirectoryName(item.FullName) + "\\" + newName;
                            File.Move(item.FullName, newPath);
                            for (int d = 0; d < indexFileTXT.Count; d++)
                            {
                                string line = indexFileTXT[d];
                                string[] temp2 = line.Split('>');
                                string temp22 = line.Split('>')[1].Replace("+", "");
                                double mileTxt = double.Parse(temp22.Substring(0, temp22.Length - 1));
                                if (mileTxt == 1005)
                                {

                                }
                                double mile2 = mile * 1000;
                                if (Math.Abs(mileTxt - mile2) <= 0.0001)
                                {
                                    string newLine = "";
                                    for (int ddd = 0; ddd < temp2.Length - 1; ddd++)
                                    {
                                        newLine += temp2[ddd] + ">";
                                    }
                                    newLine += newName;
                                    indexFileTXT[d] = newLine;
                                    break;
                                }
                            }
                        }

                        File.WriteAllLines(indexStreetFile.FullName, indexFileTXT);
                    }
                    File.Delete(GPSFile.FullName);

                    #endregion

                    #region 更新病害表
                    foreach (var item in files.Where(t => { return Path.GetFileNameWithoutExtension(t.FullName).Split('-').Length == 6 && t.Name.Contains("DR"); }).ToList())
                    {
                        string fileName = Path.GetFileNameWithoutExtension(item.Name);
                        string[] silit = fileName.Split('-');
                        string tempOutName = "";
                        for (int d = 0; d < silit.Length; d++)
                        {
                            if (d == silit.Length - 2)
                            {


                            }
                            else if (d == silit.Length - 1)
                            {
                                tempOutName += silit[d];
                            }
                            else
                            {
                                tempOutName += silit[d] + "-";

                            }
                        }
                        string outName = disDir.FullName + "\\" + tempOutName + ".csv";
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        if (string.IsNullOrEmpty(dt.Rows[0][1].ToString()))
                        {
                            dt.Rows[0].Delete();
                        }
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;
                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();

                        }
                        //    CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                        File.Delete(item.FullName);
                    }
                    #endregion
                    #region 更新空间定位表
                    //空间定位表 获得 桩号与gps信息

                    foreach (var item in files.Where(t => { return Path.GetFileNameWithoutExtension(t.FullName).Split('-').Last().Length == 8 && t.Name.Contains("LBI"); }).ToList())
                    {
                        string fileName = Path.GetFileNameWithoutExtension(item.FullName) + timeMille + ".csv";
                        string outName = lbiDir.FullName + "\\" + fileName;
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        dt.Columns.RemoveAt(3);
                        dt.Columns.RemoveAt(3);
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        // 在第0列插入一列
                        dt.Columns.Add("检测方向", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测方向";
                        dt.Columns.Add("检测日期", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测日期";

                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;

                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();
                            dt.Rows[t][4] = dirction;
                            dt.Rows[t][5] = time;


                        }



                        //  CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                        File.Delete(item.FullName);
                    }
                    #endregion


                    #region 更新平整度表
                    foreach (var item in files.Where(t =>
                    {
                        return Path.GetFileNameWithoutExtension(t.FullName).Split('-').Last().Length == 8 && t.Name.Contains("IRI");
                    }).ToList())
                    {
                        // string fileName = item.Name;
                        string fileName = Path.GetFileNameWithoutExtension(item.FullName) + timeMille + ".csv";
                        string outName = iriDir.FullName + "\\" + fileName;
                        DataTable dt = CsvHelper.ReadDataFromCsv(item.FullName);
                        if (string.IsNullOrEmpty(dt.Rows[0][1].ToString()))
                        {
                            dt.Rows[0].Delete();
                        }
                        // 在第0列插入一列
                        dt.Columns.Add("路线编码", typeof(string));
                        // 将新列移动到第0列
                        dt.Columns["路线编码"].SetOrdinal(0);

                        dt.Rows[0][0] = "路线编码";
                        // 在第0列插入一列
                        dt.Columns.Add("检测方向", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测方向";
                        dt.Columns.Add("检测日期", typeof(string));
                        dt.Rows[0][dt.Columns.Count - 1] = "检测日期";

                        for (int t = 1; t < dt.Rows.Count; t++)
                        {
                            dt.Rows[t][0] = roadCode;

                            string mile = dt.Rows[t][1].ToString();
                            dt.Rows[t][1] = (float.Parse(mile.Replace("K", "").Replace("+", "")) / 1000).ToString();
                            dt.Rows[t][4] = dirction;
                            dt.Rows[t][5] = time;

                        }
                        //  CsvHelper.WriteDataToCsv(dt, item.FullName);
                        CsvHelper.WriteDataToCsv(dt, outName);
                        File.Delete(item.FullName);
                    }
                    foreach (var dirsName in baseDir.GetDirectories())
                    {
                        if (!needDirPath.Contains(dirsName.Name))
                        {
                            Directory.Delete(dirsName.FullName, true);
                        }

                    }

                    #endregion


                }
                catch (Exception ex)
                {

                    throw new Exception(ex.Message + "\n" + proName + "工程解析出错");
                }


            }
            MessageBox.Show("所有项目处理完毕请检查!");

        }




        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择农村路规范导出的结果数据文件夹";
            var result = fd.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }


            //获取上下行 
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
            DirectoryInfo[] diss = dir.GetDirectories();

            if (dir.Name != "结果数据")
            {
                MessageBox.Show("请选择农村路规范导出的结果数据文件夹");
                return;
            }
            foreach (var disStr in diss)
            {
                DirectoryInfo[] dis = disStr.GetDirectories();
                foreach (var dirPath in dis)
                {
                    var tempPath = dirPath.Name;
                    if (tempPath.Contains("Images"))
                    {
                        DirectoryInfo tempDir = dirPath.GetDirectories().FirstOrDefault();
                        if (tempDir.Name.Length == 14)
                        {
                            string dirName = tempDir.Name.Substring(0, tempDir.Name.Length - 2);

                            Microsoft.VisualBasic.FileIO.FileSystem.RenameDirectory(tempDir.FullName, dirName);

                        }

                    }
                    if (tempPath == "DR")
                    {
                        FileInfo[] files = dirPath.GetFiles();
                        foreach (var item in files)
                        {
                            string names = Path.GetFileNameWithoutExtension(item.FullName);
                            string[] temps = names.Split('-');
                            string newFile = "";
                            if (temps.Length == 6)
                            {
                                if (names.Contains("水泥"))
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》水泥路面损坏自动化检测数据", temps[4]) + ".txt";
                                }
                                else if (names.Contains("沥青"))
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》沥青路面损坏自动化检测数据", temps[4]) + ".txt";

                                }
                                else
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》砂石路面损坏自动化检测数据", temps[4]) + ".txt";

                                }

                                newFile = Path.GetDirectoryName(item.FullName) + "\\" + newFile;

                                File.Move(item.FullName, newFile);
                            }
                        }
                    }
                }
            }


            MessageBox.Show("所有项目处理完毕请检查!");
        }

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
            string helpfile = Application.StartupPath + "\\软件说明\\12 产品内业软件使用手册V1.1-2024.docx";
            System.Diagnostics.Process.Start(helpfile);
        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择农村路规范导出的结果数据文件夹";
            var result = fd.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }


            //获取上下行 
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
            DirectoryInfo[] diss = dir.GetDirectories();

            if (dir.Name != "结果数据")
            {
                MessageBox.Show("请选择农村路规范导出的结果数据文件夹");
                return;
            }
            foreach (var disStr in diss)
            {
                DirectoryInfo[] dis = disStr.GetDirectories();
                foreach (var dirPath in dis)
                {
                    var tempPath = dirPath.Name;
                    if (tempPath.Contains("Images"))
                    {
                        DirectoryInfo tempDir = dirPath.GetDirectories().FirstOrDefault();
                        if (tempDir.Name.Length == 14)
                        {
                            string dirName = tempDir.Name.Substring(0, tempDir.Name.Length - 2);

                            Microsoft.VisualBasic.FileIO.FileSystem.RenameDirectory(tempDir.FullName, dirName);

                        }

                    }
                    if (tempPath == "DR")
                    {
                        FileInfo[] files = dirPath.GetFiles();
                        foreach (var item in files)
                        {
                            string names = Path.GetFileNameWithoutExtension(item.FullName);
                            string[] temps = names.Split('-');
                            string newFile = "";
                            if (temps.Length == 6)
                            {
                                if (names.Contains("水泥"))
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》水泥路面损坏自动化检测数据", temps[4]) + ".txt";
                                }
                                else if (names.Contains("沥青"))
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》沥青路面损坏自动化检测数据", temps[4]) + ".txt";

                                }
                                else
                                {
                                    newFile = string.Join("-", temps[0], double.Parse(temps[2]).ToString("0.000"), "《指南》砂石路面损坏自动化检测数据", temps[4]) + ".txt";

                                }

                                newFile = Path.GetDirectoryName(item.FullName) + "\\" + newFile;

                                File.Move(item.FullName, newFile);
                            }
                        }
                    }
                }
            }


            MessageBox.Show("所有项目处理完毕请检查!");
        }

        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {

            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择农村路规范导出的结果数据文件夹";
            var result = fd.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }
            //处理图片
            DialogResult dialogResult = MessageBox.Show("选择\"是\"进行图像名称修改,选否进行单张图片提取", "模式选择", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                //依据合肥规范进行 图片名称 和index.txt文件的修改
                DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                DirectoryInfo[] diss = dir.GetDirectories();

                for (int i = 0; i < diss.Length; i++)
                {
                    // DirectoryInfo dirInfo = diss[i];

                    foreach (var disStr in diss)
                    {
                        string roadCode = disStr.Name;
                        DirectoryInfo[] dis = disStr.GetDirectories();


                        foreach (var dirNow in dis)
                        {
                            List<string> lines = new List<string>();
                            var tempPath = dirNow.Name;
                            if (tempPath == "Images" || tempPath == "ViewImages")
                            {
                                string indexPath = Directory.GetFiles(dirNow.FullName, "fileindex.txt", SearchOption.AllDirectories).First();
                                //FileInfo indexFile = new FileInfo(indexPath);
                                string[] dirctoryFiles = Directory.GetFiles(dirNow.FullName, "*.jpg", SearchOption.AllDirectories);
                                progressBar1.Value = 0;
                                progressBar1.Maximum = dirctoryFiles.Length;
                                for (int d = 0; d < dirctoryFiles.Length; d++)
                                {
                                    progressBar1.Value = d + 1;
                                    string item = dirctoryFiles[d];
                                    FileInfo pciFile = new FileInfo(item);
                                    string name = Path.GetFileNameWithoutExtension(pciFile.FullName);
                                    int mile = int.Parse((double.Parse(name.Split('_').Last()) * 1000).ToString());
                                    string mileStr = Form1.ConvertIntToFormattedString(mile);
                                    string mileStr1 = mile.ToString("0+000");
                                    string picName = roadCode + "-" + mileStr + "-" + mileStr + ".jpg";
                                    string number = pciFile.Directory.Name;
                                    string lineStr = string.Format($"{roadCode}->{mileStr1}->{number}->{picName}");
                                    pciFile.MoveTo(pciFile.Directory.FullName + "\\" + picName);
                                    lines.Add(lineStr);
                                }
                                File.WriteAllLines(indexPath, lines);

                            }



                        }
                    }

                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //景观和路面图像仅提取一张图片 

                FolderBrowserDialog fd1 = new FolderBrowserDialog();
                fd1.Description = "请选择输出地址";
                var result1 = fd1.ShowDialog();

                if (result1 != DialogResult.OK)
                {
                    return;
                }
                string resultPath = fd1.SelectedPath + "\\影像结果数据\\";


                DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
                DirectoryInfo[] diss = dir.GetDirectories();

                for (int i = 0; i < diss.Length; i++)
                {
                    DirectoryInfo dirInfo = diss[i];

                    foreach (var disStr in diss)
                    {
                        DirectoryInfo[] dis = disStr.GetDirectories();


                        foreach (var dirNow in dis)
                        {
                            List<string> lines = new List<string>();
                            var tempPath = dirNow.Name;
                            if (tempPath == "Images" || tempPath == "ViewImages")
                            {
                                try
                                {
                                    string indexPath = Directory.GetFiles(dirNow.FullName, "fileindex.txt", SearchOption.AllDirectories).First();
                                    //FileInfo indexFile = new FileInfo(indexPath);
                                    string[] dirctoryFiles = Directory.GetFiles(dirNow.FullName, "*.jpg", SearchOption.AllDirectories);

                                    //复制index文件
                                    string targetIndexFilePath = resultPath + indexPath.Replace(fd.SelectedPath, "");
                                    FileInfo tempFile = new FileInfo(targetIndexFilePath);
                                    Directory.CreateDirectory(tempFile.DirectoryName);
                                    File.Copy(indexPath, targetIndexFilePath, true);

                                    //获取第一张图片地址
                                    if (dirctoryFiles.Length > 0)
                                    {
                                        FileInfo tempFile1 = new FileInfo(dirctoryFiles.FirstOrDefault());
                                        string targetPath = resultPath + tempFile1.DirectoryName.Replace(fd.SelectedPath, "");
                                        Directory.CreateDirectory(targetPath);
                                        File.Copy(tempFile1.FullName, targetPath + "\\" + tempFile1.Name);
                                    }
                                }
                                catch (Exception)
                                {

                                    MessageBox.Show("工程" + dirInfo.Name + "缺少必须图片或index文件请检查!");
                                }


                            }



                        }
                    }

                }
            }

            MessageBox.Show("处理完成!");

        }

        private void barButtonItem14_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择农村路规范导出的结果数据文件夹";

            var result = fd.ShowDialog();
            if (result != DialogResult.OK)
            {
                return;
            }
            //处理图片
            //依据合肥规范进行 图片名称 和index.txt文件的修改
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
            DirectoryInfo[] diss = dir.GetDirectories();

            var result1 = MessageBox.Show("修改路面图像索引选择是,修改景观图像索引选择否", "提示", MessageBoxButtons.YesNo);

            ////修改文件夹名称 
            //foreach (var pro in diss)
            //{
            //    string proName = pro.Name;
            //    string newName = proName.Substring(0, 4) + proName.Last() + "-" + proName.Substring(4, 6);
            //    Microsoft.VisualBasic.FileIO.FileSystem.RenameDirectory(pro.FullName,newName);

            //}
            diss = dir.GetDirectories();
            //删除多余数据
            foreach (var item in diss)
            {
                DirectoryInfo[] dirChild = item.GetDirectories();
                foreach (var child in dirChild)
                {
                    string fix = "";
                    if (result1 == DialogResult.Yes)
                    {
                        fix = "Images";
                    }
                    else
                    {
                        fix = "ViewImages";
                    }
                    if (!child.Name.Equals(fix))
                    {
                        Directory.Delete(child.FullName, true);

                    }
                    else
                    {
                        //将里面的内容提取出去
                        DirectoryInfo dirCC = child.GetDirectories().First();
                        DirectoryInfo[] resultDir = dirCC.GetDirectories();
                        foreach (var itemDir in resultDir)
                        {
                            Directory.Move(itemDir.FullName, item.FullName + "/" + itemDir.Name);
                        }
                        FileInfo resultFile = dirCC.GetFiles().First();
                        File.Move(resultFile.FullName, item.FullName + "/" + resultFile.Name);
                        Directory.Delete(child.FullName, true);
                    }
                }
            }

            foreach (var pro in diss)
            {
                //更改index文件
                FileInfo indexFile = pro.GetFiles("fileindex.txt").First();
                List<string> indexTxts = File.ReadLines(indexFile.FullName).ToList();
                //根据各地标准模板_《2024-03-04检测图像格式的要求-wqq.docx》进行修改
                List<string> newIndexTxts = new List<string>();
                foreach (var line in indexTxts)
                {
                    string newLine = "";
                    string tempStr = line.Replace("-", string.Empty);
                    string mileMsg = tempStr.Split('>')[1];
                    newLine += mileMsg + ".jpg" + "->" + tempStr.Split('>').Last() + "->" + "\\" + tempStr.Split('>')[2];
                    newIndexTxts.Add(newLine);
                }
                File.WriteAllLines(indexFile.FullName, newIndexTxts);
            }
            MessageBox.Show("处理完成");
        }

        private void barButtonItem15_ItemClick(object sender, ItemClickEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择农村路规范导出的结果数据文件夹";
            var result = fd.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }

            //获取上下行 
            DirectoryInfo dir = new DirectoryInfo(fd.SelectedPath);
            DirectoryInfo[] diss = dir.GetDirectories();

            this.treeList2.Nodes.Clear();
            foreach (var pro in diss)
            {
                TreeListNode node = this.treeList2.AppendNode(null, null);
                // node.SetValue("name",pro._DataDir.Name) ;
                node.SetValue("data", pro.Name);
            }

            if (dir.Name != "结果数据")
            {
                MessageBox.Show("请选择农村路规范导出的结果数据文件夹");
                return;
            }
            foreach (var disStr in diss)
            {
                DirectoryInfo[] dis = disStr.GetDirectories();
                foreach (var dirPath in dis)
                {
                    var tempPath = dirPath.Name;
                    if (tempPath == "DR")
                    {
                        //去除文件名中 最后的中文字符 如 指南水泥
                        FileInfo[] files = dirPath.GetFiles();
                        foreach (var item in files)
                        {
                            string names = Path.GetFileNameWithoutExtension(item.FullName);
                            string newName = names.Substring(0, names.LastIndexOf('-')) + ".txt";
                            //第一行桩号改成起点桩号

                            string[] lines = File.ReadAllLines(item.FullName);
                            lines[0] = lines.First().Replace("桩号", "起点桩号");
                            File.WriteAllLines(item.FullName, lines);
                            Microsoft.VisualBasic.FileIO.FileSystem.RenameFile(item.FullName, newName);



                        }
                    }
                    if (tempPath == "LBI")
                    {
                        FileInfo[] files = dirPath.GetFiles();
                        foreach (var item in files)
                        {
                            string names = Path.GetFileNameWithoutExtension(item.FullName);

                            //第一行桩号改成起点桩号

                            string[] lines = File.ReadAllLines(item.FullName);
                            lines[0] = "起点桩号(km),经度,纬度";
                            for (int i = 1; i < lines.Length; i++)
                            {
                                lines[i] = lines[i].Substring(0, lines[i].LastIndexOf(','));
                            }

                            File.WriteAllLines(item.FullName, lines);
                        }
                    }

                    if (tempPath == "RIFile")
                    {
                        FileInfo[] files = dirPath.GetFiles();
                        foreach (var item in files)
                        {
                            string names = Path.GetFileNameWithoutExtension(item.FullName);

                            //第一行桩号改成起点桩号

                            string[] lines = File.ReadAllLines(item.FullName);
                            lines[0] = "起点桩号(km),左高程(mm),右高程(mm),速度(m/s)";

                            File.WriteAllLines(item.FullName, lines);
                        }
                    }
                }
            }

            string oldDirectoryName = fd.SelectedPath; // 旧文件夹路径
            string newDirectoryName = fd.SelectedPath.Substring(0, fd.SelectedPath.LastIndexOf("\\")) + "\\交通部二次修改"; // 新文件夹路径

            try
            {
                // 检查旧文件夹是否存在
                if (Directory.Exists(oldDirectoryName))
                {
                    // 修改文件夹名称
                    Directory.Move(oldDirectoryName, newDirectoryName);
                    Console.WriteLine("文件夹名称修改成功！");
                }
                else
                {
                    Console.WriteLine("指定的旧文件夹不存在！");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("修改文件夹名称时发生错误: " + ex.Message);
            }

            MessageBox.Show("所有项目处理完毕请检查!");
        }

        private void ribbonControl1_Click(object sender, EventArgs e)
        {

        }



        private void barButtonItem16_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void barButtonItem18_ItemClick(object sender, ItemClickEventArgs e)
        {
            currentStandard = (CityModelItem)Enum.Parse(typeof(CityModelItem), barEditItem1.EditValue.ToString());
            string helpfile = "";
            switch (currentStandard)
            {
                case CityModelItem.等级公路5210与农村路5211标准模板导出_2025年:

                    helpfile = System.Windows.Forms.Application.StartupPath + "\\规范2025\\2025年度农村公路技术状况检测评定数据报送工作.pdf";

                    break;
                case CityModelItem.交通部2024规范:
                    break;
                case CityModelItem.河南省单位一农村路定制:
                    break;
                case CityModelItem.湖南省单位一定制:
                    helpfile = System.Windows.Forms.Application.StartupPath + "\\规范2025\\2025湖南高速检测设备提交数据格式要求.docx";

                    break;
                case CityModelItem.重庆市单位一定制:
                    break;
                case CityModelItem.甘肃省单位一定制:
                    break;
                case CityModelItem.河北省单位一定制:
                    break;
                case CityModelItem.河北省单位二定制:
                    break;
                case CityModelItem.江苏省单位一定制:
                    break;
                case CityModelItem.安徽省单位一定制:
                    break;
                case CityModelItem.广东省单位一定制:
                    break;
                case CityModelItem.河北省单位定制:
                    helpfile = System.Windows.Forms.Application.StartupPath + "\\规范2025\\附件2：农村公路技术状况数据存储报送技术.pdf";
                    break;
                case CityModelItem.辽宁省2025单位定制:
                    helpfile = System.Windows.Forms.Application.StartupPath + "\\规范2025\\辽宁原始数据提交格式要求.pdf";
                    break;
                case CityModelItem.农养国省道路况检测数据提交格式_2026年:
                    helpfile = System.Windows.Forms.Application.StartupPath + "\\规范2026\\2026年农养国省道路况检测数据提交格式要求.docx";

                    break;
                default:
                    break;
            }
            System.Diagnostics.Process.Start(helpfile);
        }
        List<GJProject> GjProjects = new List<GJProject>();

        private void barButtonItem19_ItemClick(object sender, ItemClickEventArgs e)
        {
            InputGjProjects();
        }

        private void barButtonItem20_ItemClick(object sender, ItemClickEventArgs e)
        {
            VistaFolderBrowserDialog dlg;
            if (Directory.Exists(_config.UserPath))
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择检查结果输出文件夹",

                    SelectedPath = _config.UserPath,
                    ShowNewFolderButton = true
                };
            }
            else
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择检查结果输出文件夹",


                    ShowNewFolderButton = true
                };
            }
            if (dlg.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            double dis = 0.25;
            DialogResult result = MessageBox.Show(
      "请选择间隔：\n\n是(Y) - 0.1\n否(N) - 0.25",
      "间隔选择",
      MessageBoxButtons.YesNo,
      MessageBoxIcon.Question
  );

            dis = result == DialogResult.Yes ? 0.1 : 0.25;

            if (dlg.SelectedPath != String.Empty)
            {


                string selectOutPath = dlg.SelectedPath;

                this.progressBar1.Value = 0;
                this.progressBar1.Minimum = 0;
                this.progressBar1.Maximum = GjProjects.Count;
                for (int j = 0; j < GjProjects.Count; j++)
                {

                    var pro = GjProjects[j];
                    try
                    {
                        pro.CheckIirValue(selectOutPath, dis);
                        this.progressBar1.Value = j + 1;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }

            }
            MessageBox.Show("所有项目处理完毕请检查!");
        }

        private void barButtonItem21_ItemClick(object sender, ItemClickEventArgs e)
        { 
                if (GjProjects.Count != 1)
                {
                    MessageBox.Show("该功能仅支持导入一个工程使用!请使用国检结果数据导入导入一个工程");
                    return;
                }
                LpFileCalculateIri checkIri = new LpFileCalculateIri(GjProjects[0]);
                checkIri.ShowDialog(); 
        }

        private string InputGjProjects()
        {
            string returnSelectPath = "";
            treeList2.Nodes.Clear();
            List<DirectoryInfo> projects = new List<DirectoryInfo>();



            VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog
            {
                Description = "选择结果文件夹",
                SelectedPath = _config.UserPath,
                ShowNewFolderButton = true
            };
            // 使用 FolderBrowserDialog 替代 CommonOpenFileDialog
            //using (FolderBrowserDialog dlg = new FolderBrowserDialog
            //{
            //    Description = "选择结果文件夹",
            //    SelectedPath = _config.UserPath,
            //    ShowNewFolderButton = true
            //})
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) // 传入父窗体，确保居中
                    return returnSelectPath;

                if (!string.IsNullOrEmpty(dlg.SelectedPath))
                {
                    returnSelectPath = dlg.SelectedPath;
                    string selectPath = dlg.SelectedPath.TrimEnd('\\');
                    Chktxt_Path = selectPath;
                    _config.UserPath = selectPath;
                    ConfigManager.SaveConfig();

                    // 获取国检工程
                    DirectoryInfo dir = new DirectoryInfo(selectPath);
                    DirectoryInfo[] folders = null;
                    GjProjects.Clear();

                    if (dir.Name.Length == 11)
                    {
                        folders = new DirectoryInfo[] { dir };
                    }
                    else
                    {
                        folders = dir.GetDirectories()
                            .Where(f => f.Name.Length == 11)
                            .ToArray();
                    }

                    for (int i = 0; i < folders.Length; i++)
                    {
                        GJProject gJProject = new GJProject(folders[i].FullName);
                        GjProjects.Add(gJProject);
                    }

                    foreach (var pro in GjProjects)
                    {
                        TreeListNode node = treeList2.AppendNode(null, null);
                        node.SetValue("data", pro.GjProjectName);
                    }
                }
            }
            return returnSelectPath;
        }



        // C#

        /// <summary>
        /// 图片处理按钮的点击事件处理程序。
        /// 这个方法负责UI交互和启动后台任务，它本身不执行任何耗时操作。
        /// </summary>
        private async void barButtonItem22_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            // --- 1. 准备阶段：加载数据和获取用户输入 ---

            // 加载项目数据（这部分逻辑与你原来的一样）
            string userSelectPath = InputGjProjects();
            if (GjProjects.Count <= 0)
            {
                MessageBox.Show(this, "请先导入国检结果数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            foreach (var item in GjProjects)
            {
                item.InitRoadPictureFileList();
            }
            // 将所有需要处理的图片路径收集到一个列表中
            List<string> allImagePaths = GjProjects.SelectMany(p => p.RoadPictures).ToList();
            int totalImageCount = allImagePaths.Count;
            if (totalImageCount == 0)
            {
                MessageBox.Show(this, "在项目中未找到任何图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 获取用户设置的目标尺寸
            (int originalWidth, int originalHeight) = GjProjects.First().GetPcitureSize();
            using (var form = new ChangeRoadPictureSizeForm(originalWidth, originalHeight) { StartPosition = FormStartPosition.CenterParent })
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                    return;

                (int targetWidth, int targetHeight) = form.getUserSize();

                // --- 2. 确认阶段：向用户展示信息并请求确认 ---

                var confirmResult = MessageBox.Show(this,
                    $"即将开始处理图片，目标尺寸: {targetWidth}x{targetHeight}。\n" +
                    $"共需处理 {totalImageCount} 张图片。\n\n" +
                    "请注意：此操作将直接修改原始图片，建议提前备份！\n\n是否继续？",
                    "操作确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirmResult != DialogResult.Yes)
                    return;

                // --- 3. 执行阶段：设置UI并调用后台核心任务 ---

                // 禁用按钮，防止重复点击
                barButtonItem22.Enabled = false;

                // 初始化进度条和标签
                progressBar1.Maximum = totalImageCount;
                progressBar1.Value = 0;
                label2.Text = $"0% (0 / {totalImageCount})";

                // 创建一个 Progress 对象，用于从后台线程安全地更新UI
                // 每次后台报告 "1" (表示完成一张图片)，这里的委托就会在UI线程上执行
                var progressReporter = new Progress<int>(processedCount =>
                {
                    progressBar1.Value = processedCount;
                    double percentage = (double)processedCount / totalImageCount * 100;
                    label2.Text = $"{percentage:F1}% ({processedCount} / {totalImageCount})";
                });

                // 存储处理过程中发生的错误
                ConcurrentBag<string> errorLog = new ConcurrentBag<string>();

                try
                {
                    // **核心步骤**: 调用并等待后台任务完成
                    // 所有耗时操作都在这个方法里，UI线程在这里被释放，界面保持流畅
                    await ProcessImagesInBackgroundAsync(allImagePaths, targetWidth, targetHeight, progressReporter, errorLog);

                    // --- 4. 完成阶段：处理结果 ---
                    progressBar1.Value = totalImageCount; // 确保进度条达到100%
                    label2.Text = $"100% ({totalImageCount} / {totalImageCount})";

                    // 根据是否有错误日志，显示不同的提示信息
                    if (errorLog.IsEmpty)
                    {
                        MessageBox.Show(this, "所有图片尺寸已成功调整！", "处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string logPath = Path.Combine(userSelectPath, "RoadImgHandel_error_log.txt");
                        try
                        {
                            File.WriteAllLines(logPath, errorLog);
                            MessageBox.Show(this, $"图片处理完成，但有 {errorLog.Count} 个错误发生。\n详情请查看日志文件：\n{logPath}", "部分成功", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, $"图片处理完成，但写入错误日志失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 捕获意料之外的全局异常
                    MessageBox.Show(this, $"处理过程中发生严重错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // 无论成功与否，最后都要恢复按钮可用状态
                    barButtonItem22.Enabled = true;
                }
            }
        }



        // C#
        /// <summary>
        /// 在后台并发处理所有图片的核心方法。
        /// 这个版本使用了 SemaphoreSlim 和 Task.WhenAll，以兼容旧版 .NET Framework。
        /// 它与 Parallel.ForEachAsync 的效果相同：并发执行任务，同时限制并发数量。
        /// </summary>
        /// <param name="imagePaths">所有待处理图片的完整路径列表。</param>
        /// <param name="targetWidth">目标宽度。</param>
        /// <param name="targetHeight">目标高度。</param>
        /// <param name="progress">用于向UI线程报告进度的IProgress对象。</param>
        /// <param name="errorLog">用于记录错误的线程安全的集合。</param>
        //private async Task ProcessImagesInBackgroundAsync(List<string> imagePaths, int targetWidth, int targetHeight, IProgress<int> progress, ConcurrentBag<string> errorLog)
        //{
        //    // 获取JPEG编码器，用于设置保存质量
        //    ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(codec => codec.MimeType == "image/jpeg");
        //    if (jpegCodec == null)
        //    {
        //        errorLog.Add("系统错误：未能找到JPEG图片编码器。");
        //        return;
        //    }

        //    // 设置图片保存质量 (90L 表示 90% 的质量)
        //    using (EncoderParameters encoderParams = new EncoderParameters(1))
        //    {
        //        encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);

        //        int processedCount = 0; // 已处理的图片计数器

        //        // 1. 创建一个 SemaphoreSlim 对象来控制并发数量。
        //        // 初始值为CPU核心数，意味着最多允许N个任务同时运行，N=CPU核心数。
        //        var semaphore = new SemaphoreSlim(Environment.ProcessorCount);

        //        // 2. 创建一个任务列表，用于存放所有图片的处理任务
        //        List<Task> tasks = new List<Task>();

        //        // 3. 遍历所有图片路径，为每张图片创建一个处理任务
        //        foreach (var imagePath in imagePaths)
        //        {
        //            // a. 异步等待一个信号量"名额"。如果当前正在运行的任务已达上限，代码会在这里暂停，直到有任务完成并释放名额。
        //            await semaphore.WaitAsync();

        //            // b. 启动一个新的任务来处理图片。
        //            // Task.Run 确保图片处理的全部逻辑都在后台线程池中执行。
        //            tasks.Add(Task.Run(async () =>
        //            {


        //                try
        //                {
        //                    // 提前检查文件存在
        //                    if (!File.Exists(imagePath))
        //                    {
        //                        errorLog.Add($"文件不存在: {imagePath}");
        //                        return;
        //                    }

        //                    // LoadAsync 使用 Stream
        //                    using var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        //                    using (var image = await SixLabors.ImageSharp.Image.LoadAsync(stream))
        //                    {
        //                        if (image.Width == targetWidth && image.Height == targetHeight)
        //                        {
        //                            return; // 尺寸符合，跳过处理
        //                        }

        //                        image.Mutate(x => x.Resize(new ResizeOptions
        //                        {
        //                            Size = new SixLabors.ImageSharp.Size(targetWidth, targetHeight),
        //                            Mode = ResizeMode.Stretch,
        //                            Sampler = KnownResamplers.Bicubic
        //                        }));

        //                        string tempFilePath = imagePath + ".tmp";
        //                        // SaveAsync 使用 Stream
        //                        using var outputStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
        //                        await image.SaveAsync(outputStream, new JpegEncoder { Quality = 90 });

        //                        // 确保 image 和 stream 已释放
        //                        outputStream.Close(); // 显式关闭输出流
        //                        stream.Close(); // 显式关闭输入流

        //                        // 删除原始文件并替换
        //                        try
        //                        {
        //                            File.Delete(imagePath);
        //                            File.Move(tempFilePath, imagePath);
        //                        }
        //                        catch (IOException ioEx)
        //                        {
        //                            errorLog.Add($"文件操作失败: {imagePath} | 错误: {ioEx.Message}");
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    errorLog.Add($"处理失败: {imagePath} | 错误: {ex.Message}");
        //                }
        //                finally
        //                {
        //                    // c. 关键步骤！无论成功还是失败，都必须在 finally 块中释放信号量名额，
        //                    // 以便等待中的其他任务可以开始执行。
        //                    semaphore.Release();

        //                    // 更新进度 (与之前相同)
        //                    int currentCount = Interlocked.Increment(ref processedCount);
        //                    progress.Report(currentCount);
        //                }

        //            }));
        //        }

        //        // 4. 等待所有已启动的任务全部完成
        //        await Task.WhenAll(tasks);
        //    }
        //}


        /// <summary>
        /// 在后台并发处理所有图片的核心方法。
        /// 使用 SemaphoreSlim 和 Task.WhenAll，兼容 .NET Framework 4.8。
        /// 功能：检查图片尺寸，如果不符合目标宽高，则缩放并以 JPEG 质量 90 保存。
        /// </summary>
        /// <param name="imagePaths">所有待处理图片的完整路径列表。</param>
        /// <param name="targetWidth">目标宽度。</param>
        /// <param name="targetHeight">目标高度。</param>
        /// <param name="progress">用于向 UI 线程报告进度的 IProgress 对象。</param>
        /// <param name="errorLog">用于记录错误的线程安全集合。</param>
        private async Task ProcessImagesInBackgroundAsync(
            List<string> imagePaths,
            int targetWidth,
            int targetHeight,
            IProgress<int> progress,
            ConcurrentBag<string> errorLog)
        {
            if (imagePaths == null || imagePaths.Count == 0)
            {
                return;
            }

            ImageCodecInfo jpegCodec = ImageCodecInfo.GetImageEncoders()
                .FirstOrDefault(codec => codec.MimeType == "image/jpeg");

            if (jpegCodec == null)
            {
                errorLog.Add("系统错误：未能找到 JPEG 图片编码器。");
                return;
            }

            int processedCount = 0;

            // 不建议直接使用 Environment.ProcessorCount。
            // 图片处理同时涉及 CPU + 磁盘 IO，尤其客户电脑如果是机械硬盘，并发太高反而更慢。
            int maxDegreeOfParallelism = Math.Max(1, Math.Min(Environment.ProcessorCount, 4));

            using (SemaphoreSlim semaphore = new SemaphoreSlim(maxDegreeOfParallelism))
            {
                List<Task> tasks = new List<Task>();

                foreach (string imagePath in imagePaths)
                {
                    await semaphore.WaitAsync();

                    Task task = Task.Run(() =>
                    {
                        string tempFilePath = null;

                        try
                        {
                            if (string.IsNullOrWhiteSpace(imagePath))
                            {
                                errorLog.Add("图片路径为空。");
                                return;
                            }

                            if (!File.Exists(imagePath))
                            {
                                errorLog.Add("文件不存在: " + imagePath);
                                return;
                            }

                            tempFilePath = imagePath + "." + Guid.NewGuid().ToString("N") + ".tmp";

                            using (FileStream inputStream = new FileStream(
                                imagePath,
                                FileMode.Open,
                                FileAccess.Read,
                                FileShare.Read))
                            using (Image sourceImage = Image.FromStream(inputStream, true, false))
                            {
                                if (sourceImage.Width == targetWidth && sourceImage.Height == targetHeight)
                                {
                                    return;
                                }

                                using (Bitmap destBitmap = new Bitmap(targetWidth, targetHeight, PixelFormat.Format24bppRgb))
                                {
                                    destBitmap.SetResolution(sourceImage.HorizontalResolution, sourceImage.VerticalResolution);

                                    using (Graphics graphics = Graphics.FromImage(destBitmap))
                                    {
                                        graphics.Clear(Color.White);

                                        graphics.CompositingQuality = CompositingQuality.HighQuality;
                                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                        graphics.SmoothingMode = SmoothingMode.HighQuality;
                                        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                                        graphics.DrawImage(
                                            sourceImage,
                                            new Rectangle(0, 0, targetWidth, targetHeight),
                                            new Rectangle(0, 0, sourceImage.Width, sourceImage.Height),
                                            GraphicsUnit.Pixel);
                                    }

                                    using (EncoderParameters encoderParams = new EncoderParameters(1))
                                    {
                                        encoderParams.Param[0] = new EncoderParameter(
                                            System.Drawing.Imaging.Encoder.Quality,
                                            90L);

                                        destBitmap.Save(tempFilePath, jpegCodec, encoderParams);
                                    }
                                }
                            }

                            // 到这里 sourceImage / inputStream / destBitmap 都已经释放了，可以安全替换原文件
                            File.Delete(imagePath);
                            File.Move(tempFilePath, imagePath);
                            tempFilePath = null;
                        }
                        catch (Exception ex)
                        {
                            errorLog.Add("处理失败: " + imagePath + " | 错误: " + ex.Message);
                        }
                        finally
                        {
                            // 如果中途失败，清理临时文件
                            if (!string.IsNullOrEmpty(tempFilePath))
                            {
                                try
                                {
                                    if (File.Exists(tempFilePath))
                                    {
                                        File.Delete(tempFilePath);
                                    }
                                }
                                catch
                                {
                                    // 清理临时文件失败，不影响主流程
                                }
                            }

                            semaphore.Release();

                            int currentCount = Interlocked.Increment(ref processedCount);

                            if (progress != null)
                            {
                                progress.Report(currentCount);
                            }
                        }
                    });

                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);
            }
        }


        private void barButtonItem23_ItemClick(object sender, ItemClickEventArgs e)
        {
            // --- 1. 准备阶段：加载数据和获取用户输入 ---

            // 加载项目数据（这部分逻辑与你原来的一样）
            string userSelectPath = InputGjProjects();
            if (GjProjects.Count <= 0)
            {
                MessageBox.Show(this, "请先导入国检结果数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            foreach (var item in GjProjects)
            {
                item.InitRoadPictureFileList();
            }
            // 将所有需要处理的图片路径收集到一个列表中
            List<string> allImagePaths = GjProjects.SelectMany(p => p.RoadPictures).ToList();
            int totalImageCount = allImagePaths.Count;
            if (totalImageCount == 0)
            {
                MessageBox.Show(this, "在项目中未找到任何图片。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 获取用户设置的目标尺寸
            (int originalWidth, int originalHeight) = GjProjects.First().GetPcitureSize();
            Dictionary<GJProject, List<string>> errorDic = new Dictionary<GJProject, List<string>>();
            using (var form = new ChangeRoadPictureSizeForm(originalWidth, originalHeight) { StartPosition = FormStartPosition.CenterParent })
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                    return;

                (int targetWidth, int targetHeight) = form.getUserSize();


                for (int i = 0; i < GjProjects.Count; i++)
                {
                    var curPorject = GjProjects[i];
                    if (curPorject.RoadPictures.Count == 0)
                    {

                        if (errorDic.ContainsKey(curPorject))
                        {

                            errorDic[curPorject].Add("未找到任何图片;");

                        }
                        else
                        {

                            errorDic[curPorject] = new List<string>() { "未找到任何图片;" };
                        }

                    }
                    else
                    {
                        for (int t = 0; t < curPorject.RoadPictures.Count; t++)
                        {
                            string curPath = curPorject.RoadPictures[t];
                            FileInfo curFile = new FileInfo(curPath);

                            try
                            {
                                // 检查文件是否存在
                                if (!curFile.Exists)
                                {

                                    if (errorDic.ContainsKey(curPorject))
                                    {

                                        errorDic[curPorject].Add($"图片{curPath}不存在;");

                                    }
                                    else
                                    {

                                        errorDic[curPorject] = new List<string>() { $"图片{curPath}不存在;" };
                                    }

                                    continue;
                                }

                                // 获取图片尺寸
                                using (Image image = Image.FromFile(curPath))
                                {
                                    int width = image.Width;
                                    int height = image.Height;

                                    // 这里可以添加对图片尺寸的处理逻辑
                                    // 例如检查是否符合要求尺寸
                                    if (width != targetWidth || height != targetHeight)
                                    {
                                        if (errorDic.ContainsKey(curPorject))
                                        {

                                            errorDic[curPorject].Add($"{curPath}图片尺寸不符合: w:{width}h:{height};");



                                        }
                                        else
                                        {

                                            errorDic[curPorject] = new List<string>() { $"{curPath}图片尺寸不符合: w:{width}h:{height};" };
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (errorDic.ContainsKey(curPorject))
                                {

                                    errorDic[curPorject].Add($"{curPath}读取图片失败:");
                                }
                                else
                                {

                                    errorDic[curPorject] = new List<string>() { $"{curPath}读取图片失败:" };
                                }
                            }
                        }
                    }

                    if (errorDic.Count > 0)
                    {
                        VistaFolderBrowserDialog dlg = new VistaFolderBrowserDialog
                        {
                            Description = "选择错误信息文件夹",
                            SelectedPath = _config.UserPath,
                            ShowNewFolderButton = true
                        };
                        if (dlg.ShowDialog() != DialogResult.OK)
                        {
                            return;
                        }
                        if (Directory.Exists(dlg.SelectedPath))
                        {
                            string errorPath = dlg.SelectedPath + "\\RoadImgCheck_error_log.txt";

                            List<string> errorLines = new List<string>();
                            foreach (var item in errorDic)
                            {
                                errorLines.Add($"工程{item.Key.GjProjectName}错误信息:");
                                errorLines.AddRange(item.Value);
                                errorLines.Add("---------------------------------------------------");
                            }
                            File.WriteAllLines(errorPath, errorLines);
                            MessageBox.Show($"检查完成,请查看错误日志文件:{errorPath}");


                        }
                    }
                    else
                    {
                        MessageBox.Show("工程所有图片已符合用户设置！");
                    }




                }

            }
        }

        private void barButtonItem26_ItemClick(object sender, ItemClickEventArgs e)
        {

        }
        /// <summary>
        /// 获得工程信息
        /// </summary>
        /// <param name="readExcelResult">是否寻找excel结果文件夹</param>
        private void getProject(bool readExcelResult)
        {
            this.treeList2.Nodes.Clear();
            List<DirectoryInfo> projects = new List<DirectoryInfo>();
            VistaFolderBrowserDialog dlg;
            if (Directory.Exists(_config.UserPath))
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",

                    SelectedPath = _config.UserPath,
                    ShowNewFolderButton = true
                };
            }
            else
            {
                dlg = new VistaFolderBrowserDialog
                {
                    Description = "选择结果文件夹",


                    ShowNewFolderButton = true
                };
            }
            if (dlg.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            if (dlg.SelectedPath != String.Empty)
            {
                string selectPath = dlg.SelectedPath;
                if (selectPath.Substring(dlg.SelectedPath.Length - 1) == "\\")
                {
                    selectPath = selectPath.Remove(dlg.SelectedPath.Length - 1);
                }
                this.Chktxt_Path = selectPath;
                _config.UserPath = selectPath;
                ConfigManager.SaveConfig();
                projects = GetAllProjectPath(selectPath);

                 
                foreach (DirectoryInfo dir in projects)
                {

                    ProjectInfo proj = new ProjectInfo(readExcelResult, dir.FullName);
                    proj._DataDir = dir;
                    _Projects.Add(proj);
                }

                _Projects.Sort((a, b) => StrCmpLogicalW(a._DataDir.Name, b._DataDir.Name)); 
                foreach (var pro in _Projects)
                {
                    TreeListNode node = this.treeList2.AppendNode(null, null);
                    // node.SetValue("name",pro._DataDir.Name) ;
                    node.SetValue("data", pro._DataDir.Name);
                }
            }
        }

        private void barButtonItem24_ItemClick(object sender, ItemClickEventArgs e)
        {
            _Projects.Clear();
            getProject(false);
        }
        XRSetting _Setting = XRSetting.GetInstance();


 

        private void barButtonItem25_ItemClick(object sender, ItemClickEventArgs e)
        {
            progressBar1.Value = 0;
           // progressBar1.Maximum = 200;  
            getProject(true);
            this.treeList2.ClearNodes();

            gjModelSelectForm form = new gjModelSelectForm();
            //解析数据 
           
            if (form.ShowDialog() == DialogResult.OK)
            {
                if (_Setting.gjStandardNew == hnEnumTools.CityModelItem.等级公路5210与农村路5211标准模板导出_2025年)
                {
                    string cityCode = "";
                    foreach (var pro in _Projects)
                    {
                      
                        string proName = "";

                        StandardParmType ParmStyle = pro.Standard;
                        int SelectDrawDis = pro.DrawType; 
                        _Setting.ParmStyle = ParmStyle;
                        _Setting.SelectDrawDis = SelectDrawDis;
                        CreateConventSource_Universality(pro, out proName, ref cityCode, _Setting.gjStandardNew);


                        var roadParts = pro.getMiles();

                        var datas = pro.GetAllDisease();
                        //pro.Analysis23DExcelData(Bar1process);

                        switch (ParmStyle)
                        {
                            case StandardParmType.DegreeRoad2018:
                                { 
                                    switch (SelectDrawDis)
                                    {
                                        case 0: 
                                            {
                                              
                                            
                                            }
                                            break;
                                        case 1:
                                            {
                                               
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                break;
                              
                            case StandardParmType.RuralRoadlowLevel:
                                {

                                    switch (SelectDrawDis)
                                    {
                                        case 0: 
                                            {
                                             }
                                            break;
                                        case 1:
                                            { 
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                                break;
                            case StandardParmType.DegreeRoad2007:
                                break;
                            case StandardParmType.CityRoad:
                                break;
                            case StandardParmType.RuralRoadBeijing:
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
                                break;

                            default:
                                break;
                        }
                    }



                }
                MessageBox.Show("所有通用性文件已输出，请使用《国检转换软件》选择具体【国检标准】\n《国检转换软件》将进行【文件格式转换】【图片规范输出】等操作！");
            }



        }
        Dictionary<string, DirectoryInfo> ConverDic = new Dictionary<string, DirectoryInfo>();
        /// <summary>
        /// 创建统一模板目录 写入配置参数
        /// </summary>
        /// <param name="pro"></param>
        /// <param name="RoadName"></param>
        /// <param name="cityCode"></param>
        /// <param name="standard"></param>
        /// <returns></returns>
        private bool CreateConventSource_Universality(ProjectInfo pro, out string RoadName, ref string cityCode, hnEnumTools.CityModelItem standard)
        {

            ConverDic.Clear();
            string dirc = pro._DirectionInt > 0 ? "A" : "B";
            //检查是否具有县级行政代码 

            if (string.IsNullOrWhiteSpace(cityCode))
            {
                if (string.IsNullOrEmpty(pro._RoadCode ))
                {
                    MessageBox.Show("请检查" + pro. _RoadName + "道路未设置道路编号!");
                    RoadName = "";
                    return false;
                }
                ConventGetRoadNumberForm form = new ConventGetRoadNumberForm(pro. _RoadCode);
                form.ShowDialog();
                if (form.Ok)
                {
                    RoadName = $"{form.Name}{dirc}"; //路线编码+县级行政代码+方向
                    string tmppath = pro._DataDir.FullName + @"\ProjectInfo.txt";
                    pro._CityCode = form.CityNum;
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
                string temp = pro. _RoadCode.Substring(0, 4);
                RoadName = $"{temp}{cityCode}{dirc}"; //路线编码+县级行政代码+方向
            }
            string timeYeals = pro. _DataDate;
            string timeHours = pro. _DataTime;
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

        public  void ConvertAllTxtInFolder(string folderPath, string searchPattern = "*.txt")
        {
            Encoding utf8 = Encoding.UTF8;
            Encoding ansi = Encoding.Default;  // Windows 下的 ANSI，即当前系统代码页（如中文系统为 GBK/CP936）
            Encoding gbk = Encoding.GetEncoding("GBK");
            foreach (string filePath in Directory.GetFiles(folderPath, searchPattern, SearchOption.AllDirectories))
            {
                 
                string tempFile = filePath + ".tmp";  // 临时文件，避免直接覆盖时出错

                try
                {

                    string[] lines =   File.ReadAllLines(filePath); // 测试是否能用 UTF-8 读取

                     File.WriteAllLines(tempFile, lines, ansi); // 使用 ANSI 编码写入临时文件

                    string[] tempFileStrs = File.ReadAllLines(tempFile, ansi);
                    // 步骤3: 替换原文件（确保转换成功后再覆盖）
                    File.Delete(filePath);          // 删除原文件
                    File.Move(tempFile, filePath);  // 临时文件改名为原文件

                    Console.WriteLine($"已转换: {filePath}");
                }
                catch (Exception ex)
                {
                    // 如果出错，尝试清理临时文件
                    if (File.Exists(tempFile))
                    {
                        try { File.Delete(tempFile); } catch { }
                    }

                    Console.WriteLine($"转换失败 {filePath}: {ex.Message}");
                }
            }
        }

        private void barButtonItem27_ItemClick(object sender, ItemClickEventArgs e)
        {
            string userSelectPath = InputGjProjects();
            if (GjProjects.Count <= 0)
            {
                MessageBox.Show(this, "请先导入国检结果数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

          var result =   MessageBox.Show("即将将所有检测到的文本文件转换为ANSI编码，确定执行吗？", "提示", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
            {
                for (int i = 0; i < GjProjects.Count; i++)
                { 
                    ConvertAllTxtInFolder(GjProjects[i].GjDirPath);
                }
                MessageBox.Show("已完成转换。\n注意！由于Windows记事本默认以UTF-8打开文件，其对于中文文件编码检测可能存在问题，导致乱码，可以notepad++打开查看转换后的文件!");
            }
            else
            {

            }
          
        }
    }
}
