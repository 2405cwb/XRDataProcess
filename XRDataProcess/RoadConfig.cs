using System;
using System.IO;
using Framework.Other;
using OperateIniFile;

namespace XRDataProcess
{
    class RoadConfig
    {
        private RoadConfig() { }

        private static RoadConfig singleInstance = null;
        public static RoadConfig GetInstance()
        {
            return singleInstance;
        }

        static RoadConfig()
        {
            singleInstance = new RoadConfig();
        }

        // ==================== 常量 & 路径 ====================
        private const string IniFileName = "RoadConfig.ini";
        private const string IniDefaultFileName = "RoadConfig_Default.ini"; // 安装目录默认模板

        /// <summary>
        /// 获取用户可写配置路径
        /// </summary>
        private static string GetUserIniPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);
            return Path.Combine(appFolder, IniFileName);
        }

        /// <summary>
        /// 获取安装目录默认配置路径
        /// </summary>
        private static string GetDefaultIniPath()
        {
            return Path.Combine(System.Windows.Forms.Application.StartupPath, IniDefaultFileName);
        }

        /// <summary>
        /// 首次运行时复制默认配置
        /// </summary>
        private static void EnsureDefaultIniCopied()
        {
            string userPath = GetUserIniPath();
            string defaultPath = GetDefaultIniPath();

            if (!File.Exists(userPath) && File.Exists(defaultPath))
            {
                try { File.Copy(defaultPath, userPath); }
                catch { /* 静默忽略 */ }
            }
        }

        // ==================== 读取配置 ====================
        public void ReadData()
        {
            // 1. 确保默认配置已复制
            EnsureDefaultIniCopied();

            // 2. 读取用户配置
            string userIniPath = GetUserIniPath();
            IniFiles inisetting = new IniFiles(userIniPath);

            ImageWidth = inisetting.ReadInteger("ImageInfo", "ImageWidth", 0);
            ImageHeight = inisetting.ReadInteger("ImageInfo", "ImageHeight", 0);
            RealWidth = inisetting.ReadDouble("ImageInfo", "RealWidth", 0);
            RealHeight = inisetting.ReadDouble("ImageInfo", "RealHeight", 0);

            if (RealWidth == 0) RealWidth = 3.2;
            if (RealHeight == 0) RealHeight = 3.2;

            DetectWidth = inisetting.ReadDouble("ImageInfo", "DetectWidth", 0);
            if (DetectWidth == 0) DetectWidth = RealWidth; // 兼容旧逻辑

            WidthScale = RealWidth * 1.0 / ImageWidth;
            HeightScale = RealHeight * 1.0 / ImageHeight;
            PartWidthNum = (int)(RealWidth * 10);
            PartHeightNum = (int)(RealHeight * 10);

            PartImgWidth = ImageWidth > 0 ? Convert.ToInt32(ImageWidth * 1.0 / PartWidthNum) : 0;
            PartImgHeight = ImageHeight > 0 ? Convert.ToInt32(ImageHeight * 1.0 / PartHeightNum) : 0;
        }

        // ==================== 写入配置 ====================
        public void WriteData()
        {
            try
            {
                string userIniPath = GetUserIniPath();
                IniFiles inisetting = new IniFiles(userIniPath);

                inisetting.WriteInteger("ImageInfo", "ImageWidth", ImageWidth);
                inisetting.WriteInteger("ImageInfo", "ImageHeight", ImageHeight);
                inisetting.WriteDouble("ImageInfo", "RealWidth", RealWidth);
                inisetting.WriteDouble("ImageInfo", "RealHeight", RealHeight);
                inisetting.WriteDouble("ImageInfo", "DetectWidth", RealWidth); // 保持原逻辑
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("RoadConfig WriteData 失败: " + ex.Message);
                // 如使用 log4net：log.Error("RoadConfig 保存失败", ex);
            }
        }

        // ==================== 配置字段 ====================
        public int ImageWidth;
        public int ImageHeight;
        public double RealWidth;
        public double RealHeight;
        public double DetectWidth;
        public double WidthScale;
        public double HeightScale;
        public int PartWidthNum;
        public int PartHeightNum;
        public int PartImgWidth;
        public int PartImgHeight;
    }
}