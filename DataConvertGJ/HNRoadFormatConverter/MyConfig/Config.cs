using HNRoadFormatConverter.Commons;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.MyConfig
{
    public class Config
    {
        /// <summary>
        /// 皮肤
        /// </summary>
        public string DefaultSkin { get; set; }
        
        /// <summary>
        /// 用户文件夹
        /// </summary>
        public string UserPath { get; set; }

        public string NowModel { get; set; }
    }
    public class ConfigManager
    {
    //   static private LogHelper _log = new LogHelper(type: typeof(Config));
        public static Config Config { get; set; }
        private static string GetConfigPath()
        {
            // 推荐使用 LocalAppData（不随用户漫游，适合本地配置）
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string company = "夕睿光电";
            string appName = "国检转换软件";

            string configDir = Path.Combine(appData, company, appName, "config");

            if (!Directory.Exists(configDir))
                Directory.CreateDirectory(configDir);

            return Path.Combine(configDir, "config.json");
        }
        private static readonly string jsonConfigPath = GetConfigPath();
        static ConfigManager()
        {
            try
            {
                if (File.Exists(jsonConfigPath))
                {
                    string jsonTxt = File.ReadAllText(jsonConfigPath);
                    Config = JsonHelper.FromJSON<Config>(jsonTxt);
                }
                else
                {
                    // 如果文件不存在，创建默认配置
                    Config = new Config
                    {
                        DefaultSkin = "default",
                        UserPath = "",
                        NowModel = ""
                        // 设置其他默认值
                    };
                    SaveConfig(); // 首次创建
                }
            }
            catch (Exception ex)
            {
                // log...
                Config = new Config(); // 兜底
            }

        }
        public static Config GetConfig()
        {
            return Config;
        }

        public static void SaveConfig()
        {
            string jsonStr= JsonHelper.ToJSON(Config);

            FileHerper.WriteFileMemory(FileHerper.getMemoryStream(jsonStr), jsonConfigPath, typeof(ConfigManager).ToString());
        }
        private ConfigManager()
        {

        }
    }

}
