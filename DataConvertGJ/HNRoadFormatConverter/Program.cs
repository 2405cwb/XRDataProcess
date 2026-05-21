using DevExpress.LookAndFeel;
using DevExpress.Skins;
using DevExpress.UserSkins;
using System;
using System.IO;
using System.Windows.Forms;

namespace HNRoadFormatConverter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            BonusSkins.Register();

            string selectedSource = null;
            using (var dlg = new SelectSourceForm())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    selectedSource = dlg.SelectedSource;
                }
                else
                {
                    // 用户点 X 关闭窗口 → 退出程序
                    return;
                }
            }
            // DevExpress 设置存储路径重定向到用户目录
            string userSettingsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "夕睿光电",
                "国检转换软件",
                "Settings"
            ); 
            if (!Directory.Exists(userSettingsPath))
                Directory.CreateDirectory(userSettingsPath);
            // 把结果传给主窗体
            Application.Run(new Form1(selectedSource));
        }
    }
}
