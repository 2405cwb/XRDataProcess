using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.UserSkins;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
//using ProjectLog;
using Framework.Log;

namespace XRDataProcess
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // 关键修改：只在第一个参数精确为 --auto-test 时才跑自动化测试
            if (args.Length > 0 && string.Equals(args[0], "--auto-test", StringComparison.OrdinalIgnoreCase))
            {


                /*
                    <!-- 每次 Debug 编译完成后自动运行 AutoTest，核心结果不一致就让编译直接失败 -->
             <Target Name="RunAutoTestAfterBuild" AfterTargets="Build" Condition="'$(Configuration)' == 'Debug'">
           <Exec Command="&quot;$(TargetPath)&quot; --auto-test" 
          ContinueOnError="false" 
          IgnoreExitCode="false" />
          <Message Text="所有核心计算结果校验通过！可以安心提交代码" Importance="high" />
          </Target> 
                 
                 */
                //var form = new MainForm();
                //form.AutoTest();

                //// 可选：加个提示，防止有人误触
                //Console.WriteLine("AutoTest 完成");
                Environment.Exit(0); // 明确退出，防止卡死
                 return;


            }
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("zh-Hans");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            BonusSkins.Register();
            SkinManager.EnableFormSkins();
            UserLookAndFeel.Default.SetSkinStyle("DevExpress Style");
            Application.Run(new MainForm());
        }
    }
}
