using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Commons
{
    class Global
    {
        public class _UpdateInfo
        {
            public int disidx { get; set; }
            public string disname { get; set; }
            public string disremark { get; set; }
        }

        public static Mutex g_mutex = new Mutex();
        //  public static ProjectTraceListener g_log = null;
        public static void CmdExe(string cmd)
        {
            Process p = new Process();
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            p.StartInfo.RedirectStandardOutput = true;//由调用程序获取输出信息
            p.StartInfo.RedirectStandardError = true;//重定向标准错误输出
            p.StartInfo.CreateNoWindow = true;//不显示程序窗口
            p.Start();//启动程序

            //向cmd窗口发送输入信息
            p.StandardInput.WriteLine(cmd + "&exit");
            p.StandardInput.AutoFlush = true;
        }

        public static string[] g_ParmStyles = { "等级公路2007", "城镇道路", "北京农村公路", "等级公路2018", "等级公路2001", "上海城市道路", "辽宁农村公路", "广西农村路", "重庆农村路", "湖南农村路", "低等级农村公路" };
    }
}
