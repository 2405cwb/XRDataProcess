using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace XRDataProcess
{
    public partial class WinGDProcessBar : Form
    {
        private delegate void ProMainTextHandle(int val);
        private delegate void ProMainLogHandle(string log);
        private ProMainTextHandle MyMainTexthandle = null;
        private ProMainLogHandle MyMainLoghandle = null;
        public delegate void callBack(float index, string mes);
        [DllImport(@"hnCalcuMethod", EntryPoint = "setCallBack", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AutoExtractFeature(callBack cb = null);
        public callBack iriCall = null;//回调标志
        public WinGDProcessBar()
        {
            InitializeComponent();
        }
        public static int _mainstep = 0;
        private List<SingleProject> _projects;
        public WinGDProcessBar(List<SingleProject> pros)
        {
            InitializeComponent();
            _projects = pros;
            bar_main.Maximum = _projects.Count;
            bar_main.Value = 0;
           _mainstep = 0;
           MyMainTexthandle = new ProMainTextHandle(AddMainVal); 
            MyMainLoghandle = new ProMainLogHandle(TextInfoAdd);
        }
        public void TextInfoAdd(string info)
        {
            if (tableLayoutPanel1.InvokeRequired)
            {
                tableLayoutPanel1.Invoke(MyMainLoghandle, info);
            }
            else
            {
                loginfo.Text += string.Format("{0:yyyy-MM-dd hh:mm:ss}\t{1}\r\n", DateTime.Now, info);
            }
        }
        public void SetMainMax(int val)
        {
            bar_main.Tag = val;
            AddMainVal(0);
        }
        public void AddMainVal(int step)
        {
            if (tableLayoutPanel1.InvokeRequired)
            {
                tableLayoutPanel1.Invoke(MyMainTexthandle, step);
            }
            else
            {
                _mainstep += step;
                bar_main.Value = bar_main.Maximum * _mainstep / (int)bar_main.Tag;
                //textcnt.Text = string.Format("({0}/{1})", _mainstep, bar_main.Tag);
            }
        }
      
        private string tempStr = "";
        private System.IO.DirectoryInfo di;
        private System.IO.FileInfo[] files;
        public void AddMainVal1(float step, string mes)
        {
            //progressBar1.Value += (int)step;
            //label1.Text = mes;
            if (bar_iri.InvokeRequired)
            {
                bar_iri.Invoke(iriCall, step, mes);
            }
            else
            {
                bar_iri.Value = (int)(bar_iri.Maximum * step);
            }
            if (loginfo.InvokeRequired)
            {
                loginfo.Invoke(iriCall, step, mes);
            }
            else
            {
                if (tempStr != mes)
                {
                    loginfo.Text += string.Format("{0:yyyy-MM-dd hh:mm:ss}\t{1}\r\n", DateTime.Now, mes);
                    tempStr = mes;
                }

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            iriCall = new callBack(AddMainVal1);
            bar_iri.Value = 0;
            AutoExtractFeature(iriCall);
        }
    }
}
    
             