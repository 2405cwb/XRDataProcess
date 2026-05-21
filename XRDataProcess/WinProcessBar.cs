using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace XRDataProcess
{
    public partial class WinProcessBar : Form
    {
        private delegate void ProMainTextHandle(int val);
        private delegate void ProMainLogHandle(string log);
        private delegate void ProMainValHandle(double val);

        private ProMainTextHandle MyMainTexthandle = null;
        private ProMainLogHandle MyMainLoghandle = null;
        private ProMainValHandle myMainValhandle = null;

        private ProMainValHandle myIRIValhandle = null;
        private ProMainValHandle myMTDValhandle = null;
        private ProMainValHandle myRutValhandle = null;
        private ProMainValHandle myMPDValhandle = null;
        private ProMainValHandle myGeoAligValhandle = null;

        private List<SingleProject> _Projects;
        public WinProcessBar(List<SingleProject> projects)
        {
            InitializeComponent();
            _Projects = projects;
            bar_main.Maximum = _Projects.Count;
            bar_main.Value = 0;
            _mainstep = 0;

            MyMainTexthandle = new ProMainTextHandle(AddMainVal);
            MyMainLoghandle = new ProMainLogHandle(TextInfoAdd);
            myMainValhandle = new ProMainValHandle(SetMainBar);

            myIRIValhandle = new ProMainValHandle(SetIRIBar);
            myMTDValhandle = new ProMainValHandle(SetMTDBar);
            myRutValhandle = new ProMainValHandle(SetRutBar);
            myMPDValhandle = new ProMainValHandle(SetMPDBar);
            myGeoAligValhandle = new ProMainValHandle(SetGeoAligBar);
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

        public static int _mainstep = 0;
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
                textcnt.Text = string.Format("({0}/{1})", _mainstep, bar_main.Tag);
            }
        }

        private void SetIRIBar(double val)
        {
            if (bar_iri.InvokeRequired)
            {
                bar_iri.Invoke(myIRIValhandle, val);
            }
            else
            {
                bar_iri.Value = (int)(bar_iri.Maximum * val);
            }
        }
        private void SetMTDBar(double val)
        {
            if (bar_mtd.InvokeRequired)
            {
                bar_mtd.Invoke(myMTDValhandle, val);
            }
            else
            {
                bar_mtd.Value = (int)(bar_mtd.Maximum * val);
            }
        }
        private void SetMPDBar(double val)
        {
            if (bar_mpd.InvokeRequired)
            {
                bar_mpd.Invoke(myMPDValhandle, val);
            }
            else
            {
                bar_mpd.Value = (int)(bar_mpd.Maximum * val);
            }
        }
        private void SetRutBar(double val)
        {
            if (bar_rut.InvokeRequired)
            {
                bar_rut.Invoke(myRutValhandle, val);
            }
            else
            {
                bar_rut.Value = (int)(bar_rut.Maximum * val);
            } 
        }
        private void SetGeoAligBar(double val)
        {
            if (bar_geoalig.InvokeRequired)
            {
                bar_geoalig.Invoke(myGeoAligValhandle, val);
            }
            else
            {
                bar_geoalig.Value = (int)(bar_geoalig.Maximum * val);
            }
        }
        private void SetMainBar(double val)
        { 
            if (bar_main.InvokeRequired)
            {
                bar_main.Invoke(myMainValhandle, val);
            }
            else
            {
                bar_main.Value += (int)((bar_iri.Value * 0.2 / bar_iri.Maximum
                    + bar_mtd.Value * 0.2 / bar_mtd.Maximum
                    + bar_rut.Value * 0.2 / bar_rut.Maximum
                    + bar_mpd.Value * 0.2 / bar_mpd.Maximum
                    + bar_geoalig.Value * 0.2 / bar_geoalig.Maximum)
                    * bar_main.Maximum / (double)bar_main.Tag);
            }
        }

        public void SetIRIVal(double percent)
        {
            SetIRIBar(percent);
        }
        
        public void SetMTDVal(double percent)
        {
            SetMTDBar(percent);
        }

        public void SetRutVal(double percent)
        {
            SetRutBar(percent);
        }

        public void SetMPDVal(double percent)
        {
            SetMPDBar(percent);
        }

        public void SetGeoAlig(double percent)
        {
            SetGeoAligBar(percent);
        }
        
    }
}
