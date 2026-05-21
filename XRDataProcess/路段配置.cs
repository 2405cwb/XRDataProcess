using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Framework.Other;
using OperateIniFile;

namespace XRDataProcess
{
    public partial class 路段配置 : Form
    {
        public bool _IsSet = false;
        public int _roadlinenum = 8;//一个路段的路线数量
        public 路段配置()
        {
            InitializeComponent();
        }

        private void RoadSet_Load(object sender, EventArgs e)
        {            
            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
            textBox_roadpart.Text = inisetting.ReadString("Road", "RoadPName", "123").Replace("\0", "");
            textBox_roadline.Text = inisetting.ReadString("Road", "RoadLName", "123").Replace("\0", "");
            textBox_roadtype.Text = inisetting.ReadString("Road", "RoadType", "123").Replace("\0", "");
            _IsSet = false;
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                if (ctl.Tag == null)
                {
                    continue;
                }
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (ctl is CheckBox)
                {
                    CheckBox cbox = ctl as CheckBox;
                    cbox.Checked = inisetting.ReadBool("Road", string.Format("RoadLine{0}{1}", idx / 10, idx % 10), false);
                    if (!_IsSet)
                    {
                        _IsSet = cbox.Checked;
                    };            
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
            inisetting.WriteString("Road", "RoadPName", textBox_roadpart.Text);
            inisetting.WriteString("Road", "RoadLName", textBox_roadline.Text);
            inisetting.WriteString("Road", "RoadType", textBox_roadtype.Text);
            _IsSet = false;
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                if (ctl.Tag == null)
                {
                    continue;
                }
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (ctl is CheckBox)
                {
                    CheckBox cbox = ctl as CheckBox;
                    inisetting.WriteBool("Road", string.Format("RoadLine{0}{1}", idx / 10, idx % 10), cbox.Checked);
                    if (!_IsSet)
                    {
                        _IsSet = cbox.Checked;
                    };
                }
            }
            this.Close();
        }

        private void RoadSet_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (textBox_roadpart.Text == string.Empty)
            {
                _IsSet = false;
            }
            if (textBox_roadline.Text == string.Empty)
            {
                _IsSet = false;
            }
        }
    }

    public class MyRoadSide
    {
        public bool _bRight;
        public bool _bLeft;

        public string _sLeft;
        public string _sRight;
    }
}
