using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OperateIniFile;

namespace XRDataProcess
{
    public partial class 统计单元长度 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        public 统计单元长度()
        {
            InitializeComponent();
        }

        public bool _IsOK = false;
        public static int _DisUnitLen = 100;
        public static int _IndexUnitLen = 100;

        private void button_Yes_Click(object sender, EventArgs e)
        {
            _IsOK = true;

            _DisUnitLen = int.Parse(comboBox1.Text);
            _IndexUnitLen = int.Parse(comboBox2.Text);
            
            _Setting.DetectYear = comboBox_DetectYear.Text;
            _Setting.DetectNum = comboBox_DetectNum.Text;
            _Setting.DistrictCode = textBox_DistrictCode.Text;
            _Setting.WriteData();

            this.Close();
        }

        private void button_No_Click(object sender, EventArgs e)
        {
            _IsOK = false;
            this.Close();
        }

        private void 统计单元长度_Load(object sender, EventArgs e)
        {
            comboBox_DetectYear.Text = _Setting.DetectYear;
            comboBox_DetectNum.Text = _Setting.DetectNum;
            textBox_DistrictCode.Text = _Setting.DistrictCode;
        }
    }
}
