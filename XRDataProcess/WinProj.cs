using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OperateIniFile;
using System.Xml;
using System.Runtime.InteropServices;
using System.IO;
using Framework.Other.MyGlobal;
using XRDataProcess;

namespace RoadStreet
{
    public partial class WinProj : Form
    {
        public event EventHandler EventUpdateProjectInfo;
        private ProjectInfo _ProjectInfo;
        private bool _IsChanged = false;
        public WinProj(ProjectInfo pinfo)
        {
            InitializeComponent();
            _ProjectInfo = pinfo;
            LoadOriPrjInfo();
        }

        private void LoadOriPrjInfo()
        {
            comboBox_sheng.Text = _ProjectInfo._Province;
            comboBox_shi.Text = _ProjectInfo._City;
            comboBox_xian.Text = _ProjectInfo._District;
            textBox1.Text = _ProjectInfo._RoadNum;
            textBox_roadnumber.Text = _ProjectInfo._RoadCode;
            textBox_roadname.Text = _ProjectInfo._RoadName;

            maskedTextBox_projectstartpoint.Text = _ProjectInfo._StartMile.ToString("K0000+000");
            if (_ProjectInfo._Direction == -1)
            {
                comboBox_DriveDirection.Text = "下行";
            }
            else if (_ProjectInfo._Direction == 1)
            {
                comboBox_DriveDirection.Text = "上行";
            }
            dateTimePicker_date.Value = DateTime.ParseExact(_ProjectInfo._DataDate, "yyyyMMdd", null);
            dateTimePicker_time.Value = DateTime.ParseExact(_ProjectInfo._DataTime, "HHmmss", null);
            textBox_roadgrade.Text = _ProjectInfo._RoadGrade;
            comboBox_pavement.SelectedIndex = _ProjectInfo._RoadType;

            textBox_detectorname.Text = _ProjectInfo._DataPerson;
            comboBox_weather.Text = _ProjectInfo._DataWeather;

            maskedTextBox_end.Text = _ProjectInfo._EndMile.ToString("K0000+000");
            textBox_mile.Text = _ProjectInfo._EndDmi.ToString() + "米";
        }

        private void button_No_Click(object sender, EventArgs e)
        {
            LoadOriPrjInfo();
        }

        private void SetPrjInfo()
        {
            _ProjectInfo._Province = comboBox_sheng.Text;
            _ProjectInfo._City = comboBox_shi.Text;
            _ProjectInfo._District = comboBox_xian.Text;
            _ProjectInfo._RoadNum = textBox1.Text;
            _ProjectInfo._RoadCode = textBox_roadnumber.Text;
            _ProjectInfo._RoadName = textBox_roadname.Text;

            int temp = Convert.ToInt32(maskedTextBox_projectstartpoint.Text.Replace("K", "").Replace("+", ""));
            if(temp != _ProjectInfo._StartMile)
            {
                _IsChanged = true;
                _ProjectInfo._StartMile = temp;
            }

            if (comboBox_DriveDirection.Text == "下行")
            {
                if (_ProjectInfo._Direction != -1)
                {
                    _IsChanged = true;
                    _ProjectInfo._Direction = -1;
                }
                
            }
            else if (comboBox_DriveDirection.Text == "上行")
            {
                if (_ProjectInfo._Direction != 1)
                {
                    _IsChanged = true;
                    _ProjectInfo._Direction = 1;
                }
            }
            _ProjectInfo._DataDate = dateTimePicker_date.Value.ToString("yyyyMMdd");
            _ProjectInfo._DataTime = dateTimePicker_time.Value.ToString("HHmmss");

            if (_ProjectInfo._RoadGrade != textBox_roadgrade.Text)
            {
                _IsChanged = true;
                _ProjectInfo._RoadGrade = textBox_roadgrade.Text;
            }

            if (_ProjectInfo._RoadType != Convert.ToInt16(comboBox_pavement.SelectedIndex))
            {
                _IsChanged = true;
                _ProjectInfo._RoadType = Convert.ToInt16(comboBox_pavement.SelectedIndex);
            }

            _ProjectInfo._DataPerson = textBox_detectorname.Text;
            _ProjectInfo._DataWeather = comboBox_weather.Text;

            temp = Convert.ToInt32(maskedTextBox_end.Text.Replace("K", "").Replace("+", ""));
            if (_ProjectInfo._EndMile != temp)
            {
                _IsChanged = true;
                _ProjectInfo._EndMile = temp;
            }

            temp = Convert.ToInt32(textBox_mile.Text.Replace("米", ""));
            if (_ProjectInfo._EndDmi != temp)
            {
                _IsChanged = true;
                _ProjectInfo._EndDmi = temp;
            }
        }

        private void button_Yes_Click(object sender, EventArgs e)
        {
            SetPrjInfo();
            _ProjectInfo.SavePrjInfo();
            if (_IsChanged)
            {
                MessageBox.Show("修改工程信息成功，即将重置该工程！");
                EventUpdateProjectInfo(null, null);
            }
        }
        XRSetting _Setting = XRSetting.GetInstance();
        private void WinProj_Load(object sender, EventArgs e)
        {
            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel||_Setting.ParmStyle == StandardParmType.RuralRoadChongqing)
            {
                comboBox_pavement.Items.Clear();
                comboBox_pavement.Items.Add("沥青");
                comboBox_pavement.Items.Add("水泥");
                comboBox_pavement.Items.Add("砂石");
                comboBox_pavement.SelectedIndex = _ProjectInfo._RoadType;
            }
            else
            {
                comboBox_pavement.Items.Clear();
                comboBox_pavement.Items.Add("沥青");
                comboBox_pavement.Items.Add("水泥");
                comboBox_pavement.SelectedIndex = _ProjectInfo._RoadType;
            }
            
        }

        private void textBox_roadnumber_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
