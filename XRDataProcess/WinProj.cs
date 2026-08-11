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
            comboBox_roadgrade.SelectedItem = _ProjectInfo._RoadGrade;
            comboBox_pavement.SelectedIndex = _ProjectInfo._HasInitialRoadType ? _ProjectInfo._RoadType : -1;

            textBox_detectorname.Text = _ProjectInfo._DataPerson;
            comboBox_weather.Text = _ProjectInfo._DataWeather;

            maskedTextBox_end.Text = _ProjectInfo._EndMile.ToString("K0000+000");
            textBox_mile.Text = _ProjectInfo._EndDmi.ToString() + "米";
        }

        private void button_No_Click(object sender, EventArgs e)
        {
            LoadOriPrjInfo();
        }

        private bool SetPrjInfo()
        {
            int startMile;
            int endMile;
            int endDmi;
            if (!TryParseMile(maskedTextBox_projectstartpoint.Text, out startMile)
                || !TryParseMile(maskedTextBox_end.Text, out endMile))
            {
                MessageBox.Show("起点桩号和终点桩号必须为 K0+000 格式。", "工程信息");
                return false;
            }
            if (!int.TryParse(textBox_mile.Text.Replace("米", "").Trim(), out endDmi) || endDmi < 0)
            {
                MessageBox.Show("总里程必须为不小于 0 的整数（米）。", "工程信息");
                return false;
            }
            int direction = comboBox_DriveDirection.Text == "下行" ? -1 : 1;
            if ((direction > 0 && endMile < startMile) || (direction < 0 && endMile > startMile))
            {
                MessageBox.Show("起终点桩号与行车方向不一致，请检查。", "工程信息");
                return false;
            }
            if (comboBox_roadgrade.SelectedIndex < 0 || comboBox_pavement.SelectedIndex < 0)
            {
                MessageBox.Show("请从下拉框中选择公路等级和起始路面材质。未指定起始材质时不能绘制病害。", "工程信息");
                return false;
            }

            _ProjectInfo._Province = comboBox_sheng.Text;
            _ProjectInfo._City = comboBox_shi.Text;
            _ProjectInfo._District = comboBox_xian.Text;
            _ProjectInfo._RoadNum = textBox1.Text;
            _ProjectInfo._RoadCode = textBox_roadnumber.Text;
            _ProjectInfo._RoadName = textBox_roadname.Text;

            int temp = startMile;
            if(temp != _ProjectInfo._StartMile)
            {
                _IsChanged = true;
                _ProjectInfo._StartMile = temp;
            }

            if (direction == -1)
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

            if (_ProjectInfo._RoadGrade != comboBox_roadgrade.Text)
            {
                _IsChanged = true;
                _ProjectInfo._RoadGrade = comboBox_roadgrade.Text;
            }

            if (_ProjectInfo._RoadType != Convert.ToInt16(comboBox_pavement.SelectedIndex))
            {
                _IsChanged = true;
                _ProjectInfo._RoadType = Convert.ToInt16(comboBox_pavement.SelectedIndex);
            }

            _ProjectInfo._DataPerson = textBox_detectorname.Text;
            _ProjectInfo._DataWeather = comboBox_weather.Text;

            temp = endMile;
            if (_ProjectInfo._EndMile != temp)
            {
                _IsChanged = true;
                _ProjectInfo._EndMile = temp;
            }

            temp = endDmi;
            if (_ProjectInfo._EndDmi != temp)
            {
                _IsChanged = true;
                _ProjectInfo._EndDmi = temp;
            }
            return true;
        }

        private void button_Yes_Click(object sender, EventArgs e)
        {
            _IsChanged = false;
            if (!SetPrjInfo()) return;
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
            comboBox_roadgrade.Items.Clear();
            if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
            {
                comboBox_roadgrade.Items.AddRange(new object[] { "快速路", "主干路", "次干路", "支路" });
            }
            else
            {
                comboBox_roadgrade.Items.AddRange(new object[] { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" });
            }
            comboBox_roadgrade.SelectedItem = _ProjectInfo._RoadGrade;
            if (comboBox_roadgrade.SelectedIndex < 0 && comboBox_roadgrade.Items.Count > 0)
                comboBox_roadgrade.SelectedIndex = 0;

            if (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel||_Setting.ParmStyle == StandardParmType.RuralRoadChongqing)
            {
                comboBox_pavement.Items.Clear();
                comboBox_pavement.Items.Add("沥青");
                comboBox_pavement.Items.Add("水泥");
                comboBox_pavement.Items.Add("砂石");
                comboBox_pavement.SelectedIndex = _ProjectInfo._HasInitialRoadType ? _ProjectInfo._RoadType : -1;
            }
            else
            {
                comboBox_pavement.Items.Clear();
                comboBox_pavement.Items.Add("沥青");
                comboBox_pavement.Items.Add("水泥");
                comboBox_pavement.SelectedIndex = _ProjectInfo._HasInitialRoadType ? _ProjectInfo._RoadType : -1;
            }
            
        }

        private void textBox_roadnumber_TextChanged(object sender, EventArgs e)
        {

        }

        private static bool TryParseMile(string value, out int mile)
        {
            mile = 0;
            if (string.IsNullOrWhiteSpace(value)) return false;
            string text = value.Trim().ToUpperInvariant();
            if (!text.StartsWith("K")) return false;
            string[] parts = text.Substring(1).Split('+');
            int kilometer;
            int meter;
            return parts.Length == 2 && int.TryParse(parts[0], out kilometer) && kilometer >= 0
                && int.TryParse(parts[1], out meter) && meter >= 0 && meter < 1000
                && (mile = kilometer * 1000 + meter) >= 0;
        }
    }
}
