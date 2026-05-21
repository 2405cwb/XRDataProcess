using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Windows.Forms;
using XRDataProcess; 

namespace RuralPavementDetect
{
    public partial class RoadPavementPanel : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public String RoadDiseaseType = "";
        public String RoadDiseaseRemarks = "";
        public bool IsDisease = false;
        public static bool IsDiseaseRemark = false;

        public RoadPavementPanel(Dictionary<string, int>[] DiseaseIdx, RoadDiseaseType[][] roaddis, int type)
        {
            InitializeComponent();

            AddDiseaseControls(roaddis[type], tableLayoutPanel1); 

            groupBox1.Text = GlobalExcel._RoadTypeStr[type] + "路面";

            if (_Setting.SelectDrawDis == 1 && Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle] == "等级公路2018")
            {
                label1.Text = "方格个数";
                label2.Text = "病害面积(m^2)";
            }
        }

        private void AddDiseaseControls(RoadDiseaseType[] roaddis, TableLayoutPanel panel)
        {
            
            int ccnt = 0;
            int rcnt = 0;

            for (int j = 0; j < roaddis.Length; ++j)
            {
                if (roaddis[j].isshow)
                {
                    RadioButton type = new RadioButton()
                    {
                        Text = string.Format("{0}(&{1})", roaddis[j].disname, roaddis[j].shortcut),
                    };
                    type.Dock = DockStyle.Fill;
                    type.CheckedChanged += new EventHandler(ckb_CheckedChanged);
                    panel.Controls.Add(type, ccnt, rcnt++);
                    if (rcnt >= (roaddis.Length + 1) / 2)
                    {
                        rcnt = 0;
                        ccnt++;
                    }
                }
            }
        }

        void ckb_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rbt = (RadioButton)sender;
            RoadDiseaseType = rbt.Text.Substring(0, rbt.Text.Length - 4);

            if (IsDiseaseRemark)
            {
                if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.SelectDrawDis == 0)
                {
                    radioButton1.Visible = false;
                    radioButton2.Visible = false;
                    checkBox1.Visible = false;

                    if (RoadDiseaseType.Contains("裂缝"))
                    {
                        checkBox1.Visible = true;
                    }
                    else if (RoadDiseaseType.Contains("修补.条状"))
                    {
                        radioButton1.Visible = true;
                        radioButton2.Visible = true;
                    }
                    else
                    {
                        button_confirm_Click(null, null);
                    }
                }
                else
                {
                    button_confirm_Click(null, null);
                }
            }
            else
            {
                button_confirm_Click(null, null);
            }
        }

        public void SetNumArea(int num, double aire)
        {
            this.textBox_width.Text = aire.ToString("0.000"); //面积
            this.textBox_DiseaseLength.Text = num.ToString("0.000");//个数
        }
        public void SetLengthWidth(double length, double width)
        {
            this.textBox_width.Text = (length*_RoadConfig.WidthScale).ToString("0.000");
            this.textBox_DiseaseLength.Text = (width*_RoadConfig.HeightScale).ToString("0.000");
        }
        public void SetRealLengthWidth(double length, double width)
        {
            this.textBox_width.Text = length.ToString("0.000");
            this.textBox_DiseaseLength.Text = width.ToString("0.000");
        }

        private void button_confirm_Click(object sender, EventArgs e)
        {
            IsDisease = true;

            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2018 && _Setting.SelectDrawDis == 0)
            {
                if (RoadDiseaseType.Contains("裂缝"))
                {
                    if (checkBox1.Checked)
                    {
                        RoadDiseaseRemarks = checkBox1.Text;
                    }
                }
                else if (RoadDiseaseType.Contains("修补.条状"))
                {
                    if (radioButton1.Checked)
                    {
                        RoadDiseaseRemarks = radioButton1.Text;
                    }
                    if (radioButton2.Checked)
                    {
                        RoadDiseaseRemarks = radioButton2.Text;
                    }
                }
            }

            this.Close();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            IsDisease = false;
            this.Close();
        }

    }
}
