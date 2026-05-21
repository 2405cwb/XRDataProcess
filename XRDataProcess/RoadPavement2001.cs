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
using Framework.Other.MyGlobal;

namespace RuralPavementDetect
{
    public partial class RoadPavement2001 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        /// <summary>
        /// 路面类型，0-沥青，1-水泥
        /// </summary>
        public int roadtype = 0;
        public String RoadDiseaseType;
        public bool IsDisease = false;

        public RoadPavement2001(Dictionary<string, int>[] DiseaseIdx, RoadDiseaseType[][] roaddis)
        {
            InitializeComponent();

            AddDiseaseControls(roaddis[0], tableLayoutPanel1, false);
            AddDiseaseControls(roaddis[1], tableLayoutPanel2, true);

            if (_Setting.SelectDrawDis == 1 && Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle] == "等级公路2018")
            {
                label1.Text = "方格个数";
                label2.Text = "病害面积(m)";
            }

        }
        private void AddDiseaseControls(RoadDiseaseType[] roaddis, TableLayoutPanel panel, bool fg)
        {
            int i = 0;
            int ccnt = 0;
            int rcnt = 0;

            for (int j = 0; j < roaddis.Length; ++j)
            {
                if (roaddis[j].isshow)
                {
                    RadioButton type = new RadioButton()
                    {
                        Text = string.Format("{0}(&{1})", roaddis[j].disname, roaddis[j].shortcut),
                        Top = 22 + i++ * 20,
                        Left = 20,
                        Width = 200
                    };
                    type.CheckedChanged += new EventHandler(ckb_CheckedChanged);
                    panel.Controls.Add(type, ccnt, rcnt++);
                    if (fg)
                    {
                        if (rcnt >= (roaddis.Length + 2) / 3)
                        {
                            rcnt = 0;
                            ccnt++;
                        }
                    }
                    else
                    {
                        if (rcnt >= (roaddis.Length + 1) / 2)
                        {
                            rcnt = 0;
                            ccnt++;
                        }
                    }
                }
            }

        }

        void ckb_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rbt = (RadioButton)sender;
            RoadDiseaseType = rbt.Text.Substring(0, rbt.Text.Length - 4);
            button_confirm_Click(null, null);
        }

        public void SetNumArea(int num, double aire)
        {
            this.textBox_width.Text = aire.ToString("0.000"); //面积
            this.textBox_DiseaseLength.Text = num.ToString("0.000");//个数
        }
        public void SetLengthWidth(double length, double width)
        {
            this.textBox_width.Text = (length * _RoadConfig.WidthScale).ToString("0.000");
            this.textBox_DiseaseLength.Text = (width * _RoadConfig.HeightScale).ToString("0.000");
        }
        public void JudgeRoadType(int type)
        {
            switch (type)
            {
                //水泥
                case 1:
                    this.splitContainer_RoadTypes.Panel1Collapsed = true;
                    this.splitContainer_RoadTypes.Panel2Collapsed = false;
                    roadtype = 1;
                    break;
                //沥青
                case 0:
                    this.splitContainer_RoadTypes.Panel1Collapsed = false;
                    this.splitContainer_RoadTypes.Panel2Collapsed = true;
                    roadtype = 0;
                    break;
                default: break;
            }
        }
        private void button_confirm_Click(object sender, EventArgs e)
        {
            IsDisease = true;
            this.Close();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            IsDisease = false;
            this.Close();
        }

    }
}
