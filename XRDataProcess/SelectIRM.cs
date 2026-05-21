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
    public partial class SelectIRM : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        private int _type = 0;
        public SelectIRM(int type)
        {
            InitializeComponent();
            _type = type;
        }
        
        /// <summary>
        /// 平整度IRI、车辙Rut、构造深度SMTD、构造深度MPD、几何线形
        /// </summary>
        public static bool[] irm = { true, true, true, true, false };

        private bool _IsYes = false;
        private void button_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button_confirm_Click(object sender, EventArgs e)
        {
            double val = 0;
            try
            {
                val = double.Parse(textBox_Thresh0.Text);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("参数只能填入数字！");
                return;
            }

            try
            {
                val = double.Parse(textBox_Thresh1.Text);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("参数只能填入数字！");
                return;
            }

            irm[0] = chk_IRI.Checked;
           irm[1] = chk_RUT.Checked;
            irm[2] = chk_MTD.Checked;
            irm[3] = chk_MPD.Checked;
            irm[4] = chk_Geoalig.Checked;
            _IsYes = true;

            _Setting.Las_Filter = chk_LasFilter.Checked;

            _Setting.outMoHaoData = checkBox_mohao.Checked;

            _Setting.Las_Filter_Thresh0 = Convert.ToDouble(textBox_Thresh0.Text);
            _Setting.Las_Filter_Thresh1 = Convert.ToDouble(textBox_Thresh1.Text);
            _Setting.WriteData();

            this.Close();
        }

        private void SelectIRM_Load(object sender, EventArgs e)
        {

            if (_type == 0)
            {
                this.Text = "选择计算指标";
                chk_IRI.Text = "计算平整度IRI";
                chk_RUT.Text = "计算车辙Rut";
                chk_MTD.Text = "计算构造深度SMTD";
                chk_MPD.Text = "计算构造深度MPD";
                chk_Geoalig.Text = "计算几何线形";
            }
            else if(_type == 1)
            {
                this.Text = "选择清除指标";
                chk_IRI.Text = "清除平整度IRI中间结果";
                chk_RUT.Text = "清除车辙Rut中间结果";
                chk_MTD.Text = "清除构造深度SMTD中间结果";
                chk_MPD.Text = "清除构造深度MPD中间结果";
                chk_Geoalig.Text = "清除几何线形中间结果";
            }

            chk_IRI.Checked = irm[0];
            chk_RUT.Checked = irm[1];
            chk_MTD.Checked = irm[2];
            chk_MPD.Checked = irm[3];
            chk_Geoalig.Checked = irm[4];

            chk_LasFilter.Checked = _Setting.Las_Filter;
            checkBox_mohao.Checked = _Setting.outMoHaoData;
            textBox_Thresh0.Text = _Setting.Las_Filter_Thresh0.ToString();
            textBox_Thresh1.Text = _Setting.Las_Filter_Thresh1.ToString();
        }

        public bool IsYes()
        {
            return _IsYes;
        }

        private void chk_IRI_CheckedChanged(object sender, EventArgs e)
        {
            if(!chk_IRI.Checked)
            {
                chk_LasFilter.Checked = false;
            }
        }
    }
}
