
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OperateIniFile;
using DevExpress.XtraBars;
using System.IO;
using Spire.Xls;
using XRDataProcess.Properties;
using DevExpress.XtraCharts;
using NPOI.SS.Formula.Functions;

namespace XRDataProcess
{
    public partial class 软件设置 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public 软件设置()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_confirm_Click(object sender, EventArgs e)
        {


            _Setting.is5211MergeArea500 = is5211MergeArea500_ck.Checked;
           

            if (radioButton37.Checked)
            {
                _Setting.IRIAlgorithmInterval = 0.1;
            }
            else
            {
                _Setting.IRIAlgorithmInterval = 0.25;
            }

             _Setting.hefei2MinSplit = int.Parse(textBox1.Text);
            _Setting.splitExcelDh = ck_th.Checked; 

            _Setting.SplitPartDistance = int.Parse( txt_SplitDistance.Text);

            if (checkBox6.Checked)
            {
                _Setting.outSmallDisCalculateArea = true;
            }
            else
            {
                _Setting.outSmallDisCalculateArea=false;
            }
            if (checkBox8.Checked)
            {
                _Setting.zcSplit = true;
            }
            else
            {
                _Setting.zcSplit = false;
            }
            if (checkBox7.Checked )
            {
                _Setting.JSAverageType = true;
            }
            else
            {
                _Setting.JSAverageType = false;
            }
            
            if (radioButton33.Checked)
            {
                _Setting.czJudgeType = 0;
            }
            else if (radioButton39.Checked)
            {
                _Setting.czJudgeType = 1;
            }
            else if (radioButton40.Checked)
            {
                 _Setting.czJudgeType = 2;
            }
            if (checkBox5.Checked)
            {
                _Setting.outDaqAccelerate = true;
            }
            else
            {
                _Setting.outDaqAccelerate = false;
            }
            if (checkBox4.Checked)
            {
                _Setting.gjLbiOutHight = true;
            }
            else
            {
                _Setting.gjLbiOutHight = false;
            }
           
            _Setting.splitExcelDh = ck_th.Checked;
            if (radioButton16.Checked)
            {
                if (int.Parse(txb_cmoprow.Text) < 31)
                {
                    MessageBox.Show("设置失败：人工调查模式时，CPMS调查表每页行数设置必须大于30");
                    return;
                }
            }

            int old_drawdis = _Setting.SelectDrawDis;

            if (radioButton_dmi.Checked)
            {
                _Setting.PartType = 1;
            }
            if (radioButton_mile.Checked)
            {
                _Setting.PartType = 0;
            }
            if (rb_roadCrossing.Checked)
            {
                _Setting.roadCrossingShow = false;
            }
            else
            {
                _Setting.roadCrossingShow = true;
            }
            _Setting.PartType_Dmi_Len = Convert.ToInt32(textBox_DmiLen.Text);

            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.ExcelType = Convert.ToInt32(tctl.Tag);
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tct2 = ctl as RadioButton;
                    if (tct2.Checked)
                    {
                        _Setting.OutRut = Convert.ToInt16(tct2.Tag);
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.Qufen_dis_degree = Convert.ToInt16(tctl.Tag);
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel8.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.RQIJudgeType = Convert.ToInt16(tctl.Tag);
                        break;
                    }
                }
            }

            _Setting.Out_roadimg = checkBox_out_imgname.Checked ? 1 : 0;
            _Setting.Out_roadinfo = checkBox_out_info.Checked ? 1 : 0;
            _Setting.sheetRoundingOffType = radioButton12.Checked ? 0 : 1;
            _Setting.sheetRoundingOffNum = Convert.ToInt16(numericUpDown1.Value);
            _Setting.IsOutputDisAreaSubtotal = checkBox3.Checked;

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.Is_Multfolder = Convert.ToInt16(tctl.Tag);
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel6.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.SelectDrawDis = Convert.ToInt16(tctl.Tag);
                        break;
                    }

                }
            }

            foreach (Control ctl in flowLayoutPanel7.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.Is_SnCarve = Convert.ToInt16(tctl.Tag);
                        break;
                    }

                }
            }

            foreach (Control ctl in groupBox3.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.BrokenPlatetype = Convert.ToInt16(tctl.Tag);
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel11.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    if (tctl.Checked)
                    {
                        _Setting.IRIExcelSide = Convert.ToInt16(tctl.Tag);
                        break;
                    }
                }
            }

            //水泥板块面积参数
            _Setting.BrokenPlatetype = radioButton18.Checked ? 0 : 1;

            _Setting.PlateWidth = Convert.ToDouble(txb_bkwidth.Text);
            _Setting.PlateLength = Convert.ToDouble(txb_bklength.Text);

            bool oldforbid = _Setting.IsForbidOverLapping;
            bool oldremark = _Setting.IsCrackRemark;
            _Setting.IsForbidOverLapping = checkBox1.Checked;
            _Setting.IsCrackRemark = checkBox2.Checked;

            double oldRealWidth = _RoadConfig.RealWidth;
            double oldRealHeight = _RoadConfig.RealHeight;

            string imgwidth = txb_ImageWidth.Text;
            string imgheight = txb_ImageHeight.Text;
            string realwidth = txb_RealWidth.Text;
            string realheight = txb_RealHeight.Text;
            _RoadConfig.ImageWidth = int.Parse(imgwidth);
            _RoadConfig.ImageHeight = int.Parse(imgheight);
            _RoadConfig.RealWidth = double.Parse(realwidth);
            _RoadConfig.RealHeight = double.Parse(realheight);
            _Setting.cmop_rows = int.Parse(txb_cmoprow.Text);
            if (_Setting.ParmStyle != StandardParmType.DegreeRoad2018 && _Setting.SelectDrawDis != 1)
            {
                if (_Setting.cmop_rows < 31)
                {
                    _Setting.cmop_rows = 33;
                }
            }
            _Setting.WriteData();
            _RoadConfig.WriteData();
            if (old_drawdis != _Setting.SelectDrawDis
                || oldforbid != _Setting.IsForbidOverLapping
                || oldremark != _Setting.IsCrackRemark
                || oldRealWidth != double.Parse(realwidth)
                || oldRealHeight != double.Parse(realheight))
            {
                MessageBox.Show("病害框图发生变化，即将重启软件生效！");

                Application.Exit();
                System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);
                this.Close();
            }
            this.Focus();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radioButton19_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton19.Checked)
            {
                label6.Visible = true;
                txb_bkwidth.Visible = true;
                label7.Visible = true;
                txb_bklength.Visible = true;
            }
            else
            {
                label6.Visible = false;
                txb_bkwidth.Visible = false;
                label7.Visible = false;
                txb_bklength.Visible = false;
            }
        }

        private void radioButton_dmi_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_dmi.Checked)
            {
                radioButton_PG.Enabled = true;
                radioButton_ZY.Enabled = false;
                if (radioButton_ZY.Checked)
                    radioButton1.Checked = true;

                label_DmiLen.Enabled = true;
                textBox_DmiLen.Enabled = true;
            }
            else
            {
                radioButton_ZY.Enabled = true;
                radioButton_PG.Enabled = false;
                if (radioButton_PG.Checked)
                    radioButton1.Checked = true;

                label_DmiLen.Enabled = false;
                textBox_DmiLen.Enabled = false;
            }
        }

        private void LoadRoad2007()
        {
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                if (ctl is RadioButton)
                {
                    if (Convert.ToInt32(ctl.Tag) == _Setting.ExcelType)
                    {
                        RadioButton tctl = (RadioButton)ctl;
                        tctl.Checked = true;
                    }
                }
            }
            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.OutRut)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }

            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Qufen_dis_degree)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {

                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }
        }

        private void LoadCity()
        {
            
            radioButton_ZY.Visible = true;
            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    if (ttagval != 0 && ttagval != 4 && ttagval != 6 && ttagval != 7 && ttagval!=13&&ttagval!=12)
                    {
                        RadioButton tctl = (RadioButton)ctl;
                        tctl.Visible = false;
                    }
                    else
                    {
                        if (ttagval == 0 && _Setting.ExcelType == 0)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                        else if (ttagval == 4 && _Setting.ExcelType == 4)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                        else if (ttagval == 6 && _Setting.ExcelType == 6)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                        else if (ttagval == 7 && _Setting.ExcelType == 7)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                        
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.OutRut)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                    if (ttagval == 1)
                    {
                        tctl.Text = "导出车辙病害";
                    }
                    if (ttagval == 2)
                    {
                        tctl.Visible = false;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Qufen_dis_degree)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }
        }

        private void LoadRoad2018()
        {
            if (_Setting.SelectDrawDis == 1)
            {
                radioButton35.Visible = false;
            }
            else
            {
                radioButton35.Visible = true;
            }
            int ttagval = 0;
            if (_Setting.SelectDrawDis == 0)
            {
                radioButton30.Hide(); 
            }
            else
            {
                radioButton30.Show(); 
            }
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {

                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    if (ttagval == _Setting.ExcelType)
                    {
                        RadioButton tctl = (RadioButton)ctl;
                        tctl.Checked = true;
                        tctl.Visible = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.OutRut)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Qufen_dis_degree)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }
        }

        private void LoadRoad2001()
        {
            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.ExcelType)
                    {

                        tctl.Checked = true;
                        tctl.Visible = true;
                    }
                    else
                    {
                        tctl.Visible = false;
                    }
                }
            }
            groupBox5.Visible = false;
            groupBox6.Visible = false;
            groupBox7.Visible = false;

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }
        }

        private void LoadCityShangHai()
        {
            radioButton_ZY.Visible = false;
            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    if (ttagval != 0 && ttagval != 7)
                    {
                        RadioButton tctl = (RadioButton)ctl;
                        tctl.Visible = false;
                    }
                    else
                    {
                        if (ttagval == 0 && _Setting.ExcelType == 0)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                        else if (ttagval == 7 && _Setting.ExcelType == 7)
                        {
                            RadioButton tctl = (RadioButton)ctl;
                            tctl.Checked = true;
                        }
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.OutRut)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                    if (ttagval == 1)
                    {
                        tctl.Text = "导出车辙病害";
                    }
                    if (ttagval == 2)
                    {
                        tctl.Visible = false;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Qufen_dis_degree)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;

                    }
                }
            }
        }

        private void LoadRoadBeiJin()
        {
            groupBox4.Visible = false;

        }

        private void LoadRoadLiaoNing()
        {
            groupBox4.Visible = false;
        }

        private void LoadRoadGuangXi()
        {
            groupBox4.Visible = true;

            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.ExcelType)
                    {
                        tctl.Checked = true;
                        tctl.Visible = true;
                    }
                    if (ttagval < 2)
                    {
                        tctl.Visible = true;
                    }
                    else
                    {
                        tctl.Visible = false;
                    }
                }
            }
        }

        private void LoadRoadChongQing()
        {
            groupBox4.Visible = true;

            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.ExcelType)
                    {
                        tctl.Checked = true;
                        tctl.Visible = true;
                    }
                    if (ttagval < 1)
                    {
                        tctl.Visible = true;
                    }
                    else
                    {
                        tctl.Visible = false;
                    }
                }
            }
        }
        private void LoadLowVillageRoad()
        {
            is5211MergeArea500_ck.Visible = true;
            if (_Setting.is5211MergeArea500)
            {
                is5211MergeArea500_ck.Checked = true;
            }
            else
            {
                is5211MergeArea500_ck.Checked = false;
            }
            if (_Setting.SelectDrawDis == 1)
            {
                radioButton35.Visible = false;
            }
            else
            {
                radioButton35.Visible = true;
            }
         
            if (_Setting.SelectDrawDis == 0)
            {
                radioButton30.Hide();
            }
            else
            {
                radioButton30.Show();
            }
            int ttagval = 0;
            groupBox4.Visible = true;

            this.radioButton28.Visible = true;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.ExcelType)
                    {
                        tctl.Checked = true;
                        tctl.Visible = true;
                    }
                    if (ttagval <= 1)
                    {

                        tctl.Visible = true;
                    }
                    else if (ttagval == 2)
                    {
                        tctl.Text = "自动化报表输出";
                        tctl.Visible = true;
                    }
                    else if (ttagval == 3)
                    {
                        tctl.Text = "多车道统计定制输出";
                        tctl.Visible = true;
                    }
                    else if (ttagval == 4)
                    {

                        tctl.Text = "报表定制模板"; //长沙理工检测咨询有限公司    湖南定制
                        tctl.Visible = true;
                    }
                    else if (ttagval == 5)
                    { 
                        tctl.Visible = false;
                    }
                    else if (ttagval == 6)
                    {
                        if (_Setting.SelectDrawDis ==1)
                        {
                            tctl.Text = "报表定制(合并)"; //
                            tctl.Visible = false;
                        }
                        else
                        {
                            tctl.Text = "报表定制(合并)"; //
                            tctl.Visible = true;
                        }

                          
                    }
                    else if (ttagval == 7)
                    { 
                        tctl.Text = "资产表定制"; //合肥&资产表定制
                        tctl.Visible = true;
                    }
                    else if (ttagval == 8)
                    {

                        tctl.Text = "病害调绘表"; //
                        tctl.Visible = true;
                    }
                    else if (ttagval == 9)
                    {

                        tctl.Text = "报表定制"; //重庆报表定制
                        tctl.Visible = false;
                    }
                    else if (ttagval == 10)
                    {

                        tctl.Text = "导入模板"; //甘肃导入模板
                        tctl.Visible = true;
                    }
                    else if (ttagval == 11)
                    {

                        tctl.Text = "原始数据报送"; // 2024甘肃原始数据报送
                        tctl.Visible = true;
                    } 
                    else if(ttagval == 12)
                    {
                        tctl.Text = "贵州定制报表";
                        tctl.Visible = false;
                    }
                    else if(ttagval == 13)
                    {
                        tctl.Text = "江西车检2024上传";
                        tctl.Visible = true;
                    }
                    else if (ttagval == 14)
                    {
                        tctl.Text = "csv表格输出";
                        tctl.Visible = true;
                    }
                    else if (ttagval == 15)
                    {
                        tctl.Text = "村名组，景观图片等提交模板";
                        tctl.Visible = true;
                    }
                    else
                    {
                        tctl.Visible = false;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel6.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton rd = ctl as RadioButton;
                    ttagval = Convert.ToInt16(ctl.Tag);
                    //if (ttagval == 1)
                    //{
                    //    rd.Visible = false;
                    //}
                }
            }

            foreach (Control ctl in flowLayoutPanel2.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.OutRut)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel3.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Qufen_dis_degree)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel5.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    if (ttagval == _Setting.Is_Multfolder)
                    {
                        RadioButton tctl2 = (RadioButton)ctl;
                        tctl2.Checked = true;
                    }
                }
            }
        }

        private void LoadRoadHuNan()
        {
            groupBox4.Visible = true;

            int ttagval = 0;
            foreach (Control ctl in flowLayoutPanel1.Controls)
            {
                ttagval = Convert.ToInt16(ctl.Tag);
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;

                    if (ttagval == _Setting.ExcelType)
                    {
                        tctl.Checked = true;
                        tctl.Visible = true;

                    }
                    if (ttagval < 1)
                    {
                        tctl.Visible = true;
                    }
                    else
                    {
                        tctl.Visible = false;
                    }
                }
            }
        }
        private void configFormControlState()
        {



            if (_Setting.SelectDrawDis == 0)
            {
                checkBox6.Visible = false;
            }
            else 
            {
                checkBox6.Visible = true;
            }

              this.txt_SplitDistance.Text = _Setting.SplitPartDistance.ToString();

            if (_Setting.outSmallDisCalculateArea)
            {
                checkBox6.Checked = true;

            }
            else
            {
                checkBox6.Checked = false;
            }

            if (_Setting.czJudgeType == 0)
            {
                radioButton33.Checked = true;
            }
            else if (_Setting.czJudgeType == 1)
            {
                radioButton39.Checked = true;
            }
            else if (_Setting.czJudgeType == 2)
            {
                radioButton40.Checked = true;
            }
            if (_Setting.outDaqAccelerate)
            {
                checkBox5.Checked = true;
            }
            else
            {
                checkBox5.Checked =false;
            }

            if (_Setting.gjLbiOutHight)
            {
                checkBox4.Checked = true;
            }
            else
            {
                checkBox4.Checked = false;
            }
    

            this.radioButton28.Visible = false;
            if (_Setting.ParmStyle != StandardParmType.CityRoad && _Setting.ParmStyle != StandardParmType.CityRoadShanghai)
            {
                groupBox12.Visible = false;
            }
            else
            {
                groupBox12.Visible = true;
                if (_Setting.PartType == 0)
                {
                    radioButton_mile.Checked = true;
                    radioButton_dmi.Checked = false;
                }
                else
                {
                    radioButton_mile.Checked = false;
                    radioButton_dmi.Checked = true;
                }
                radioButton_dmi_CheckedChanged(null, null);
            }

        }
        private void 软件设置_Load(object sender, EventArgs e)
        {
            configFormControlState();
         
            switch (_Setting.ParmStyle)
            {
                case StandardParmType.DegreeRoad2007: LoadRoad2007(); break;
                case StandardParmType.CityRoad: LoadCity(); break;
                case StandardParmType.RuralRoadBeijing: LoadRoadBeiJin(); break;
                case StandardParmType.DegreeRoad2018: LoadRoad2018(); break;
                case StandardParmType.DegreeRoad2001: LoadRoad2001(); break;
                case StandardParmType.CityRoadShanghai: LoadCityShangHai(); break;
                case StandardParmType.RuralRoadLiaoning: LoadRoadLiaoNing(); break;
                case StandardParmType.RuralRoadGuangxi: LoadRoadGuangXi(); break;
                case StandardParmType.RuralRoadChongqing: LoadRoadChongQing(); break;
                case StandardParmType.RuralRoadHunan: LoadRoadHuNan(); break;
                case StandardParmType.RuralRoadlowLevel: LoadLowVillageRoad(); break;
                default: break;
            }     
            int heFeiSplit = _Setting.hefei2MinSplit;
            textBox1.Text = heFeiSplit.ToString();
            if (_Setting.zcSplit)
            {
                checkBox8.Checked = true;
            }
            else
            {
                checkBox8.Checked = false;
            }
            if (_Setting.splitExcelDh)
            {
                ck_th.Checked = true;
            }
            else
            {
                ck_th.Checked = false;
            }
            if (_Setting.JSAverageType)
            {
                checkBox7.Checked = true;
            }
            else
            {
                checkBox7.Checked = false;
            }
            if (_Setting.SelectDrawDis == 0)
            {
                radioButton32.Hide();
            }
            if (_Setting.Out_roadimg == 0)
            {
                checkBox_out_imgname.Checked = false;
            }
            else if (_Setting.Out_roadimg == 1)
            {
                checkBox_out_imgname.Checked = true;
            }

            if (_Setting.Out_roadinfo == 0)
            {
                checkBox_out_info.Checked = false;
            }
            else if (_Setting.Out_roadinfo == 1)
            {
                checkBox_out_info.Checked = true;
            }

            if (_Setting.sheetRoundingOffType == 0)
            {
                radioButton12.Checked = true;
            }
            else if (_Setting.sheetRoundingOffType == 1)
            {
                radioButton13.Checked = true;
            }
            if (_Setting.roadCrossingShow)
            {
                //不过滤
                rb_roadCrossing.Checked = false;
            }
            else
            {
                //过滤
                rb_roadCrossing.Checked = true;
            }

            numericUpDown1.Value = _Setting.sheetRoundingOffNum;

            textBox_DmiLen.Text = _Setting.PartType_Dmi_Len.ToString();
            int g_ttagval = 0;
            foreach (Control ctl in flowLayoutPanel6.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    g_ttagval = Convert.ToInt16(tctl.Tag);
                    if (g_ttagval == _Setting.SelectDrawDis)
                    {
                        tctl.Checked = true;
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel7.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = ctl as RadioButton;
                    g_ttagval = Convert.ToInt16(tctl.Tag);
                    if (g_ttagval == _Setting.Is_SnCarve)
                    {
                        tctl.Checked = true;
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel8.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    g_ttagval = Convert.ToInt16(tctl.Tag);
                    if (g_ttagval == _Setting.RQIJudgeType)
                    {
                        tctl.Checked = true;
                        break;
                    }
                }
            }

            foreach (Control ctl in flowLayoutPanel11.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tctl = (RadioButton)ctl;
                    g_ttagval = Convert.ToInt16(tctl.Tag);
                    if (g_ttagval == _Setting.IRIExcelSide)
                    {
                        tctl.Checked = true;
                        break;
                    }
                }
            }

            txb_ImageWidth.Text = _RoadConfig.ImageWidth.ToString();
            txb_ImageHeight.Text = _RoadConfig.ImageHeight.ToString();


            txb_RealWidth.Text = _RoadConfig.RealWidth.ToString();
            txb_RealHeight.Text = _RoadConfig.RealHeight.ToString();

            checkBox1.Checked = _Setting.IsForbidOverLapping;
            checkBox2.Checked = _Setting.IsCrackRemark;
            checkBox3.Checked = _Setting.IsOutputDisAreaSubtotal;

            if (_Setting.BrokenPlatetype == 0)
            {
                radioButton18.Checked = true;
                radioButton19.Checked = false;
            }
            else
            {
                radioButton18.Checked = false;
                radioButton19.Checked = true;
            }

            if (_Setting.IRIAlgorithmInterval == 0.1)
            {
                radioButton37.Checked = true;
            }
            else
            {
                radioButton38.Checked = true;

            }
            txb_bklength.Text = _Setting.PlateLength.ToString();
            txb_bkwidth.Text = _Setting.PlateWidth.ToString();

            txb_cmoprow.Text = _Setting.cmop_rows.ToString();

        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton13.Checked)
            {
                label8.Enabled = true;
                numericUpDown1.Enabled = true;
            }
            else
            {
                label8.Enabled = false;
                numericUpDown1.Enabled = false;
            }
        }

        private void groupBox12_Enter(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton43_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void rb_roadCrossing_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton43_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void ck_th_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void radioButton36_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
