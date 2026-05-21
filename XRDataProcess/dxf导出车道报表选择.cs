using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OperateIniFile;

namespace XRDataProcess
{
    public partial class dxf导出车道报表选择 : Form
    {
        public dxf导出车道报表选择()
        {
            InitializeComponent();
            textBox_beginMile.Text = "0";
            textBox_endMile.Text = "0";
        }

        /// <summary>
        /// 上行的报表文件：车道、指标
        /// </summary>
        public string[][] _UpXlsFiles = null;

        /// <summary>
        /// 下行的报表文件：车道、指标
        /// </summary>
        public string[][] _DownXlsFiles = null;

        /// <summary>
        /// 上行的报表文件：车道、指标
        /// </summary>
        public List<string> _UpXlsFileList = null;

        /// <summary>
        /// 下行的报表文件：车道、指标
        /// </summary>
        public List<string> _DownXlsFileList = null;

         /// <summary>
        /// 保存路径
        /// </summary>
        public string _savePath = null;

        /// <summary>
        /// 是否处理
        /// </summary>
        public bool _IsOK = false;

        public double _BeginMile;

        public double _EndMile;

        /// <summary>
        /// 是否国省道
        /// </summary>
        public bool _IsProvinceRoad = false;


        void bbox_Click_Up(object sender, EventArgs e)
        {
            Button bbox = (Button)sender;
            int idx = Convert.ToInt16(bbox.Tag);

            TextBox tbox = null;
            foreach (Control ctl in tableLayoutPanel_Up.Controls)
            {
                if (ctl is TextBox)
                {
                    tbox = ctl as TextBox;
                    int idxt = Convert.ToInt16(tbox.Tag);
                    if (idx == idxt)
                    {
                        break;
                    }
                }
            }

            if (tbox == null)
                return;

            tbox.Text = "";

            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
            fd.RestoreDirectory = true;
            fd.FilterIndex = 1;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                tbox.Text = fd.FileName;
                toolTip1.SetToolTip(tbox, tbox.Text);

                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt == idx + 1)
                        {
                            bbox.Enabled = true;
                        }
                        else if (idxt >= idx + 1)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
            else
            {
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt > idx)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
        }

        void bbox_Click_Down(object sender, EventArgs e)
        {
            Button bbox = (Button)sender;
            int idx = Convert.ToInt16(bbox.Tag);

            TextBox tbox = null;
            foreach (Control ctl in tableLayoutPanel_Down.Controls)
            {
                if (ctl is TextBox)
                {
                    tbox = ctl as TextBox;
                    int idxt = Convert.ToInt16(tbox.Tag);
                    if (idx == idxt)
                    {
                        break;
                    }
                }
            }

            if (tbox == null)
                return;

            tbox.Text = "";

            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
            fd.RestoreDirectory = true;
            fd.FilterIndex = 1;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                tbox.Text = fd.FileName;        
                toolTip1.SetToolTip(tbox, tbox.Text);

                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt == idx + 1)
                        {
                            bbox.Enabled = true;
                        }
                        else if (idxt >= idx + 1)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
            else
            {
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt > idx)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
        }

        private void button_Yes_Click(object sender, EventArgs e)
        {
          
            string[] upxlsdirpaths = new string[tableLayoutPanel_Up.Controls.Count];
            string[] downxlsdirpaths = new string[tableLayoutPanel_Down.Controls.Count];
            _BeginMile = Double.Parse(textBox_beginMile.Text);
            _EndMile = Double.Parse(textBox_endMile.Text);
            _IsProvinceRoad = checkBox_ProvinceRoad.Checked;
            _UpXlsFileList = new List<string>();
            _DownXlsFileList = new List<string>();

            if (checkBox_ProvinceRoad.Checked)
            {
                if (radioButton_full.Checked)
                {
                    foreach (Control ctl in tableLayoutPanel_Down.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            downxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < downxlsdirpaths.Length; ++i)
                    {
                        if (downxlsdirpaths[i] != string.Empty)
                        {
                            _UpXlsFileList.Add(downxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                    foreach (Control ctl in tableLayoutPanel_Up.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            upxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < upxlsdirpaths.Length; ++i)
                    {
                        if (upxlsdirpaths[i] != string.Empty)
                        {
                            _UpXlsFileList.Add(upxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else if (radioButton_halfDown.Checked)
                {
                    foreach (Control ctl in tableLayoutPanel_Down.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            downxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < downxlsdirpaths.Length; ++i)
                    {
                        if (downxlsdirpaths[i] != string.Empty)
                        {
                            _DownXlsFileList.Add(downxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                else if (radioButton_halfUp.Checked)
                {
                    foreach (Control ctl in tableLayoutPanel_Up.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            upxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < upxlsdirpaths.Length; ++i)
                    {
                        if (upxlsdirpaths[i] != string.Empty)
                        {
                            _UpXlsFileList.Add(upxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            else
            {
                if (radioButton_full.Checked || radioButton_halfUp.Checked)
                {
                    foreach (Control ctl in tableLayoutPanel_Up.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            upxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < upxlsdirpaths.Length; ++i)
                    {
                        if (upxlsdirpaths[i] != string.Empty)
                        {
                            _UpXlsFileList.Add(upxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                if (radioButton_full.Checked || radioButton_halfDown.Checked)
                {
                    foreach (Control ctl in tableLayoutPanel_Down.Controls)
                    {
                        if (ctl is TextBox)
                        {
                            TextBox tbox = ctl as TextBox;
                            int idxt = Convert.ToInt16(tbox.Tag);
                            downxlsdirpaths[idxt] = tbox.Text;
                        }
                    }
                    for (int i = 0; i < downxlsdirpaths.Length; ++i)
                    {
                        if (downxlsdirpaths[i] != string.Empty)
                        {
                            _DownXlsFileList.Add(downxlsdirpaths[i]);
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                if (radioButton_full.Checked)
                {
                    if (_UpXlsFileList.Count == 0 || _DownXlsFileList.Count == 0)
                    {
                        MessageBox.Show("请选择全幅的上下行车道报表！");
                        return;
                    }
                }
                else if (radioButton_halfUp.Checked)
                {
                    if (_UpXlsFileList.Count == 0)
                    {
                        MessageBox.Show(string.Format("已勾选【上行半幅】，没有选择上行半幅车道报表，请检查！",
                            _UpXlsFileList.Count, _DownXlsFileList.Count));
                        return;
                    }
                }
                else if (radioButton_halfDown.Checked)
                {
                    if (_DownXlsFileList.Count == 0)
                    {
                        MessageBox.Show(string.Format("已勾选【下行半幅】，没有选择下行半幅车道报表，请检查！",
                            _UpXlsFileList.Count, _DownXlsFileList.Count));
                        return;
                    }
                }
            }
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择dxf导出路径：";
    
             fd.ShowDialog();
             if (fd.SelectedPath != string.Empty)
             {
                 if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                 {
                     fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                 }
                 _savePath = fd.SelectedPath;
                 _IsOK = true;
                 this.Close();
                 return;
             }
            _IsOK = false;
            this.Close();
        }

        private void button_Esc_Click(object sender, EventArgs e)
        {
            _IsOK = false;
            this.Close();
        }

        private void radioButton_full_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = true;
            groupBox_Down.Enabled = true;
        }

        private void radioButton_halfUp_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = true;
            groupBox_Down.Enabled = false;
        }

        private void radioButton_halfDown_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = false;
            groupBox_Down.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void 车道报表选择_Load(object sender, EventArgs e)
        {
            foreach (Control ctl in tableLayoutPanel_Up.Controls)
            {
                if (ctl is Button)
                {
                    Button bbox = ctl as Button;
                    bbox.Click += new EventHandler(bbox_Click_Up);
                }
                else if (ctl is TextBox)
                {
                    //TextBox tbox = ctl as TextBox;
                    //int idx = Convert.ToInt32(tbox.Tag);
                    //tbox.Text = inisetting.ReadString("MergeInfo", "UpPath" + idx.ToString(), "");
                    //if (!Directory.Exists(tbox.Text))
                    //{
                    //    tbox.Text = "";
                    //}
                }
            }

            foreach (Control ctl in tableLayoutPanel_Down.Controls)
            {
                if (ctl is Button)
                {
                    Button bbox = ctl as Button;
                    bbox.Click += new EventHandler(bbox_Click_Down);
                }
                else if (ctl is TextBox)
                {
                    //TextBox tbox = ctl as TextBox;
                    //int idx = Convert.ToInt32(tbox.Tag);
                    //tbox.Text = inisetting.ReadString("MergeInfo", "DownPath" + idx.ToString(), "");
                    //if (!Directory.Exists(tbox.Text))
                    //{
                    //    tbox.Text = "";
                    //}
                }
            }
        }

     
    }

}
