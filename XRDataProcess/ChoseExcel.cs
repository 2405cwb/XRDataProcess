using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XRDataProcess
{
    public partial class ChoseExcel : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        public bool _IsOK = false;
        public string _leftpath, _rightpath, _destpath;
        public ChoseExcel()
        {
            InitializeComponent();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (_Setting.ExcelType == 2)
            {
                OpenFileDialog fd = new OpenFileDialog();
                fd.Title = "请选择左幅行车道报表";
                fd.Filter = "Excel文件(*.xlsx)|*.xlsx";
                if (fd.ShowDialog() == DialogResult.OK)
                {
                    _leftpath = System.IO.Path.GetFullPath(fd.FileName);
                    textBox1.Text = _leftpath;
                }
            }
            else if (_Setting.ExcelType == 3)
            {
                FolderBrowserDialog fd = new FolderBrowserDialog();
                fd.Description = "请选择原始报表";
                fd.ShowDialog();
                if (fd.SelectedPath != string.Empty)
                {
                    if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                    {
                        fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                    }

                    _leftpath = fd.SelectedPath;
                    textBox1.Text = _leftpath;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (_Setting.ExcelType == 2)
            {
                OpenFileDialog fd = new OpenFileDialog();
                fd.Title = "请选择右幅行车道报表";
                fd.Filter = "Excel文件(*.xlsx)|*.xlsx";
                if (fd.ShowDialog() == DialogResult.OK)
                {
                    _rightpath = System.IO.Path.GetFullPath(fd.FileName);
                    textBox2.Text = _rightpath;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (_Setting.ExcelType == 2)
            {
                if (_leftpath != null && _rightpath != null && _destpath != null)
                {
                    _IsOK = true;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("路径选择不全，请检查！");
                }
            }
            else if (_Setting.ExcelType == 3)
            {
                if (_leftpath != null && _destpath != null)
                {
                    _IsOK = true;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("路径选择不全，请检查！");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择合并后报表放置位置：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }

                _destpath = fd.SelectedPath;
                textBox3.Text = _destpath;
            }
        }

        private void ChoseExcel_Load(object sender, EventArgs e)
        {
            if (_Setting.ExcelType == 3)
            {
                button2.Visible = false;
                textBox2.Visible = false;
                button1.Text = "原始报表";
            }
        }
    }
}
