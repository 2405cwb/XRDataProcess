using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace XRDataProcess
{
    public partial class 畸变矫正窗体 : DevExpress.XtraEditors.XtraForm
    {
        public 畸变矫正窗体()
        {
            InitializeComponent();
        }
        private string _width;
        private string _hight;
        public new string Width
        {
            get
            { return _width;}
          
        }
        public string hight { get { return _hight; } }
        public bool isOk;
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                this.textEdit1.Text = "3.75";
                this.textEdit2.Text = "2.00";
            }
        }

        private void 畸变矫正窗体_Load(object sender, EventArgs e)
        {
            this.textEdit1.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.textEdit2.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this._width = "";
            this._hight = "";
            isOk = false;
            this.radioButton1.Checked = true;
            this.textEdit1.Text = "3.75";
            this.textEdit2.Text = "2.00";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            isOk = true;
           this._width =this.textEdit1.Text;
            this._hight = this.textEdit2.Text;
            this.Hide();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                this.textEdit1.Text = "3.23";
                this.textEdit2.Text = "2.09";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            isOk = false;
            this.Hide();
        }
    }
}