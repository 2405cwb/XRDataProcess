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
    public partial class 输出报告类型 : Form
    {
        public bool _IsOK = false;

        /// <summary>
        /// 选择的报告类型，0-百米公里报告，1-十米公里报告
        /// </summary>
        public int _WordType = -1;

        public 输出报告类型()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _IsOK = true;
            if (radioButton1.Checked)
                _WordType = 0;
            if (radioButton2.Checked)
                _WordType = 1;

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _IsOK = false;
            this.Close();
        }
    }
}
