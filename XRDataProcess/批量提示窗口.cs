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
    public partial class 批量提示窗口 : Form
    {
        public 批量提示窗口()
        {
            InitializeComponent();
        }

        public bool _IsOK = false;

        private void button1_Click(object sender, EventArgs e)
        {
            _IsOK = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _IsOK = false;
            this.Close();
        }
    }
}
