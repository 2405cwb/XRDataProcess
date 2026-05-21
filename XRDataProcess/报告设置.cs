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
    public partial class 报告设置 : Form
    {
        public 报告设置()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(radioButton1.Checked)
            {
                MyWordCity._YHTypeDoc = 0;
            }

            if (radioButton2.Checked)
            {
                MyWordCity._YHTypeDoc = 1;
            }
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 报告设置_Load(object sender, EventArgs e)
        {
            if (MyWordCity._YHTypeDoc == 0)
            {
                radioButton1.Checked = true;
            }
            else if (MyWordCity._YHTypeDoc == 1)
            {
                radioButton2.Checked = true;
            }
        }
    }
}
