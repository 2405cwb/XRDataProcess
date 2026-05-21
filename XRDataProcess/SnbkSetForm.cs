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
    public partial class SnbkSetForm : Form
    {
        public SnbkSetForm()
        {
            InitializeComponent();
        }
        public static int bknum = -1;
        public bool falg = false;
        private void btn_confirm_Click(object sender, EventArgs e)
        {     
            try
            {
                bknum = int.Parse(txb_bknum.Text.ToString());
            }
            catch
            {
                MessageBox.Show("板块编号格式不正确,请输入整数");
                return;
            }
            falg = true;
            this.Close();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
