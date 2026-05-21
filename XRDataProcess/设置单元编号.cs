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
    public partial class 设置单元编号 : Form
    {
        public static bool fg = false;
        public static string unitnum = "100000";
        public static string roadnum = "1";
        public 设置单元编号()
        {
            InitializeComponent();
        }

        private void btn_confirm_Click(object sender, EventArgs e)
        {
            if (txb_unitnum.Text == null || txb_RoadNum.Text == null)
            {
                MessageBox.Show("请输入单元编号或检测车道数。例如：1 ");
                return;
            }
            unitnum = txb_unitnum.Text.ToString();
            roadnum = txb_RoadNum.Text.ToString();
            fg = true;
            this.Close();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            fg = false;
            this.Close();
        }
    }
}
