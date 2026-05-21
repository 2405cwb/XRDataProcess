using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using hnPanoShowAPI;

namespace hnPanoShowTest
{
    public partial class Form1 : Form
    {
        public PanoControlPanel panoCtrl;

        public Form1()
        {
            InitializeComponent();

            panoCtrl = new PanoControlPanel();
          //  panoCtrl.ThrowPanoPoint += new PanoPointEventHandler(panoCtrl_ThrowPanoPoint);
          //  panoCtrl.ThrowMouseLeave += new hnShowPano.PanoPointLeaveEventHandler(panoCtrl_ThrowMouseLeave);

            this.panel1.Controls.Add(panoCtrl);
            panoCtrl.Dock = DockStyle.Fill;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panoCtrl.Dock = DockStyle.Fill;
            this.panel1.Controls.Add(panoCtrl);

            
        }

        private void 加载影像ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // 读取影像
            string strImagePath = "I:\\DATA\\iScan-00000000-20200723171912\\Image\\1\\HDPano\\00000000-01-20200723172428764.jpg";

            panoCtrl.addImage(strImagePath);
        }

    }
}
