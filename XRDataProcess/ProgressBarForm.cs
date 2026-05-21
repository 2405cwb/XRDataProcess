using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XRDataProcess
{
    public partial class ProgressBarForm : Form
    {
        public ProgressBarForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;
            this.progressBar1.Maximum = 100;
            this.TopMost = true;
        }
        /// <summary>
        /// 设置提示信息
        /// </summary>
        public string MessageInfo
        {
            set { this.label1.Text = value; }
        }
        /// <summary>
        /// 设置进度条显示值
        /// </summary>
        public int ProcessValue
        {
            set { this.progressBar1.Value = value; }
        }

        public ProgressBarStyle ProcessStyle
        {
            set { this.progressBar1.Style = value; }
        }

        private void ProgressBarForm_Load(object sender, EventArgs e)
        {
            this.progressBar1.Style = ProgressBarStyle.Continuous;

        }
    }
}
