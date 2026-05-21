using DevExpress.XtraEditors;
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
    public partial class MyTipForm : DevExpress.XtraEditors.XtraForm
    {
        public MyTipForm()
        {
            InitializeComponent();
        }

        public string FilterNameStr = "";
         
        public int StartRowIndex = 0;

        public bool NeedInstallRoadCode =false;

        public string TargetSheetName = "";
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            FilterNameStr = textEdit1.Text;
            StartRowIndex = int.Parse(textEdit2.Text);
            NeedInstallRoadCode = checkBox1.Checked;
            TargetSheetName = textEdit3.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void textEdit1_EditValueChanged(object sender, EventArgs e)
        {
            FilterNameStr  = textEdit1.Text;
        }

        private void textEdit2_EditValueChanged(object sender, EventArgs e)
        {
            StartRowIndex = int.Parse(textEdit2.Text);
        }
    }
}