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

namespace HNRoadFormatConverter
{
    public partial class SelectSourceForm : DevExpress.XtraEditors.XtraForm
    {
        public string SelectedSource { get; private set; }
        public SelectSourceForm()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
            // 任一 RadioButton 被点击 → 立即选中 + 关闭
            radioButton1.CheckedChanged += RadioButton_CheckedChanged;
            radioButton2.CheckedChanged += RadioButton_CheckedChanged;
        }
        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            var radio = sender as RadioButton;
            if (radio?.Checked == true)
            {
                SelectedSource = radio.Text;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}