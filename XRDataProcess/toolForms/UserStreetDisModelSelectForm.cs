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

namespace XRDataProcess.toolForms
{
    public partial class UserStreetDisModelSelectForm : DevExpress.XtraEditors.XtraForm
    {
        public UserStreetDisModelSelectForm()
        {
            InitializeComponent();
        }
        public int SelectedModelIndex { get; private set; } = -1;

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            SelectedModelIndex = 1;
            this.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            SelectedModelIndex = 2;
            this.Close();
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            SelectedModelIndex = 3;
            this.Close();
        }

        private void UserStreetDisModelSelectForm_KeyDown(object sender, KeyEventArgs e)
        {
            // 假设你的按钮分别叫 btnMode1, btnMode2, btnMode3

            switch (e.KeyCode)
            {
                case Keys.D1:    // 或者是 Keys.D1 (对应键盘上方的数字1)
                case Keys.NumPad1: // 对应小键盘的1
                    simpleButton1.PerformClick(); // 假装点了一下按钮1
                    e.Handled = true; // 告诉系统“我知道了，不用再传给别人了”
                    break;

                case Keys.D2:
                case Keys.NumPad2:
                    simpleButton2.PerformClick();
                    e.Handled = true;
                    break;

                case Keys.D3:
                case Keys.NumPad3:
                    simpleButton3.PerformClick();
                    e.Handled = true;
                    break;

                case Keys.Escape: // 顺便做个按 ESC 关闭窗口
                    this.Close();
                    break;
            }
        }
    }
}