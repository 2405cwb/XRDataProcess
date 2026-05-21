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
    public partial class 四川公路院出表选择 : DevExpress.XtraEditors.XtraForm
    {
        private 四川公路院出表选择()
        {
            InitializeComponent();
            
        }
        private static 四川公路院出表选择 _form = null;
        public static 四川公路院出表选择 getInstance()
        {
            if (_form == null)
            {
                _form = new 四川公路院出表选择();
            }
            _form.StartPosition = FormStartPosition.CenterParent;
            _form.TopMost = true;
            return _form;
        }
        public List<int> getUserSelect()
        {
            List<int> result = new List<int>();
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                if (checkedListBoxControl1.Items[i].CheckState == CheckState.Checked)
                {
                        result.Add(i);
                }
               
            }
            return result;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                checkedListBoxControl1.Items[i].CheckState = CheckState.Checked;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                checkedListBoxControl1.Items[i].CheckState = CheckState.Unchecked;
            }
        }
    }
}