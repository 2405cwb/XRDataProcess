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
using Framework.Other.MyGlobal;
namespace HNRoadFormatConverter.toolForms
{

    
    public partial class gjModelSelectForm : DevExpress.XtraEditors.XtraForm
    {
        Farmework.Other.XRSetting _Setting = Farmework.Other.XRSetting.GetInstance();
        public gjModelSelectForm() 
        {  
            this.StartPosition = FormStartPosition.CenterParent;
            InitializeComponent();
           
            
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            _Setting.gjStandardNew = (Farmework.Other.enumTools.hnEnumTools.CityModelItem)Enum.Parse(typeof( Farmework.Other.enumTools.hnEnumTools.CityModelItem), comboBoxEdit1.EditValue?.ToString());
           
            this.Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void gjModelSelectForm_Load(object sender, EventArgs e)
        {
            this.comboBoxEdit1.Properties.Items.Clear();
            this.comboBoxEdit1.Properties.Items.Add("等级公路5210与农村路5211标准模板导出_2025年");
           
            for (int i = 0; i < this.comboBoxEdit1.Properties.Items.Count; i++)
            {
                string curStr = this.comboBoxEdit1.Properties.Items[i].ToString();
                if (curStr == _Setting.gjStandardNew.ToString())
                {
                    this.comboBoxEdit1.SelectedIndex = i;
                    break;
                }
            }
        }
    }
}