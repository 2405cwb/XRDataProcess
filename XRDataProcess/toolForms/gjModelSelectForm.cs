using DevExpress.XtraEditors;
using Farmework.Other.enumTools;
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

    
    public partial class gjModelSelectForm : DevExpress.XtraEditors.XtraForm
    {
        XRSetting _Setting = XRSetting.GetInstance();
        public gjModelSelectForm() 
        {  
            this.StartPosition = FormStartPosition.CenterParent;
            InitializeComponent();
           
            
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            _Setting.gjStandardNew = (hnEnumTools.CityModelItem)Enum.Parse(typeof( hnEnumTools.CityModelItem), comboBoxEdit1.EditValue?.ToString());
           
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
            
            //this.comboBoxEdit1.Properties.Items.Add("交通部2024规范");
            this.comboBoxEdit1.Properties.Items.Add("河南省单位一农村路定制");
            this.comboBoxEdit1.Properties.Items.Add("农养国省道路况检测数据提交格式_2026年"); //2026年国省道路况检测数据提交格式
             // this.comboBoxEdit1.Properties.Items.Add("湖南省单位一定制");
             //this.comboBoxEdit1.Properties.Items.Add("重庆市单位一定制");
             //this.comboBoxEdit1.Properties.Items.Add("甘肃省单位一定制");
             // this.comboBoxEdit1.Properties.Items.Add("河北省单位一定制");
             // this.comboBoxEdit1.Properties.Items.Add("河北省单位二定制"); 
             //this.comboBoxEdit1.Properties.Items.Add("安徽省单位一定制");
             //this.comboBoxEdit1.Properties.Items.Add("广东省单位一定制");
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