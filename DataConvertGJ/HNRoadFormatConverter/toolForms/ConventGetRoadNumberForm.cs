using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace HNRoadFormatConverter.toolForms
{
    public partial class ConventGetRoadNumberForm : DevExpress.XtraEditors.XtraForm
    {
        public ConventGetRoadNumberForm()
        {
            InitializeComponent();
            
        }
        private string _roadNum;

        public string RoadNum
        {
            get { return _roadNum; }
            set { _roadNum = value; }
        }

        private string name;

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        private string cityNum;

        public string CityNum
        {
            get { return cityNum; }
            set { cityNum = value; }
        }

        public ConventGetRoadNumberForm(string roadNum)
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterParent;
            string roadStr = roadNum.Substring(0, 4);
            RoadNum = roadStr;
            this.textEdit1.EditValue = roadStr;
        }
        public bool Ok { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {
            name = RoadNum + this.textEdit2.EditValue;
            CityNum = this.textEdit2.EditValue.ToString();
            Ok = true;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Ok = false;
            this.Close();
        }
    }
}