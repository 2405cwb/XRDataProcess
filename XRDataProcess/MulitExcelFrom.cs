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

namespace XRDataProcess
{
    public partial class MulitExcelFrom : DevExpress.XtraEditors.XtraForm
    {
        XRSetting _Setting = XRSetting.GetInstance();
        public MulitExcelFrom()
        {
            InitializeComponent();
        }

        private void MulitExcelFrom_Load(object sender, EventArgs e)
        {
             
            switch (_Setting.ParmStyle)
            {
                case StandardParmType.DegreeRoad2007:
                    break;
                case StandardParmType.CityRoad:
                    break;
                case StandardParmType.RuralRoadBeijing:
                    break;
                case StandardParmType.DegreeRoad2018:
                    break;
                case StandardParmType.DegreeRoad2001:
                    break;
                case StandardParmType.CityRoadShanghai:
                    break;
                case StandardParmType.RuralRoadLiaoning:
                    break;
                case StandardParmType.RuralRoadGuangxi:
                    break;
                case StandardParmType.RuralRoadChongqing:
                    break;
                case StandardParmType.RuralRoadHunan:
                    break;
                case StandardParmType.RuralRoadlowLevel:
                    loadLowExcelType();
                    break;
                default:
                    break;
            }

        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            _Setting.multiExcelMergeType = this.radioGroup1.SelectedIndex;
            this.DialogResult = DialogResult.OK;
            this.Close();
            return;
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
            return;
        }
        private void loadLowExcelType()
        {
            List<string> radiosName = new List<string> { "多车道统计", "咸宁路况统计", "孝感路况统计", "南昌农村路统计" };
            for (int i = 0; i < radiosName.Count; i++)
            {
                var r = new DevExpress.XtraEditors.Controls.RadioGroupItem("单表统计表格", radiosName[i]);
                this.radioGroup1.Properties.Items.Add(r);

            }
            try
            {
                this.radioGroup1.SelectedIndex = _Setting.multiExcelMergeType;
            }
            catch (Exception)
            {
                this.radioGroup1.SelectedIndex = 0;


            }

        }
    }
}