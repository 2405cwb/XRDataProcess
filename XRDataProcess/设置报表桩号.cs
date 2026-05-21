using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XRDataProcess
{
    
    public partial class 设置报表桩号 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        public 设置报表桩号()
        {
            InitializeComponent();
      

           // dataGridView1.Columns[0].DefaultCellStyle.Format = "k000_000";

        }
        private void getDatas()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("起始桩号", typeof(int));
            dt.Columns.Add("结束桩号", typeof(int));
            if (_Setting.needSub)
            {
              
                string subDatas = _Setting.subData;
                if (string.IsNullOrEmpty(subDatas))
                {
                    dataGridView1.DataSource = dt;
                    return;
                }
                string[] data = subDatas.Split(',');
                
               
                for (int i = 0; i < data.Length - 1; i += 2)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = data[i];
                    dr[1] = data[i+1];
                    dt.Rows.Add(dr);
                }

            }
            dataGridView1.DataSource = dt;
        }

        private bool yes = false;

        public bool Yes
        {
            get
            {
                return yes;
            } 
           
        }

        private void button_No_Click(object sender, EventArgs e)
        {
            yes = false;
            this.Close();
        }
        public List<int> getUserSortValue()
        {
            //获取用户输入数据
           List<int> sortInts = new List<int>();
            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    sortInts.Add(int.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString()));
                }
            }

            return sortInts;
        }
        private void button_Yes_Click(object sender, EventArgs e)
        {
            var date = getUserSortValue();
            if (date.Count<=0)
            {
                yes = false;
            }
            else
            {
                 yes = true;
            }
            //对数据进行返回
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CLearnDataGridView();
        }

        /// <summary>
        /// 清空data并设置取消分段出表
        /// </summary>
        private void CLearnDataGridView()
        {
            var dataTable = (DataTable)dataGridView1.DataSource;
            if (dataTable == null)
            {
                return;
            }
            dataTable.Rows.Clear();
            dataGridView1.DataSource = dataTable;
            yes = false;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("数据格式应为整数，请修改错误");
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {
            if (sender is Button)
            {
                Button btn = (Button)sender;
                if (btn.Name.Equals("button_Yes"))
                {
                    this.toolTip1.Show("设置分段出表", btn);
                }
            }
        }

        private void 设置报表桩号_Load(object sender, EventArgs e)
        {
            getDatas();
        }
    }
}
