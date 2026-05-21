using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Framework.Other.MyGlobal;
namespace XRDataProcess
{
    public partial class WinStreetDisList : Form
    {
        public event EventHandler EventJump2Dis;

        public static bool _IsShowAuto = false;
        public WinStreetDisList()
        {
            InitializeComponent();
        }
   
        public void DeleteDis(object sender, EventArgs e)
        {
            StreetDisRecord Record = (StreetDisRecord)sender;
            int idx = 0;
            foreach (DataGridViewRow trow in dataGridView_Dislist.Rows)
            {
                if (trow.Cells[0].Value.ToString() == Record._mile
                    && trow.Cells[1].Value.ToString() == Record._disname
                    && Convert.ToInt32(trow.Cells[2].Value) == Record._score)
                {
                    dataGridView_Dislist.Rows.RemoveAt(idx);
                    break;
                }
                idx++;
            }
        }

        public void UpdateDisList(object sender, EventArgs e)
        {
            StreetDisRecord Record = (StreetDisRecord)sender;
            dataGridView_Dislist.Rows.Add(new object[] { Record._mile, Record._disname, Record._score });
        }

        public void LoadDisList(object sender, EventArgs e)
        {
            List<StreetDisRecord> disRecord = (List<StreetDisRecord>)sender;
            foreach (StreetDisRecord temp in disRecord)
            {
                dataGridView_Dislist.Rows.Add(new object[] { temp._mile, temp._disname, temp._score });
            }
        }

        private void dataGridView_Dislist_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridViewSelectedRowCollection selectrow = dataGridView_Dislist.SelectedRows;
                int mile = int.Parse(selectrow[0].Cells[0].Value.ToString().Replace("K", "").Replace("+", ""));
                EventJump2Dis(mile, EventArgs.Empty);
            }
            catch { }
        }

        private void button_update_Click(object sender, EventArgs e)
        {
            LoadDisList(sender,e);
        }

        private void WinStreetDisList_Load(object sender, EventArgs e)
        {
            LoadDisList(sender, e);
        }

        private void WinStreetDisList_Load_1(object sender, EventArgs e)
        {
         
        }
    }
}
