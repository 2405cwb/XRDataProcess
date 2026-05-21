using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OperateIniFile;

namespace XRDataProcess
{
    public partial class 导出景观报表区间 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        public bool _IsExcel = false;
        public 导出景观报表区间()
        {
            InitializeComponent();
        }

        private void button_yes_Click(object sender, EventArgs e)
        {
            foreach(Control ctl in tableLayoutPanel1.Controls)
            {
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (idx >= 0)
                {
                    if (ctl is CheckBox)
                    {
                        CheckBox cbox = ctl as CheckBox;
                        _Setting.StreetIsExcel[idx] = cbox.Checked;
                    }
                    else if (ctl is ComboBox)
                    {
                        ComboBox cbox = ctl as ComboBox;
                        _Setting.StreetLenExcel[idx] = cbox.Text;
                    }
                }
            }
            _Setting.WriteData();
            _IsExcel = true;
            this.Close();
        }

        private void button_no_Click(object sender, EventArgs e)
        {
            _IsExcel = false;
            this.Close();
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (idx >= 0)
                {
                    if (ctl is CheckBox)
                    {
                        CheckBox cbox = ctl as CheckBox;
                        cbox.Checked = checkBox7.Checked;
                    }
                }
            }
        }

        private void ExcelDis_Load(object sender, EventArgs e)
        {
            
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                    
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                try
                {
                    if (idx >= 0)
                    {
                        if (ctl is CheckBox)
                        {
                            CheckBox cbox = ctl as CheckBox;
                            cbox.Checked = _Setting.StreetIsExcel[idx];
                        }
                        else if (ctl is ComboBox)
                        {
                            ComboBox cbox = ctl as ComboBox;
                            cbox.Text = _Setting.StreetLenExcel[idx];
                        }
                    }
                }
                catch (Exception)
                {

                    continue;
                }
                   
             
               
            }

            if (_Setting.ParmStyle != StandardParmType.RuralRoadlowLevel
                &&
                _Setting.ParmStyle != StandardParmType.RuralRoadHunan
                )
            {
                RemoveRow(this.tableLayoutPanel1, 5);
                RemoveRow(this.tableLayoutPanel1, 5);
            }
            else
            {

            }
              
        }
        private void RemoveRow(System.Windows.Forms.TableLayoutPanel tableLayoutPanel, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= tableLayoutPanel.RowCount)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index is out of range.");
            }

            // 移除该行中的所有控件
            for (int col = 0; col < tableLayoutPanel.ColumnCount; col++)
            {
                Control control = tableLayoutPanel.GetControlFromPosition(col, rowIndex);
                if (control != null)
                {
                    tableLayoutPanel.Controls.Remove(control);
                }
            }

            // 调整行定义
            tableLayoutPanel.RowStyles.RemoveAt(rowIndex);
            tableLayoutPanel.RowCount--;

            // 重新排列剩余的控件
            for (int row = rowIndex; row < tableLayoutPanel.RowCount; row++)
            {
                for (int col = 0; col < tableLayoutPanel.ColumnCount; col++)
                {
                    Control control = tableLayoutPanel.GetControlFromPosition(col, row + 1);
                    if (control != null)
                    {
                        tableLayoutPanel.SetCellPosition(control, new TableLayoutPanelCellPosition(col, row));
                    }
                }
            }
        }
        private List<int> subData = null;

        /// <summary>
        /// 是否需要分段
        /// </summary>
        private bool needSub = false;

        public bool NeedSub
        {
            get
            {
                return needSub;
            }

        }

        public List<int> SubData
        {
            get
            {
                return subData;
            }


        }
        private void setSubBtnColor(bool need = false)
        {
            if (need)
            {
                this.button_SetExcelMile.ForeColor = Color.Green;
            }
            else
            {
                this.button_SetExcelMile.ForeColor = Color.Gray;
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            设置报表桩号 setExcelMile = new 设置报表桩号();

            setExcelMile.ShowDialog();
            if (setExcelMile.Yes)
            {

                List<int> data = setExcelMile.getUserSortValue();

                if (data.Count % 2 != 0)
                {
                    MessageBox.Show("请您保证分段桩号是成对的且不为空！");
                    this.needSub = false;
                    setSubBtnColor(false);
                }
                else
                {
                    subData = data;
                    this.needSub = true;
                    setSubBtnColor(true);
                }
            }
            else
            {
                this.needSub = false;
                setSubBtnColor(false);
            }
        }
    }
}
