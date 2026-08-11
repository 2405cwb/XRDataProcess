using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms; 
using OperateIniFile;

namespace XRDataProcess
{
    public partial class 导出路面报表区间 : Form
    {

        // public Action<bool>  ChangeSubColor = null;
        //  public event Action<bool> changeSubColorEvent = null;
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public bool _IsExcel = false;

        public 导出路面报表区间(int ProjectNum)
        {
            InitializeComponent();
            if (ProjectNum > 1)
            {
                button_SetExcelMile.Visible = false;
            }

        }


        private void button_yes_Click(object sender, EventArgs e)
        {
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (idx >= 0)
                {
                    if (ctl is CheckBox)
                    {
                        CheckBox cbox = ctl as CheckBox;
                        _Setting.IsExcel[idx] = cbox.Checked;
                        if (cbox.Checked)
                        {
                            _IsExcel = true;
                        }
                    }
                    else if (ctl is ComboBox)
                    {
                        ComboBox cbox = ctl as ComboBox;
                        _Setting.LenExcel[idx] = cbox.Text;
                    }
                }
            }
            _Setting.DutyUnit = textBox_DutyUnit.Text;
            _Setting.RoadSideType = textBox_RoadSideType.Text;
            _Setting.DetectYear = comboBox_DetectYear.Text;
            _Setting.DetectNum = comboBox_DetectNum.Text;
            _Setting.DistrictCode = textBox_DistrictCode.Text;
            _Setting.WriteData();
            _RoadConfig.DetectWidth = double.Parse(textBox_detectwidth.Text);
            _RoadConfig.WriteData();

            this.Close();
        }

        private void button_no_Click(object sender, EventArgs e)
        {
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
        /// <summary>
        /// 当分段激活时文字为绿色
        /// </summary>
        /// <param name="need"></param>
        private void setSubBtnColor()
        {
            if (_Setting.needSub)
            {
                this.needSub = true;
                string[] subStr = _Setting.subData.Split(',');
                List<int> temp = new List<int>();
                foreach (var item in subStr)
                {
                    temp.Add(int.Parse(item));
                }
                this.subData = temp;
                this.button_SetExcelMile.ForeColor = Color.Green;
            }
            else if (!_Setting.needSub)
            {
                this.button_SetExcelMile.ForeColor = Color.Gray;
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
        private void ExcelDis_Load(object sender, EventArgs e)
        {
            setSubBtnColor();

            textBox_DutyUnit.Text = _Setting.DutyUnit;
            textBox_RoadSideType.Text = _Setting.RoadSideType;
            comboBox_DetectYear.Text = _Setting.DetectYear;
            comboBox_DetectNum.Text = _Setting.DetectNum;
            textBox_DistrictCode.Text = _Setting.DistrictCode;

            textBox_detectwidth.Text = _RoadConfig.DetectWidth.ToString();

            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                int idx = Convert.ToInt32(ctl.Tag.ToString());
                if (idx >= 0)
                {
                    if (ctl is CheckBox)
                    {
                        CheckBox cbox = ctl as CheckBox;
                        cbox.Checked = _Setting.IsExcel[idx];
                        cbox.Visible = false;
                    }
                    else if (ctl is ComboBox)
                    {
                        ComboBox cbox = ctl as ComboBox;
                        cbox.Text = _Setting.LenExcel[idx].ToString();
                        cbox.Visible = false;
                    }
                }
            }

            switch (_Setting.ParmStyle)
            {
                case StandardParmType.DegreeRoad2007: SetRoad2007ExcelShow(_Setting.ExcelType); break;
                case StandardParmType.CityRoad: SetCity2016ExcelShow(_Setting.ExcelType, _Setting.PartType); break;
                case StandardParmType.RuralRoadBeijing: SetBJExcelShow(_Setting.ExcelType); break;
                case StandardParmType.DegreeRoad2018: SetRoad2018ExcelShow(_Setting.ExcelType); break;
                case StandardParmType.DegreeRoad2001: SetRoad2001ExcelShow(_Setting.ExcelType); break;
                case StandardParmType.CityRoadShanghai: SetSHCity2013ExcelShow(_Setting.ExcelType, _Setting.PartType); break;
                case StandardParmType.RuralRoadLiaoning: SetLNExcelShow(_Setting.ExcelType); break;
                case StandardParmType.RuralRoadGuangxi: SetGXExcelShow(_Setting.ExcelType); break;
                case StandardParmType.RuralRoadChongqing: SetCQExcelShow(_Setting.ExcelType); break;
                case StandardParmType.RuralRoadlowLevel: SetNCExcelShow(_Setting.ExcelType); break;
                case StandardParmType.RuralRoadHunan: SetHNExcelShow(_Setting.ExcelType); break;
                default: break;
            }

        }

        private void SetRoad2007ExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 7 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 1:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 2:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 3:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 4:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 5:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                                if (idx == 0)
                                {
                                    if (ExcelType == 1 || ExcelType == 2 || ExcelType == 4 || ExcelType == 5)
                                    {
                                        cbox.Text = "综合报表区间距离";
                                    }
                                }
                                else if (idx == 4)
                                {
                                    if (ExcelType == 3)
                                    {
                                        cbox.Text = "路面高程、GPS报表";
                                    }
                                }
                                else if (idx == 5)
                                {
                                    if (ExcelType == 3)
                                    {
                                        cbox.Text = "综合大表";
                                    }
                                }
                            }
                            else if (ctl is ComboBox)
                            {

                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetCity2016ExcelShow(int ExcelType, int PartType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 6, 7 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 4:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 5:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 6:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 7:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 12:
                    {
                        int[] tmpidx = { 0,1 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 13:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 14:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                                if (idx == 0)
                                {
                                    if (ExcelType == 4 || ExcelType == 5 || ExcelType == 7)
                                    {
                                        cbox.Text = "综合报表区间距离";
                                    }
                                }
                                if (idx == 6)
                                {
                                    if (ExcelType == 0)
                                    {
                                        cbox.Text = "路面磨耗评价等级统计表";
                                    }
                                }
                                if(idx==7)
                                {
                                    if (ExcelType == 0)
                                    {

                                        cbox.Text = "路面平整度评价等级记录表(带车速)";

                                    }
                                }
                                 if (ExcelType == 12)
                                {
                                    if (idx == 0)
                                    {
                                        cbox.Text = "上海惠浦:病害表格";

                                    }
                                    if (idx == 1)
                                    {
                                        cbox.Text = "上海惠浦:表格iri";

                                    }
                                }
                                if (ExcelType == 14)
                                {
                                    if (idx == 0)
                                    {
                                        cbox.Text = "公里指标报表"; 
                                    }

                                }


                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                if (PartType == 0)
                                {
                                    if(ExcelType == 13)
                                    {
                                        cbox.Visible = false;
                                    }
                                   
                                    else if (ExcelType == 12)
                                    {
                                        if (idx == 0)
                                        {
                                            cbox.Visible = false;

                                        }
                                        else
                                        {
                                            cbox.Visible = true;

                                        }
                                    }
                                    else if (ExcelType == 14)
                                    {
                                        if (idx == 0)
                                        {
                                            cbox.Visible = true;

                                        }
                                      
                                    }
                                    else if (ExcelType != 4)
                                    {
                                        cbox.Visible = true;
                                    }
                                    
                                }
                               
                                else
                                {
                                    cbox.Visible = false;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetBJExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 2, 3, 4, 5, 7 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetLNExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 2, 3, 4, 5, 8 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
                    }
                }
            }
        }


        private void SetRoad2018ExcelShow(int ExcelType)
        {
            HashSet<int> showIdxs = new HashSet<int>();
            List<string> modelNames = new List<string>();

            // 1. 根据 ExcelType 初始化需要显示的索引和名称配置
            switch (ExcelType)
            {
                case 0: showIdxs = new HashSet<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 }; break;
                case 1: showIdxs = _Setting.SelectDrawDis == 0 ? new HashSet<int> { 0, 1 } : new HashSet<int> { 0, 1, 2 }; break;
                case 2:
                case 4:
                case 5:
                case 15:
                case 17: showIdxs = new HashSet<int> { 0 }; break;
                case 3: showIdxs = new HashSet<int> { 0, 1, 2, 3, 4, 5, 6 }; break;
                case 8: showIdxs = new HashSet<int> { 0, 1, 2, 3, 5, 6, 9 }; break;
                case 9: showIdxs = new HashSet<int> { 0, 1 }; break;
                case 10: showIdxs = new HashSet<int> { 4 }; break;
                case 12: showIdxs = new HashSet<int> { 0, 1, 2, 3, 4, 5 }; break;
                case 13:
                    showIdxs = _Setting.SelectDrawDis == 1 ? new HashSet<int> { 0, 1, 2, 3, 4, 5, 6, 7, 8 } : new HashSet<int> { 0, 1, 2, 3, 4, 5, 6, 7 };
                    modelNames.AddRange(new[] { "空间定位数据", "路面平整度自动化检测数据", "路面平整度自动化检测原始数据", "DR破损率csv检测结果数据表格", "IRI平整度csv检测结果数据表格", "IRI平整度csv检测原始数据表格", "空间定位数据csv检测原始数据表格", "车辙数据0.1米导出" });
                    if (_Setting.SelectDrawDis == 1) modelNames.Add("四川定位数据定制");
                    break;
                case 14:
                    if (_Setting.SelectDrawDis == 1) showIdxs = new HashSet<int> { 0 };
                    break;
                case 18:
                    showIdxs = new HashSet<int> { 0 };
                    modelNames.Add("病害调绘报表");
                    break;
                case 19:
                    showIdxs = new HashSet<int> { 0, 1 };
                    modelNames.AddRange(new[] { "平整度报表", "路面病害汇总报表" });
                    break;
                case 20:
                    showIdxs = new HashSet<int> { 0 };
                    modelNames.Add("重庆道路病害统计定制表");
                    break;
                case 21:
                    showIdxs = new HashSet<int> { 0, 1, 2, 3, 4, 5 };
                    modelNames.AddRange(new[] { "高等级沥青路面破损", "高等级水泥路面破损", "路面平整度", "路面磨耗", "路面跳车", "路面车辙" });
                    break;
            }

            if (showIdxs.Count == 0) return;

            // 2. 遍历 UI 控件并应用状态
            foreach (Control ctl in tableLayoutPanel1.Controls)
            {
                if (ctl.Tag == null || !int.TryParse(ctl.Tag.ToString(), out int idx)) continue;

                // 如果当前控件的索引在需要显示的集合中
                if (showIdxs.Contains(idx))
                {
                    if (ctl is CheckBox cbox)
                    {
                        cbox.Visible = true;
                        string newText = GetCheckboxText(ExcelType, idx, modelNames);
                        if (!string.IsNullOrEmpty(newText))
                        {
                            cbox.Text = newText;
                        }
                    }
                    else if (ctl is ComboBox combo)
                    {
                        combo.Visible = IsComboBoxVisible(ExcelType, idx);
                    }
                }
            }
        }

        // 辅助方法 1：获取 CheckBox 应该显示的文本
        private string GetCheckboxText(int excelType, int idx, List<string> modelNames)
        {
            // 如果有动态加载的名称（如 case 13, 18, 19, 20, 21），优先使用
            if (modelNames != null && idx < modelNames.Count)
            {
                return modelNames[idx];
            }

            // 处理静态的特定规则
            switch (idx)
            {
                case 0:
                    if (excelType == 1 || excelType == 2 || excelType == 4 || excelType == 5) return "综合报表区间距离";
                    if (excelType == 9) return "综合PQI评定表格";
                    if (excelType == 12) return "上海惠浦:病害表格";
                    if (excelType == 14) return "公路院数据汇总";
                    if (excelType == 15 || excelType == 17) return "合肥路况";
                    break;
                case 1:
                    if (excelType == 9) return "路况调查表及评定汇总";
                    if (excelType == 12) return "上海惠浦:表格iri";
                    return "路面磨耗评价等级统计表"; // 默认兜底
                case 2:
                    if (excelType == 1) return "公路技术状况评定结果";
                    if (excelType == 12) return "上海惠浦:表格PBI";
                    break;
                case 3:
                    if (excelType == 12) return "上海惠浦:表格Rut";
                    break;
                case 4:
                    if (excelType == 12) return "芜湖:表格DR";
                    if (excelType == 3) return "路面高程、GPS报表";
                    break;
                case 5:
                    if (excelType == 12) return "芜湖:表格IRI";
                    if (excelType == 3) return "综合大表";
                    break;
            }

            return null; // 返回 null 表示保持原有 Text 不变
        }

        // 辅助方法 2：判断 ComboBox 是否应该显示
        private bool IsComboBoxVisible(int excelType, int idx)
        {
            if (excelType == 12)
            {
                if (idx == 0 || idx == 4 || idx==5) return false; 

            }
            if ( excelType == 18) return false;
            if (excelType == 1 && idx == 2) return false;
            if (excelType == 13 && (idx == 2 || idx == 7)) return false;

            return true; // 默认显示
        }

        private void SetRoad2018ExcelShow1(int ExcelType)
        {
            int[] showidxs = null; List<string> modelNames = new List<string>();
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 1:
                    {
                       
                        int[] tmpidx = { 0, 1,2 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        if (_Setting.SelectDrawDis == 0)
                        {
                            tmpidx = new int[] { 0, 1 };
                            showidxs = new int[tmpidx.Length];
                            Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        }
                    }
                    break;
                case 2:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 3:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 6 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 4:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 5:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 8:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 5, 6, 9 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 9:
                    {
                        int[] tmpidx = { 0, 1 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 10:
                    {
                        int[] tmpidx = { 4 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                //上海惠浦
                case 12:
                    {
                        int[] tmpidx = { 0, 1, 2,3 ,4,5};
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 13:
                    {
                        int[] tmpidx;
                        if (_Setting.SelectDrawDis == 1)
                        {
                            tmpidx = new int[] { 0, 1, 2, 3, 4, 5, 6,7,8 };

                        }
                        else
                        {
                              tmpidx = new int[] { 0, 1, 2, 3, 4, 5, 6, 7, };

                        }
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("空间定位数据");
                        modelNames.Add("路面平整度自动化检测数据");
                        modelNames.Add("路面平整度自动化检测原始数据");

                        modelNames.Add("DR破损率csv检测结果数据表格");
                        modelNames.Add("IRI平整度csv检测结果数据表格");
                        modelNames.Add("IRI平整度csv检测原始数据表格");
                        modelNames.Add("空间定位数据csv检测原始数据表格");
                        modelNames.Add("车辙数据0.1米导出"); 
                        if (_Setting.SelectDrawDis == 1)
                        {
                            modelNames.Add("四川定位数据定制");

                        }

                    }
                    break;
                case 14:
                    {
                        if (_Setting.SelectDrawDis==1)
                        {
                            int[] tmpidx = { 0 };
                            showidxs = new int[tmpidx.Length];
                            Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        }
                        
                    }
                    break;
                case 15:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 17:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 18:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("病害调绘报表");
                    }
                    break;
                case 19:
                    {
                        int[] tmpidx = { 0 ,1};
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("平整度报表");
                        modelNames.Add("路面病害汇总报表");
                    }
                    break;
                case 20:
                    {
                        int[] tmpidx = { 0, };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("重庆道路病害统计定制表");
                    }
                    break;
                case 21:
                    {
                        int[] tmpidx = { 0, 1,2,3,4,5 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("高等级沥青路面破损");
                        modelNames.Add("高等级水泥路面破损");
                        modelNames.Add("路面平整度");
                        modelNames.Add("路面磨耗");
                        modelNames.Add("路面跳车");
                        modelNames.Add("路面车辙");

                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                                if (idx == 0)
                                {
                                    if (ExcelType == 1 || ExcelType == 2 || ExcelType == 4 || ExcelType == 5)
                                    {
                                        cbox.Text = "综合报表区间距离";
                                    }
                                    else if (ExcelType == 9)
                                    {
                                        cbox.Text = "综合PQI评定表格";
                                    }
                                    else if (ExcelType == 12)
                                    {
                                        cbox.Text = "上海惠浦:病害表格";
                                    }
                                    else if (ExcelType == 14)
                                    {
                                        cbox.Text = "公路院数据汇总";
                                    }
                                    else if (ExcelType == 15 || ExcelType == 17)
                                    {
                                        cbox.Text = "合肥路况";
                                    }
                                }
                                else if (idx == 1)
                                {
                                    if (ExcelType == 9)
                                    {
                                        cbox.Text = "路况调查表及评定汇总";
                                    }
                                    else if (ExcelType == 12)
                                    {
                                        cbox.Text = "上海惠浦:表格iri";
                                    }
                                    else
                                    {
                                        cbox.Text = "路面磨耗评价等级统计表";
                                    }

                                }
                                else if (idx ==2)
                                {
                                    if (ExcelType == 1)
                                    {
                                        cbox.Text = "公路技术状况评定结果";
                                    }
                                    if (ExcelType == 12)
                                    {
                                        cbox.Text = "上海惠浦:表格PBI";
                                    }
                                }
                                else if (idx == 3 && ExcelType == 12)
                                    cbox.Text = "上海惠浦:表格Rut";
                                else if (idx == 4 && ExcelType == 12)
                                    cbox.Text = "芜湖:表格DR";
                                else if (idx == 5 && ExcelType == 12)
                                    cbox.Text = "芜湖:表格IRI";
                                else if (idx == 4)
                                {
                                    if (ExcelType == 3)
                                    {
                                        cbox.Text = "路面高程、GPS报表";
                                    }
                                }
                                else if (idx == 5)
                                {
                                    if (ExcelType == 3)
                                    {
                                        cbox.Text = "综合大表";
                                    }
                                }
                            }
                            if (idx == showidxs[tt])
                            { 
                                if (ExcelType == 13)
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;
                                        cbox.Text = modelNames[idx];
                                        cbox.Visible = true;
                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        if (idx == 7 ||idx==2)
                                        {

                                            cbox.Visible = false;
                                        }
                                        else
                                        {
                                            cbox.Visible = true;

                                        }

                                    }

                                }
                                else if (ExcelType == 19)
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;
                                        cbox.Text = modelNames[idx];
                                        cbox.Visible = true;

                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        cbox.Visible = true;
                                    }
                                }
                                else if (ExcelType == 20)
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;
                                        cbox.Text = modelNames[idx];
                                        cbox.Visible = true;

                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        cbox.Visible = true;
                                    }
                                }
                                else if (ExcelType == 18)
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;
                                        cbox.Text = modelNames[idx];
                                        cbox.Visible = true;
                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        cbox.Visible = false;
                                    }

                                }
                                else if(ExcelType == 21)
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;
                                        cbox.Text = modelNames[idx];
                                        cbox.Visible = true;
                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        cbox.Visible = true;
                                    }
                                }
                                else
                                {
                                    if (ctl is CheckBox)
                                    {
                                        CheckBox cbox = ctl as CheckBox;

                                        cbox.Visible = true;

                                    }
                                    else if (ctl is ComboBox)
                                    {
                                        ComboBox cbox = ctl as ComboBox;
                                        cbox.Visible = true;
                                    }

                                }
                            }
                            if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                
                                if (ExcelType ==1)
                                {
                                    if (idx == 2)
                                    {
                                        cbox.Visible = false;
                                    }
                                }

                                else if (ExcelType == 12)
                                {
                                    cbox.Visible = false;
                                }
                                else
                                {
                                    cbox.Visible = true;
                                }

                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetRoad2001ExcelShow(int ExcelType)
        {
            int[] showidxs = { 2, 3, 4 };
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetSHCity2013ExcelShow(int ExcelType, int PartType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 6 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 7:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                                if (idx == 0)
                                {
                                    if (ExcelType == 7)
                                    {
                                        cbox.Text = "综合报表区间距离";
                                    }
                                }
                                if (idx == 6)
                                {
                                    if (ExcelType == 0)
                                    {
                                        cbox.Text = "路面磨耗评价等级统计表";
                                    }
                                }
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                if (PartType == 0)
                                {
                                    if (ExcelType != 4)
                                    {
                                        cbox.Visible = true;
                                    }
                                }
                                else
                                {
                                    cbox.Visible = false;
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetGXExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 2, 3, 4, 5, 7, 8 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 1:
                    {
                        int[] tmpidx = { 0, 1 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;

                                if (idx == 0)
                                {
                                    if (ExcelType == 1)
                                    {
                                        cbox.Text = "综合报表区间距离";
                                    }
                                }
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
                    }
                }
            }
        }

        private void SetCQExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {

                        int[] tmpidx = { 2, 3, 4, 5, 7, 8 };

                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;
                                cbox.Visible = true;
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }

                    }
                }

            }
        }
        private void SetNCExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            List<string> modelNames = new List<string>();
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = {0,1, 2, 3, 4, 5, 7, 8,9,10,11 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;
                case 1:
                    {
                        int[] tmpidx = { 0, 1 ,2};
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("综合报表");
                        modelNames.Add("路面磨耗评价等级统计表");
                        modelNames.Add("农村公路技术状况数据汇总表");
                    }
                    break;
                case 2:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("空间定位数据");
                        modelNames.Add("沥青路面损坏自动化检测数据");
                        modelNames.Add("水泥路面损坏自动化检测数据");
                        modelNames.Add("路面平整度自动化检测数据");
                        modelNames.Add("路面平整度自动化检测原始数据");
                        modelNames.Add("DR破损率csv检测结果数据表格");
                        modelNames.Add("IRI平整度csv检测结果数据表格");
                        modelNames.Add("IRI平整度csv检测原始数据表格");
                        modelNames.Add("空间定位数据csv检测原始数据表格");
                        modelNames.Add("空间定位数据txt检测原始数据表格"); 
                    }
                    break;
                case 4:
                    {
                        int[] tmpidx = { 0, 1, 2, 3, 4,5 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("空间定位表格");
                        modelNames.Add("水泥病害统计表格");
                        modelNames.Add("沥青路面病害统计表格");
                        modelNames.Add("病害流水表表格"); 
                        modelNames.Add("导入模板(包含各指标及病害面积情况)");
                        modelNames.Add("重庆道路病害统计定制表");
                    }

                    break;
                case 5:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length]; 
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("合肥路况");
                    }
                    break;
                case 6:
                    {

                        int[] tmpidx = { 0 };

                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        if (_Setting.SelectDrawDis== 0)
                        {
                            modelNames.Add("贵州省农村公路路况检测照片交换模板");

                        }

                    }
                    break;
                case 7:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];

                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("合肥路况");
                    }
                    break;
                case 8:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];

                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("病害调绘表");
                    }
                    break;
                case 9:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];

                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("公里指标导入模板");
                    }
                    break;
                case 10:
                    {
                        int[] tmpidx = { 0, 1 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("平整度报表");
                        modelNames.Add("路面病害汇总报表");
                    }
                    break;
                case 11:
                    {
                        int[] tmpidx = { 0, 1 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("空间定位原始数据LBI文本");
                        modelNames.Add("平整度原始数据LP文本"); 
                    }
                    break;
                case 12:
                    {
                        int[] tmpidx = { 0, 1, 2};
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("农村公路 路况检测照片");
                        modelNames.Add("农村公路 检测数据明细表");
                        modelNames.Add("农村公路 路况检测轨迹");
                    }
                    break;
                case 13:
                    {
                        int[] tmpidx = { 0, 1, 2 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("低等级沥青路面破损率");
                        modelNames.Add("低等级水泥路面破损率");
                        modelNames.Add("低等级路面平整度");
                    }
                    break;

                case 14:
                    {
                        int[] tmpidx = { 0 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("上海惠浦: 病害表格"); 
                  
                    }
                    break;
                case 15:
                    {
                        int[] tmpidx = { 0,1,2,3,4,5,6 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                        modelNames.Add("附件4：交通安全设施排查表");
                        modelNames.Add("附件5：村民组数据交换模板2024");
                        modelNames.Add("附件6：路线轨迹数据交换模板");
                        modelNames.Add("附件7：安全隐患数据交换模板");
                        modelNames.Add("附件8：POI交换模板");
                        modelNames.Add("附件9：检测照片交换模板");
                        modelNames.Add("空间定位数据txt检测原始数据表格");

                    }
                    break;
                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            //模块化  自动化报表
                            if (ExcelType == 0)
                            {
                                if (ctl is CheckBox)
                                {
                                    CheckBox cbox = ctl as CheckBox;
                                    if (idx == 0)
                                    {
                                        cbox.Text = "(5211)路面破损评定汇总表";
                                    }
                                    if (idx ==1)
                                    {
                                        cbox.Text = "(5211)路面平整度评定汇总表";
                                    }
                                    if (idx == 9)
                                    {
                                        cbox.Text = "(5211)路面病害面积统计表";
                                    }
                                    if (idx == 10)
                                    {
                                        cbox.Text = "(5211)技术状况评定汇总表";
                                    }
                                    if (idx == 11)
                                    {
                                        cbox.Text = "(5211)技术状况评定明细表";
                                    }
                                    cbox.Visible = true;
                                }
                                else if (ctl is ComboBox)
                                {
                                    ComboBox cbox = ctl as ComboBox;
                                    if (idx == 0 || idx ==10 || idx == 1)
                                    {
                                        cbox.Visible = false;

                                    }
                                    else
                                    {
                                        cbox.Visible = true;

                                    }
                                }


                            }
                            else if (ExcelType == 2)
                            {
                                if (ctl is CheckBox)
                                {
                                    CheckBox cbox = ctl as CheckBox;
                                    cbox.Text = modelNames[idx];
                                    cbox.Visible = true;

                                }
                                else if (ctl is ComboBox)
                                {
                                    ComboBox cbox = ctl as ComboBox;
                                    if (idx == 4)
                                    {
                                        cbox.Visible = false;
                                    }
                                    else
                                    {
                                        cbox.Visible = true;

                                    }
                                }

                            }

                            else if (ExcelType == 8 || ExcelType == 11 || ExcelType == 12 || ExcelType == 15)
                            {
                                if (ctl is CheckBox)
                                {
                                    CheckBox cbox = ctl as CheckBox;
                                    cbox.Text = modelNames[idx];
                                    cbox.Visible = true;

                                }
                                else if (ctl is ComboBox)
                                {
                                    ComboBox cbox = ctl as ComboBox;
                                    cbox.Visible = false;
                                }
                            }
                           
                            else
                            {
                                if (ctl is CheckBox)
                                {
                                    CheckBox cbox = ctl as CheckBox;
                                    cbox.Text = modelNames[idx];
                                    cbox.Visible = true;

                                }
                                else if (ctl is ComboBox)
                                {
                                    ComboBox cbox = ctl as ComboBox;
                                    cbox.Visible = true;
                                }

                            } 
                            break;
                        }

                    }
                }

            }
        }

        private void SetHNExcelShow(int ExcelType)
        {
            int[] showidxs = null;
            switch (ExcelType)
            {
                case 0:
                    {
                        int[] tmpidx = { 2, 3, 4, 5, 7, 8, 9, 10, 11 };
                        showidxs = new int[tmpidx.Length];
                        Array.Copy(tmpidx, showidxs, tmpidx.Length);
                    }
                    break;

                default: break;
            }
            if (showidxs != null && showidxs.Length > 0)
            {
                foreach (Control ctl in tableLayoutPanel1.Controls)
                {
                    int idx = Convert.ToInt32(ctl.Tag.ToString());
                    for (int tt = 0; tt < showidxs.Length; ++tt)
                    {
                        if (idx == showidxs[tt])
                        {
                            //模块化  自动化报表
                            if (ExcelType == 0)
                            {
                                if (ctl is CheckBox)
                                {
                                    CheckBox cbox = ctl as CheckBox;
                                    if (idx == 9)
                                    {
                                        cbox.Text = "(5211)路面病害面积统计表";
                                    }
                                    if (idx == 10)
                                    {
                                        cbox.Text = "(5211)技术状况评定汇总表";
                                    }
                                    if (idx == 11)
                                    {
                                        cbox.Text = "(5211)技术状况评定明细表";
                                    }
                                    cbox.Visible = true;
                                }
                                else if (ctl is ComboBox)
                                {
                                    ComboBox cbox = ctl as ComboBox;
                                    cbox.Visible = true;
                                }
                            }

                            if (ctl is CheckBox)
                            {
                                CheckBox cbox = ctl as CheckBox;

                                cbox.Visible = true; 
                            }
                            else if (ctl is ComboBox)
                            {
                                ComboBox cbox = ctl as ComboBox;
                                cbox.Visible = true;
                            }
                            break;
                        }
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

        private void button_SetExcelMile_Click(object sender, EventArgs e)
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
