using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace XRDataProcess
{
    public partial class 采集打标列表 : Form
    {
        public event EventHandler EventUpdateProjectInfo;
        public event EventHandler EventJumpMark;
        public event EventHandler EventUpdateRoadPart;
        private ProjectInfo _ProjectInfo;
        private string _ProjPath;
        private List<MarkInfo> _MarkInfo = null;
        private Dictionary<MarkInfo, string> _MarkSources = new Dictionary<MarkInfo, string>();
        private bool _IsMarkClear;

        private List<DmiMile> _DmiMileList = null;

        private sealed class MarkRowSource
        {
            public string RawLine;
        }

        public 采集打标列表(ProjectInfo proinfo, string ppath)
        {
            InitializeComponent();
            _ProjectInfo = proinfo;
            _ProjPath = ppath;
            _MarkInfo = new List<MarkInfo>();
            _DmiMileList = new List<DmiMile>();
            LoadAllMark(false);
            LoadAllDmiMileCali(false);
        }

        /// <summary>
        /// 加载所有打标数据
        /// </summary>
        /// <param name="isclear"></param>
        public void LoadAllMark(bool isclear)
        {
            _IsMarkClear = isclear;
            _MarkInfo.Clear();
            _MarkSources.Clear();
            dataGridView_Mark.Rows.Clear();

            bool flagcalidmi = File.Exists(_ProjPath + "\\MileStoneCaliInfo.txt");
            string filename = _ProjPath + "\\RoadStatuMarkInfo.txt";
            if (File.Exists(filename))
            {
                string[] infos = File.ReadAllLines(filename);
                foreach (string info in infos)
                {
                    if (info.Length < 1)
                        continue;

                    MarkInfo tmark = new MarkInfo(info);
                    if (!flagcalidmi)
                    {
                        string[] s = info.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        if (s.Length > 1)
                        {
                            int dmi = Convert.ToInt32(s[2]);
                            tmark._Mile = _ProjectInfo.Dmi2Mile(dmi);
                            //tmark._Mile = Convert.ToInt32(s[0]);
                        }
                }
                _MarkInfo.Add(tmark);
                _MarkSources[tmark] = info;
                }
            }
            if (_ProjectInfo._Direction > 0)//升序
            {
                _MarkInfo.Sort(delegate(MarkInfo x, MarkInfo y) { return x._Mile.CompareTo(y._Mile); });
            }
            else if (_ProjectInfo._Direction < 0)//降序
            {
                _MarkInfo.Sort(delegate(MarkInfo x, MarkInfo y) { return y._Mile.CompareTo(x._Mile); });
            }
            foreach (MarkInfo mark in _MarkInfo)
            {
                object[] var = new object[3];
                var[0] = mark._Mile;
                var[1] = mark._Type;
                var[2] = mark._Info;
                int rowIndex = dataGridView_Mark.Rows.Add(var);
                // 原始行是删除时的唯一凭据，避免按排序后的行号误删其它打标。
                dataGridView_Mark.Rows[rowIndex].Tag = new MarkRowSource
                {
                    RawLine = _MarkSources[mark]
                };
            }
        }

        /// <summary>
        /// 加载所有校桩数据
        /// </summary>
        public void LoadAllDmiMileCali(bool isclear)
        {
            _DmiMileList.Clear();
            dataGridView1.Rows.Clear();

            string fname = _ProjPath + "\\MileStoneCaliInfo.txt";
            bool flagcalidmi = File.Exists(fname);
            if (!flagcalidmi)
                return;

            string[] sinfo = File.ReadAllLines(fname);
            foreach (string s in sinfo)
            {
                string[] str = s.Split(' ');
                if (str.Length > 1)
                {
                    _DmiMileList.Add(new DmiMile(int.Parse(str[0]), int.Parse(str[1])));
                }
            }

            if (_ProjectInfo._Direction > 0)//升序
            {
                _DmiMileList.Sort(delegate(DmiMile x, DmiMile y) { return x._Mile.CompareTo(y._Mile); });
            }
            else if (_ProjectInfo._Direction < 0)//降序
            {
                _DmiMileList.Sort(delegate(DmiMile x, DmiMile y) { return y._Mile.CompareTo(x._Mile); });
            }
            foreach (DmiMile mark in _DmiMileList)
            {
                object[] var = new object[2];
                var[0] = mark._Mile;
                var[1] = mark._Dmi;
                dataGridView1.Rows.Add(var);
            }
        }

        private void dataGridView_Mark_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                DataGridViewSelectedRowCollection selectrow = dataGridView_Mark.SelectedRows;
                int markmile = int.Parse(selectrow[0].Cells[0].Value.ToString());
                EventJumpMark(markmile, EventArgs.Empty);
            }
            catch { }
        }

        private void dataGridView_Mark_Sorted(object sender, EventArgs e)
        {
            _MarkInfo.Clear();
            foreach (DataGridViewRow row in dataGridView_Mark.Rows)
            {
                _MarkInfo.Add(new MarkInfo(row));
            }
        }

        private void button_add_Click(object sender, EventArgs e)
        {
            int mile = -1;
            int dmi = -1;
            if (tbox_mile.Text.Length > 0)
            {
                if(!int.TryParse(tbox_mile.Text.Replace("K", "").Replace("k", "").Replace("+", ""), out mile))
                {
                    MessageBox.Show("输入的【桩号】不合法，请重新输入！\r\n注意格式：K0+000");
                    return;
                }
            }
            if (tbox_dmi.Text.Length > 0)
            {
                if(!int.TryParse(tbox_dmi.Text.Replace("K", "").Replace("k", "").Replace("+", ""), out dmi))
                {
                    MessageBox.Show("输入的【里程】不合法，请重新输入！\r\n注意格式：K0+000");
                    return;
                }
            }

            if (tbox_mark.Text.Length == 0)
            {
                MessageBox.Show("【标记】内容不能为空，请重新输入！");
                return;
            }

            if (tbox_mile.Text.Length > 0 && tbox_dmi.Text.Length > 0)
            {
                MessageBox.Show("只能输入【里程】或【桩号】，请勿同时输入！");
                return;
            }

            if (mile > 0)
            {
                if ((_ProjectInfo._Direction > 0 && (mile < _ProjectInfo._StartMile || mile > _ProjectInfo._EndMile))
                    || (_ProjectInfo._Direction < 0 && (mile > _ProjectInfo._StartMile || mile < _ProjectInfo._EndMile)))
                {
                    MessageBox.Show("输入的【桩号】不在工程区间范围内，请重新输入！");
                    return;
                }
                dmi = _ProjectInfo.Mile2Dmi(mile);
                AddRoadUnit(mile, dmi, tbox_mark.Text, comboBox1.SelectedIndex);
            }
            else if (dmi > 0)
            {
                if (dmi > _ProjectInfo._EndDmi)
                {
                    MessageBox.Show("输入的【里程】不在工程区间范围内，请重新输入！");
                    return;
                }
                mile = _ProjectInfo.Dmi2Mile(dmi);
                AddRoadUnit(mile, dmi, tbox_mark.Text, comboBox1.SelectedIndex);
            }
        }

        /// <summary>
        /// 添加一行打标
        /// </summary>
        /// <param name="mile">桩号</param>
        /// <param name="dmi">里程</param>
        /// <param name="info">打标记录</param>
        /// <param name="type">打标类型，0-路面单元，1-路面情况</param>
        private void AddRoadUnit(int mile, int dmi, string info, int type)
        {
            foreach (MarkInfo mark in _MarkInfo)
            {
                if (mark._Type == "路面单元")
                {
                    if (mark._Mile == mile)
                    {
                        MessageBox.Show(string.Format("已经存在【K{0}+{1}】的划分单元！", mile / 1000, mile % 1000));
                        return;
                    }
                }
                if (mark._Type == "路面情况")
                {
                    if (mark._Mile == mile)
                    {
                        MessageBox.Show(string.Format("已经存在【K{0}+{1}】的路面情况！", mile / 1000, mile % 1000));
                        return;
                    }
                }
            }

            string fpath = _ProjPath + "\\RoadStatuMarkInfo.txt";
            List<string> infolist = new List<string>();
            string[] infos = null;
            if (File.Exists(fpath))
            {
                infos = File.ReadAllLines(fpath);
                infolist.AddRange(infos);
            }
            if (type == 0)
            {
                infolist.Add(string.Format("{0} {0} {1} 路面单元:{2}", mile, dmi, info));
            }
            else if (type == 1)
            {
                infolist.Add(string.Format("{0} {0} {1} 路面情况:{2}", mile, dmi, info));
            }
            infos = infolist.ToArray();
            File.WriteAllLines(fpath, infos, Encoding.UTF8);

            LoadAllMark(true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            批量提示窗口 tbox = new 批量提示窗口();
            tbox.ShowDialog();
            if (!tbox._IsOK)
                return;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel文件|*.xlsx|Excel文件|*.xls";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.Cursor = Cursors.WaitCursor;
                string fName = openFileDialog.FileName;
                MSExcel.Application excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };
                MSExcel.Workbook workbook = excelApp.Workbooks.Open(fName, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MSExcel.Worksheet excelsheet = workbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                int excelrow = GlobalExcel.judegeusedrow(excelsheet, 3, 1);
                if (excelrow > 0)
                {
                    MSExcel.Range excelrange = excelsheet.get_Range(String.Format("A1:D{0}", excelrow));
                    object[,] prj = (object[,])excelrange.Value2;
                    for (int i = 1; i <= excelrow; ++i)
                    {
                        if ((prj[i, 1] == null && prj[i, 2] == null)
                            || (prj[i, 1] != null && prj[i, 2] != null))
                            continue;

                        int mile = 0;
                        int dmi = 0;

                        if (prj[i, 3] == null)
                            continue;
                        int type = 1;
                        string typestr = prj[i, 3].ToString();
                        if (typestr == "路面情况")
                            type = 1;
                        else if (typestr == "路面单元")
                            type = 0;
                        else
                            continue;

                        if (prj[i, 4] == null)
                            continue;
                        string info = prj[i, 4].ToString();

                        if (prj[i, 1] != null)
                        {
                            string milestr = prj[i, 1].ToString().Replace("K", "").Replace("k", "").Replace("+", "");
                            try
                            {
                                mile = Convert.ToInt32(milestr);
                            }
                            catch (System.Exception ex)
                            {
                                continue;
                            }
                            if ((_ProjectInfo._Direction > 0 && (mile < _ProjectInfo._StartMile || mile > _ProjectInfo._EndMile))
                                || (_ProjectInfo._Direction < 0 && (mile > _ProjectInfo._StartMile || mile < _ProjectInfo._EndMile)))
                                continue;
                            dmi = _ProjectInfo.Mile2Dmi(mile);
                        }
                        if (prj[i, 2] != null)
                        {
                            string dmistr = prj[i, 2].ToString().Replace("K", "").Replace("k", "").Replace("+", "");
                            try
                            {
                                dmi = Convert.ToInt32(dmistr);
                            }
                            catch (System.Exception ex)
                            {
                                continue;
                            }
                            if (dmi < 0 || dmi > _ProjectInfo._EndDmi)
                                continue;
                            mile = _ProjectInfo.Dmi2Mile(dmi);
                        }

                        AddRoadUnit(mile, dmi, info, type);
                    }
                }
                workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                int generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                excelApp.Quit();
                this.Cursor = Cursors.Default;
            }
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                DataGridViewSelectedRowCollection selectrow = dataGridView1.SelectedRows;
                int markmile = int.Parse(selectrow[0].Cells[0].Value.ToString());
                EventJumpMark(markmile, EventArgs.Empty);
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int mile = -1;
            int dmi = -1;
            if (textBox1.Text.Length > 0)
            {
                if (!int.TryParse(textBox1.Text.Replace("K", "").Replace("k", "").Replace("+", ""), out mile))
                {
                    MessageBox.Show("输入的【桩号】不合法，请重新输入！\r\n注意格式：K0+000");
                    return;
                }
            }
            if (textBox2.Text.Length > 0)
            {
                if (!int.TryParse(textBox2.Text.Replace("K", "").Replace("k", "").Replace("+", ""), out dmi))
                {
                    MessageBox.Show("输入的【里程】不合法，请重新输入！\r\n注意格式：K0+000");
                    return;
                }
            }

            if (textBox1.Text.Length == 0 || textBox2.Text.Length == 0)
            {
                MessageBox.Show("输入的【里程】和【桩号】不能为空！");
                return;
            }

            AddRoadCali(mile, dmi);
        }

        private void AddRoadCali(int mile, int dmi)
        {
            foreach (DmiMile tdmimile in _DmiMileList)
            {
                if (tdmimile._Mile == mile || tdmimile._Dmi == dmi)
                {
                    MessageBox.Show(string.Format("已经存在该校桩数据，请重新输入！"));
                    return;
                }
            }
            _DmiMileList.Add(new DmiMile(dmi, mile));

            object[] var = new object[2];
            var[0] = mile;
            var[1] = dmi;
            dataGridView1.Rows.Add(var);
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            _DmiMileList.Clear();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                _DmiMileList.Add(new DmiMile(row));
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (_DmiMileList.Count < 1)
                return;

            // 将校桩数据存入文件
            List<string> strs = new List<string>();
            foreach (DmiMile tdmimile in _DmiMileList)
            {
                strs.Add(tdmimile.ToString());
            }
            string fname = _ProjPath + @"\MileStoneCaliInfo.txt";
            File.WriteAllLines(fname, strs, Encoding.UTF8);

            // 重新加载工程
            MessageBox.Show("修改校桩文件成功，即将重置该工程！");
            EventUpdateProjectInfo(null, null);
        }

        private void 采集打标列表_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void dataGridView_Mark_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button_deleteMark_Click(object sender, EventArgs e)
        {
            DataGridViewRow[] selected = dataGridView_Mark.SelectedRows.Cast<DataGridViewRow>().ToArray();
            if (selected.Length == 0) return;
            if (MessageBox.Show("确定删除所选的 " + selected.Length + " 条打标吗？", "删除打标", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            string fileName = Path.Combine(_ProjPath, "RoadStatuMarkInfo.txt");
            List<string> lines = File.Exists(fileName) ? File.ReadAllLines(fileName, Encoding.UTF8).ToList() : new List<string>();
            foreach (DataGridViewRow row in selected)
            {
                MarkRowSource source = row.Tag as MarkRowSource;
                if (source == null) continue;
                int index = lines.FindIndex(line => string.Equals(line.Trim(), source.RawLine.Trim(), StringComparison.Ordinal));
                if (index >= 0) lines.RemoveAt(index);
            }
            WriteLinesAtomically(fileName, lines);
            LoadAllMark(true);
            if (EventUpdateRoadPart != null) EventUpdateRoadPart(null, EventArgs.Empty);
        }

        private void button_deleteCali_Click(object sender, EventArgs e)
        {
            DataGridViewRow[] selected = dataGridView1.SelectedRows.Cast<DataGridViewRow>().ToArray();
            if (selected.Length == 0) return;
            if (MessageBox.Show("确定删除所选的 " + selected.Length + " 条校桩吗？保存并重载工程后生效。", "删除校桩", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            foreach (DataGridViewRow row in selected)
            {
                int mile = Convert.ToInt32(row.Cells[0].Value);
                int dmi = Convert.ToInt32(row.Cells[1].Value);
                _DmiMileList.RemoveAll(item => item._Mile == mile && item._Dmi == dmi);
                dataGridView1.Rows.Remove(row);
            }
        }

        private static void WriteLinesAtomically(string path, IEnumerable<string> lines)
        {
            string temporaryPath = path + ".tmp";
            File.WriteAllLines(temporaryPath, lines, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temporaryPath, path, path + ".bak", true);
            else File.Move(temporaryPath, path);
        }
    }
}

