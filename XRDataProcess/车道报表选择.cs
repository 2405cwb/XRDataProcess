using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OperateIniFile;
using Framework.Other;

namespace XRDataProcess
{
    public partial class 车道报表选择 : Form
    {
        public 车道报表选择()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 上行的报表文件：车道、指标
        /// </summary>
        public string[][] _UpXlsFiles = null;

        /// <summary>
        /// 下行的报表文件：车道、指标
        /// </summary>
        public string[][] _DownXlsFiles = null;

        /// <summary>
        /// 上行的公里报表文件：车道、指标
        /// </summary>
        public string[][] _UpXlsFilesKM = null;

        /// <summary>
        /// 下行的公里报表文件：车道、指标
        /// </summary>
        public string[][] _DownXlsFilesKM = null;

        /// <summary>
        /// 是否处理
        /// </summary>
        public bool _IsOK = false;

        /// <summary>
        /// 上下行合并类型，0-全幅，1-上行半幅，2-下行半幅
        /// </summary>
        public int _MergeType = 0;

        /// <summary>
        /// 合并报表的类型，0-四川振兴模板1
        /// </summary>
        public int _ExcelType = 0;
        
        /// <summary>
        /// 指标是否需要合并：PCI\RQI\RDI\PBI\PWI\SMTD\PQI
        /// </summary>
        public MergeIndexInfo[] _MergeIdxInfo = new MergeIndexInfo[7];

        /// <summary>
        /// 合并报表放置的文件夹路径
        /// </summary>
        public string _OutputPath = null;

        void bbox_Click_Up(object sender, EventArgs e)
        {
            Button bbox = (Button)sender;
            int idx = Convert.ToInt16(bbox.Tag);

            TextBox tbox = null;
            foreach (Control ctl in tableLayoutPanel_Up.Controls)
            {
                if (ctl is TextBox)
                {
                    tbox = ctl as TextBox;
                    int idxt = Convert.ToInt16(tbox.Tag);
                    if (idx == idxt)
                    {
                        break;
                    }
                }
            }

            if (tbox == null)
                return;

            tbox.Text = "";

            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择车道报表文件夹：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                tbox.Text = fd.SelectedPath;
                toolTip1.SetToolTip(tbox, tbox.Text);

                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt == idx + 1)
                        {
                            bbox.Enabled = true;
                        }
                        else if (idxt >= idx + 1)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
            else
            {
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt > idx)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
        }

        void bbox_Click_Down(object sender, EventArgs e)
        {
            Button bbox = (Button)sender;
            int idx = Convert.ToInt16(bbox.Tag);

            TextBox tbox = null;
            foreach (Control ctl in tableLayoutPanel_Down.Controls)
            {
                if (ctl is TextBox)
                {
                    tbox = ctl as TextBox;
                    int idxt = Convert.ToInt16(tbox.Tag);
                    if (idx == idxt)
                    {
                        break;
                    }
                }
            }

            if (tbox == null)
                return;

            tbox.Text = "";

            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择车道报表文件夹：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                tbox.Text = fd.SelectedPath;
                toolTip1.SetToolTip(tbox, tbox.Text);

                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt == idx + 1)
                        {
                            bbox.Enabled = true;
                        }
                        else if (idxt >= idx + 1)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
            else
            {
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is Button)
                    {
                        bbox = ctl as Button;
                        int idxt = Convert.ToInt16(bbox.Tag);
                        if (idxt > idx)
                        {
                            bbox.Enabled = false;
                        }
                    }
                }
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is TextBox)
                    {
                        tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        if (idxt > idx)
                        {
                            tbox.Text = string.Empty;
                            toolTip1.SetToolTip(tbox, string.Empty);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获取用户专属的布局文件完整路径（%LocalAppData%）
        /// </summary>
        private string GetUserLayoutPath()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appFolder = Path.Combine(localAppData, "夕睿光电", "内业数据处理软件");
            Directory.CreateDirectory(appFolder);               // 确保目录存在
            return Path.Combine(appFolder, _layoutFileName);
        }
        private string _layoutFileName = "MergeExcel.ini";

        private void button_Yes_Click(object sender, EventArgs e)
        {
            MessageBox.Show("提示：请确认上下行所有车道的单元区间桩号相同！若不相同均以一车道单元区间桩号为多车道合并后的单元区间桩号！");
            string layoutPath = GetUserLayoutPath();
            if (!File.Exists(layoutPath))
            {
                File.Copy(System.Windows.Forms.Application.StartupPath + @"\MergeExcel.ini", layoutPath);
            }

            IniFiles inisetting = new IniFiles(layoutPath);
            string[] upxlsdirpaths = new string[tableLayoutPanel_Up.Controls.Count];
            string[] downxlsdirpaths = new string[tableLayoutPanel_Down.Controls.Count];
            List<string> upxlsdirpathlist = new List<string>();
            List<string> downxlsdirpathlist = new List<string>();

            if (radioButton_full.Checked || radioButton_halfUp.Checked)
            {
                foreach (Control ctl in tableLayoutPanel_Up.Controls)
                {
                    if (ctl is TextBox)
                    {
                        TextBox tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        upxlsdirpaths[idxt] = tbox.Text;
                    }
                }
                for (int i = 0; i < upxlsdirpaths.Length; ++i)
                {
                    if (upxlsdirpaths[i] != string.Empty)
                    {
                        upxlsdirpathlist.Add(upxlsdirpaths[i]);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            if (radioButton_full.Checked || radioButton_halfDown.Checked)
            {
                foreach (Control ctl in tableLayoutPanel_Down.Controls)
                {
                    if (ctl is TextBox)
                    {
                        TextBox tbox = ctl as TextBox;
                        int idxt = Convert.ToInt16(tbox.Tag);
                        downxlsdirpaths[idxt] = tbox.Text;
                    }
                }
                for (int i = 0; i < downxlsdirpaths.Length; ++i)
                {
                    if (downxlsdirpaths[i] != string.Empty)
                    {
                        downxlsdirpathlist.Add(downxlsdirpaths[i]);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            if (radioButton_full.Checked)
            {
                if (upxlsdirpathlist.Count != downxlsdirpathlist.Count)
                {
                    MessageBox.Show(string.Format("已勾选【全幅】，上行选择了前{0}个车道，下行选择了前{1}个车道，上下行车道数量不一致，请检查！",
                        upxlsdirpathlist.Count, downxlsdirpathlist.Count));
                    return;
                }

                if (upxlsdirpathlist.Count == 0 || downxlsdirpathlist.Count == 0)
                {
                    MessageBox.Show("请选择全幅的上下行车道报表！");
                    return;
                }
            }
            else if(radioButton_halfUp.Checked)
            {
                if (upxlsdirpathlist.Count == 0)
                {
                    MessageBox.Show(string.Format("已勾选【上行半幅】，没有选择上行半幅车道报表，请检查！",
                        upxlsdirpathlist.Count, downxlsdirpathlist.Count));
                    return;
                }
            }
            else if(radioButton_halfDown.Checked)
            {
                if(downxlsdirpathlist.Count == 0)
                {
                    MessageBox.Show(string.Format("已勾选【下行半幅】，没有选择下行半幅车道报表，请检查！",
                        upxlsdirpathlist.Count, downxlsdirpathlist.Count));
                    return;
                }
            }

            if (radioButton_full.Checked || radioButton_halfUp.Checked)
            {
                _UpXlsFiles = new string[upxlsdirpathlist.Count][];
                for (int i = 0; i < upxlsdirpathlist.Count; ++i)
                {
                    _UpXlsFiles[i] = new string[_MergeIdxInfo.Length];
                    for (int k = 0; k < _MergeIdxInfo.Length; ++k)
                    {
                        _UpXlsFiles[i][k] = null;
                    }
                }

                _UpXlsFilesKM = new string[upxlsdirpathlist.Count][];
                for (int i = 0; i < upxlsdirpathlist.Count; ++i)
                {
                    _UpXlsFilesKM[i] = new string[_MergeIdxInfo.Length];
                    for (int k = 0; k < _MergeIdxInfo.Length; ++k)
                    {
                        _UpXlsFilesKM[i][k] = null;
                    }
                }
            }

            if (radioButton_full.Checked || radioButton_halfDown.Checked)
            {
                _DownXlsFiles = new string[downxlsdirpathlist.Count][];
                for (int i = 0; i < downxlsdirpathlist.Count; ++i)
                {
                    _DownXlsFiles[i] = new string[_MergeIdxInfo.Length];
                    for (int k = 0; k < _MergeIdxInfo.Length; ++k)
                    {
                        _DownXlsFiles[i][k] = null;
                    }
                }

                _DownXlsFilesKM = new string[downxlsdirpathlist.Count][];
                for (int i = 0; i < downxlsdirpathlist.Count; ++i)
                {
                    _DownXlsFilesKM[i] = new string[_MergeIdxInfo.Length];
                    for (int k = 0; k < _MergeIdxInfo.Length; ++k)
                    {
                        _DownXlsFilesKM[i][k] = null;
                    }
                }
            }

            bool isIdxChecked = false;
            foreach (Control ctl in tableLayoutPanel_Index.Controls)
            {
                if (ctl is CheckBox)
                {
                    CheckBox cbox = ctl as CheckBox;
                    int tidx = Convert.ToInt16(cbox.Tag);
                    _MergeIdxInfo[tidx]._IsMergeIdx = cbox.Checked;
                    inisetting.WriteBool("MergeInfo", "IndexChecked" + tidx.ToString(), _MergeIdxInfo[tidx]._IsMergeIdx);
                    if (cbox.Checked)
                    {
                        isIdxChecked = true;
                    }
                }
                else if (ctl is ComboBox)
                {
                    ComboBox cbox = ctl as ComboBox;
                    int tidx = Convert.ToInt16(cbox.Tag);
                    _MergeIdxInfo[tidx]._OriUnitLen = Convert.ToInt32(cbox.Text);
                    inisetting.WriteInteger("MergeInfo", "OriUnitLen" + tidx.ToString(), _MergeIdxInfo[tidx]._OriUnitLen);
                }
            }

            if (!isIdxChecked)
            {
                MessageBox.Show("没有选择要合并的技术指标，请检查！");
                return;
            }

            bool IsXlsOK = true;
            if (radioButton_full.Checked || radioButton_halfUp.Checked)
            {
                for (int di = 0; di < upxlsdirpathlist.Count; ++di)
                {
                    DirectoryInfo tdir = new DirectoryInfo(upxlsdirpathlist[di]);
                    FileInfo[] tfiles = tdir.GetFiles("*.xlsx");

                    for (int i = 0; i < _MergeIdxInfo.Length; ++i)
                    {
                        if (!_MergeIdxInfo[i]._IsMergeIdx)
                        {
                            continue;
                        }
                        bool isHasFile = false;
                        bool isHasFileKM = false;
                        string subfname = string.Format("_{0}_{1}m.xlsx", _MergeIdxInfo[i]._ExcelSubName, _MergeIdxInfo[i]._OriUnitLen);
                        string subfnameKM = string.Format("_{0}_1000m.xlsx", _MergeIdxInfo[i]._ExcelSubName);
                        foreach (FileInfo tfile in tfiles)
                        {
                            if (tfile.Name.Contains(subfname))
                            {
                                isHasFile = true;
                                _UpXlsFiles[di][i] = tfile.FullName;
                            }
                             if (tfile.Name.Contains(subfnameKM))
                            {
                                isHasFileKM = true;
                                _UpXlsFilesKM[di][i] = tfile.FullName;
                            }
                            if (isHasFile && isHasFileKM)
                            {
                                break;
                            }
                        }
                        if (!isHasFile)
                        {
                            MessageBox.Show(string.Format("没有找到【{0}】的报表【{1}】，请检查导出的原始报表是否齐全！",
                                upxlsdirpathlist[di], subfname));
                            IsXlsOK = false;
                        }
                        if (!isHasFileKM)
                        {
                            MessageBox.Show(string.Format("没有找到【{0}】的报表【{1}】，请检查导出的原始报表是否齐全！",
                                upxlsdirpathlist[di], subfnameKM));
                            IsXlsOK = false;
                        }
                    }
                }
            }

            if (radioButton_full.Checked || radioButton_halfDown.Checked)
            {
                for (int di = 0; di < downxlsdirpathlist.Count; ++di)
                {
                    DirectoryInfo tdir = new DirectoryInfo(downxlsdirpathlist[di]);
                    FileInfo[] tfiles = tdir.GetFiles("*.xlsx");

                    for (int i = 0; i < _MergeIdxInfo.Length; ++i)
                    {
                        if (!_MergeIdxInfo[i]._IsMergeIdx)
                        {
                            continue;
                        }

                        bool isHasFile = false;
                        bool isHasFileKM = false;
                        string subfname = string.Format("_{0}_{1}m.xlsx", _MergeIdxInfo[i]._ExcelSubName, _MergeIdxInfo[i]._OriUnitLen);
                        string subfnameKM = string.Format("_{0}_1000m.xlsx", _MergeIdxInfo[i]._ExcelSubName);
                        foreach (FileInfo tfile in tfiles)
                        {
                            if (tfile.Name.Contains(subfname))
                            {
                                isHasFile = true;
                                _DownXlsFiles[di][i] = tfile.FullName;
                            }
                             if (tfile.Name.Contains(subfnameKM))
                            {
                                isHasFileKM = true;
                                _DownXlsFilesKM[di][i] = tfile.FullName;
                            }
                            if (isHasFile && isHasFileKM)
                            {
                                break;
                            }
                        }
                        if (!isHasFile)
                        {
                            MessageBox.Show(string.Format("没有找到【{0}】的报表【{1}】，请检查导出的原始报表是否齐全！",
                                downxlsdirpathlist[di], subfname));
                            IsXlsOK = false;
                        }
                        if (!isHasFileKM)
                        {
                            MessageBox.Show(string.Format("没有找到【{0}】的报表【{1}】，请检查导出的原始报表是否齐全！",
                                downxlsdirpathlist[di], subfnameKM));
                            IsXlsOK = false;
                        }
                    }
                }
            }

            if (!IsXlsOK)
            {
                return;
            }

            if (radioButton_full.Checked)
            {
                _MergeType = 0;
            }
            else if (radioButton_halfUp.Checked)
            {
                _MergeType = 1;
            }
            else if (radioButton_halfDown.Checked)
            {
                _MergeType = 2;
            }

            foreach (Control ctl in groupBox_ExcelType.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton rbox = ctl as RadioButton;
                    if (rbox.Checked)
                    {
                        _ExcelType = Convert.ToInt16(rbox.Tag);
                        break;
                    }
                }
            }
            
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "请选择合并报表放置文件夹：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                _OutputPath = fd.SelectedPath;
                _IsOK = true;
            }
            else
            {
                _IsOK = false;
            }

            foreach (Control ctl in tableLayoutPanel_Up.Controls)
            {
                if (ctl is TextBox)
                {
                    TextBox tbox = ctl as TextBox;
                    int idx = Convert.ToInt16(tbox.Tag);
                    inisetting.WriteString("MergeInfo", "UpPath" + idx.ToString(), tbox.Text);
                }
            }

            foreach (Control ctl in tableLayoutPanel_Down.Controls)
            {
                if (ctl is TextBox)
                {
                    TextBox tbox = ctl as TextBox;
                    int idx = Convert.ToInt16(tbox.Tag);
                    inisetting.WriteString("MergeInfo", "DownPath" + idx.ToString(), tbox.Text);
                }
            }

            foreach (Control ctl in tableLayoutPanel_Index.Controls)
            {
                if (ctl is FlowLayoutPanel)
                {
                    FlowLayoutPanel fbox = ctl as FlowLayoutPanel;
                    int idx = Convert.ToInt16(fbox.Tag);
                    foreach (Control tctl in fbox.Controls)
                    {
                        if (tctl is RadioButton)
                        {
                            RadioButton rbox = tctl as RadioButton;
                            if (rbox.Checked)
                            {
                                _MergeIdxInfo[idx]._ThreshType = Convert.ToInt16(rbox.Tag);
                            }
                        }
                        else if(tctl is TextBox)
                        {
                            TextBox tbox = tctl as TextBox;
                            _MergeIdxInfo[idx]._ThreshVal = Convert.ToDouble(tbox.Text);
                        }
                    }
                    inisetting.WriteInteger("MergeInfo", "IndexThreshType" + idx.ToString(), _MergeIdxInfo[idx]._ThreshType);
                    inisetting.WriteString("MergeInfo", "IndexThreshVal" + idx.ToString(), _MergeIdxInfo[idx]._ThreshVal.ToString());
                }
            }

            foreach (Control ctl in groupBox_ExcelType.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton rbox = ctl as RadioButton;
                    if (rbox.Checked)
                    {
                        inisetting.WriteInteger("MergeInfo", "MergeExcelType", Convert.ToInt16(rbox.Tag));
                        break;
                    }
                }
            }

            this.Close();
        }

        private void button_Esc_Click(object sender, EventArgs e)
        {
            _IsOK = false;
            this.Close();
        }

        private void 车道报表选择_Load(object sender, EventArgs e)
        {
            string layoutPath = GetUserLayoutPath();
            if (!File.Exists(layoutPath))
            {
                File.Copy(System.Windows.Forms.Application.StartupPath + @"\MergeExcel.ini", layoutPath);
            }

            IniFiles inisetting = new IniFiles(layoutPath);
            foreach (Control ctl in tableLayoutPanel_Up.Controls)
            {
                if (ctl is Button)
                {
                    Button bbox = ctl as Button;
                    bbox.Click += new EventHandler(bbox_Click_Up);
                }
                else if (ctl is TextBox)
                {
                    TextBox tbox = ctl as TextBox;
                    int idx = Convert.ToInt32(tbox.Tag);
                    tbox.Text = inisetting.ReadString("MergeInfo", "UpPath"+idx.ToString(), "");
                    if (!Directory.Exists(tbox.Text))
                    {
                        tbox.Text = "";
                    }
                }
            }

            foreach (Control ctl in tableLayoutPanel_Down.Controls)
            {
                if (ctl is Button)
                {
                    Button bbox = ctl as Button;
                    bbox.Click += new EventHandler(bbox_Click_Down);
                }
                else if (ctl is TextBox)
                {
                    TextBox tbox = ctl as TextBox;
                    int idx = Convert.ToInt32(tbox.Tag);
                    tbox.Text = inisetting.ReadString("MergeInfo", "DownPath" + idx.ToString(), "");
                    if (!Directory.Exists(tbox.Text))
                    {
                        tbox.Text = "";
                    }
                }
            }
            
            for (int i = 0; i < _MergeIdxInfo.Length; ++i )
            {
                _MergeIdxInfo[i] = new MergeIndexInfo(
                    inisetting.ReadBool("MergeInfo", "IndexChecked" + i.ToString(), false),
                    inisetting.ReadInteger("MergeInfo", "IndexExcelStartRow" + i.ToString(), 4),
                    inisetting.ReadString("MergeInfo", "IndexExcelMergeName" + i.ToString(), "").Replace("\0", ""),
                    Convert.ToDouble(inisetting.ReadString("MergeInfo", "IndexThreshVal" + i.ToString(), "0")),
                    inisetting.ReadInteger("MergeInfo", "IndexThreshType" + i.ToString(), 0),
                    inisetting.ReadString("MergeInfo", "IndexExcelSubName" + i.ToString(), "").Replace("\0", ""),
                    inisetting.ReadInteger("MergeInfo", "OriUnitLen" + i.ToString(), 0));
            }

            foreach (Control ctl in tableLayoutPanel_Index.Controls)
            {
                if (ctl is CheckBox)
                {
                    CheckBox cbox = ctl as CheckBox;
                    int idx = Convert.ToInt16(cbox.Tag);
                    cbox.Checked = _MergeIdxInfo[idx]._IsMergeIdx;
                }
                else if (ctl is FlowLayoutPanel)
                {
                    FlowLayoutPanel fbox = ctl as FlowLayoutPanel;
                    int idx = Convert.ToInt16(fbox.Tag);
                    foreach (Control tctl in fbox.Controls)
                    {
                        if (tctl is RadioButton)
                        {
                            RadioButton rbox = tctl as RadioButton;
                            if (_MergeIdxInfo[idx]._ThreshType == Convert.ToInt16(rbox.Tag))
                            {
                                rbox.Checked = true;
                            }
                        }
                        else if (tctl is TextBox)
                        {
                            TextBox tbox = tctl as TextBox;
                            tbox.Text = _MergeIdxInfo[idx]._ThreshVal.ToString();
                        }
                    }                    
                }
                else if (ctl is ComboBox)
                {
                    ComboBox cbox = ctl as ComboBox;
                    int idx = Convert.ToInt16(cbox.Tag);
                    cbox.Text = _MergeIdxInfo[idx]._OriUnitLen.ToString();
                }
            }

            int mergeexceltype = inisetting.ReadInteger("MergeInfo", "MergeExcelType", 0);
            foreach (Control ctl in groupBox_ExcelType.Controls)
            {
                if (ctl is RadioButton)
                {
                    RadioButton tbox = ctl as RadioButton;
                    tbox.CheckedChanged += new EventHandler(Index_CheckedChanged);
                    int tidx = Convert.ToInt16(tbox.Tag);
                    if (tidx == mergeexceltype)
                    {
                        tbox.Checked = true;
                    }
                }
            }
        }

        void Index_CheckedChanged(object sender, EventArgs e)
        {
            bool[] idxchk = new bool[_MergeIdxInfo.Length];
            for (int i = 0; i < _MergeIdxInfo.Length; ++i)
            {
                idxchk[i] = false;
            }

            RadioButton rbox = (RadioButton)sender;
            int idx = Convert.ToInt16(rbox.Tag);
            switch(idx)
            {
                case 0: idxchk[1] = true; idxchk[5] = true; break;
                default: break;
            }

            foreach (Control ctl in tableLayoutPanel_Index.Controls)
            {
                if(ctl is CheckBox)
                {
                    CheckBox cbox = ctl as CheckBox;
                    int cidx = Convert.ToInt16(cbox.Tag);
                    cbox.Enabled = idxchk[cidx];
                    if (!idxchk[cidx])
                    {
                        cbox.Checked = idxchk[cidx];
                    }
                }
                else if (ctl is FlowLayoutPanel)
                {
                    FlowLayoutPanel fbox = ctl as FlowLayoutPanel;
                    int cidx = Convert.ToInt16(fbox.Tag);
                    fbox.Enabled = idxchk[cidx];
                }
            }
        }

        private void radioButton_full_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = true;
            groupBox_Down.Enabled = true;
        }

        private void radioButton_halfUp_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = true;
            groupBox_Down.Enabled = false;
        }

        private void radioButton_halfDown_CheckedChanged(object sender, EventArgs e)
        {
            groupBox_Up.Enabled = false;
            groupBox_Down.Enabled = true;
        }
    }
    public class MergeIndexInfo
    {
        public MergeIndexInfo()
        {
            _IsMergeIdx = false;
            _ExcelStartRow = 3;
            _ExcelMergeName = null;
            _ThreshVal = 0.0;
            _ThreshType = 0;
            _OriUnitLen = 10;
        }

        public MergeIndexInfo(bool ismergeidx, int excelstartrow, string excelmergename, double threshval, 
            int threshtype, string excelsubname, int oriunitlen)
        {
            _IsMergeIdx = ismergeidx;
            _ExcelStartRow = excelstartrow;
            _ExcelMergeName = excelmergename;
            _ThreshVal = threshval;
            _ThreshType = threshtype;
            _ExcelSubName = excelsubname;
            _OriUnitLen = oriunitlen;
        }

        /// <summary>
        /// 指标是否需要合并：PCI\RQI\RDI\PBI\PWI\SMTD\PQI
        /// </summary>
        public bool _IsMergeIdx = false;

        /// <summary>
        /// 报表的sheet1开始有内容的行数PCI\RQI\RDI\PBI\PWI\SMTD\PQI
        /// </summary>
        public int _ExcelStartRow = 3;

        /// <summary>
        /// 输出的合并报表的名称
        /// </summary>
        public string _ExcelMergeName = null;

        /// <summary>
        /// 对应基础指标合格值
        /// </summary>
        public double _ThreshVal = 0.0;

        /// <summary>
        /// 对应基础指标合格区间，0-大于等于合格值，1-小于等于合格值
        /// </summary>
        public int _ThreshType = 0;

        /// <summary>
        /// 原始车道报表的名称后缀
        /// </summary>
        public string _ExcelSubName = null;

        /// <summary>
        /// 要合并的原始单元区间长度
        /// </summary>
        public int _OriUnitLen = 10;
    }

}
