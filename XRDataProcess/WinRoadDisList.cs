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
using RuralPavementDetect;

namespace XRDataProcess
{
    public partial class WinRoadDisList : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        public event EventHandler EventUpdateDis;
        public event EventHandler EventJump2Dis;

        private ProjectInfo _ProjectInfo;
        private string _ProjPath;

        /// <summary>
        /// 图片浏览方式，0-连续，1-识别有病害，2-识别无病害，3-指定病害图像
        /// </summary>
        public static int _BrowserType = 0;

        public static string _BrowserDisName = null;

        /// <summary>
        /// true-小方格，false-大方框
        /// </summary>
        private bool _DisType = false;

        public WinRoadDisList(ProjectInfo proinfo, string ppath)
        {
            InitializeComponent();
            diseaseNum = 0;
            comboBox_BSType.SelectedIndex = 0;
            _ProjectInfo = proinfo;
            _ProjPath = ppath;
            if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.DegreeRoad2018)
                _DisType = true;
            else if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel)
                _DisType = true;
            else if (_Setting.SelectDrawDis == 1 && _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
                _DisType = true;
            else
                _DisType = false;
        }

        public void UpDateCurDis(object disinfo)
        {
            object[] var1 = new object[6];
            dataGridView_curdis.Rows.Clear();

            if (_DisType)
            {
                List<SmalRectDisease> DisInfoList = (List<SmalRectDisease>)(disinfo);
                foreach (SmalRectDisease tdis in DisInfoList)
                {

                    var1[0] = tdis.m_mile;
                    var1[1] = tdis.RoadDisType;
                    var1[4] = tdis.Area.ToString("0.000");
                    var1[5] = tdis.RoadType;
                    dataGridView_curdis.Rows.Add(var1);
                }
            }
            else
            {
                List<Disease> DisInfoList = (List<Disease>)(disinfo);
                foreach (Disease tdis in DisInfoList)
                {
                    var1[0] = tdis.m_mile;
                    var1[1] = tdis.RoadDisType;
                    var1[2] = tdis.calcheight.ToString("0.000");
                    var1[3] = tdis.calcwidth.ToString("0.000");
                    var1[4] = tdis.Area.ToString("0.000000");
                    var1[5] = tdis.RoadType;
                    dataGridView_curdis.Rows.Add(var1);
                }
            }
        }
        private int diseaseNum = 0;

        /// <summary>
        /// 0 当前病害列表
        /// 1 所有病害列表
        /// </summary>
        /// <param name="index"></param>
        public void LoadAllDis()
        {

            diseaseNum = 0;
            this.Cursor = Cursors.WaitCursor;

            dataGridView_Dislist.Rows.Clear();
            dataGridView_MLDisImgList.Rows.Clear();
            dataGridView_MLNoDisImgList.Rows.Clear();

            object[] var1New = new object[6];
            object[] var2 = new object[3];
            string[] ImgMilestr = File.ReadAllLines(_ProjPath + "\\RoadImg\\Camera0\\Road2Mile.txt");

            foreach (string infostr in ImgMilestr)
            {

                string[] s = null;
                int curmile = 0;
                string curimgname = null;

                try
                {
                    s = infostr.Split(' ');
                    curmile = (int)Math.Round(Convert.ToDouble(s[0]));
                    curimgname = s[1];
                }
                catch (System.Exception ex)
                {
                    continue;
                }

                var2[1] = curmile;
                var2[2] = curimgname;

                string disfile = string.Empty;
                if (_DisType)
                {

                    disfile = string.Format("{0}\\RoadImg\\Camera0{1}_PartClass.txt", _ProjPath, curimgname);
                    if (File.Exists(disfile))
                    {
                        string[] dises = File.ReadAllLines(disfile);
                        foreach (string dis in dises)
                        {
                            SmalRectDisease tdis = null;
                            try
                            {
                                tdis = new SmalRectDisease(dis, curmile);
                                if (!tdis.isDiseaseOK)
                                {
                                    continue;
                                }
                            }
                            catch
                            {
                                continue;
                            }

                            if (tdis.Area > 0)
                            {
                                if (tdis.FirstRectNum > (_RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2))
                                {
                                    tdis.m_mile += _ProjectInfo._Direction;
                                }
                                var1New[0] = diseaseNum++;
                                var1New[1] = tdis.m_mile;
                                var1New[2] = tdis.RoadDisType;
                                var1New[5] = tdis.Area.ToString("0.00");

                                dataGridView_Dislist.Invoke(new Action(() =>
                                {
                                    dataGridView_Dislist.Rows.Add(var1New);
                                })
                       );


                            }
                        }

                        dataGridView_MLDisImgList.Invoke(new Action(() =>
                        {
                            dataGridView_MLDisImgList.Rows.Add(var2);
                        })
                     );
                   
                    }
                    else
                    {
                        dataGridView_MLNoDisImgList.Invoke(new Action(() =>
                        {
                            dataGridView_MLNoDisImgList.Rows.Add(var2);
                        })
               );
                       
                    }
                }
                else
                {

                    disfile = string.Format("{0}\\RoadImg\\Camera0{1}.txt", _ProjPath, curimgname);
                    if (File.Exists(disfile))
                    {

                        string[] dises = File.ReadAllLines(disfile);
                        foreach (string dis in dises)
                        {
                            Disease tdis = null;
                            try
                            {
                                tdis = new Disease(dis, curmile);
                            }
                            catch
                            {

                                continue;
                            }
                            if (tdis.Area > 0)
                            {
                                if (!tdis.RoadDisType.Contains("破碎板") && !tdis.RoadDisType.Contains("松散") && !tdis.RoadDisType.Contains("露骨"))
                                {
                                    if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
                                    {
                                        tdis.m_mile += _ProjectInfo._Direction;
                                    }

                                }
                                else
                                {
                                    if (tdis.Area <= _RoadConfig.DetectWidth * 2 * 2 / 3)
                                    {
                                        if (tdis.rect.Y > (_RoadConfig.ImageHeight - tdis.rect.Height) / 2)
                                        {
                                            tdis.m_mile = tdis.m_mile + _ProjectInfo._Direction;
                                        }
                                    }

                                }

                                var1New[0] = diseaseNum++;
                                var1New[1] = tdis.m_mile;
                                var1New[2] = tdis.RoadDisType;
                                var1New[3] = tdis.calcheight.ToString("0.000");
                                var1New[4] = tdis.calcwidth.ToString("0.000");
                                var1New[5] = tdis.Area.ToString("0.000000");

                                dataGridView_Dislist.Invoke(new Action(() =>
                                {
                                    dataGridView_Dislist.Rows.Add(var1New);
                                })
                     );
                             
                            }
                        }

                        dataGridView_MLDisImgList.Invoke(new Action(() =>
                        {
                            dataGridView_MLDisImgList.Rows.Add(var2);
                        })
                     );
                        
                    }
                    else
                    {
                        dataGridView_MLNoDisImgList.Invoke(new Action(() =>
                        {
                            dataGridView_MLNoDisImgList.Rows.Add(var2);
                        })
                 );
                        
                    }
                }
            }

            this.Cursor = Cursors.Default;
        }

        private void dataGridView_Dislist_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView tmpgrid = (DataGridView)sender;
                DataGridViewSelectedRowCollection selectrow = tmpgrid.SelectedRows;
                int markmile = int.Parse(selectrow[0].Cells[1].Value.ToString());
                EventJump2Dis(markmile, EventArgs.Empty);
            }
            catch { }
        }

        private void button_update_Click(object sender, EventArgs e)
        {
            LoadAllDis();
        }

        private void WinRoadDisList_Load(object sender, EventArgs e)
        {
       
            //LoadAllDis();
        }

        private void comboBox_BSType_SelectedIndexChanged(object sender, EventArgs e)
        {
            _BrowserType = comboBox_BSType.SelectedIndex;
            if (_BrowserType == 3)
            {
                _BrowserDisName = textBox_DisName.Text;
            }
        }
        [Obsolete("bug")]
        private void dataGridView_curdis_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Global._UpdateInfo tupdate = new Global._UpdateInfo();
            DataGridView tmpgrid = (DataGridView)sender;
            DataGridViewSelectedRowCollection selectrow = tmpgrid.SelectedRows;
            int markmile = int.Parse(selectrow[0].Cells[0].Value.ToString());

            int roadtype = 0;
            if (selectrow[0].Cells.Count >= 6)
            {
                if (selectrow[0].Cells[5].Value.ToString().Contains("沥青"))
                    roadtype = 0;
                else if (selectrow[0].Cells[5].Value.ToString().Contains("水泥"))
                    roadtype = 1;
                else if (selectrow[0].Cells[5].Value.ToString().Contains("砂石"))
                    roadtype = 2;
            }
            RoadPavementPanel PavementDisease = new RoadPavementPanel(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis, roadtype);
            if (_DisType)
            {
                PavementDisease.SetNumArea((int)(Convert.ToDouble(selectrow[0].Cells[4].Value) * 100), Convert.ToDouble(selectrow[0].Cells[4].Value));
            }
            else
            {
                PavementDisease.SetRealLengthWidth(Convert.ToDouble(selectrow[0].Cells[2].Value), Convert.ToDouble(selectrow[0].Cells[3].Value));
            }
            PavementDisease.ShowDialog();

            if (PavementDisease.IsDisease)
            {
                tupdate.disidx = e.RowIndex;
                tupdate.disname = PavementDisease.RoadDiseaseType;
                tupdate.disremark = PavementDisease.RoadDiseaseRemarks;
                EventUpdateDis(tupdate, EventArgs.Empty);
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadAllDis();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
