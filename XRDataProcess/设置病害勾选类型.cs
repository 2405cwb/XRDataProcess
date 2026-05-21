using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Framework.Other.MyGlobal;
using OperateIniFile;

namespace XRDataProcess
{
    public partial class 设置病害勾选类型 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        public String RoadDiseaseType;
        public 设置病害勾选类型(Dictionary<string, int>[] DiseaseIdx, RoadDiseaseType[][] roaddis)
        {
            InitializeComponent();

            AddDiseaseControls(roaddis[0], tableLayoutPanel1);
            AddDiseaseControls(roaddis[1], tableLayoutPanel2);

            if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing||_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel
                /*||_Setting.ParmStyle==  StandardParmType.RuralRoadHunan2024*/)
            {
                tableLayoutPanel1.ColumnCount = 2;
                tableLayoutPanel2.ColumnCount = 2;
                tableLayoutPanel3.ColumnCount = 2;
                AddDiseaseControls(roaddis[2], tableLayoutPanel3);
            }
            else
            {
                tableLayoutPanel3.Visible = false;
                tableLayoutPanel0.ColumnCount = 2;
            }
        }

        private void AddDiseaseControls(RoadDiseaseType[] roaddis, TableLayoutPanel panel)
        {
            int i = 0;
            int ccnt = 0;
            int rcnt = 0;
            for (int j = 0; j < roaddis.Length; ++j)
            {
                CheckBox type = new CheckBox()
                {
                    Text = roaddis[j].disname,
                    Checked = roaddis[j].isshow
                };
                type.Dock = DockStyle.Fill;

                TextBox tbox = new TextBox();
                tbox.Text = roaddis[j].shortcut;
                tbox.Tag = roaddis[j].disname;
                tbox.Dock = DockStyle.Fill;

                panel.Controls.Add(type, ccnt, rcnt);
                panel.Controls.Add(tbox, ccnt+1, rcnt);
                ++rcnt;

                if ((rcnt >= (roaddis.Length + 1) / 2))
                {
                    rcnt = 0;
                    ccnt += 2;
                }
            }
        }

        private void button_confirm_Click(object sender, EventArgs e)
        {
            TableLayoutPanel[] tablelayout = { tableLayoutPanel1, tableLayoutPanel2, tableLayoutPanel3 };

            string fpath = Application.StartupPath + @"\ParaVal.xml";
            XmlDocument Doc = new XmlDocument();
            Doc = new XmlDocument();
            XmlElement Elem;
            Doc.Load(fpath);    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点 
            foreach (XmlNode rootchild in Elem.ChildNodes)
            {
                if (rootchild.Name == Global.g_ParmStyles[(int)_Setting.ParmStyle])
                {
                    foreach (XmlNode subnode in rootchild.ChildNodes)
                    {
                        if (subnode.Name.Contains("路面病害类型"))
                        {
                            int idx = RoadDiseaseTypes.roadtypedict[subnode.Name.Substring(0, 2)];
                            foreach (Control tctl in tablelayout[idx].Controls)
                            {
                                if (tctl is CheckBox)
                                {
                                    CheckBox tcheck = tctl as CheckBox;
                                    foreach (XmlNode node in subnode.ChildNodes)
                                    {
                                        if (tcheck.Text == node.Name)
                                        {
                                            XmlElement enode = (XmlElement)node;
                                            enode.SetAttribute("显示", tcheck.Checked ? "1" : "0");
                                            break;
                                        }
                                    }
                                }
                                else if (tctl is TextBox)
                                {
                                    TextBox tbox = tctl as TextBox;
                                    string disname = tbox.Tag.ToString();
                                    foreach (XmlNode node in subnode.ChildNodes)
                                    {
                                        if (disname == node.Name)
                                        {
                                            XmlElement enode = (XmlElement)node;
                                            enode.SetAttribute("快捷键", tbox.Text);
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Doc.Save(fpath);
                }
            }        
            MessageBox.Show("修改病害勾选类型成功，即将重启软件！");
            Application.Exit();
            System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);
            this.Close();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
