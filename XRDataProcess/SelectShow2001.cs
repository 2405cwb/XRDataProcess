using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml; 
using OperateIniFile;

namespace XRDataProcess
{
    public partial class SelectShow2001 : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        public String RoadDiseaseType;
        public SelectShow2001(Dictionary<string, int>[] DiseaseIdx, RoadDiseaseType[][] roaddis)
        {
            InitializeComponent();
            AddDiseaseControls(roaddis[0], tableLayoutPanel1, false);
            AddDiseaseControls(roaddis[1], tableLayoutPanel2, true);

        }

        private void AddDiseaseControls(RoadDiseaseType[] roaddis, TableLayoutPanel panel,bool fg)
        {
            int i = 0;
            int ccnt = 0;
            int rcnt = 0;
            for (int j = 0; j < roaddis.Length; ++j)
            {
                CheckBox type = new CheckBox()
                {
                    Text = roaddis[j].disname,
                    Top = 22 + i++ * 20,
                    Left = 20,
                    Width = 200,
                    Checked = roaddis[j].isshow
                };
                panel.Controls.Add(type, ccnt, rcnt++);
                if (fg)
                {
                    if (rcnt >= (roaddis.Length + 2) / 3)
                    {
                        rcnt = 0;
                        ccnt++;
                    }
                }
                else
                {
                    if (rcnt >= (roaddis.Length + 1) / 2)
                    {
                        rcnt = 0;
                        ccnt++;
                    }
                }
            }
        }

        private void button_confirm_Click(object sender, EventArgs e)
        {
            string fpath = Application.StartupPath + @"\ParaVal.xml";
            XmlDocument Doc = new XmlDocument();
            Doc = new XmlDocument();
            XmlElement Elem;
            Doc.Load(fpath);    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点 
            foreach (XmlNode rootchild in Elem.ChildNodes)
            {
                if (rootchild.Name == Framework.Other.MyGlobal.Global.g_ParmStyles[(int)_Setting.ParmStyle])
                {
                    foreach (XmlNode subnode in rootchild.ChildNodes)
                    {
                        if (subnode.Name == "沥青路面病害类型")
                        {
                            foreach (Control tctl in tableLayoutPanel1.Controls)
                            {
                                if (tctl is CheckBox)
                                {
                                    CheckBox tcheck = (CheckBox)tctl;
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
                            }
                        }
                        if (subnode.Name == "水泥路面病害类型")
                        {
                            foreach (Control tctl in tableLayoutPanel2.Controls)
                            {
                                if (tctl is CheckBox)
                                {
                                    CheckBox tcheck = (CheckBox)tctl;
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
