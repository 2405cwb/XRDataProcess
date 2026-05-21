using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using Framework.Other.MyGlobal;

namespace XRDataProcess
{
    public partial class 病害类型转换 : Form
    {
        List<SingleProject> _Projects = null;
        private Dictionary<string, string>[] DisTypeDict = new Dictionary<string, string>[2];
        private Dictionary<string, int> RoadTypeDict = new Dictionary<string, int>();
        public 病害类型转换()
        {
            InitializeComponent();
            for (int i = 0; i < 2; ++i)
            {
                DisTypeDict[i] = new Dictionary<string,string>();
            }

            RoadTypeDict.Add("沥青", 0);
            RoadTypeDict.Add("水泥", 1);
        }

        public 病害类型转换(List<SingleProject> projects)
        {
            InitializeComponent();
            for (int i = 0; i < 2; ++i)
            {
                DisTypeDict[i] = new Dictionary<string, string>();
            }

            RoadTypeDict.Add("沥青", 0);
            RoadTypeDict.Add("水泥", 1);

            _Projects = projects;
        }

        private void 病害类型转换_Load(object sender, EventArgs e)
        {
            comboBox_Src.SelectedIndex = 0;
            comboBox_Dest.SelectedIndex = 0;
        }

        private void button_TranDis_Click(object sender, EventArgs e)
        {
            textBox_log.Text = "开始病害类型转换：【" + comboBox_Src.Text + "】转换为【" + comboBox_Dest.Text + "】\r\n";
            if (_Projects == null)
            {
                textBox_log.Text = textBox_log.Text + "没有待转换工程!\r\n";
                return;
            }

            int prjnum = _Projects.Count;
            if (prjnum == 0)
            {
                textBox_log.Text = textBox_log.Text + "没有待转换工程!\r\n";
                return;
            }

            for (int i = 0; i < 2; ++i )
            {
                DisTypeDict[i].Clear();
            }

            string trantype = comboBox_Src.Text + "-" + comboBox_Dest.Text;
            ReadXml(trantype, ref DisTypeDict);

            progressBar1.Value = 0;
            progressBar1.Maximum = prjnum;
            foreach (SingleProject tprj in _Projects)
            {
                string roadpath = tprj._DataDir.FullName + "\\RoadImg\\Camera0";
                if (!Directory.Exists(roadpath))
                    return;

                DirectoryInfo roaddir = new DirectoryInfo(roadpath);
                DirectoryInfo[] srcdirs = roaddir.GetDirectories();
                foreach (DirectoryInfo tdir in srcdirs)
                {
                    FileInfo[] srcfiles = tdir.GetFiles("*.jpg");
                    foreach (FileInfo tfile in srcfiles)
                    {
                        string txtfpath = tfile.FullName + ".txt";
                        if (File.Exists(txtfpath))
                        {
                            bool istran = false;
                            string[] dises = File.ReadAllLines(txtfpath);
                            int txtlen = dises.Length;
                            for (int i = 0; i < txtlen; ++i )
                            {
                                Disease tdis = new Disease();
                                tdis.SetDisInfoValFromTXT(dises[i]);

                                try
                                {
                                    tdis.RoadDisType = DisTypeDict[RoadTypeDict[tdis.RoadType]][tdis.RoadDisType];
                                    istran = true;
                                }
                                catch (System.Exception ex)
                                { }

                                dises[i] = tdis.GetDisInfoStr();
                            }

                            if (istran)
                            {
                                string newdir = tdir.FullName.Replace("\\RoadImg\\Camera0", "\\RoadImg\\TranCamera0");
                                if (!Directory.Exists(newdir))
                                {
                                    Directory.CreateDirectory(newdir);
                                }
                                string fpath = string.Format("{0}\\{1}.txt", newdir, tfile.Name);
                                if (File.Exists(fpath))
                                {
                                    File.Delete(fpath);
                                }
                                File.Move(txtfpath, fpath);
                                File.WriteAllLines(txtfpath, dises, Encoding.UTF8);
                            }
                        }
                    }
                }

                ++progressBar1.Value;
                textBox_log.Text = textBox_log.Text + tprj._DataDir.FullName + "转换完成！\r\n";
            }

            MessageBox.Show("病害类型转换完成！");
        }

        private void ReadXml(string trantype, ref Dictionary<string, string>[] tranDict)
        {
            XmlDocument Doc = new XmlDocument();
            Doc = new XmlDocument();
            XmlElement Elem;
            XmlNodeList xmlNodes;

            //读取病害类型
            Doc.Load(System.Windows.Forms.Application.StartupPath + "\\TranPara.xml");    //加载Xml文件  
            Elem = Doc.DocumentElement;   //获取根节点  
            xmlNodes = Elem.ChildNodes;

            for (int i = 0; i < 2; i++)
            {
                foreach (XmlNode rootchild in Elem.ChildNodes)
                {
                    if (rootchild.Name == trantype)
                    {
                        foreach (XmlNode subnode in rootchild.ChildNodes)
                        {
                            if (subnode.Name == GlobalExcel._RoadTypeStr[i] + "路面病害类型")
                            {
                                foreach (XmlNode node in subnode.ChildNodes)
                                {
                                    tranDict[i].Add(node.Name, node.InnerText);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
