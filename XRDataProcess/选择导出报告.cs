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
    public partial class 选择导出报告 : Form
    {
        /// <summary>
        /// 是否选择了有效的导出报告
        /// </summary>
        public bool m_ischek = false;

        /// <summary>
        /// 选中的检测报告ID
        /// </summary>
        public string m_reportID = null;

        /// <summary>
        /// 选中的项目ID
        /// </summary>
        public string m_prjID = null;

        private List<ProjectProjectClass> _SrcProjectinfoList = null;
        public 选择导出报告()
        {
            InitializeComponent();
        }

        public 选择导出报告(List<ProjectProjectClass> SrcProjectinfoList)
        {
            InitializeComponent();

            _SrcProjectinfoList = SrcProjectinfoList;
        }

        private void 选择导出报告_Load(object sender, EventArgs e)
        {
            if (_SrcProjectinfoList != null)
            {
                foreach (ProjectProjectClass tproject in _SrcProjectinfoList)
                {
                    TreeNode prjnode = new TreeNode();
                    prjnode.Text = tproject.m_project.m_id;
                    foreach (ReportProjectClass treport in tproject.m_reportlist)
                    {
                        TreeNode reportnode = new TreeNode();
                        reportnode.StateImageIndex = 0;
                        reportnode.Text = treport.m_report.m_id;
                        foreach (RoadPartProjectClass troadpart in treport.m_roadpartlist)
                        {
                            TreeNode roadpartnode = new TreeNode();
                            roadpartnode.Text = troadpart.m_roadpart.m_id + "："
                                + troadpart.m_roadpart.m_roadinfo.m_code + "_"
                                + troadpart.m_roadpart.m_roadinfo.m_name + "_"
                                + troadpart.m_roadpart.m_startlocation + "-" + troadpart.m_roadpart.m_endlocation + "_";
                            foreach (LaneProjectClass tlane in troadpart.m_lanelist)
                            {
                                TreeNode lanenode = new TreeNode();
                                lanenode.Text = tlane.m_lane.m_id + "："
                                    + tlane.m_lane.m_direction + "_" + tlane.m_lane.m_lanenum;
                                foreach(string prj in tlane.m_projectdatapathlist)
                                {
                                    TreeNode prjpathnode = new TreeNode();
                                    prjpathnode.Text = prj;
                                    lanenode.Nodes.Add(prjpathnode);                                
                                }
                                roadpartnode.Nodes.Add(lanenode);
                            }
                            reportnode.Nodes.Add(roadpartnode);
                        }
                        prjnode.Nodes.Add(reportnode);
                    }
                    treeView1.Nodes.Add(prjnode);
                }
            }
        }

        private void button_No_Click(object sender, EventArgs e)
        {
            m_ischek = false;
            this.Close();
        }

        private void button_Yes_Click(object sender, EventArgs e)
        {
            if (m_ischek == false)
            {
                MessageBox.Show("没有选择要导出的检测报告！");
            }
            else
            {
                this.Close();
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Level == 1)
            {
                for (int i = 0; i < treeView1.Nodes[e.Node.Parent.Index].Nodes.Count; ++i)
                {
                    if (i != e.Node.Index)
                    {
                        treeView1.Nodes[e.Node.Parent.Index].Nodes[i].StateImageIndex = 0;
                    }
                    else
                    {
                        treeView1.Nodes[e.Node.Parent.Index].Nodes[i].StateImageIndex = 1;
                        label_log.Text = "已选择要导出的检测报告ID=" + treeView1.Nodes[e.Node.Parent.Index].Nodes[i].Text;

                        m_ischek = true;
                        m_prjID = treeView1.Nodes[e.Node.Parent.Index].Text;
                        m_reportID = treeView1.Nodes[e.Node.Parent.Index].Nodes[i].Text;
                    }
                }
            }
            else
            {
                for (int i = 0; i < treeView1.Nodes.Count; ++i)
                {
                    for (int j = 0; j < treeView1.Nodes[i].Nodes.Count; ++j)
                    {
                        treeView1.Nodes[i].Nodes[j].StateImageIndex = 0;
                    }
                }
                label_log.Text = "没有选择要导出的检测报告";

                m_ischek = false;
                m_reportID = null;
                m_prjID = null;
            }
        }
    }
}
