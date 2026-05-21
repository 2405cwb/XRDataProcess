using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace XRDataProcess
{
    public partial class GetDiseaseFiles : Form
    {
        private string selectPath;
        public string openPath;
        public List<string> srcPaths;

        public GetDiseaseFiles()
        {
            InitializeComponent();
        }

        private void button_Select_Click(object sender, EventArgs e)
        {
            MessageBox.Show("请保证选择的目的文件夹所在磁盘剩余空间充足！", "注意");
            FolderBrowserDialog fd = new FolderBrowserDialog();
            fd.Description = "选择导出病害文件放置位置：";
            fd.ShowDialog();
            if (fd.SelectedPath != string.Empty)
            {
                if (fd.SelectedPath.Substring(fd.SelectedPath.Length - 1) == "\\")
                {
                    fd.SelectedPath = fd.SelectedPath.Remove(fd.SelectedPath.Length - 1);
                }

                if (fd.SelectedPath == openPath)
                {
                    MessageBox.Show("导出病害文件放置位置不能和已打开的工程位置相同，请重新选择！", "注意");
                }
                else
                {
                    selectPath = fd.SelectedPath;
                    textBox_Path.Text = selectPath;
                    button_Start.Enabled = true;
                }
            }
        }

        private void button_Start_Click(object sender, EventArgs e)
        {
            MessageBox.Show("开始提取工程中的病害文件，请关闭此窗口，并等待弹窗提示完成！", "提示");

            this.Cursor = Cursors.WaitCursor;
            foreach(string str in srcPaths)
            {
                CopyFiles(selectPath, str);
            }
            this.Cursor = Cursors.Default;

            MessageBox.Show("提取工程中的病害文件完成！", "提示");
            MessageBox.Show("要将导出的病害文件和原始工程文件合并时，请注意需要在合并后的新工程中手动编辑【沥青/水泥/砂石】材质变化的位置！！！", "提示");
            this.Close();
        }

        private void CopyFiles(string destPath, string srcPrjPath)
        {
            string roadpath = srcPrjPath + "\\RoadImg\\Camera0";
            if (!Directory.Exists(roadpath))
                return;

            DirectoryInfo Prjdir = new DirectoryInfo(srcPrjPath);
            string destPrjPath = destPath + "\\" + Prjdir.Name + "\\RoadImg\\Camera0";
            Directory.CreateDirectory(destPrjPath);

            DirectoryInfo roaddir = new DirectoryInfo(roadpath);
            DirectoryInfo[] srcdirs = roaddir.GetDirectories();
            foreach(DirectoryInfo tdir in srcdirs)
            {
                string tpath = destPrjPath + "\\" + tdir.Name;
                Directory.CreateDirectory(tpath);

                FileInfo[] srcfiles = tdir.GetFiles("*.txt");
                foreach(FileInfo tfile in srcfiles)
                {
                    string newfilepath = tpath + "\\" + tfile.Name;
                    File.Copy(tfile.FullName, newfilepath, true);
                }
            }
        }

        private void button_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void GetDiseaseFiles_Load(object sender, EventArgs e)
        {
            button_Start.Enabled = false;
        }
    }
}
