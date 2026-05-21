using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OperateIniFile;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using hnPanoShowAPI;

namespace XRDataProcess
{
    public partial class WinPanoImg : Form
    {
        public event EventHandler EventUpdateMile;

        private const int m_imgh = 8192;
        private const int m_imgw = 4096;
        private ProjectInfo _ProjectInfo;
        private string _ProjPath;
        public List<MyImgMile> _ImgPath = null;
        private int _curidx = 0;

        private double _dmival = 0;
        private int _mileval = 0;
        public bool _IsInitLoad = false;
        public bool _IsActivated = false;

        private PanoControlPanel pictureBox_Img;        

        public WinPanoImg(ProjectInfo pinfo, string ppath)
        {
            InitializeComponent();
            _ProjectInfo = pinfo;
            _ProjPath = ppath;
            _ImgPath = new List<MyImgMile>();
            
            pictureBox_Img = new PanoControlPanel();
            this.panel1.Controls.Add(pictureBox_Img);
            pictureBox_Img.Dock = DockStyle.Fill;

            InitPano();
        }

        public void InitPano()
        {
            _ImgPath.Clear();
            GetAllImg(_ProjPath + "\\PanoImg\\Camera0", ref _ImgPath);
        }

        private void ShowImg(MyImgMile path, int idx)
        {
            if (idx == 0)
            {
                _mileval = (int)(Convert.ToDouble(path.imgmile.ToString()));
                _dmival = _ProjectInfo.Mile2Dmi(_mileval);
                EventUpdateMile(_mileval, null);
                textBox_mile.Text = _mileval.ToString();
            }

            string strImagePath = string.Format(@"{0}\PanoImg\Camera{1}{2}", _ProjPath, idx, path.imgpath);
            pictureBox_Img.addImage(strImagePath);
            label_ImgPath.Text = strImagePath;

            textBox_dmi.Text = _dmival.ToString();
            progressBar_per.Value = _curidx;
        }

        private void GetAllImg(string path, ref List<MyImgMile> imgs)
        {
            if (File.Exists(path + "\\Pano2Mile.txt"))
            {
                string[] imgsinfo = File.ReadAllLines(path + "\\Pano2Mile.txt");
                foreach (string str in imgsinfo)
                {
                    imgs.Add(new MyImgMile(str));
                }
            }
        }

        public int BinSearch(double x, ref List<MyImgMile> imgmile, int direction)
        {
            int mid = 0, beg = 0, last = imgmile.Count - 1, miles = 0, milee = 0;
            if (beg > last)
            {
                return -1;
            }
            while (beg <= last)
            {
                mid = (beg + last) / 2;
                miles = Convert.ToInt32(imgmile[mid].imgmile);
                milee = Convert.ToInt32(imgmile[mid + 1 > last ? last : mid + 1].imgmile);
                if (x >= miles && x < milee || x <= miles && x > milee || mid == beg || last == beg)
                {
                    return mid;
                }
                else
                {
                    if ((miles - x) * direction < 0)
                    {
                        beg = mid;
                    }
                    else
                    {
                        last = mid;
                    }
                }
            }
            return -1;
        }

        public void ShowJumpImg(double jval)
        {
            if (jval <= _ImgPath[0].imgmile && jval >= _ImgPath[_ImgPath.Count - 1].imgmile
                || jval >= _ImgPath[0].imgmile && jval <= _ImgPath[_ImgPath.Count - 1].imgmile)
            {
                _curidx = BinSearch(jval, ref _ImgPath, _ProjectInfo._Direction);
                if (_curidx >= 0 && _curidx < _ImgPath.Count)
                {
                    ShowImg(_ImgPath[_curidx], 0);
                }
            }
        }
        public void ShowImg(double jval)
        {
            if (jval <= _ImgPath[0].imgmile && jval >= _ImgPath[_ImgPath.Count - 1].imgmile
                || jval >= _ImgPath[0].imgmile && jval <= _ImgPath[_ImgPath.Count - 1].imgmile)
            {
                _curidx = BinSearch(jval, ref _ImgPath, _ProjectInfo._Direction);
                if (_curidx >= 0 && _curidx < _ImgPath.Count)
                {
                    ShowImg(_ImgPath[_curidx], 0);
                }
            }
        }

        private void label_ImgPath_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowDefalutSystemImg(label_ImgPath.Text);
        }

        private void ShowDefalutSystemImg(string fpath)
        {           
            System.Diagnostics.Process process = new System.Diagnostics.Process();            
            process.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + @"\360度全景浏览工具.exe";
            process.StartInfo.Arguments = fpath;
            process.StartInfo.UseShellExecute = true;
            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized;
            process.Start();
            process.Close();
        }

        private void pictureBox_Img_MouseEnter(object sender, EventArgs e)
        {
            this.pictureBox_Img.Focus();
        }

        #region 里程和桩号互转
        public double[] m_d2m_dA = null;
        public double[] m_d2m_mA = null;
        static double _V2T(double v, double[] A)
        {
            if (A == null || A.Length == 0)
                return v;
            int i = A.Length - 1;
            double s, e;
            if (i == 0)
                return 0;
            while (--i >= 0)
            {
                s = A[i];
                e = A[i + 1];
                if (v > s && s > e || v < s && v < e)
                    continue;
                return i + (v - s) / (e - s);
            }
            i = A.Length - 1;
            if (A[0] < A[i]) return v <= A[0] ? 0 : i;
            else return v <= A[0] ? i : 0;
        }
        static double _T2V(double v, double[] A)
        {
            if (A == null || A.Length == 0)
                return v;
            if (v <= 0)
                return A[0];
            if (v >= A.Length - 1) return A[A.Length - 1];
            int i = (int)v; double t;
            v = v - i;
            t = A[i];
            return t + (A[i + 1] - t) * v;
        }
        public double DMI2Mile(double dmi)
        {
            return _T2V(_V2T(dmi, m_d2m_dA), m_d2m_mA);
        }
        public double Mile2DMI(double mile)
        {
            return _T2V(_V2T(mile, m_d2m_mA), m_d2m_dA);
        }
        private void InitProject(string projectpath)
        {
            m_d2m_mA = m_d2m_dA = null;
            //DoDMI2Mile(projectpath);

            string fpath = projectpath + "\\RoadImage\\MergeImg\\Dmi2Mile.txt";
            if (File.Exists(fpath))
            {
                List<double> dA = new List<double>();
                List<double> mA = new List<double>();
                char[] sp = new char[] { ' ', '\t', ',' };
                double d, m;
                foreach (String line in System.IO.File.ReadAllLines(fpath))
                {
                    String[] A = line.Split(sp, StringSplitOptions.RemoveEmptyEntries);
                    if (A.Length != 2) continue;
                    if (!double.TryParse(A[0], out d))
                        continue;
                    if (!double.TryParse(A[1], out m))
                        continue;
                    dA.Add(d); mA.Add(m);
                }
                if (dA.Count > 1)
                {
                    m_d2m_dA = dA.ToArray();
                    m_d2m_mA = mA.ToArray();
                }
            }
            else
            {
                //  MessageBox.Show("缺少 Dmi2Mile.txt 文件");
                // return;
            }
        }
        #endregion
        private void button_next_Click(object sender, EventArgs e)
        {
            ShowNextImg();
        }
        private void button_last_Click(object sender, EventArgs e)
        {
            ShowLastImg();
        }

        public bool _IsRoadAutoPlay = false;
        private void button_play_Click(object sender, EventArgs e)
        {
            if (!_IsRoadAutoPlay)
            {
                timer_roadplay.Start();
                button_play.ImageIndex = 6;
                _IsRoadAutoPlay = true;
            }
            else
            {
                timer_roadplay.Stop();
                button_play.ImageIndex = 7;
                _IsRoadAutoPlay = false;
            }
        }

        private void button_speedadd_Click(object sender, EventArgs e)
        {
            bool flag = false;
            if (timer_roadplay.Enabled)
            {
                flag = true;
                timer_roadplay.Stop();
            }
            int timeval = timer_roadplay.Interval / 2;
            if (timeval > 0)
            {
                timer_roadplay.Interval = timeval;
            }
            if (flag)
            {
                timer_roadplay.Start();
            }
        }
        private void button_speedsub_Click(object sender, EventArgs e)
        {
            bool flag = false;
            if (timer_roadplay.Enabled)
            {
                flag = true;
                timer_roadplay.Stop();
            }
            int timeval = timer_roadplay.Interval * 2;
            if (timeval > 0)
            {
                timer_roadplay.Interval = timeval;
            }
            if (flag)
            {
                timer_roadplay.Start();
            }
        }

        private void button_jump_Click(object sender, EventArgs e)
        {
            int temp = 0;
            try
            {
                temp = _ProjectInfo.Mile2Dmi(Convert.ToInt32(textBox_mile.Text));
            }
            catch
            {
                return;
            }

            if (_dmival != temp)
            {
                ShowJumpImg(temp);
                return;
            }

            try
            {
                temp = _ProjectInfo.Dmi2Mile((int)Math.Ceiling(Mile2DMI(Convert.ToDouble(textBox_dmi.Text))));
            }
            catch
            {
                return;
            }
            if (_dmival != temp)
            {
                ShowJumpImg(temp);
                return;
            }
        }
        public void ShowNextImg()
        {
            if (_ProjectInfo._IsPano)
            {
                if (_curidx + 1 < _ImgPath.Count)
                {
                    ShowImg(_ImgPath[++_curidx], 0);
                }
                else
                {
                    MessageBox.Show("已经是最后一张图像！");
                }
            }
        }
        public void ShowLastImg()
        {
            if (_ProjectInfo._IsPano)
            {
                if (_curidx > 0)
                {
                    ShowImg(_ImgPath[--_curidx], 0);
                }
                else
                {
                    MessageBox.Show("已经是第一张图像！");
                }
            }
        }
        
        private void WinPanoImg_Load(object sender, EventArgs e)
        {
            progressBar_per.Maximum = _ImgPath.Count;

            if (_ImgPath.Count > 0)
            {
                ShowImg(_ImgPath[_curidx], 0);
            }
            _IsInitLoad = true;
        }

    }
}
