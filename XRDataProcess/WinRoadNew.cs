
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using RuralPavementDetect;
using System.Xml;
using OperateIniFile;
using System.Text;
using System.Threading;
using Framework.Other.MyGlobal;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Diagnostics;
using XRDataProcess;
using System.Threading.Tasks;
using SqlSugar.Extensions;

namespace XRDataProcess
{
    public partial class WinRoadNew : WinRoad
    //public partial class WinRoadNew : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();

        override public event EventHandler EventUpdateDis;
        override public event EventHandler EventChangeType;
        override public event EventHandler EventUpdateMile;
        override public event EventHandler EventUpdateDmi;
        override public event EventHandler EventUpdateYG;
        override public event EventHandler EventUpdateFullImg;
        override public event EventHandler EventUpdateFullPoint;
        //[DllImport("YuGuang.dll", EntryPoint = "YG_PaintImg")]
        //static extern int YG_PaintImg(IntPtr hMain);

        [DllImport("YuGuang.dll", EntryPoint = "YG_LoadImg")]
        static extern int YG_LoadImg(string fpath, ref int imgw, ref int imgh);

        [DllImport("YuGuang.dll", EntryPoint = "YG_GetImgBuf")]
        static extern void YG_GetImgBuf([In, Out] IntPtr destp);

        [DllImport("YuGuang.dll", EntryPoint = "YG_GetImgBufNew")]
        static extern void YG_GetImgBufNew([In, Out] IntPtr destp);

        private ProjectInfo _ProjectInfo;
        private string _ProjPath;
        public List<MyImgMile> _ImgPath = null;

        public int _RoadType = 1;
        public int _RoadTypeOld = -1;

        private int _curidx = 0;
        private int _oldidx = 0;

        private List<MilePart> _RoadPart = new List<MilePart>();

        public static string _ImgName = null;

        private int _oldimgwidth = 0;
        private int _oldimgheight = 0;

        // 减少判断点数，计算点X坐标最小步数
        private int _PartStep = 0;
        //图片通道信息
        private int picPassagewayType;
        ColorPalette m_palette;
        private Bitmap _image = null;
        private BitmapData m_OriData = null;
        private Bitmap m_NewImg = null;
        private BitmapData m_NewData = null;

        private double _dmival = 0;
        private double _mileval = 0;

        private Rectangle RoadImgRect = new Rectangle();
        /// <summary>
        /// 小框人工病害地址
        /// </summary>
        private string unAutoDisPath; 
        
        private string deleteDisPath;
        private List<PartRectInfo> _PartRectInfos = new List<PartRectInfo>();
        List<SmalRectDisease> DisInfoList = new List<SmalRectDisease>();
        public List<int> CurRectIdx = new List<int>(); //记录单次病害的矩形

        public List<int> LineOneStepRectIdx = new List<int>(); //记录线性绘制模式两点之间的矩形

        public WinRoadNew(ProjectInfo pinfo, string ppath)
        {
            InitializeComponent();
            _ProjectInfo = pinfo;
            _ProjPath = ppath;
            _RoadType = _ProjectInfo._RoadType;

            _ImgPath = new List<MyImgMile>();

            string projectName = pinfo._PrjPath.Split('\\').Last();
            //小框人工病害记录文本
            //大框人工病害记录文本
            if (Directory.Exists(_Setting.outHumanDeleteDiseasePath))
            {
                try
                {
                    unAutoDisPath = string.Format(@"{0}\{1}\HumanSmallDisMessage.txt", _Setting.outHumanDeleteDiseasePath, projectName);
                    // 获取文件的上级目录路径
                    string directoryPath = Path.GetDirectoryName(unAutoDisPath);

                    // 创建上级目录（如果它们不存在）
                    Directory.CreateDirectory(directoryPath);

                    // 创建文件（如果它不存在）
                    if (!File.Exists(unAutoDisPath))
                    {
                        File.Create(unAutoDisPath).Close();
                    }
                    deleteDisPath = string.Format(@"{0}\{1}\SmallDeleteDisPath.txt", _Setting.outHumanDeleteDiseasePath, projectName);
                    // 获取文件的上级目录路径
                    directoryPath = Path.GetDirectoryName(deleteDisPath);
                    // 创建上级目录（如果它们不存在）
                    Directory.CreateDirectory(directoryPath);
                    // 创建文件（如果它不存在）
                    if (!File.Exists(deleteDisPath))
                    {
                        File.Create(deleteDisPath).Close();
                    }

                }
                catch (Exception ex)
                {
                }

            }

            string imgDir = string.Format("{0}\\RoadImg\\Camera0\\Image_0000", ppath);
            var timgname = Directory.GetFiles(imgDir, "*.jpg").FirstOrDefault();

            if (File.Exists(timgname))
            {
                using (FileStream fs = new FileStream(timgname, FileMode.Open, FileAccess.Read))
                {
                    System.Drawing.Image _image = System.Drawing.Image.FromStream(fs);
                    var picType = _image.PixelFormat;
                    if (picType == PixelFormat.Format24bppRgb)
                    {
                        picPassagewayType = 3;
                    }
                    _image.Dispose();
                    _image = null;
                }
            }

            if (picPassagewayType == 3)
            {
                _image = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format24bppRgb);
                m_NewImg = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format24bppRgb);
            }
            else
            {
                _image = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                m_NewImg = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                m_palette = _image.Palette;
                for (int i = 0; i < 256; i++)
                {
                    m_palette.Entries[i] = Color.FromArgb(i, i, i);
                }
                _image.Palette = m_palette;
                m_NewImg.Palette = m_palette;
            }

            pictureBox_road.MouseWheel += new MouseEventHandler(pictureBox_road_MouseWheel);
        }
        public override void SaveDisease()
        {
            ClearAllDiseaseInfoBox(true);

        }
        /// <summary>
        /// 初始化
        /// 读取文本中记录的病害
        /// </summary>
        private void readHumanDis()
        {
            //路径
            //   File.ReadLines()
        }
        override public void UpdateYG(object YGStatu)
        {
            bool IsOriImg = (bool)YGStatu;
            if (!IsOriImg)
            {
                //开始调整
                if (picPassagewayType == 3)
                {
                    m_NewData = m_NewImg.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight),
                 ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
                }
                else
                {
                    m_NewData = m_NewImg.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight),
                   ImageLockMode.ReadWrite, PixelFormat.Format8bppIndexed);
                }

                unsafe
                {
                    IntPtr pin = m_NewData.Scan0;
                    YG_GetImgBufNew(pin);
                }
                m_NewImg.UnlockBits(m_NewData);
                pictureBox_road.Image = m_NewImg;
            }
            else
            {
                pictureBox_road.Image = _image;
            }
            EventUpdateFullImg(pictureBox_road.Image, EventArgs.Empty);
        }

        private void WinRoadNew_Load(object sender, EventArgs e)
        {
           
            drawModel_Combox.SelectedIndex = _Setting.SmallDiseaseDrawType;
            drawModel_Combox.DropDownStyle = ComboBoxStyle.DropDownList;
            GetAllImg(_ProjPath + "\\RoadImg\\Camera0", ref _ImgPath);
            progressBar_per.Maximum = _ImgPath.Count;
            progressBar_per.Value = 0;

            _curidx = 0;
            RoadImgRect.X = 0;
            RoadImgRect.Y = 0;
            RoadImgRect.Width = pictureBox_road.Width;
            RoadImgRect.Height = pictureBox_road.Height;
            GetTypeMilePart(_ProjPath, _ProjectInfo._Direction);
            _IsInitLoad = true;

            if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing ||
                _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel ||
                _Setting.ParmStyle == StandardParmType.RuralRoadHunan)
                button_SS.Visible = true;
            else
            {
                button_SS.Visible = false;
            }
        }

        private void InitPartRects()
        {
            int len = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum;
            for (int i = 0; i < len; ++i)
            {
                _PartRectInfos.Add(new PartRectInfo(i, new Rectangle()));
            }
        }

        //获取所有图像
        private void GetAllImg(string path, ref List<MyImgMile> imgs)
        {
            string[] imgsinfo = File.ReadAllLines(_ProjPath + "\\RoadImg\\Camera0\\Road2Mile.txt");
            foreach (string str in imgsinfo)
            {
                _ImgPath.Add(new MyImgMile(str));
            }
        }
        private void ShowImg(MyImgMile path)
        {
            progressBar_per.Value = _curidx;
            _dmival = _curidx * _ProjectInfo._RoadImgDis;
            textBox_dmi.Text = _dmival.ToString();
            EventUpdateDmi(_dmival, null);

            _mileval = (int)(Convert.ToDouble(path.imgmile.ToString()));
            textBox_mile.Text = _mileval.ToString();
            EventUpdateMile(_mileval, null);

            int len = _RoadPart.Count - 1;
            for (int i = 0; i < len; ++i)
            {
                if ((_mileval - _RoadPart[i].mile) * _ProjectInfo._Direction >= 0
                    && (_mileval - _RoadPart[i + 1].mile) * _ProjectInfo._Direction < 0)
                {
                    _RoadType = _RoadPart[i].roadtype;
                    break;
                }
            }

            if (_RoadType != _RoadTypeOld)
            {
                foreach (Control tctl in flowLayoutPanel2.Controls)
                {
                    if (tctl is Button)
                    {
                        Button btn = tctl as Button;
                        if (btn.Tag != null)
                        {
                            int idx = Convert.ToInt16(btn.Tag);
                            if (idx >= 0)
                            {
                                if (_RoadType == idx)
                                {
                                    btn.BackColor = Color.CadetBlue;
                                    toolTip1.SetToolTip(btn, "当前为" + btn.Text + "路面");
                                }
                                else
                                {
                                    btn.BackColor = SystemColors.Control;
                                    toolTip1.SetToolTip(btn, "切换为" + btn.Text + "路面");
                                }
                            }
                        }
                    }
                }
                _RoadTypeOld = _RoadType;
            }
            //_PartClass.txt
            _ImgName = string.Format(@"{0}\RoadImg\Camera0{1}", _ProjPath, path.imgpath);
            label_imgpath.Text = _ImgName;
            if (MainForm._IsSaveDisImg)
            {
                if (!File.Exists(_ImgName + "_PartClass.txt"))
                {
                    return;
                }
            }

            picPassagewayType = YG_LoadImg(_ImgName, ref _RoadConfig.ImageWidth, ref _RoadConfig.ImageHeight);

            if (_image.Width != _RoadConfig.ImageWidth || _image.Height != _RoadConfig.ImageHeight)
            {
                _image.Dispose();
                m_NewImg.Dispose();
                if (picPassagewayType == 3)
                {
                    _image = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format24bppRgb);
                    m_NewImg = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format24bppRgb);
                    //  m_OriData = _image.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

                }
                else
                {
                    _image = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                    m_NewImg = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                    _image.Palette = m_palette;
                    m_NewImg.Palette = m_palette;
                    //  m_OriData = _image.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), ImageLockMode.ReadWrite, PixelFormat.Format8bppIndexed);

                }

            }
            //      if (!_Setting.hasCamsetting)
            {
                _RoadConfig.WidthScale = _RoadConfig.RealWidth * 1.0 / _RoadConfig.ImageWidth;
                _RoadConfig.HeightScale = _RoadConfig.RealHeight * 1.0 / _RoadConfig.ImageHeight;
            }

            if (picPassagewayType == 3)
            {
                m_OriData = _image.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            }
            else
            {
                m_OriData = _image.LockBits(new Rectangle(0, 0, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), ImageLockMode.ReadWrite, PixelFormat.Format8bppIndexed);

            }



            unsafe
            {
                IntPtr pdest = m_OriData.Scan0;
                YG_GetImgBuf(pdest);
                if (picPassagewayType == 3)
                {
                    int bytes = Math.Abs(m_OriData.Stride) * _RoadConfig.ImageHeight;
                    byte[] rgbValues = new byte[bytes];
                    Marshal.Copy(pdest, rgbValues, 0, bytes);
                    for (int counter = 0; counter < rgbValues.Length - 3; counter += 3)
                    {
                        var value = rgbValues[counter];
                        rgbValues[counter] = rgbValues[counter + 2];
                        rgbValues[counter + 2] = value;
                    }
                    Marshal.Copy(rgbValues, 0, pdest, bytes);
                }
            }
            _image.UnlockBits(m_OriData);
            //_image.Save("D:\\tst\\cshape.jpg", ImageFormat.Jpeg);
            EventUpdateYG(_image, EventArgs.Empty);
            EventUpdateFullImg(_image, EventArgs.Empty);

            LoadRecInfo(_ImgName + "_PartClass.txt");

            InitPartRectInfo();
            _oldidx = _curidx;
        }

        /// <summary>
        /// 初始化小方格
        /// </summary>
        private void InitPartRectInfo()
        {
            if (_PartRectInfos.Count == 0)
            {
                InitPartRects();
            }

            //减少反复初始化小方格矩形，提高效率
            if (_oldimgwidth != _RoadConfig.ImageWidth || _oldimgheight != _RoadConfig.ImageHeight)
            {
                InitPartRectrect();
                _oldimgwidth = _RoadConfig.ImageWidth;
                _oldimgheight = _RoadConfig.ImageHeight;
            }
            else
            {
                //初始化全部未选中
                int len = _PartRectInfos.Count;
                for (int i = 0; i < len; ++i)
                {
                    _PartRectInfos[i].SetCheck(false);
                }
            }

            //设置记录的小方格为选中状态
            try
            {
                foreach (SmalRectDisease rect in DisInfoList)
                {
                    string[] s = rect.dispos.Split('-');
                    s = s.Select(oneStr => oneStr.Split('.').First()).ToArray();
                    for (int i = 0; i < s.Length - 1; ++i)
                    {
                        _PartRectInfos[int.Parse(s[i])].SetCheck(true);
                    }
                }
            }
            catch (ArgumentOutOfRangeException ex)
            {
                if (tipTemp)
                {
                    MessageBox.Show("该工程内病害小方格索引超出最大界");
                    tipTemp = false;
                }

                // throw ex;
            }

        }

        private bool tipTemp = true;
        /// <summary>
        /// 初始化小方格矩形信息
        /// </summary>
        private void InitPartRectrect()
        {
            _RoadConfig.PartImgWidth = (int)(_RoadConfig.ImageWidth * 1.0 / _RoadConfig.PartWidthNum);
            _RoadConfig.PartImgHeight = (int)(_RoadConfig.ImageHeight * 1.0 / _RoadConfig.PartHeightNum);

            for (int i = 0; i < _RoadConfig.PartHeightNum; ++i)
            {
                for (int j = 0; j < _RoadConfig.PartWidthNum; ++j)
                {
                    Rectangle timgrect = new Rectangle(j * _RoadConfig.PartImgWidth, i * _RoadConfig.PartImgHeight, _RoadConfig.PartImgWidth, _RoadConfig.PartImgHeight);
                    //从图像坐标系转换到控件坐标系
                    Rectangle tpicrect = Img2Box(timgrect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                    _PartRectInfos[i * _RoadConfig.PartWidthNum + j].SetRect(tpicrect);
                }
            }
        }

        private void LoadRecInfo(string RecInfoFilename)
        {
            if (File.Exists(RecInfoFilename))
            {
                FileStream fr = File.OpenRead(RecInfoFilename);
                StreamReader sr = new StreamReader(fr);
                String strline;
                while ((strline = sr.ReadLine()) != null)
                {
                    SmalRectDisease temp = new SmalRectDisease(strline, (int)Math.Round(_ImgPath[_curidx].imgmile));



                    if (temp.isDiseaseOK)
                    {
#if 辽宁建祥3m
                        int splitY1 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2 / 3;
                        int splitY2 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2 * 2 / 3;
                        if (temp.FirstRectNum > splitY1 && temp.FirstRectNum < splitY2)
                    {
                        temp.m_mile += _ProjectInfo._Direction;
                    }
                        else if (temp.FirstRectNum > splitY2)
                    {
                        temp.m_mile = temp.m_mile + _ProjectInfo._Direction * 2;
                    }
#else
                        if (temp.FirstRectNum > (_RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2))
                        {
                            temp.m_mile += _ProjectInfo._Direction;
                        }
#endif

                        DisInfoList.Add(temp);
                    }
                }

                sr.Close();
                fr.Close();
            }
        }

        public Rectangle Box2Img(Rectangle boxrect, int boxw, int boxh, int imgw, int imgh)
        {
            Rectangle imgrect = new Rectangle();
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                double scalew = (double)imgw / boxw;
                double scaleh = (double)imgh / boxh;
                imgrect.X = (int)(boxrect.X * scalew);
                imgrect.Y = (int)(boxrect.Y * scaleh);
                imgrect.Width = (int)(boxrect.Width * scalew);
                imgrect.Height = (int)(boxrect.Height * scaleh);
            }
            return imgrect;
        }

        public Rectangle Img2Box(Rectangle imgrect, int boxw, int boxh, int imgw, int imgh)
        {
            Rectangle boxrect = new Rectangle();
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                double scalew = (double)boxw / imgw;
                double scaleh = (double)boxh / imgh;
                boxrect.X = (int)(imgrect.X * scalew);
                boxrect.Y = (int)(imgrect.Y * scaleh);
                boxrect.Width = (int)(imgrect.Width * scalew);
                boxrect.Height = (int)(imgrect.Height * scaleh);
            }
            return boxrect;
        }

        /// <summary>
        /// 将PicBox上的某一个点转换为小方格的序号索引
        /// </summary>
        /// <param name="point">坐标点</param>
        /// <returns>小方格索引</returns>
        public int BoxPoint2RectIdx(Point boxpoint, int boxw, int boxh, int imgw, int imgh)
        {
            int idx = 0;
            Point imgpoint = new Point();
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                double scalew = (double)imgw / boxw;
                double scaleh = (double)imgh / boxh;
                imgpoint.X = (int)(boxpoint.X * scalew);
                imgpoint.Y = (int)(boxpoint.Y * scaleh);
                idx = imgpoint.Y / _RoadConfig.PartImgHeight * _RoadConfig.PartWidthNum + imgpoint.X / _RoadConfig.PartImgWidth;
            }
            return idx;
        }

        public Point BoxPoint2RectPoint(Point boxpoint, int boxw, int boxh, int imgw, int imgh)
        {
            Point imgpoint = new Point();
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                double scalew = (double)imgw / boxw;
                double scaleh = (double)imgh / boxh;
                imgpoint.X = (int)(boxpoint.X * scalew);
                imgpoint.Y = (int)(boxpoint.Y * scaleh);
            }
            return imgpoint;
        }

        private void pictureBox_road_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta > 0)
            {
                ShowLastImg();
            }
            else if (e.Delta < 0)
            {
                ShowNextImg();
            }
        }

        private void pictureBox_road_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Pen m_pen1 = new Pen(Color.Red, 1);
            Pen m_pen2 = new Pen(Color.Blue, 3);

            m_pen1.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            m_pen1.Width = 1;
            if (_IsDrawLineModel)
            {
                {
                    //绘制虚拟直线
                    if (lineModelPoints.Count > 0)
                    {
                        g.DrawLine(m_pen1, lineModelPoints.Last(), curpoint);

                    } 
                    foreach (int idx in LineOneStepRectIdx)
                    {
                        g.DrawRectangle(m_pen2, _PartRectInfos[idx].GetRect());
                    }

                }
            }


            foreach (PartRectInfo tprectinfo in _PartRectInfos)
            {
                if (tprectinfo.GetChek())
                {
                    g.DrawRectangle(m_pen1, tprectinfo.GetRect());
                }
            }
            var s = _PartRectInfos.Where(t => t.GetChek());
            foreach (int idx in CurRectIdx)
            {
                g.DrawRectangle(m_pen2, _PartRectInfos[idx].GetRect());
            }

            int x = 0, y = 0;
            foreach (SmalRectDisease tdis in DisInfoList)
            {
                if (tdis.dispos != "")
                {
                    //防止超出索引的情况         
                    if (tdis.rectarry[0] < _PartRectInfos.Count)
                    {
                        x = _PartRectInfos[tdis.rectarry[0]].GetRect().X;
                        y = _PartRectInfos[tdis.rectarry[0]].GetRect().Y;
                    }


                    Brush color = null;
                    if (tdis.selectfg) color = Brushes.Yellow;
                    else color = Brushes.GreenYellow;

                    g.DrawString(GlobalExcel._RoadTypeStr[_RoadType] + "." + tdis.GetRectInfoStr(),
                        new Font("宋体", 10, FontStyle.Regular),
                        color, x, y);
                }
            }
            EventUpdateDis(DisInfoList, null);
        }

        private void pictureBox_road_Resize(object sender, EventArgs e)
        {
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                RoadImgRect.Width = pictureBox_road.Width;
                RoadImgRect.Height = pictureBox_road.Height;

                if (_PartRectInfos.Count == 0)
                {
                    InitPartRects();
                }

                for (int i = 0; i < _RoadConfig.PartHeightNum; ++i)
                {
                    for (int j = 0; j < _RoadConfig.PartWidthNum; ++j)
                    {
                        Rectangle timgrect = new Rectangle(j * _RoadConfig.PartImgWidth, i * _RoadConfig.PartImgHeight, _RoadConfig.PartImgWidth, _RoadConfig.PartImgHeight);
                        //从图像坐标系转换到控件坐标系
                        Rectangle tpicrect = Img2Box(timgrect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                        _PartRectInfos[i * _RoadConfig.PartWidthNum + j].SetRect(tpicrect);
                    }
                }
                _PartStep = Math.Min(_PartRectInfos[0].GetRect().Width, _PartRectInfos[0].GetRect().Height);
            }
        }

        // 定义与画图相关的控件
        private bool mouseStatus = false;//鼠标状态，false为松开
        private bool mouseMove = false;//鼠标状态，false为松开
        private Point lastpoint, curpoint;

        // 状态变量

        private List<Point> lineModelPoints = new List<Point>();
        /// <summary>
        /// 鼠标按下后触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pictureBox_road_MouseDown(object sender, MouseEventArgs e)
        {
            if (_IsDrawLineModel && e.Button == MouseButtons.Left)
            {
                
                mouseStatus = true;
                lineModelPoints.Add(e.Location);
                CurRectIdx.AddRange(LineOneStepRectIdx);
            }
            else
            {
                mouseStatus = true;
                CurRectIdx.Clear();
                lastpoint = e.Location;
            }

        }
        // 辅助交换函数
        private void Swap(ref int a, ref int b)
        {
            int temp = a;
            a = b;
            b = temp;
        }
        private void pictureBox_road_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Location.X > 0 && e.Location.Y < pictureBox_road.Width && e.Location.Y > 0 && e.Location.Y < pictureBox_road.Height)
            {
                EventUpdateFullPoint(BoxPoint2RectPoint(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), EventArgs.Empty);
            }
            if (mouseStatus)
            {
                mouseMove = true;
                if (e.Location.X > RoadImgRect.Width || e.Location.Y > RoadImgRect.Height || e.Location.X < 0 || e.Location.Y < 0)
                    return;

                int idx = BoxPoint2RectIdx(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                if (idx < 0 || idx >= _PartRectInfos.Count)
                    return;

                List<Point> linepoints = new List<Point>();
                List<Point> removePoints = new List<Point>();//供线性模式使用
                if (_IsDrawRect)
                {
                    //获取起止点之间的所有点
                    curpoint = e.Location;
                    linepoints.Add(curpoint);
                    GetRectPoint(ref linepoints);
                }
                else if (_IsDrawLineModel)
                {
                    LineOneStepRectIdx.Clear();
                    curpoint = e.Location;
                    if (lineModelPoints.Count > 0)
                    {
                        // 计算从_startPoint到curpoint的直线上的所有点
                        linepoints.Add(lineModelPoints.Last()); // 添加起点

                        // 使用Bresenham算法获取直线上的所有点
                        int x0 = lineModelPoints.Last().X;
                        int y0 = lineModelPoints.Last().Y;
                        int x1 = curpoint.X;
                        int y1 = curpoint.Y;

                        bool steep = Math.Abs(y1 - y0) > Math.Abs(x1 - x0);
                        if (steep)
                        {
                            // 交换x和y
                            Swap(ref x0, ref y0);
                            Swap(ref x1, ref y1);
                        }

                        if (x0 > x1)
                        {
                            // 确保从左到右绘制
                            Swap(ref x0, ref x1);
                            Swap(ref y0, ref y1);
                        }

                        int dx = x1 - x0;
                        int dy = Math.Abs(y1 - y0);
                        int error = dx / 2;
                        int ystep = (y0 < y1) ? 1 : -1;
                        int y = y0;

                        for (int x = x0; x <= x1; x++)
                        {
                            Point pt = steep ? new Point(y, x) : new Point(x, y);

                            // 确保点在有效范围内
                            if (pt.X >= 0 && pt.X < RoadImgRect.Width &&
                                pt.Y >= 0 && pt.Y < RoadImgRect.Height)
                            {
                                // 避免重复添加起点
                                if (pt != lineModelPoints.Last())
                                {
                                    linepoints.Add(pt);
                                }
                            }

                            error -= dy;
                            if (error < 0)
                            {
                                y += ystep;
                                error += dx;
                            }
                        }

                        // 添加终点（当前鼠标位置）
                        if (curpoint != lineModelPoints.Last())
                        {
                            linepoints.Add(curpoint);
                        }
                    }
                    else
                    {
                        curpoint = e.Location;
                        removePoints.Add(curpoint);
                        GetRectPoint(ref removePoints);
                    }

                    
                }
                else
                {
                    // 获取直线上的所有点
                    curpoint = e.Location;
                    linepoints.Add(curpoint);
                    GetLinePoint(ref linepoints);
                }

                //新增和修改新增视同一样
                if (e.Button == System.Windows.Forms.MouseButtons.Left
                    || _IsDrawLineModel)
                {
                    foreach (Point tpt in linepoints)
                    {
                        double maxWidth = 0;
                        maxWidth = _RoadConfig.PartWidthNum * _RoadConfig.PartImgWidth;
                        Rectangle timgrect = new Rectangle(_RoadConfig.PartWidthNum * _RoadConfig.PartImgWidth, _RoadConfig.PartHeightNum * _RoadConfig.PartImgHeight, _RoadConfig.PartImgWidth, _RoadConfig.PartImgHeight);
                        //从图像坐标系转换到控件坐标系
                        Rectangle tpicrect = Img2Box(timgrect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                        if (tpt.X <= tpicrect.X && tpt.Y <= tpicrect.Y)
                        {
                            //得到小方格个数时进行了取整，导致如果鼠标绘制位置位于图像最右边时  还是能够查到索引，这里加一个判断这种情况不进行绘制

                            idx = BoxPoint2RectIdx(tpt, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                            if (idx >= 0 && idx < _PartRectInfos.Count)
                            {
                                if (!_PartRectInfos[idx].GetChek() && !CurRectIdx.Contains(idx))
                                {

                                    if (_IsDrawLineModel)
                                    {
                                        if (!LineOneStepRectIdx.Contains(idx))
                                        {
                                            LineOneStepRectIdx.Add(idx);

                                        }

                                    }
                                    else
                                    {
                                        CurRectIdx.Add(idx);
                                    }
                                }
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                // 删除整个病害和删除病害的某部分
                else if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    foreach (Point tpt in linepoints)
                    {
                        int i = 0, j = 0;
                        bool bflag = false;
                        int dlen = DisInfoList.Count;
                        int plen = 0;
                        List<int> tarr;
                        idx = BoxPoint2RectIdx(tpt, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                        if (idx >= 0 && idx < _PartRectInfos.Count)
                        {
                            for (i = 0; i < dlen; ++i)
                            {
                                bflag = false;
                                tarr = DisInfoList[i].rectarry;
                                plen = tarr.Count;
                                for (j = 0; j < plen; ++j)
                                {
                                    if (tarr[j] == idx)
                                    {
                                        _PartRectInfos[idx].SetCheck(false);
                                        bflag = true;
                                        break;
                                    }
                                }
                                if (bflag)
                                {
                                    break;
                                }
                            }
                            if (bflag)
                            {
                                DisInfoList[i].rectarry.RemoveAt(j);
                                if (DisInfoList[i].rectarry.Count > 0)
                                {
                                    DisInfoList[i].Update();
                                }
                                else
                                {
                                    DisInfoList.RemoveAt(i);
                                }
                            }
                        }
                    }
                }

                 if (e.Button == System.Windows.Forms.MouseButtons.Right&& _IsDrawLineModel)
                {
                    foreach (Point tpt in removePoints)
                    {
                        int i = 0, j = 0;
                        bool bflag = false;
                        int dlen = DisInfoList.Count;
                        int plen = 0;
                        List<int> tarr;
                        idx = BoxPoint2RectIdx(tpt, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                        if (idx >= 0 && idx < _PartRectInfos.Count)
                        {
                            for (i = 0; i < dlen; ++i)
                            {
                                bflag = false;
                                tarr = DisInfoList[i].rectarry;
                                plen = tarr.Count;
                                for (j = 0; j < plen; ++j)
                                {
                                    if (tarr[j] == idx)
                                    {
                                        _PartRectInfos[idx].SetCheck(false);
                                        bflag = true;
                                        break;
                                    }
                                }
                                if (bflag)
                                {
                                    break;
                                }
                            }
                            if (bflag)
                            {
                                DisInfoList[i].rectarry.RemoveAt(j);
                                if (DisInfoList[i].rectarry.Count > 0)
                                {
                                    DisInfoList[i].Update();
                                }
                                else
                                {
                                    DisInfoList.RemoveAt(i);
                                }
                            }
                        }
                    }
                }

            }
            pictureBox_road.Invalidate();
        }
        private void pictureBox_road_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left && !_IsDrawLineModel)
            {
                if (mouseMove && mouseStatus)
                {
                    mouseMove = false;
                    mouseStatus = false;
                    if (CurRectIdx.Count == 0 || !NewDrawDis()) //不保存
                    {
                        CurRectIdx.Clear();
                        pictureBox_road.Invalidate();
                    }
                    else
                    {
                        foreach (int idx in CurRectIdx)
                        {
                            _PartRectInfos[idx].SetCheck(true);
                        }
                    }
                }
                pictureBox_road.Invalidate();
            }

        }
        private void GetLinePoint(ref List<Point> linepoints)
        {
            if (Math.Abs(curpoint.X - lastpoint.X) > Math.Abs(curpoint.Y - lastpoint.Y))
            {
                float k = ((float)(curpoint.Y - lastpoint.Y)) / (curpoint.X - lastpoint.X);
                if (curpoint.X < lastpoint.X)
                {
                    for (int x = curpoint.X + _PartStep / 3; x < lastpoint.X; x += _PartStep)
                    {
                        int y = (int)Math.Round(k * (x - curpoint.X) + curpoint.Y);
                        if (y > RoadImgRect.Height || y < 0)
                            continue;
                        linepoints.Add(new Point(x, y));
                    }
                }
                else
                {
                    for (int x = lastpoint.X + _PartStep / 3; x < curpoint.X; x += _PartStep)
                    {
                        int y = (int)Math.Round(k * (x - lastpoint.X) + lastpoint.Y);
                        if (y > RoadImgRect.Height || y < 0)
                            continue;
                        linepoints.Add(new Point(x, y));
                    }
                }
            }
            else
            {
                float k = ((float)(curpoint.X - lastpoint.X)) / (curpoint.Y - lastpoint.Y);
                if (curpoint.Y < lastpoint.Y)
                {
                    for (int y = curpoint.Y + _PartStep / 3; y < lastpoint.Y; y += _PartStep)
                    {
                        int x = (int)Math.Round(k * (y - curpoint.Y) + curpoint.X);
                        if (x > RoadImgRect.Width || x < 0)
                            continue;
                        linepoints.Add(new Point(x, y));
                    }
                }
                else
                {
                    for (int y = lastpoint.Y + _PartStep / 3; y < curpoint.Y; y += _PartStep)
                    {
                        int x = (int)Math.Round(k * (y - lastpoint.Y) + lastpoint.X);
                        if (x > RoadImgRect.Width || x < 0)
                            continue;
                        linepoints.Add(new Point(x, y));
                    }
                }
            }
            lastpoint = curpoint;
        }
        private void GetRectPoint(ref List<Point> linepoints)
        {
            if (curpoint.X < lastpoint.X)
            {
                if (curpoint.Y < lastpoint.Y)
                {
                    for (int x = curpoint.X + _PartStep / 3; x < lastpoint.X; x += _PartStep)
                    {
                        for (int y = curpoint.Y + _PartStep / 3; y < lastpoint.Y; y += _PartStep)
                        {
                            linepoints.Add(new Point(x, y));
                        }
                    }
                }
                else
                {
                    for (int x = curpoint.X + _PartStep / 3; x < lastpoint.X; x += _PartStep)
                    {
                        for (int y = lastpoint.Y + _PartStep / 3; y < curpoint.Y; y += _PartStep)
                        {
                            linepoints.Add(new Point(x, y));
                        }
                    }
                }
            }
            else
            {
                if (curpoint.Y < lastpoint.Y)
                {
                    for (int x = lastpoint.X + _PartStep / 3; x < curpoint.X; x += _PartStep)
                    {
                        for (int y = curpoint.Y + _PartStep / 3; y < lastpoint.Y; y += _PartStep)
                        {
                            linepoints.Add(new Point(x, y));
                        }
                    }
                }
                else
                {
                    for (int x = lastpoint.X + _PartStep / 3; x < curpoint.X; x += _PartStep)
                    {
                        for (int y = lastpoint.Y + _PartStep / 3; y < curpoint.Y; y += _PartStep)
                        {
                            linepoints.Add(new Point(x, y));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 鼠标按下并释放后触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pictureBox_road_MouseClick(object sender, MouseEventArgs e)
        {
            foreach (SmalRectDisease dis in DisInfoList)
            {
                dis.selectfg = false;
            }

            pictureBox_road.Invalidate();
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (e.Location.X > RoadImgRect.Width || e.Location.Y > RoadImgRect.Height || e.Location.X < 0 || e.Location.Y < 0)
                    return;
                _idx = BoxPoint2RectIdx(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                foreach (SmalRectDisease dis in DisInfoList)
                {
                    if (dis.rectarry.Contains(_idx))
                    {
                        dis.selectfg = true;

                    }
                }
            }

        }
        private void pictureBox_road_MouseEnter(object sender, EventArgs e)
        {
            this.pictureBox_road.Focus();
        }
        private void pictureBox_road_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (e.Location.X > RoadImgRect.Width || e.Location.Y > RoadImgRect.Height || e.Location.X < 0 || e.Location.Y < 0)
                    return;
                int idx = BoxPoint2RectIdx(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                if (idx < 0 || idx >= _PartRectInfos.Count)
                    return;

                bool flag = false;

                foreach (SmalRectDisease dis in DisInfoList)
                {
                    if (dis.rectarry.Contains(idx))
                    {
                        DisInfoList.Remove(dis);
                        if (_Setting.outHumanDeleteDisease)
                        {
                            string deleteTxt = string.Format("{0} 桩号:{1} {2} {3}\n{4}", dis.RoadDisType, dis.m_mile.ToString("K0+000"), dis.RoadType, dis.dispos, _ImgPath[_curidx].imgpath);

                            try
                            {
                                using (StreamWriter fs = new StreamWriter(deleteDisPath, true))
                                {
                                    fs.WriteLine(deleteTxt);
                                }

                                string picFile = string.Format(@"{0}\RoadImg\Camera0{1}", _ProjPath, _ImgPath[_oldidx].imgpath);
                                string saveRecordPic = string.Format(@"{0}{1}", deleteDisPath.Substring(0, deleteDisPath.LastIndexOf("\\")), _ImgPath[_oldidx].imgpath);
                                // 获取文件的上级目录路径
                                string directoryPath = Path.GetDirectoryName(saveRecordPic);

                                // 创建上级目录（如果它们不存在）
                                Directory.CreateDirectory(directoryPath);

                                // 创建文件（如果它不存在）
                                if (!File.Exists(saveRecordPic))
                                {
                                    File.Copy(picFile, saveRecordPic);
                                }
                            }
                            catch (Exception ex)
                            {
                            }
                        }

                        foreach (int tidx in dis.rectarry)
                        {
                            _PartRectInfos[tidx].SetCheck(false);
                        }
                        flag = true;
                        break;
                    }
                }
                if (flag)
                {
                    this.pictureBox_road.Invalidate();
                }
            }
        }

        public void ShowNextImg()
        {
            if (WinRoadDisList._BrowserType == 1)
            {
                while (true)
                {
                    if (_curidx + 1 < _ImgPath.Count)
                    {
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
                        {
                            ClearAllDiseaseInfoBox(false);
                            ShowImg(_ImgPath[_curidx]);
                            break;
                        }
                    }
                    else if (_curidx + 1 == _ImgPath.Count)
                    {
                        ClearAllDiseaseInfoBox(true);
                        _IsRoadAutoPlay = false;
                        MessageBox.Show("已经是最后一张自动识别有病害图像！");
                        break;
                    }
                }
            }
            else if (WinRoadDisList._BrowserType == 2)
            {
                while (true)
                {
                    if (_curidx + 1 < _ImgPath.Count)
                    {
                        if (!File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
                        {
                            ClearAllDiseaseInfoBox(false);
                            ShowImg(_ImgPath[_curidx]);
                            break;
                        }
                    }
                    else if (_curidx + 1 == _ImgPath.Count)
                    {
                        ClearAllDiseaseInfoBox(true);
                        _IsRoadAutoPlay = false;
                        MessageBox.Show("已经是最后一张自动识别无病害图像！");
                        break;
                    }
                }
            }
            else if (WinRoadDisList._BrowserType == 3)
            {
                while (true)
                {
                    if (_curidx + 1 < _ImgPath.Count)
                    {
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
                        {
                            string strs = File.ReadAllText(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[_curidx].imgpath));
                            if (strs.Contains(WinRoadDisList._BrowserDisName))
                            {
                                ClearAllDiseaseInfoBox(false);
                                ShowImg(_ImgPath[_curidx]);
                                break;
                            }
                        }
                    }
                    else if (_curidx + 1 == _ImgPath.Count)
                    {
                        ClearAllDiseaseInfoBox(true);
                        _IsRoadAutoPlay = false;
                        MessageBox.Show("已经是最后一张自动识别有病害图像！");
                        break;
                    }
                }
            }
            else
            {
                LabelDistoImg();
                if (_curidx + 1 < _ImgPath.Count)
                {
                    ClearAllDiseaseInfoBox(false);
                    ShowImg(_ImgPath[++_curidx]);
                }
                else if (_curidx + 1 == _ImgPath.Count)
                {
                    ClearAllDiseaseInfoBox(true);
                    _IsRoadAutoPlay = false;
                    MessageBox.Show("已经是最后一张图像！");
                }
            }
        }
        public void ShowLastImg()
        {
            if (WinRoadDisList._BrowserType == 1)
            {
                while (true)
                {
                    if (_curidx > 0)
                    {
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
                        {
                            ClearAllDiseaseInfoBox(false);
                            ShowImg(_ImgPath[_curidx]);
                            break;
                        }
                    }
                    else if (_curidx == 0)
                    {
                        ClearAllDiseaseInfoBox(true);
                        MessageBox.Show("已经是第一张自动识别有病害图像！");
                        break;
                    }
                }
            }
            else if (WinRoadDisList._BrowserType == 2)
            {
                while (true)
                {
                    if (_curidx > 0)
                    {
                        if (!File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
                        {
                            ClearAllDiseaseInfoBox(false);
                            ShowImg(_ImgPath[_curidx]);
                            break;
                        }
                    }
                    else if (_curidx == 0)
                    {
                        ClearAllDiseaseInfoBox(true);
                        MessageBox.Show("已经是第一张自动识别无病害图像！");
                        break;
                    }
                }
            }
            else if (WinRoadDisList._BrowserType == 3)
            {
                while (true)
                {
                    if (_curidx > 0)
                    {
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
                        {
                            string strs = File.ReadAllText(string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[_curidx].imgpath));
                            if (strs.Contains(WinRoadDisList._BrowserDisName))
                            {
                                ClearAllDiseaseInfoBox(false);
                                ShowImg(_ImgPath[_curidx]);
                                break;
                            }
                        }
                    }
                    else if (_curidx == 0)
                    {
                        ClearAllDiseaseInfoBox(true);
                        MessageBox.Show("已经是第一张自动识别有病害图像！");
                        break;
                    }
                }
            }
            else
            {
                if (_curidx > 0)
                {
                    ClearAllDiseaseInfoBox(false);
                    ShowImg(_ImgPath[--_curidx]);
                }
                else if (_curidx == 0)
                {
                    ClearAllDiseaseInfoBox(true);
                    MessageBox.Show("已经是第一张图像！");
                }
            }
        }

        public void LabelDistoImg()
        {
            if (MainForm._IsSaveDisImg && DisInfoList.Count > 0)
            {
                Bitmap bit = new Bitmap(pictureBox_road.Width, pictureBox_road.Height);
                pictureBox_road.DrawToBitmap(bit, new Rectangle(new Point(0, 0), new Size(pictureBox_road.Width, pictureBox_road.Height)));

                string newpath = _ImgName.Substring(_ProjPath.Length + 1);
                newpath = newpath.Replace("RoadImg", "RoadDisImg");
                string[] s = newpath.Split('\\');
                string ts = _ProjPath;
                for (int i = 0; i < s.Length - 1; i++)
                {
                    ts = string.Format("{0}\\{1}", ts, s[i]);
                    if (!Directory.Exists(ts))
                    {
                        Directory.CreateDirectory(ts);
                    }
                }
                newpath = string.Format("{0}\\{1}", _ProjPath, newpath);
                bit.Save(newpath, System.Drawing.Imaging.ImageFormat.Jpeg);
                bit.Dispose();
            }
        }
        public void ShowJumpImg2(double jval)
        {
            int temp = Convert.ToInt32(jval / 2);
            if (temp >= 0 && temp < _ImgPath.Count)
            {
                ClearAllDiseaseInfoBox(false);
                _curidx = temp;
                ShowImg(_ImgPath[_curidx]);
            }
            else
            {
                MessageBox.Show("跳转桩号不在工程范围内！");
            }
        }
        override public void ShowJumpImg(double jval)
        {
            if (jval <= _ImgPath[0].imgmile && jval >= _ImgPath[_ImgPath.Count - 1].imgmile
                || jval >= _ImgPath[0].imgmile && jval <= _ImgPath[_ImgPath.Count - 1].imgmile)
            {
                int temp = BinSearch(jval, ref _ImgPath, _ProjectInfo._Direction);
                if (temp >= 0 && temp < _ImgPath.Count)
                {
                    ClearAllDiseaseInfoBox(false);
                    _curidx = temp;
                    ShowImg(_ImgPath[_curidx]);
                }
                else
                {
                    MessageBox.Show("跳转桩号不在工程范围内！");
                }
            }
            else
            {
                MessageBox.Show("跳转桩号不在工程范围内！");
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
        void WinRoadNew_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {

        }
        private int _idx = 0;
        private bool _IsDrawRect = false;

        private bool _IsDrawLineModel = false; //是否是画线模式

        //按键
        private void WinRoadNew_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.D))
            {
                for (int i = DisInfoList.Count - 1; i >= 0; i--)
                {
                    foreach (int idxTemp in DisInfoList[i].rectarry)
                    {
                        DelDis(idxTemp);
                    }
                }
            }
            if (e.KeyData == Keys.Space)
            {
                ShowNextImg();
            }
            else if (e.KeyData == Keys.Up)
            {
                ShowLastImg();
            }
            else if (e.KeyData == Keys.Down)
            {
                ShowNextImg();
            }
            else if (e.KeyData == Keys.Left)
            {
                ShowLastImg();
            }
            else if (e.KeyData == Keys.Right)
            {
                ShowNextImg();
            }
            else if (e.KeyData == Keys.Escape)
            {
                ShowLastImg();
            }
            else if (e.KeyData == Keys.D)
            {
                ChangeDrawModel(1,false);
                
               
            }
            else if (e.KeyData == Keys.B)
            {
                ChangeDrawModel(2,false);
                
                
            }
            else if (e.KeyData == Keys.N)
            {
                //按下N键 代表线性绘制模式当前用户选择的病害已经结束
                if (_IsDrawLineModel)
                {

                    lineModelPoints.Clear();
                    LineOneStepRectIdx.Clear();
                   
                         
                            mouseMove = false;
                            mouseStatus = false;
                            if (CurRectIdx.Count == 0 || !NewDrawDis()) //不保存
                            {
                                CurRectIdx.Clear();
                                pictureBox_road.Invalidate();
                            }
                            else
                            {
                                foreach (int idx in CurRectIdx)
                                {
                                    _PartRectInfos[idx].SetCheck(true);
                                }
                            }
                         
                        pictureBox_road.Invalidate();
                    
                }
                else
                {


                }
            }
            else if (e.KeyData == Keys.F1)
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
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //else if (e.KeyData == Keys.A)
            //{
            //    currRect = new Rectangle(2, 2, RoadImgRect.Width - 4, RoadImgRect.Height - 4);
            //    Rectangle imgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            //    if (JudgeRectArea(currRect, imgrect, rectList, true))
            //    {
            //        UpdateBox();
            //    }
            //    else if (!DrawDis(currRect))
            //    {
            //        currRect = Rectangle.Empty;
            //        this.pictureBox_road.Invalidate();
            //    }
            //}
            //////////////////////////////////////////////////////////////////////////
            else if (e.KeyData == Keys.Delete)
            {
                DelDis(_idx);
            }
        }

        private bool NewDrawDis()
        {
            this.pictureBox_road.Invalidate();
            RoadPavementPanel PavementDisease = new RoadPavementPanel(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis, _RoadType);
            PavementDisease.SetNumArea(CurRectIdx.Count, CurRectIdx.Count * 0.01);
            PavementDisease.ShowDialog();

            if (PavementDisease.IsDisease)//保存病害信息
            {
                String tempStack = "0";
                int stacknum = (int)Math.Round(_ImgPath[_curidx].imgmile);

                if (CurRectIdx.Count > 0)
                {
                    string dispos = string.Empty;
                    foreach (int tidx in CurRectIdx)
                    {
                        dispos = string.Format("{0}{1}-", dispos, tidx);
                    }
#if 辽宁建祥3m
               int splitY1 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 3;
               int splitY2 = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum *2/ 3;
               if (CurRectIdx[0] > splitY1 && CurRectIdx[0] < splitY2)
                {
                    stacknum = stacknum + _ProjectInfo._Direction;
                }
                else if (CurRectIdx[0]>splitY2)
                {
                    stacknum = stacknum + _ProjectInfo._Direction*2;
                }
#else
                    if (CurRectIdx[0] > (_RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum / 2))
                    {
                        stacknum = stacknum + _ProjectInfo._Direction;
                    }
#endif

                    tempStack = stacknum.ToString("K0000+000");

                    string saveinfo = string.Format("{0} 桩号:{1} {2} {3}", PavementDisease.RoadDiseaseType.ToString(), tempStack,
                        GlobalExcel._RoadTypeStr[_RoadType], dispos);
                    SmalRectDisease temp = new SmalRectDisease(saveinfo, stacknum);


                    if (temp.isDiseaseOK)
                    {
                        DisInfoList.Add(temp);
                        string disPathMessage = _ImgPath[_oldidx].imgpath;

                        if (_Setting.outHumanDeleteDisease)
                        {
                            using (StreamWriter sw = new StreamWriter(unAutoDisPath, true))
                            {
                                sw.WriteLine(disPathMessage);
                                //病害文本地址 病害
                                sw.WriteLine(temp.GetDisInfoStr());
                            }  //复制图片
                            string resourcePicPath = string.Format("{0}\\RoadImg\\Camera0\\{1}", _ProjPath, _ImgPath[_curidx].imgpath);
                            string recordPicPath = string.Format("{0}\\{1}", unAutoDisPath.Substring(0, unAutoDisPath.LastIndexOf('\\')), _ImgPath[_curidx].imgpath);
                            // 获取文件的上级目录路径
                            string directoryPath = Path.GetDirectoryName(recordPicPath);
                            try
                            {
                                // 创建上级目录（如果它们不存在）
                                Directory.CreateDirectory(directoryPath);

                                // 创建文件（如果它不存在）
                                if (!File.Exists(recordPicPath))
                                {
                                    File.Copy(resourcePicPath, recordPicPath);
                                }
                            }
                            catch (Exception ex)
                            {


                            }
                        }

                    }
                }
                this.pictureBox_road.Invalidate();
            }
            return PavementDisease.IsDisease;
        }
        private void DelDis(int idx)
        {
            if (idx < 0 || idx >= _PartRectInfos.Count)
                return;

            bool flag = false;
            foreach (SmalRectDisease dis in DisInfoList)
            {
                if (dis.rectarry.Contains(idx))
                {

                    string deleteTxt = string.Format("{0} 桩号:{1} {2} {3}\n{4}", dis.RoadDisType, dis.m_mile.ToString("K0+000"), dis.RoadType, dis.dispos, _ImgPath[_curidx].imgpath);
                    if (_Setting.outHumanDeleteDisease)
                    {
                        try
                        {
                            using (StreamWriter fs = new StreamWriter(deleteDisPath, true))
                            {
                                fs.WriteLine(deleteTxt);
                            }

                            string picFile = string.Format(@"{0}\RoadImg\Camera0{1}", _ProjPath, _ImgPath[_oldidx].imgpath);
                            string saveRecordPic = string.Format(@"{0}{1}", deleteDisPath.Substring(0, deleteDisPath.LastIndexOf("\\")), _ImgPath[_oldidx].imgpath);
                            // 获取文件的上级目录路径
                            string directoryPath = Path.GetDirectoryName(saveRecordPic);

                            // 创建上级目录（如果它们不存在）
                            Directory.CreateDirectory(directoryPath);

                            // 创建文件（如果它不存在）
                            if (!File.Exists(saveRecordPic))
                            {
                                File.Copy(picFile, saveRecordPic);
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    DisInfoList.Remove(dis);
                    foreach (int tidx in dis.rectarry)
                    {
                        _PartRectInfos[tidx].SetCheck(false);
                    }
                    flag = true;
                    break;
                }
            }
            if (flag)
            {
                this.pictureBox_road.Invalidate();
            }
        }
        //保存病害信息到txt
        private void ClearAllDiseaseInfoBox(bool isside)
        {
            if (_ImgPath.Count == 0) return;

            string recInfoFile = string.Format(@"{0}\RoadImg\Camera0{1}_PartClass.txt", _ProjPath, _ImgPath[_oldidx].imgpath);
            //如果上个图片病害数量大于0
            if (DisInfoList.Count > 0)
            {
                FileStream fw = File.Open(@recInfoFile, FileMode.Create);
                StreamWriter sw = new StreamWriter(fw);
                for (int i = 0; i < DisInfoList.Count; ++i)
                {
                    sw.WriteLine(DisInfoList[i].GetDisInfoStr(), Encoding.UTF8);
                }
                sw.Close();
                fw.Close();
                writeRGXKdis(false);
            }
            else
            {
                if (!MainForm._IsSaveDisImg)
                {
                    if (File.Exists(recInfoFile))
                    {
                        FileInfo tf = new FileInfo(recInfoFile);
                        string newdir = tf.Directory.FullName.Replace("\\RoadImg\\Camera0\\", "\\RoadImg\\DelCamera0\\");
                        if (!Directory.Exists(newdir))
                        {
                            Directory.CreateDirectory(newdir);
                        }
                        string fpath = string.Format("{0}\\{1}", newdir, tf.Name);
                        if (File.Exists(fpath))
                        {
                            File.Delete(fpath);
                        }
                        File.Move(recInfoFile, fpath);

                    }
                    writeRGXKdis(true);
                }
            }
            if (!isside)
            {
                DisInfoList.Clear();
                CurRectIdx.Clear();
            }
        }

        override public void UpdateDisType(object updateinfo)
        {
            Global._UpdateInfo tupdate = (Global._UpdateInfo)(updateinfo);

            DisInfoList[tupdate.disidx].RoadDisType = tupdate.disname;

            //重绘整个界面
            this.pictureBox_road.Invalidate();
        }

        public void WinRoadNew_FormClosed(object sender, FormClosedEventArgs e)
        {
            ClearAllDiseaseInfoBox(false);
        }

        private void button_jump_Click(object sender, EventArgs e)
        {
            double temp = 0;
            try
            {
                temp = Convert.ToDouble(textBox_mile.Text);
            }
            catch
            {
                return;
            }

            if (_mileval != temp)
            {
                ShowJumpImg(Convert.ToDouble(textBox_mile.Text));
                return;
            }

            try
            {
                temp = Convert.ToDouble(textBox_dmi.Text);
            }
            catch
            {
                return;
            }
            if (_dmival != temp)
            {
                ShowJumpImg2(Convert.ToDouble(textBox_dmi.Text));
                return;
            }
        }

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

        private void button_RoadType_Click(object sender, EventArgs e)
        {
            Button bt = sender as Button;
            if (_RoadType != Convert.ToInt16(bt.Tag))
            {
                AddRoadType(bt.Text);
            }
            else
            {
                MessageBox.Show(string.Format("当前已经是{0}路面！", bt.Text));
            }
        }

        private void AddRoadType(string type)
        {
            List<string> AllMarkInfo = new List<string>();
            string fname = _ProjPath + "\\RoadStatuMarkInfo.txt";
            if (File.Exists(fname))
            {
                string headstr = string.Format("{0} {0} {1} 路面材质:", _mileval, _dmival);
                string[] markstrs = File.ReadAllLines(fname);
                AllMarkInfo = new List<string>(markstrs);
                for (int i = 0; i < AllMarkInfo.Count; ++i)
                {
                    if (AllMarkInfo[i].Contains(headstr))
                    {
                        AllMarkInfo.RemoveAt(i);
                        --i;
                    }
                }
            }
            string str = string.Format("{0} {0} {1} 路面材质:{2}", _mileval, _dmival, type);
            AllMarkInfo.Add(str);

            File.WriteAllLines(fname, AllMarkInfo.ToArray(), Encoding.UTF8);
            MessageBox.Show("修改路面材质成功！");
            GetTypeMilePart(_ProjPath, _ProjectInfo._Direction);
            EventChangeType(null, null);
        }

        private void timer_roadplay_Tick(object sender, EventArgs e)
        {
            if (_IsRoadAutoPlay)
            {
                ShowNextImg();
            }
            else
            {
                timer_roadplay.Stop();
                button_play.ImageIndex = 7;
                _IsRoadAutoPlay = false;
            }
        }

        //路面材质的里程区间间隔
        override public void GetTypeMilePart(string projectpath, int direction)
        {
            _RoadPart.Clear();
            _RoadPart.Add(new MilePart() { mile = _ProjectInfo._StartMile, roadtype = _ProjectInfo._RoadType });

            //获取打标的信息
            string filename = projectpath + "\\RoadStatuMarkInfo.txt";
            if (File.Exists(filename))
            {
                string[] disinfo = File.ReadAllLines(filename, Encoding.UTF8);
                foreach (string line in disinfo)
                {
                    if (line.Contains("路面材质"))
                    {
                        MilePart tpart = new MilePart();
                        string[] s = line.Split(" :".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                        int dmival = Convert.ToInt32(s[2]);
                        tpart.mile = _ProjectInfo.Dmi2Mile(dmival);
                        if (s[s.Length - 1] == "沥青" || s[s.Length - 1] == "水泥" || (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing && s[s.Length - 1] == "砂石")
                            || (_Setting.ParmStyle == StandardParmType.RuralRoadlowLevel && s[s.Length - 1] == "砂石")
                            || (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && s[s.Length - 1] == "砂石")
                            )
                        {
                            tpart.roadtype = RoadDiseaseTypes.roadtypedict[s[s.Length - 1]];

                            if (tpart.mile == _RoadPart[0].mile)
                            {
                                _RoadPart[0].roadtype = tpart.roadtype;
                            }
                            else
                            {
                                _RoadPart.Add(tpart);
                            }
                        }
                    }
                }
            }
            _RoadPart.Add(new MilePart() { mile = _ProjectInfo._EndMile, roadtype = _RoadPart[_RoadPart.Count - 1].roadtype });

            if (direction > 0)//升序
            {
                _RoadPart.Sort(delegate (MilePart x, MilePart y) { return x.mile.CompareTo(y.mile); });
            }
            else if (direction < 0)//降序
            {
                _RoadPart.Sort(delegate (MilePart x, MilePart y) { return y.mile.CompareTo(x.mile); });
            }
            if (_curidx < _ImgPath.Count)
            {
                ShowImg(_ImgPath[_curidx]);
            }
        }

        private void label_imgpath_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowDefalutSystemImg(label_imgpath.Text);
        }

        private void ShowDefalutSystemImg(string fpath)
        {
            //建立新的系统进程                
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            //设置图片的真实路径和文件名                
            process.StartInfo.FileName = fpath;
            //设置进程运行参数，这里以最大化窗口方法显示图片。               
            process.StartInfo.Arguments = "rundl132.exe C://WINDOWS//system32//shimgvw.dll,ImageView_Fullscreen";
            //此项为是否使用Shell执行程序，因系统默认为true，此项也可不设，但若设置必须为true                
            process.StartInfo.UseShellExecute = true;
            //此处可以更改进程所打开窗体的显示样式，可以不设               
            process.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            process.Start();
            process.Close();
        }

        private readonly static object lock1 = new object();

       private void ChangeDrawModel(int index,bool isSelectCall)
        {
            this.Cursor = System.Windows.Forms.Cursors.Default;
            
            if (!isSelectCall)
            { 
                if (index == 0)
                {
                    _IsDrawRect = false;
                    _IsDrawLineModel = false;
                }
                else if (index == 1)
                {
                    if (_IsDrawRect)
                    {
                        _IsDrawRect = false;
                        drawModel_Combox.SelectedIndex = 0; //重置下拉框
                    }
                    else
                    {
                        _IsDrawRect = true;
                        _IsDrawLineModel = false;
                        drawModel_Combox.SelectedIndex = index;
                        this.Cursor = System.Windows.Forms.Cursors.Cross;
                    }
                }
                else if (index == 2)
                {
                    if (_IsDrawLineModel)
                    {
                        _IsDrawLineModel = false;
                        drawModel_Combox.SelectedIndex = 0; //重置下拉框
                    }
                    else
                    {
                        _IsDrawLineModel = true;
                        _IsDrawRect = false;
                        drawModel_Combox.SelectedIndex = index;
                        this.Cursor = System.Windows.Forms.Cursors.Cross;
                    }
                }
               
            }
            else
            {
                if (index == 0 )
                {
                    _IsDrawRect = false;
                    _IsDrawLineModel = false;
                }
                else if (index == 1)
                { 
                        _IsDrawRect = true;
                        drawModel_Combox.SelectedIndex = index;
                        this.Cursor = System.Windows.Forms.Cursors.Cross; 
                }
                else if (index == 2)
                {
                   
                        _IsDrawLineModel = true;
                        drawModel_Combox.SelectedIndex = index;
                        this.Cursor = System.Windows.Forms.Cursors.Cross; 
                }
               
            }
            _Setting.SmallDiseaseDrawType = index;
        }


        private void drawModel_Combox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeDrawModel(drawModel_Combox.SelectedIndex,true);
        }

        /// <summary>
        /// 写人工小框病害
        /// </summary>
        /// <param name="clear">false 上张图片还有病害</param>
        private void writeRGXKdis(bool clear)
        {
            try
            {
                List<SmalRectDisease> DisInfoListCopy = new List<SmalRectDisease>();
                DisInfoList.ForEach(i => DisInfoListCopy.Add(i));
                Task.Run(() =>
                {
                    lock (lock1)
                    {
                        if (_Setting.outHumanDeleteDisease)
                        {
                            string[] strArray = null;
                            if (File.Exists(unAutoDisPath))
                                strArray = File.ReadAllLines(unAutoDisPath);
                            else
                                return;
                            List<string> realStr = new List<string>();

                            if (strArray != null)
                            {
                                if (clear)
                                {
                                    //上张图片没有病害了表示都被用户删除了
                                    for (int i = 0; i < strArray.Length; i += 2)
                                    {
                                        if (!strArray[i].Equals(_ImgPath[_oldidx].imgpath))
                                        {
                                            realStr.Add(strArray[i]);
                                            realStr.Add(strArray[i + 1]);
                                        }
                                        else
                                        {
                                            realStr.Add(strArray[i]);
                                            realStr.Add("用户绘制后清除");
                                        }
                                    }
                                }
                                else
                                {

                                    for (int i = 0; i < strArray.Length; i += 2)
                                    {
                                        if (strArray[i].Equals(_ImgPath[_oldidx].imgpath))
                                        {
                                            foreach (var disInfo in DisInfoListCopy)
                                            {
                                                if (disInfo.GetDisInfoStr().Equals(strArray[i + 1]))
                                                {
                                                    realStr.Add(strArray[i]);
                                                    realStr.Add(strArray[i + 1]);
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            realStr.Add(strArray[i]);
                                            realStr.Add(strArray[i + 1]);

                                        }
                                    }
                                }
                                File.Delete(unAutoDisPath);
                                using (StreamWriter sw1 = new StreamWriter(unAutoDisPath, true))
                                {
                                    foreach (var item in realStr)
                                    {
                                        sw1.WriteLine(item);
                                    }

                                }
                            }
                        }

                    }
                }

                    );


            }
            catch (Exception ex)
            {


            }

        }
    }
}
