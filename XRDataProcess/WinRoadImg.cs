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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using DevExpress.Map.Native; 
using AutoMapper.Mappers;

namespace XRDataProcess
{
    public partial class WinRoadImg : WinRoad
    //public partial class WinRoadImg : Form
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
        //图片通道信息
        private int picPassagewayType = 1;
        private List<Rectangle> rectList = new List<Rectangle>();//存储所有画过的矩形     
        private Point startPoint;
        private Point endPoint;
        private Rectangle currRect;
        private int minStartX, minStartY, maxEndX, maxEndY;//最大重绘矩形的上下左右的坐标，这样重绘的效率更高。
        List<Disease> DisInfoList = new List<Disease>();
        private List<MilePart> _RoadPart = new List<MilePart>();

        ColorPalette m_palette;
        private Bitmap _image = null;
        private BitmapData m_OriData = null;
        private Bitmap m_NewImg = null;
        private BitmapData m_NewData = null;

        public static string _ImgName = null;

        private double _dmival = 0;
        private double _mileval = 0;
        private bool _IsKeyRepair = false;

        private Rectangle RoadImgRect = new Rectangle();
        private int currRectIndex = -1;
        /// <summary>
        /// 人工病害地址
        /// </summary>
        private string unAutoDisPath;
        private string deleteDisPath;
        /// <summary>
        /// 坐标点经度纬度高程
        /// </summary>
        string latitude ="";
        string longitude = "" ;
        string centerH = "";
         
        public WinRoadImg(ProjectInfo pinfo, string ppath)
        {
            InitializeComponent();
            _ProjectInfo = pinfo;
            _ProjPath = ppath;
            string projectName = pinfo._PrjPath.Split('\\').Last();
            //大框人工病害记录文本
            if (Directory.Exists( _Setting.outHumanDeleteDiseasePath))
            {
                try
                {
                    unAutoDisPath = string.Format(@"{0}\{1}\HumanBigDisMessage.txt", _Setting.outHumanDeleteDiseasePath, projectName);
                    // 获取文件的上级目录路径
                    string directoryPath = Path.GetDirectoryName(unAutoDisPath);

                    // 创建上级目录（如果它们不存在）
                    Directory.CreateDirectory(directoryPath);

                    // 创建文件（如果它不存在）
                    if (!File.Exists(unAutoDisPath))
                    {
                        File.Create(unAutoDisPath).Close(); 
                    } 
                    deleteDisPath = string.Format(@"{0}\{1}\deleteDisPath.txt", _Setting.outHumanDeleteDiseasePath, projectName);
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
           
           
            _RoadType = _ProjectInfo._RoadType;

            _ImgPath = new List<MyImgMile>();

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

        private void WinRoadImg_Load(object sender, EventArgs e)
        {
            //EventSetYGPalette(m_palette, EventArgs.Empty);
            lblCoordinates = new Label();

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
            //控制砂石路面有无
            if (_Setting.ParmStyle == StandardParmType.RuralRoadChongqing ||
                _Setting.ParmStyle == StandardParmType.RuralRoadlowLevel||
                _Setting.ParmStyle == StandardParmType.RuralRoadHunan
                /*_Setting.ParmStyle== StandardParmType.RuralRoadHunan2024*/)
                button_SS.Visible = true;
            else
                button_SS.Visible = false;
            
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
            _ImgName = string.Format(@"{0}\RoadImg\Camera0{1}", _ProjPath, path.imgpath);
            label_imgpath.Text = _ImgName;
            if (MainForm._IsSaveDisImg)
            {
                if (!File.Exists(_ImgName + ".txt"))
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

                }
                else
                {
                    _image = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                    m_NewImg = new Bitmap(_RoadConfig.ImageWidth, _RoadConfig.ImageHeight, PixelFormat.Format8bppIndexed);
                    _image.Palette = m_palette;
                    m_NewImg.Palette = m_palette;
                }
            }
          //  if (!_Setting.hasCamsetting)
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
            // _image.Save("D:\\tst\\cshape0.jpg", ImageFormat.Jpeg);


            EventUpdateYG(_image, EventArgs.Empty);
            EventUpdateFullImg(_image, EventArgs.Empty);

            LoadRecInfo(_ImgName + ".txt");

            currRect = Rectangle.Empty;
            _oldidx = _curidx;
        }
        internal static IntPtr ArrayToIntptr(byte[] source)
        {
            if (source == null)
                return IntPtr.Zero;
            byte[] da = source;
            IntPtr ptr = Marshal.AllocHGlobal(da.Length);
            Marshal.Copy(da, 0, ptr, da.Length);
            return ptr;
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
                    Disease temp = new Disease(strline, (int)Math.Round(_ImgPath[_curidx].imgmile));

#if 辽宁建祥3m

                    int split1 = (_RoadConfig.ImageHeight - temp.rect.Height) / 3;
                    int split2 = (_RoadConfig.ImageHeight - temp.rect.Height) * 2 / 3;
                    if (temp.rect.Y >split1 && temp.rect.Y<split2)
                    {
                        temp.m_mile = temp.m_mile + _ProjectInfo._Direction;
                    }
                    else if (temp.rect.Y>split2)
                    {
                        temp.m_mile = temp.m_mile + _ProjectInfo._Direction*2;
                    }
#else
                    if (!temp.RoadDisType.Contains("破碎板")&&!temp.RoadDisType.Contains("松散")&&!temp.RoadDisType.Contains("露骨"))
                    {
                        if (temp.rect.Y > (_RoadConfig.ImageHeight - temp.rect.Height) / 2)
                        {
                            temp.m_mile = temp.m_mile + _ProjectInfo._Direction;
                        }
                    }
                    else
                    {
                        if (temp.Area <= _RoadConfig.DetectWidth * 2 * 2 / 3)
                        {
                            if (temp.rect.Y > (_RoadConfig.ImageHeight - temp.rect.Height) / 2)
                            {
                                temp.m_mile = temp.m_mile + _ProjectInfo._Direction;
                            }
                        }

                    }
#endif


                    DisInfoList.Add(temp);
                    rectList.Add(Img2Box(temp.rect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight));
                }
                sr.Close();
                fr.Close();
            }
        }

        #region 拉框相关鼠标消息响应函数
        public const int MinLableDistance = 12;
        public Point LableStartPoint, LableEndPoint;
        public bool LableMoveFlag = false;
        private void roadlableLeftTop_MouseDown(object sender, MouseEventArgs e)
        {
            LableStartPoint.X = e.X;
            LableStartPoint.Y = e.Y;
            minStartX = e.X;
            minStartY = e.Y;
            maxEndX = e.X;
            maxEndY = e.Y;
            LableMoveFlag = true;
        }
        private void roadlableLeftTop_MouseMove(object sender, MouseEventArgs e)
        {
            if (LableMoveFlag)
            {
                roadlableLeftTop_Move(e);
            }
        }
        private void roadlableLeftTop_MouseUp(object sender, MouseEventArgs e)
        {
            LableMoveFlag = false;
            roadlableLeftTop_Move(e);
        }
        private void roadlableLeftTop_Move(MouseEventArgs e)
        {
            LableEndPoint.X = e.X;
            LableEndPoint.Y = e.Y;
            //currRect = new Rectangle(currRect.Location.X, currRect.Y + (LableEndPoint.Y - LableStartPoint.Y),
            //        currRect.Width, currRect.Height - (LableEndPoint.Y - LableStartPoint.Y)); 
            //LableMove(currRect);

            if (roadlableLeftTop.Location.X + (LableEndPoint.X - LableStartPoint.X) + MinLableDistance < roadlableRightTop.Location.X
                && roadlableLeftTop.Location.Y + (LableEndPoint.Y - LableStartPoint.Y) + MinLableDistance < roadlableLeftBottom.Location.Y)
            {
                roadlableLeftTop.Location = new Point(roadlableLeftTop.Location.X + (LableEndPoint.X - LableStartPoint.X),
                    roadlableLeftTop.Location.Y + (LableEndPoint.Y - LableStartPoint.Y));
                roadlableRightTop.Location = new Point(roadlableRightTop.Location.X, roadlableLeftTop.Location.Y);
                roadlableLeftBottom.Location = new Point(roadlableLeftTop.Location.X, roadlableLeftBottom.Location.Y);
                LableMove();
            }
        }

        private void roadlableRightTop_MouseDown(object sender, MouseEventArgs e)
        {
            LableStartPoint.X = e.X;
            LableStartPoint.Y = e.Y;
            LableMoveFlag = true;
        }
        private void roadlableRightTop_MouseMove(object sender, MouseEventArgs e)
        {
            if (LableMoveFlag)
            {
                roadlableRightTop_Move(e);
            }
        }
        private void roadlableRightTop_MouseUp(object sender, MouseEventArgs e)
        {
            LableMoveFlag = false;
            roadlableRightTop_Move(e);
        }
        private void roadlableRightTop_Move(MouseEventArgs e)
        {
            LableEndPoint.X = e.X;
            LableEndPoint.Y = e.Y;
            //currRect = new Rectangle(currRect.Location.X, currRect.Y,
            //    currRect.Width + (LableEndPoint.X - LableStartPoint.X), currRect.Height);
            //LableMove(currRect);

            if (roadlableLeftTop.Location.X + MinLableDistance < roadlableRightTop.Location.X + (LableEndPoint.X - LableStartPoint.X)
                && roadlableRightTop.Location.Y + (LableEndPoint.Y - LableStartPoint.Y) + MinLableDistance < roadlableRightBottom.Location.Y)
            {
                roadlableRightTop.Location = new Point(roadlableRightTop.Location.X + (LableEndPoint.X - LableStartPoint.X),
                    roadlableRightTop.Location.Y + (LableEndPoint.Y - LableStartPoint.Y));
                roadlableLeftTop.Location = new Point(roadlableLeftTop.Location.X, roadlableRightTop.Location.Y);
                roadlableRightBottom.Location = new Point(roadlableRightTop.Location.X, roadlableRightBottom.Location.Y);
                LableMove();
            }
        }

        private void roadlableRightBottom_MouseDown(object sender, MouseEventArgs e)
        {
            LableStartPoint.X = e.X;
            LableStartPoint.Y = e.Y;
            LableMoveFlag = true;
        }
        private void roadlableRightBottom_MouseMove(object sender, MouseEventArgs e)
        {
            if (LableMoveFlag)
            {
                roadlableRightBottom_Move(e);
            }
        }
        private void roadlableRightBottom_MouseUp(object sender, MouseEventArgs e)
        {
            LableMoveFlag = false;
            roadlableRightBottom_Move(e);
        }
        private void roadlableRightBottom_Move(MouseEventArgs e)
        {
            LableEndPoint.X = e.X;
            LableEndPoint.Y = e.Y;
            //currRect = new Rectangle(currRect.Location.X, currRect.Y,
            //    currRect.Width, currRect.Height + (LableEndPoint.Y - LableStartPoint.Y));
            //LableMove(currRect);

            if (roadlableLeftTop.Location.X + MinLableDistance < roadlableRightBottom.Location.X + (LableEndPoint.X - LableStartPoint.X)
                && roadlableRightTop.Location.Y + MinLableDistance < roadlableRightBottom.Location.Y + (LableEndPoint.Y - LableStartPoint.Y))
            {
                roadlableRightBottom.Location = new Point(roadlableRightBottom.Location.X + (LableEndPoint.X - LableStartPoint.X),
                    roadlableRightBottom.Location.Y + (LableEndPoint.Y - LableStartPoint.Y));
                roadlableRightTop.Location = new Point(roadlableRightBottom.Location.X, roadlableRightTop.Location.Y);
                roadlableLeftBottom.Location = new Point(roadlableLeftBottom.Location.X, roadlableRightBottom.Location.Y);
                LableMove();
            }
        }

        private void roadlableLeftBottom_MouseDown(object sender, MouseEventArgs e)
        {
            LableStartPoint.X = e.X;
            LableStartPoint.Y = e.Y;
            LableMoveFlag = true;
        }
        private void roadlableLeftBottom_MouseMove(object sender, MouseEventArgs e)
        {
            if (LableMoveFlag)
            {
                roadlableLeftBottom_Move(e);
            }
        }
        private void roadlableLeftBottom_MouseUp(object sender, MouseEventArgs e)
        {
            LableMoveFlag = false;
            roadlableLeftBottom_Move(e);
        }
        private void roadlableLeftBottom_Move(MouseEventArgs e)
        {
            LableEndPoint.X = e.X;
            LableEndPoint.Y = e.Y;
            //currRect = new Rectangle(currRect.Location.X + (LableEndPoint.X - LableStartPoint.X), currRect.Y,
            //    currRect.Width - (LableEndPoint.X - LableStartPoint.X), currRect.Height);
            //LableMove(currRect);

            if (roadlableLeftBottom.Location.X + (LableEndPoint.X - LableStartPoint.X) + MinLableDistance < roadlableRightBottom.Location.X
                && roadlableLeftTop.Location.Y + MinLableDistance < roadlableLeftBottom.Location.Y + (LableEndPoint.Y - LableStartPoint.Y))
            {
                roadlableLeftBottom.Location = new Point(roadlableLeftBottom.Location.X + (LableEndPoint.X - LableStartPoint.X),
                    roadlableLeftBottom.Location.Y + (LableEndPoint.Y - LableStartPoint.Y));
                roadlableLeftTop.Location = new Point(roadlableLeftBottom.Location.X, roadlableLeftTop.Location.Y);
                roadlableRightBottom.Location = new Point(roadlableRightBottom.Location.X, roadlableLeftBottom.Location.Y);
                LableMove();
            }
        }

        private void LableMove()
        {
            Rectangle tmpRect = new Rectangle(roadlableLeftTop.Location.X + 3,
                roadlableLeftTop.Location.Y + 3,
                Math.Abs(roadlableRightTop.Location.X - roadlableLeftTop.Location.X),
                Math.Abs(roadlableLeftBottom.Location.Y - roadlableLeftTop.Location.Y));

            Rectangle imgrect = Box2Img(tmpRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            if (JudgeRectArea(tmpRect, imgrect, rectList, false))
            {
                UpdateBox();
            }
            else if (currRectIndex >= 0)
            {
                currRect = tmpRect;
                DisInfoList[currRectIndex].rect = imgrect;
                UpdateAllRectSaveInfo();
            }

        }

        //private void LableMove(Rectangle trect)
        //{
        //    Rectangle imgrect = Box2Img(trect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
        //    if (JudgeRectArea(trect, imgrect, rectList, false))
        //    {
        //        UpdateBox();
        //    }
        //    else
        //    {
        //        DisInfoList[currRectIndex].rect = imgrect;
        //        UpdateAllRectSaveInfo();
        //    }
        //}
        #endregion

        /// <summary>
        /// 新矩形不合法时，不用更新返回ture，需要更新的时候返回false
        /// </summary>
        /// <param name="tRect"></param>
        /// <param name="imgRect"></param>
        /// <param name="rectList"></param>
        /// <param name="IsNew"></param>
        /// <returns></returns>
        private bool JudgeRectArea(Rectangle tRect, Rectangle imgRect, List<Rectangle> rectList, bool IsNew)
        {
            bool flag = false;
            if (!(0 <= imgRect.Top && imgRect.Top < imgRect.Bottom && imgRect.Bottom <= _RoadConfig.ImageHeight
                && 0 <= imgRect.Left && imgRect.Left < imgRect.Right && imgRect.Right <= _RoadConfig.ImageWidth))
            {
                flag = true;
            }
            else
            {
                if (RoadImgRect.Contains(tRect))
                {
                    //for (int i = 0; i < rectList.Count; ++i)
                    //{
                    //    if (IsNew)
                    //    {
                    //        if ((tRect.Contains(rectList[i]) || rectList[i].Contains(tRect)))
                    //        {
                    //            flag = true;
                    //            break;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        if (i != currRectIndex && (tRect.Contains(rectList[i]) || rectList[i].Contains(tRect)))
                    //        {
                    //            flag = true;
                    //            break;
                    //        }
                    //    }
                    //}
                }
                else
                {
                    flag = true;
                }
            }
            return flag;
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
            int len = rectList.Count;
            for (int i = 0; i < len; ++i)
            {
                rectList[i] = Img2Box(DisInfoList[i].rect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            }

            Graphics g = e.Graphics;
            Pen m_pen1 = new Pen(Color.Red, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
            Pen m_pen2 = new Pen(Color.Blue, 3) { DashStyle = System.Drawing.Drawing2D.DashStyle.Solid };

            if (currRect != Rectangle.Empty)
            {
                //SetRectLable(new Point(currRect.Left + currRect.Width / 2 - 3, currRect.Top - 3),
                //    new Point(currRect.Left + currRect.Width / 2 - 3, currRect.Bottom - 3),
                //    new Point(currRect.Left - 3, currRect.Top + currRect.Height / 2 - 3),
                //    new Point(currRect.Right - 3, currRect.Top + currRect.Height / 2 - 3)); 

                //SetRectLable(new Point(currRect.Left - 3, currRect.Top - 3),
                //    new Point(currRect.Left - 3, currRect.Bottom - 3),
                //    new Point(currRect.Right - 3, currRect.Top - 3),
                //    new Point(currRect.Right - 3, currRect.Bottom - 3));
                g.DrawRectangle(m_pen2, currRect);

            }

            int ii = 0;
            foreach (Rectangle rect in rectList)
            {
                g.DrawRectangle(m_pen1, rect);
                g.DrawString(GlobalExcel._RoadTypeStr[_RoadType] + "." + DisInfoList[ii++].GetRectInfoStr(),
                    new Font("宋体", 10, FontStyle.Regular),
                    Brushes.GreenYellow, rect.Location.X, rect.Location.Y);
            }
            EventUpdateDis(DisInfoList, null);
        }

        private void pictureBox_road_Resize(object sender, EventArgs e)
        {
            if (pictureBox_road.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                RoadImgRect.X = 0;
                RoadImgRect.Y = 0;
                RoadImgRect.Width = pictureBox_road.Width;
                RoadImgRect.Height = pictureBox_road.Height;
            }
        }

        private void pictureBox_road_MouseDown(object sender, MouseEventArgs e)
        {
           
            //按鼠标左键，勾画病害
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                DrawDisMouseDown(e);
            }
        }
        private Label lblCoordinates;
        private int curPosX = 0;
        private int curPosY = 0;
        private void pictureBox_road_MouseMove(object sender, MouseEventArgs e)
        {

            if (e.Location.X > 0 && e.Location.X < pictureBox_road.Width && e.Location.Y > 0 && e.Location.Y < pictureBox_road.Height)
            {
                EventUpdateFullPoint(BoxPoint2RectPoint(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight), EventArgs.Empty);
            }
            EventUpdateFullImg(pictureBox_road.Image, EventArgs.Empty);
            //按鼠标左键，勾画病害    
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                DrawDisMouseMove(e);
            }

            try
            {
                if (_Setting.showGpsInfoToPicture)
                {
                    PictureBox pictureBox = sender as PictureBox;
                    
                    if (pictureBox != null)
                    {
                        int pictureW = _RoadConfig.ImageWidth;
                          int pictureH = _RoadConfig.ImageHeight;
                        // 获取鼠标相对于控件左上角的坐标
                        Point mouseLocation = e.Location;

                        // 获取鼠标相对于图片左上角的像素点坐标
                        int xRelativeToImage = (int)(((double)mouseLocation.X / pictureBox.Width) * pictureBox.Image.Width);
                        int yRelativeToImage = (int)(((double)mouseLocation.Y / pictureBox.Height) * pictureBox.Image.Height);
                        curPosX = xRelativeToImage;
                        curPosY = yRelativeToImage;
                        double dDiseaseLon = 0, dDiseaseLat = 0, dDiseaseH = 0; //当前像素
                        // 现在，xRelativeToImage 和 yRelativeToImage 分别为鼠标相对于图片左上角的像素点坐标
                        // 可以在这里使用这些坐标进行进一步的操作
                        HighAccuracyPositioning.getHighAccPosition(_Setting.gpsformat,_Setting.equipType, _ProjPath,_ImgPath, _curidx, curPosX,curPosY,pictureW, pictureH, _RoadConfig.RealWidth, _RoadConfig.RealHeight,
                             ref  dDiseaseLon, ref dDiseaseLat, ref  dDiseaseH
                            ); 

                        //Point screenPos = BoxPoint2RectPoint(e.Location, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                        lblCoordinates.BackColor = System.Drawing.Color.White;
                        lblCoordinates.ForeColor = System.Drawing.Color.Red;
                        // 模拟获取经纬度信息，这里可以替换为你实际获取经纬度的逻辑
                        latitude = dDiseaseLat.ToString();
                        longitude = dDiseaseLon.ToString();
                        centerH = dDiseaseH.ToString();

                        lblCoordinates.Text = $"坐标：({curPosX},{curPosY})\n经度：{longitude + "°E"}\n纬度：{latitude + "°N"}\n高程：{centerH + " M"}";
                        // 设置 Label 的位置在鼠标右下角

                        // 获取标签的尺寸
                        int labelWidth = lblCoordinates.Width;
                        int labelHeight = lblCoordinates.Height;

                        // 获取当前窗口（widget）的尺寸
                        int widgetWidth =this.Width;
                        int widgetHeight = this.Height;

                        // 定义基础偏移量
                        const int offsetX = 10;
                        const int offsetY = 15;

                        // 计算四个可能的位置，优先选择右下角
                        int targetX = e.X + offsetX;        // 右下角 X
                        int targetY =e.Y + offsetY;        // 右下角 Y

                        // 如果右下角会超出右边界 → 改成鼠标左侧
                        if (targetX + labelWidth > widgetWidth)
                        {
                            targetX = e.X - labelWidth - offsetX;
                        }

                        // 如果右下角会超出下边界 → 改成鼠标上方
                        if (targetY + labelHeight > widgetHeight)
                        {
                            targetY = e.Y- labelHeight - offsetY;
                        }

                        // 额外保险：防止左上角也超出（鼠标在左上角时）
                        if (targetX < 0) targetX = 0;
                        if (targetY < 0) targetY = 0;

                        lblCoordinates.Location = new Point(targetX, targetY); // 可根据实际情况调整偏移位置
                        lblCoordinates.AutoSize = true;
                        lblCoordinates.Visible = true;
                        this.panel1.Controls.Add(lblCoordinates);
                        this.lblCoordinates.BringToFront();
                        this.pictureBox_road.Invalidate();

                    }
                  
                }
                else
                {
                    lblCoordinates.Visible = false;
                }
            }
            catch (Exception)
            {

                //throw;
            }

        }
       


        private void pictureBox_road_MouseUp(object sender, MouseEventArgs e)
        {
            //按鼠标左键，勾画病害    
            bool IsUpdate = false;
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                DrawDisMouseUp(e, ref IsUpdate, _Setting.IsRepair && _IsKeyRepair);
            }

            if (IsUpdate)
            {
                currRect = Rectangle.Empty;
                this.pictureBox_road.Invalidate();
                SetRectLableVisible(false);
            }
        }

        private void DrawDisMouseDown(MouseEventArgs e)
        {
            startPoint.X = e.X;
            startPoint.Y = e.Y;            //重新一个矩形，重置最大重绘矩形的上下左右的坐标   
            minStartX = e.X;
            minStartY = e.Y;
            maxEndX = e.X;
            maxEndY = e.Y;

            if (!RoadImgRect.Contains(startPoint))
            {
                currRect = Rectangle.Empty;
                MessageBox.Show("超出图像区域！");
                this.pictureBox_road.Invalidate();
                SetRectLableVisible(false);
                return;
            }
            else
            {
                if (_Setting.IsForbidOverLapping)
                {
                    for (int i = 0; i < rectList.Count; ++i)
                    {
                        if (rectList[i].Contains(startPoint))
                        {
                            MessageBox.Show("所选病害与已有病害有重叠部分，请重新勾画病害！");
                            this.pictureBox_road.Invalidate();
                            SetRectLableVisible(false);
                            return;
                        }
                    }
                }
            }
        }
        private void DrawDisMouseMove(MouseEventArgs e)
        {
            endPoint.X = e.X;
            endPoint.Y = e.Y; //这一段是获取要绘制矩形的上下左右的坐标，如果不这样处理的话，只有从左上开始往右下角才能画出矩形。       
            //这样处理的话，可以任意方向，当然中途可以更换方向。  
            int realStartX = Math.Min(startPoint.X, endPoint.X);
            int realStartY = Math.Min(startPoint.Y, endPoint.Y);
            int realEndX = Math.Max(startPoint.X, endPoint.X);
            int realEndY = Math.Max(startPoint.Y, endPoint.Y);
            minStartX = Math.Min(minStartX, realStartX);
            minStartY = Math.Min(minStartY, realStartY);
            maxEndX = Math.Max(maxEndX, realEndX);
            maxEndY = Math.Max(maxEndY, realEndY);
            currRect = new Rectangle(realStartX, realStartY, realEndX - realStartX, realEndY - realStartY);
            //以下是为了获取最大重绘矩形。
            Rectangle refeshRect = new Rectangle(minStartX, minStartY, maxEndX - minStartX, maxEndY - minStartY);
            refeshRect.Inflate(1, 1);//重绘矩形的大小扩展1个单位 
            this.pictureBox_road.Invalidate(refeshRect);//失效一个区域，并使其重绘。
            SetRectLableVisible(true);
        }
        private void DrawDisMouseUp(MouseEventArgs e, ref bool IsUpdate, bool IsRepair)
        {
            endPoint.X = e.X;
            endPoint.Y = e.Y;
            int realStartX = Math.Min(startPoint.X, endPoint.X);
            int realStartY = Math.Min(startPoint.Y, endPoint.Y);
            int realEndX = Math.Max(startPoint.X, endPoint.X);
            int realEndY = Math.Max(startPoint.Y, endPoint.Y);
            if (Math.Abs(realStartY - realEndY) < 3 && Math.Abs(realStartX - realEndX) < 3)
            {
                UpdateBox();
                return;
            }

            try
            {
                if (!RoadImgRect.Contains(new Point(realStartX, realStartY))
                        || !RoadImgRect.Contains(new Point(realEndX, realEndY)))
                {
                    int scale = 2;
                    //IsUpdate = true;
                    //MessageBox.Show("超出图像区域！");
                    if (realEndX > RoadImgRect.Width && realEndY < RoadImgRect.Height && realStartY > 0)// 右侧中间外部
                    {
                        currRect = new Rectangle(realStartX, realStartY, RoadImgRect.Width - realStartX - scale, realEndY - realStartY - scale);
                    }
                    else if (realStartY < 0 && realEndX > RoadImgRect.Width) //右上侧
                    {
                        currRect = new Rectangle(realStartX, 0, RoadImgRect.Width - realStartX - scale, realEndY - scale);
                    }
                    else if (realEndX > RoadImgRect.Width && realEndY > RoadImgRect.Height)// 右下侧外部
                    {
                        currRect = new Rectangle(realStartX, realStartY, RoadImgRect.Width - realStartX - scale, RoadImgRect.Height - realStartY - scale);
                    }
                    else if (realStartX <= 0 && realStartY >= 0 && realEndY > 0 && realEndY < RoadImgRect.Height) //左侧中间外部
                    {
                        currRect = new Rectangle(0, realStartY, realEndX, realEndY - realStartY);
                    }
                    else if (realStartX <= 0 && realEndY >= RoadImgRect.Height)//左下侧外部
                    {
                        currRect = new Rectangle(0, realStartY, realEndX - scale, RoadImgRect.Height - realStartY - scale);
                    }
                    else if (realStartX > 0 && realStartY <= 0) //上侧外部
                    {
                        currRect = new Rectangle(realStartX, 0, realEndX - realStartX - scale, realEndY - scale);
                    }
                    else if (realStartX <= 0 && realStartY <= 0) //左上侧外部
                    {
                        currRect = new Rectangle(0, 0, realEndX - scale, realEndY - scale);
                    }
                    else if (realEndX > 0 && realEndY > RoadImgRect.Height)//下侧外部
                    {
                        currRect = new Rectangle(realStartX, realStartY, realEndX - realStartX, RoadImgRect.Height - realStartY - scale);
                    }

                    if (!DrawDis(currRect))
                    {
                        IsUpdate = true;
                    }
                }
                else if ((realEndX - realStartX) >= 10 && (realEndY - realStartY) >= 10)
                //else if ((realEndX - realStartX) >= 1 && (realEndY - realStartY) >= 1)
                {
                    currRect = new Rectangle(realStartX, realStartY, realEndX - realStartX, realEndY - realStartY);
                    Rectangle imgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                    if (JudgeRectArea(currRect, imgrect, rectList, true))
                    {
                        UpdateBox();
                    }
                    else
                    {
                        if (IsRepair)
                        {
                            SaveRectangular("修补", _RoadType);
                            IsUpdate = true;
                        }
                        else
                        {
                            if (_Setting.ParmStyle == StandardParmType.DegreeRoad2001)
                            {
                                if (!DrawDis2001(currRect))
                                {
                                    IsUpdate = true;
                                }
                            }
                            else
                            {
                                if (!DrawDis(currRect))
                                {
                                    IsUpdate = true;
                                }
                            }
                        }
                    }

                }
                else
                {
                    IsUpdate = true;
                }
                SetRectLableVisible(false);
            }
            catch
            { }
        }

        private void pictureBox_road_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                bool IsDrawRect = false;
                Point Mouse = new Point(e.X, e.Y);
                for (int i = 0; i < rectList.Count; ++i)
                {
                    if (rectList[i].Contains(Mouse))
                    {
                        currRect = rectList[i];
                        currRectIndex = i;
                        IsDrawRect = true;
                        break;
                    }
                }

                if (IsDrawRect)
                {
                    SetRectLableVisible(true);
                    //SetRectLable(new Point(currRect.Left + currRect.Width / 2 - 3, currRect.Top - 3),
                    //    new Point(currRect.Left + currRect.Width / 2 - 3, currRect.Bottom - 3),
                    //    new Point(currRect.Left - 3, currRect.Top + currRect.Height / 2 - 3),
                    //    new Point(currRect.Right - 3, currRect.Top + currRect.Height / 2 - 3)); 

                    SetRectLable(new Point(currRect.Left - 3, currRect.Top - 3),
                        new Point(currRect.Left - 3, currRect.Bottom - 3),
                        new Point(currRect.Right - 3, currRect.Top - 3),
                        new Point(currRect.Right - 3, currRect.Bottom - 3));
                }
            }
            else if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                SetRectLableVisible(false);
            }
        }
        private void pictureBox_road_MouseEnter(object sender, EventArgs e)
        {
            this.pictureBox_road.Focus();
        }

        public void SetRectLable(Point LeftTop, Point LeftBottom, Point RightTop, Point RightBottom)
        {
            roadlableLeftTop.Location = LeftTop;
            roadlableLeftBottom.Location = LeftBottom;
            roadlableRightTop.Location = RightTop;
            roadlableRightBottom.Location = RightBottom;
        }
        public void SetRectLableVisible(bool visible)
        {
            roadlableLeftTop.Visible = visible;
            roadlableRightBottom.Visible = visible;
            roadlableLeftBottom.Visible = visible;
            roadlableRightTop.Visible = visible;
        }
        public void SaveRectangular(string RoadDiseaseType, int roadType)
        {
            String tempStack = "0";
            int stacknum = (int)Math.Round(_ImgPath[_curidx].imgmile);
            if (currRect.Y > (this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 2))
            {
                stacknum = stacknum + _ProjectInfo._Direction;
            }
            tempStack = stacknum.ToString("K0000+000");

            //存储该矩形框到与该图片对应的txt中
            Rectangle curimgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            string rectInfotemp = string.Format("{0} {1} {2} {3} {4} 桩号:{5} {6}", curimgrect.X, curimgrect.Y, curimgrect.Width, curimgrect.Height,
                RoadDiseaseType, tempStack, GlobalExcel._RoadTypeStr[roadType]);
            Disease temp = new Disease(rectInfotemp, stacknum);
            DisInfoList.Add(temp);
            rectList.Add(currRect);

            this.pictureBox_road.Invalidate();
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

        int stacknum;
        string rectInfotemp;
        Disease updatatemp;
        private void UpdateAllRectSaveInfo()
        {
            //获取桩号
            String tempStack = "0";
            stacknum = Convert.ToInt32(_ImgPath[_curidx].imgmile);
            if (currRect.Y > (this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 2))
            {
                stacknum += _ProjectInfo._Direction;
            }
            tempStack = stacknum.ToString("K0000+000");

            //存储该矩形框到与该图片对应的txt中
            Rectangle trect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            rectInfotemp = string.Format("{0} {1} {2} {3} {4} 桩号:{5} {6}", trect.X, trect.Y, trect.Width, trect.Height,
                DisInfoList[currRectIndex].GetRectInfoStr().Split('\n')[0].Replace(":", ""), tempStack, GlobalExcel._RoadTypeStr[_RoadType]);
            updatatemp = new Disease(rectInfotemp, stacknum);
            DisInfoList[currRectIndex] = updatatemp;
            rectList[currRectIndex] = Img2Box(DisInfoList[currRectIndex].rect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);

            this.pictureBox_road.Invalidate();
        }

        //保存病害图片
        public void LabelDistoImg()
        {
            if (MainForm._IsSaveDisImg && rectList.Count > 0)
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

        public void ShowNextImg()
        {
            if (WinRoadDisList._BrowserType == 1)
            {
                while (true)
                {
                    if (_curidx + 1 < _ImgPath.Count)
                    {
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
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
                        if (!File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
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
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[++_curidx].imgpath)))
                        {
                            string strs = File.ReadAllText(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[_curidx].imgpath));
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
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
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
                        if (!File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
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
                        if (File.Exists(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[--_curidx].imgpath)))
                        {
                            string strs = File.ReadAllText(string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[_curidx].imgpath));
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
        public override void SaveDisease()
        {
            ClearAllDiseaseInfoBox(true);
            
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

        private void WinRoadImg_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Z)
            {
                _IsKeyRepair = false;
            }
        }
        private void WinRoadImg_KeyDown(object sender, KeyEventArgs e)
        {
            

            if (e.KeyData == (Keys.Control | Keys.D))
            {
                for (int i = rectList.Count - 1; i >= 0; --i)
                {
                    SetRectLableVisible(true);
                    DeleteCurRect(i);
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
            else if (e.KeyData == Keys.Delete)
            {
                DeleteCurRect(currRectIndex);
            }
            else if (e.KeyData == Keys.G)
            {
                if (_Setting.showGpsInfoToPicture)
                {
                    Form msg = new Form();
                    msg.Text = "坐标信息";
                    msg.Size = new System.Drawing.Size(300, 220);

                    // Create and configure the TextBox to display the msgTxt
                    TextBox infoTextBox = new TextBox();
                    //infoTextBox.Text =  $"坐标：({curPosX},{curPosY})\r\n经度：{longitude + "°E"}\r\n纬度：{latitude + "°N"}\r\n高程：{centerH + " M"}";
                    infoTextBox.Text = $"{longitude}\r\n{latitude}\r\n{centerH}";
                    infoTextBox.Multiline = true;
                    infoTextBox.ReadOnly = true;
                    infoTextBox.Dock = DockStyle.Fill;

                    msg.Controls.Add(infoTextBox);

                    msg.StartPosition = FormStartPosition.CenterParent;
                    msg.ShowDialog();
                }
            }
            else if (e.KeyData == Keys.A && this.ContainsFocus)
            {
                //currRect = new Rectangle(2, 2, RoadImgRect.Width - 4, RoadImgRect.Height - 4);
                currRect = new Rectangle(0, 0, RoadImgRect.Width , RoadImgRect.Height );
                Rectangle imgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                if (JudgeRectArea(currRect, imgrect, rectList, true))
                {
                    UpdateBox();
                }
                else if (!DrawDis(currRect))
                {
                    currRect = Rectangle.Empty;
                    this.pictureBox_road.Invalidate();
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
            else if (e.KeyData == Keys.Z)
            {
                _IsKeyRepair = true;
            }
        }

        private bool DrawDis(Rectangle tcurrect)
        {
            this.pictureBox_road.Invalidate();

            RoadPavementPanel PavementDisease = new RoadPavementPanel(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis, _RoadType);
            Rectangle img = new Rectangle();
            img = Box2Img(tcurrect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            PavementDisease.SetLengthWidth(img.Width, img.Height);
            PavementDisease.ShowDialog();

            if (PavementDisease.IsDisease)
            {
                // SaveRectangular(PavementDisease.RoadDiseaseType, _RoadType);
                String tempStack = "0";
                int stacknum = (int)Math.Round(_ImgPath[_curidx].imgmile);
#if 辽宁建祥3m

                int splitY1 = this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 3;
                int splitY2 = this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) *2/3;

                if (currRect.Y > splitY1&& currRect.Y<splitY2)
                {
                    stacknum = stacknum + _ProjectInfo._Direction;
                }
                else if (currRect.Y>splitY2)
                {
                    stacknum = stacknum + _ProjectInfo._Direction*2;
                }
#else
                    if (currRect.Y > (this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 2))
                    {
                        stacknum = stacknum + _ProjectInfo._Direction;
                    }
#endif
                tempStack = stacknum.ToString("K0000+000");

                //存储该矩形框到与该图片对应的txt中
                Rectangle curimgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                string rectInfotemp = string.Format("{0} {1} {2} {3} {4} 桩号:{5} {6} {7}",
                    curimgrect.X, curimgrect.Y, curimgrect.Width, curimgrect.Height,
                    PavementDisease.RoadDiseaseType, tempStack,
                    GlobalExcel._RoadTypeStr[_RoadType],
                    PavementDisease.RoadDiseaseRemarks);
                Disease tempDi = new Disease(rectInfotemp, stacknum);


                if (!tempDi.RoadDisType.Contains("破碎板") && !tempDi.RoadDisType.Contains("松散") && !tempDi.RoadDisType.Contains("露骨"))
                {
                    
                }
                else
                {
                    if (tempDi.Area <= _RoadConfig.DetectWidth * 2 * 2 / 3)
                    {
                         
                    }
                    else
                    {
                        if (currRect.Y > (this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 2))
                        {
                            stacknum = stacknum - _ProjectInfo._Direction;
                        }
                    }
                }


                 Disease temp = new Disease(rectInfotemp, stacknum);


                DisInfoList.Add(temp);
                rectList.Add(currRect);
                //记录人工绘制病害 
                string disPathMessage = _ImgPath[_oldidx].imgpath;

                if (_Setting.outHumanDeleteDisease)
                { 
                    using (StreamWriter sw = new StreamWriter(unAutoDisPath, true))
                    {
                        sw.WriteLine(disPathMessage);
                        //病害文本地址 病害
                        sw.WriteLine(temp.GetDisInfoStr());
                    }
                    //复制图片
                    string resourcePicPath = string.Format("{0}\\RoadImg\\Camera0\\{1}", _ProjPath, _ImgPath[_curidx].imgpath);
                    string recordPicPath = string.Format("{0}\\{1}", unAutoDisPath.Substring(0,unAutoDisPath.LastIndexOf('\\')), _ImgPath[_curidx].imgpath);
                    // 获取文件的上级目录路径
                    string directoryPath = Path.GetDirectoryName(recordPicPath);
                    try
                    {
                        // 创建上级目录（如果它们不存在）
                        Directory.CreateDirectory(directoryPath);

                        // 创建文件（如果它不存在）
                        if (!File.Exists(recordPicPath))
                        {
                            File.Copy( resourcePicPath, recordPicPath);
                        }
                    }
                    catch (Exception ex)
                    {

                        
                    }
                  
                }
             

                this.pictureBox_road.Invalidate();
            }
            return PavementDisease.IsDisease;
        }

        private bool DrawDis2001(Rectangle tcurrect)
        {
            this.pictureBox_road.Invalidate();

            RoadPavement2001 PavementDisease = new RoadPavement2001(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis);
            Rectangle img = new Rectangle();
            img = Box2Img(tcurrect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
            PavementDisease.SetLengthWidth(img.Width, img.Height);
            PavementDisease.JudgeRoadType(_RoadType);
            PavementDisease.ShowDialog();

            if (PavementDisease.IsDisease)
            {
                String tempStack = "0";
                int stacknum = (int)Math.Round(_ImgPath[_curidx].imgmile);
                if (currRect.Y > (this.pictureBox_road.Location.Y + (this.pictureBox_road.Height - currRect.Height) / 2))
                {
                    stacknum = stacknum + _ProjectInfo._Direction;
                }
                tempStack = stacknum.ToString("K0000+000");

                //存储该矩形框到与该图片对应的txt中
                Rectangle curimgrect = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                string rectInfotemp = string.Format("{0} {1} {2} {3} {4} 桩号:{5} {6}", curimgrect.X, curimgrect.Y, curimgrect.Width, curimgrect.Height,
                    PavementDisease.RoadDiseaseType.ToString(), tempStack, GlobalExcel._RoadTypeStr[_RoadType]);
                Disease temp = new Disease(rectInfotemp, stacknum);
                DisInfoList.Add(temp);
                rectList.Add(currRect);

                this.pictureBox_road.Invalidate();
            }
            return PavementDisease.IsDisease;
        }

        /// <summary>
        /// 写病害文本
        /// </summary>
        /// <param name="isside"></param>
        private void ClearAllDiseaseInfoBox(bool isside)
        {
            if (_ImgPath.Count == 0) return;

            string recInfoFile = string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, _ImgPath[_oldidx].imgpath);
           

            //  unAutoDisPath = string.Format(@"{0}\RoadImg\Camera0{1}.txt", _ProjPath, "\\HumanDisMeaage");


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
                //
                //写入的时候要找一下 有可能用户删掉了人工病害但是图片上还有其他病害
                writeRGDKdis(false); 
            }
            else
            {
                //当前图片病害列表为0
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
                        //移走用户删除的病害
                        File.Move(recInfoFile, fpath); 
                    }
                    writeRGDKdis(true);
                }
            }
            if (!isside)
            {
                DisInfoList.Clear();
                rectList.Clear();
                currRectIndex = -1;
                SetRectLableVisible(false);
            }
        }
        private readonly static object lock1 = new object();
        /// <summary>
        /// 写人工大框病害
        /// </summary>
        /// <param name="clear"> true 上一张图片没有病害 </param>
        /// <param name="realStr">更新信息</param>
        private void writeRGDKdis(bool clear)
        {
            try
            {
                List<Disease> DisInfoListCopy = new List<Disease>();
                DisInfoList.ForEach(i => DisInfoListCopy.Add(i));
                System.Threading.Tasks.Task.Run(() =>
                {
                    lock (lock1)
                    {
                        if (_Setting.outHumanDeleteDisease)
                        {
                            string[] strArray = null;
                            if (File.Exists(unAutoDisPath))
                                strArray = File.ReadAllLines(unAutoDisPath);
                            List<string> realStr = new List<string>();

                            if (strArray != null)
                            {
                                if (clear)
                                {

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
                                            realStr.Add("人工添加后删除");
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
                });
            }
            catch (Exception ex)
            {

               
            }
           
        }

        /// <summary>
        /// 清除病害
        /// </summary>
        /// <param name="_currRectIndex"></param>
        private void DeleteCurRect(int _currRectIndex)
        {
            if (roadlableLeftTop.Visible)
            {
                if (_currRectIndex >= 0)
                {
                    var dis = DisInfoList[_currRectIndex];

                    string deleteTxt = string.Format("{0} {4} {5} {6} {1} 桩号:{2} {3}\n{7}", dis.rect.X, dis.RoadDisType, dis.m_mile.ToString("K0+000"), dis.RoadType, dis.rect.Y,
                        dis.rect.Width, dis.rect.Height, _ImgPath[_curidx].imgpath);
                    if (_Setting.outHumanDeleteDisease)
                    {
                        try
                        {
                            using (StreamWriter fs = new StreamWriter(deleteDisPath, true))
                            {
                                fs.WriteLine(deleteTxt);
                            }

                            string picFile = string.Format(@"{0}\RoadImg\Camera0{1}", _ProjPath, _ImgPath[_oldidx].imgpath);
                            string saveRecordPic = string.Format(@"{0}{1}", deleteDisPath.Substring(0,deleteDisPath.LastIndexOf("\\")), _ImgPath[_oldidx].imgpath);
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





                    DisInfoList.RemoveAt(_currRectIndex);
                    rectList.RemoveAt(_currRectIndex);

                }

                for (int i = 0; i < DisInfoList.Count; ++i)
                {
                    rectList[i] = DisInfoList[i].rect;
                }
                UpdateBox();
            }
        }
        private void UpdateBox()
        {
            currRectIndex = -1;
            SetRectLableVisible(false);
            currRect = Rectangle.Empty;
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.pictureBox_road.Invalidate();
        }
        public void WinRoadImg_FormClosed(object sender, FormClosedEventArgs e)
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
                            || (_Setting.ParmStyle == StandardParmType.RuralRoadHunan && s[s.Length - 1] == "砂石"))
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

      
        private void pictureBox_road_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                bool IsCheckRect = false;
                Point Mouse = new Point(e.X, e.Y);
                for (int i = 0; i < rectList.Count; ++i)
                {
                    if (rectList[i].Contains(Mouse))
                    {
                        currRect = rectList[i];
                        currRectIndex = i;
                        IsCheckRect = true;
                        break;
                    }
                }
                if (IsCheckRect)
                {
                    RoadPavementPanel PavementDisease = new RoadPavementPanel(RoadDiseaseTypes.DiseaseTypeDict, RoadDiseaseTypes.roaddis, _RoadType);
                    Rectangle img = new Rectangle();
                    img = Box2Img(currRect, pictureBox_road.Width, pictureBox_road.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                    PavementDisease.SetLengthWidth(img.Width, img.Height);
                    PavementDisease.ShowDialog();

                    if (PavementDisease.IsDisease)
                    {
                        if (currRectIndex >= 0)
                        {
                            DisInfoList[currRectIndex].RoadDisType = PavementDisease.RoadDiseaseType;
                            DisInfoList[currRectIndex].remarks = PavementDisease.RoadDiseaseRemarks;
                            DisInfoList[currRectIndex].RoadType = GlobalExcel._RoadTypeStr[_RoadType];
                        }
                        //重绘整个界面
                        this.pictureBox_road.Invalidate();
                    }

                    return;
                }
            }
        }

        override public void UpdateDisType(object updateinfo)
        {
            Global._UpdateInfo tupdate = (Global._UpdateInfo)(updateinfo);

            DisInfoList[tupdate.disidx].RoadDisType = tupdate.disname;
            DisInfoList[tupdate.disidx].remarks = tupdate.disremark;

            //重绘整个界面
            this.pictureBox_road.Invalidate();
        }

        private void label_imgpath_DoubleClick(object sender, EventArgs e)
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
    }
}
