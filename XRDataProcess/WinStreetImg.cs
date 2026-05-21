using DevExpress.Utils.About;
using DevExpress.Utils.CodedUISupport;
using NPOI.OpenXmlFormats.Vml.Wordprocessing;
using OperateIniFile;
using Spire.Pdf.Exporting.XPS.Schema;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using XRDataProcess.toolForms;

namespace XRDataProcess
{
    public partial class WinStreetImg : WinRoad
    {
        public event EventHandler EventUpdateMile;
        public event EventHandler EventUpdateDisList;
        public event EventHandler EventLoadDisList;
        public event EventHandler EventDeleteDis;
        private ProjectInfo _ProjectInfo;
        private string _ProjPath;
        public List<MyImgMile>[] _ImgPath = null;
        private int []_curidx = {0,0};
        private PictureBox []_PicBox;
        private Label[] _PicName;

        private double _dmival = 0;
        private int _mileval = 0;
        public bool _IsInitLoad = false;
        public bool _IsActivated = false;
        private string _StreetImgSubName = "StreetImg";

        private string _UserInfoSubName = "_UserSign";

        private List<StreetDisRecord> _DisRecord = null;
        private List<StreetDisRecord> _AllDisRecord = null;

        private List<StreetDisRecord> _DisRecord_RoadBed = null;

       // private List<string> _AllUserSignRecord = null;
        private List<UserSignMsg> _UserSignRecord = null;

        private List<StreetDisRecord> _AllDisRecord_RoadBed = null;
        RoadConfig _RoadConfig = RoadConfig.GetInstance();
        private Bitmap firstPicute = null; 
        public WinStreetImg(ProjectInfo pinfo, string ppath)
        {
            InitializeComponent();
            _ProjectInfo = pinfo;
            _ProjPath = ppath;
            _ImgPath = new List<MyImgMile>[2];
            _UserSignRecord  = new List<UserSignMsg>();
            _DisRecord = new List<StreetDisRecord>();
            _AllDisRecord = new List<StreetDisRecord>();

            _DisRecord_RoadBed = new List<StreetDisRecord>();

           
            _AllDisRecord_RoadBed = new List<StreetDisRecord>();

            for (int i = 0; i < 2; i++)
            {
                _ImgPath[i] = new List<MyImgMile>();
            }
            GetAllImg(_ProjPath + "\\StreetImg\\Camera0", ref _ImgPath[0]);
            if (_ProjectInfo._IsDStreet)
            {
                GetAllImg(_ProjPath + "\\StreetImg\\Camera1", ref _ImgPath[1]);
                if (_ImgPath[1].Count==0)
                {
                    GetAllImg(_ProjPath + "\\StreetImg2\\Camera0", ref _ImgPath[1]);

                   
                }

                splitContainer1.Panel2Collapsed = false;
            }
            else
            {
                splitContainer1.Panel2Collapsed = true;
            }
            if (_ImgPath[0].Count>0)
            {
                try
                {
                    string firstPicturePath = _ProjPath + "\\StreetImg\\Camera0" + _ImgPath[0][0].imgpath;
                    firstPicute = new Bitmap(firstPicturePath);
                }
                catch (Exception)
                {

                    
                }
                
                
            } 
            pictureBox_Img.MouseWheel += new MouseEventHandler(pictureBox_Img_MouseWheel);
            LoadAllRecInfo(_ProjPath, ref _AllDisRecord);
            LoadAllRecInfo_RoadBed(_ProjPath, ref _AllDisRecord_RoadBed); 
        }



        private void LoadAllRecInfo(string ppath, ref List<StreetDisRecord> AllRecord)
        {
            AllRecord.Clear();
            //单景观病害
            if (_ImgPath[0].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[0])
                {
                    List<StreetDisRecord> trecord = new List<StreetDisRecord>();
                    string fname = string.Format(@"{0}\{1}\{2}{3}.txt", ppath, _StreetImgSubName, "Camera0", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 0);
                    AllRecord.AddRange(trecord);

                   
                }
            }
            //双景观病害
            if (_ProjectInfo._IsDStreet && _ImgPath[1].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[1])
                {
                    List<StreetDisRecord> trecord = new List<StreetDisRecord>();
                    string fname = string.Format(@"{0}\{1}\{2}\{3}.txt", ppath, _StreetImgSubName, "Camera1", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 0);
                    AllRecord.AddRange(trecord);
                    fname = string.Format(@"{0}\{1}\{2}\{3}.txt", ppath, "StreetImg2", "Camera0", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 0);
                    AllRecord.AddRange(trecord);
                }
            }
        }
        private void LoadAllRecInfo_RoadBed(string ppath, ref List<StreetDisRecord> AllRecord)
        {
            AllRecord.Clear();
            //单景观病害
            if (_ImgPath[0].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[0])
                {
                    List<StreetDisRecord> trecord = new List<StreetDisRecord>();
                    string fname = string.Format(@"{0}\{1}\{2}{3}.rbd", ppath, _StreetImgSubName, "Camera0", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 1);
                    AllRecord.AddRange(trecord);
                }
            }
            //双景观病害
            if (_ProjectInfo._IsDStreet && _ImgPath[1].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[1])
                {
                    List<StreetDisRecord> trecord = new List<StreetDisRecord>();
                    string fname = string.Format(@"{0}\{1}\{2}\{3}.rbd", ppath, _StreetImgSubName, "Camera1", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 1);
                    AllRecord.AddRange(trecord);

                    fname = string.Format(@"{0}\{1}\{2}\{3}.txt", ppath, "StreetImg2", "Camera0", path.imgpath);
                    LoadRecInfo(fname, ref trecord, path.imgmile, 1);
                    AllRecord.AddRange(trecord);
                }
            }
        }

        /// <summary>
        /// 加载用户自定义病害信息
        /// </summary>
        /// <param name="ppath"></param>
        /// <param name="AllRecord"></param>
        private void LoadUserSign_AllRecInfo(string ppath, ref List<string> AllRecord)
        {
            AllRecord.Clear();
            //单景观病害
            if (_ImgPath[0].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[0])
                {
                    List<string> msgs = new List<string>();
                    string fname = string.Format(@"{0}\{1}\{2}{3}{4}.txt", ppath, _StreetImgSubName, "Camera0", path.imgpath, _UserInfoSubName);
                    if (File.Exists(fname))
                    {

                       AllRecord.AddRange(  File.ReadAllLines(fname).ToList());

                    }


                }
            }
            //双景观病害
            if (_ProjectInfo._IsDStreet && _ImgPath[1].Count > 0)
            {
                foreach (MyImgMile path in _ImgPath[1])
                {
                    List<StreetDisRecord> trecord = new List<StreetDisRecord>();
                    string fname = string.Format(@"{0}\{1}\{2}\{3}{4}.txt", ppath, _StreetImgSubName, "Camera1", path.imgpath,_UserInfoSubName);
                    if (File.Exists(fname))
                    {

                        AllRecord.AddRange(File.ReadAllLines(fname).ToList());

                    }
                    fname = string.Format(@"{0}\{1}\{2}\{3}{4}.txt", ppath, "StreetImg2", "Camera0", path.imgpath,_UserInfoSubName);
                    if (File.Exists(fname))
                    {

                        AllRecord.AddRange(File.ReadAllLines(fname).ToList());

                    }
                }
            }
        }
        private void WinStreetImg_Load(object sender, EventArgs e)
        {
            EventLoadDisList(_AllDisRecord, null);
            EventLoadDisList(_AllDisRecord_RoadBed, null);

            progressBar_per.Maximum = _ImgPath[0].Count;
            _PicBox = new PictureBox[2];
            _PicBox[0] = pictureBox_Img;
            _PicBox[1] = pictureBox_ImgR;

            _PicName = new Label[2];
            _PicName[0] = label_imgpathL;
            _PicName[1] = label_imgpathR;

            if (_ImgPath[0].Count > 0)
            {
                ShowImg(_ImgPath[0][_curidx[0]], 0);
            }
            if (_ProjectInfo._IsDStreet && _ImgPath[1].Count > 0)
            {
                ShowImg(_ImgPath[1][_curidx[1]], 1);
            }
            _IsInitLoad = true;
            RoadImgRect.X = 0;
            RoadImgRect.Y = 0;
            RoadImgRectR.X = 0;
            RoadImgRectR.Y = 0;
            RoadImgRect.Width = pictureBox_Img.Width;
            RoadImgRect.Height = pictureBox_Img.Height;
            RoadImgRectR.Width = pictureBox_ImgR.Width;
            RoadImgRectR.Height = pictureBox_ImgR.Height;

        }
       
        /// <summary>
        /// 最后会导致问题
        /// </summary>
        /// <param name="path"></param>
        /// <param name="idx"></param>
        private void ShowImg(MyImgMile path, int idx)
        {
            int left = _ProjectInfo._StreetImgDis_Left;
            int right = _ProjectInfo._StreetImgDis_Right;

           
          
            if (left != right)
            {
                _mileval = (int)(Convert.ToDouble(path.imgmile.ToString()));
                //_dmival = (int)Math.Ceiling(Mile2DMI((Convert.ToDouble(path.imgmile.ToString()))));
                _dmival = _ProjectInfo.Mile2Dmi(_mileval);
                EventUpdateMile(_mileval, null);
                textBox_mile.Text = _mileval.ToString();
            }
            else
            {
                if (idx == 0)
                {
                    _mileval = (int)(Convert.ToDouble(path.imgmile.ToString()));
                    //_dmival = (int)Math.Ceiling(Mile2DMI((Convert.ToDouble(path.imgmile.ToString()))));
                    _dmival = _ProjectInfo.Mile2Dmi(_mileval);
                    EventUpdateMile(_mileval, null);
                    textBox_mile.Text = _mileval.ToString();
                }
            }
            // 设置图像路径
            string imagePath = string.Format(@"{0}\StreetImg\Camera{1}{2}", _ProjPath, idx, path.imgpath);
            if (!File.Exists(imagePath))
            {
                imagePath = string.Format(@"{0}\StreetImg2\Camera0{2}", _ProjPath, idx, path.imgpath);
            } 
            if (_ProjectInfo._StreetImgDis_Left != _ProjectInfo._StreetImgDis_Right)
            {
                //右边图像需要翻转90度
                if (idx == 1)
                {     // 加载原始图像
                    Bitmap originalImage = new Bitmap(imagePath);
                    _PicBox[idx].Image = originalImage; // 直接赋值给 Image，而不是使用 ImageLocation
                    RotateImage90(_PicBox[idx]);
                    _PicName[idx].Text = string.Format(@"{0}\StreetImg\Camera{1}{2}", _ProjPath, idx, path.imgpath);

                }
                else
                {
                    _PicBox[idx].ImageLocation = imagePath;

                    _PicName[idx].Text = imagePath;
                }
            }
            else
            {
                _PicBox[idx].ImageLocation = imagePath;
                 
                _PicName[idx].Text = imagePath;
            } 
            textBox_dmi.Text = _dmival.ToString();
            progressBar_per.Value = _curidx[0];
            LoadRecInfo(_PicBox[0].ImageLocation + ".txt", ref _DisRecord, path.imgmile, 0);
            LoadRecInfo(_PicBox[0].ImageLocation + ".rbd", ref _DisRecord_RoadBed, path.imgmile, 1);
            _UserSignRecord.Clear();
            LoadUserRecInfo(_PicBox[0].ImageLocation +_UserInfoSubName + ".txt", ref _UserSignRecord);
            if (_PicBox.Length ==2)
            {
                LoadUserRecInfo(_PicBox[1].ImageLocation + _UserInfoSubName + ".txt", ref _UserSignRecord);

            }
        }
        // 旋转图像 90 度
        private void RotateImage90(PictureBox pictureBox)
        {
            if (pictureBox.Image != null)
            {
                // 获取当前图像
                Image originalImage = pictureBox.Image;

                // 创建一个新的 Bitmap 对象，用于旋转后的图像
                Bitmap rotatedImage = new Bitmap(originalImage.Height, originalImage.Width);

                // 使用 Graphics 对象绘制旋转后的图像
                using (Graphics g = Graphics.FromImage(rotatedImage))
                {
                    // 旋转图像
                    g.TranslateTransform(rotatedImage.Width / 2, rotatedImage.Height / 2);
                    g.RotateTransform(90); // 顺时针旋转 90 度
                    g.TranslateTransform(-originalImage.Width / 2, -originalImage.Height / 2);

                    // 绘制图像
                    g.DrawImage(originalImage, new Point(0, 0));
                }

                // 将旋转后的图像设置回 PictureBox
                pictureBox.Image = rotatedImage; 
            }
        }
        private bool GetAllImg(string path, ref List<MyImgMile> imgs)
        {
            try
            {
                if (!File.Exists(path + "\\Street2Mile.txt"))
                {
                    return false;
                }
                // GetStreet2Dmi(path);
                string[] imgsinfo = File.ReadAllLines(path + "\\Street2Mile.txt");
                
                foreach (string str in imgsinfo)
                {
                    imgs.Add(new MyImgMile(str));
                }
                return true;
            }
            catch (Exception ex)
            {

                throw ex;
            }
          
        }

        private void pictureBox_Img_MouseWheel(object sender, MouseEventArgs e)
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
            if (_PicBox[0].ImageLocation != null)
            {
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".txt", _DisRecord,true);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".rbd", _DisRecord_RoadBed, true);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".txt", _DisRecord, false);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".rbd", _DisRecord_RoadBed, false);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + _UserInfoSubName+".txt", _UserSignRecord,true);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + _UserInfoSubName+".txt", _UserSignRecord,false);
            }
            if (jval <= _ImgPath[0][0].imgmile && jval >= _ImgPath[0][_ImgPath[0].Count - 1].imgmile
                || jval >= _ImgPath[0][0].imgmile && jval <= _ImgPath[0][_ImgPath[0].Count - 1].imgmile)
            {
                _curidx[0] = BinSearch(jval, ref _ImgPath[0], _ProjectInfo._Direction);
                if (_curidx[0] >= 0 && _curidx[0] < _ImgPath[0].Count)
                {
                    ShowImg(_ImgPath[0][_curidx[0]], 0);
                }
            }
            if (_ProjectInfo._IsDStreet)
            {
                if (jval <= _ImgPath[1][0].imgmile && jval >= _ImgPath[1][_ImgPath[1].Count - 1].imgmile
                    || jval >= _ImgPath[1][0].imgmile && jval <= _ImgPath[1][_ImgPath[1].Count - 1].imgmile)
                {
                    _curidx[1] = BinSearch(jval, ref _ImgPath[1], _ProjectInfo._Direction);
                    if (_curidx[1] >= 0 && _curidx[1] < _ImgPath[1].Count)
                    {
                        ShowImg(_ImgPath[1][_curidx[1]], 1);
                    }
                }
            }
        }

        private void WinStreetImg_SizeChanged(object sender, EventArgs e)
        {
            try
            {
                splitContainer1.SplitterDistance = (splitContainer1.Width - splitContainer1.SplitterWidth) / 2;
            }
            catch 
            {
                try
                {
                    splitContainer1.SplitterDistance = splitContainer1.Width / 2;
                }
                catch { }
            }
        }
        private static int showTimeCount = 1;




        //public void ShowNextImg()
        //{
        //    if (_PicBox[0].ImageLocation != null)
        //    {
        //        ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".txt", _DisRecord);
        //        ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".rbd", _DisRecord_RoadBed);
        //    }
        //    if (_ProjectInfo._IsDStreet)
        //    {
        //        int left = _ProjectInfo._StreetImgDis_Left;
        //        int right = _ProjectInfo._StreetImgDis_Right;
        //        if (_curidx[0] + 1 < _ImgPath[0].Count)
        //        {
        //            ShowImg(_ImgPath[0][++_curidx[0]], 0);
        //        }
        //        if (_curidx[1] + 1 < _ImgPath[1].Count)
        //        {
        //            ShowImg(_ImgPath[1][++_curidx[1]], 1);
        //        }
        //        if (_curidx[0] + 1 == _ImgPath[0].Count && _curidx[1] + 1 == _ImgPath[1].Count)
        //        {
        //            if (showTimeCount < 0)
        //            {
        //                return;
        //            }
        //            showTimeCount--;
        //            MessageBox.Show("景观已经是最后一张图像！");

        //        }
        //    }
        //    else
        //    {
        //        if (_curidx[0] + 1 < _ImgPath[0].Count)
        //        {
        //            ShowImg(_ImgPath[0][++_curidx[0]], 0);
        //        }
        //        else if (_curidx[0] + 1 == _ImgPath[0].Count)
        //        {
        //            if (showTimeCount < 0)
        //            {
        //                return;
        //            }
        //            showTimeCount--;
        //            MessageBox.Show("景观已经是最后一张图像！");

        //        }
        //    }
        //}
        //public void ShowLastImg()
        //{
        //    if (_PicBox[0].ImageLocation != null)
        //    {
        //        ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".txt", _DisRecord);
        //        ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".rbd", _DisRecord_RoadBed);
        //    }
        //    if (_ProjectInfo._IsDStreet)
        //    {
        //        int left = _ProjectInfo._StreetImgDis_Left;
        //        int right = _ProjectInfo._StreetImgDis_Right;
        //        if (_curidx[0] > 0)
        //        {
        //            ShowImg(_ImgPath[0][--_curidx[0]], 0);
        //        }
        //        if (_curidx[1] > 0)
        //        {
        //            ShowImg(_ImgPath[1][--_curidx[1]], 1);
        //        }
        //        else if (_curidx[0] == 0 && _curidx[1] == 0)
        //        {
        //            MessageBox.Show("已经是第一张图像！");
        //        }
        //    }
        //    else
        //    {
        //        if (_curidx[0] > 0)
        //        {
        //            ShowImg(_ImgPath[0][--_curidx[0]], 0);
        //        }
        //        else if (_curidx[0] == 0)
        //        {
        //            MessageBox.Show("已经是第一张图像！");
        //        }
        //    }
        //}

        public void ShowNextImg()
        {
            // 清除旧信息
            if (_PicBox[0].ImageLocation != null)
            {
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".txt", _DisRecord,true);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".rbd", _DisRecord_RoadBed, true);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + _UserInfoSubName + ".txt", _UserSignRecord, true);


                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".txt", _DisRecord,false);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".rbd", _DisRecord_RoadBed,false);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + _UserInfoSubName + ".txt", _UserSignRecord,false);
            }

            if (_ProjectInfo._IsDStreet)
            {

                // 获取拍摄间距（仅用于理解问题，此处逻辑不直接依赖它）
                int left = _ProjectInfo._StreetImgDis_Left;
                int right = _ProjectInfo._StreetImgDis_Right;

                 if(left!= right)
                {
                    // 假设 _ProjectInfo.IsAscending 表示是否上行（需根据实际情况添加此属性）
                    bool isAscending = _ProjectInfo._Direction == 1; // true 为上行，false 为下行

                    // 检查是否可以推进
                    bool canAdvanceLeft = _curidx[0] + 1 < _ImgPath[0].Count;
                    bool canAdvanceRight = _curidx[1] + 1 < _ImgPath[1].Count;

                    // 如果都无法推进，提示用户
                    if (!canAdvanceLeft && !canAdvanceRight)
                    {
                        if (showTimeCount < 0)
                        {
                            return;
                        }
                        showTimeCount--;
                        MessageBox.Show("景观已经是最后一张图像！");
                        return;
                    }

                    // 获取下一个图像的桩号，若无更多图像则设为无穷大或无穷小
                    double nextLeftMile = canAdvanceLeft ? _ImgPath[0][_curidx[0] + 1].imgmile
                                                        : (isAscending ? double.PositiveInfinity : double.NegativeInfinity);
                    double nextRightMile = canAdvanceRight ? _ImgPath[1][_curidx[1] + 1].imgmile
                                                          : (isAscending ? double.PositiveInfinity : double.NegativeInfinity);

                    // 决定推进哪个相机
                    bool advanceLeft = false;
                    bool advanceRight = false;

                    if (isAscending) // 上行：桩号递增，选择较小的桩号
                    {
                        if (nextLeftMile < nextRightMile)
                        {
                            advanceLeft = true;
                        }
                        else if (nextRightMile < nextLeftMile)
                        {
                            advanceRight = true;
                        }
                        else // 桩号相等时同时推进
                        {
                            advanceLeft = true;
                            advanceRight = true;
                        }
                    }
                    else // 下行：桩号递减，选择较大的桩号
                    {
                        if (nextLeftMile > nextRightMile)
                        {
                            advanceLeft = true;
                        }
                        else if (nextRightMile > nextLeftMile)
                        {
                            advanceRight = true;
                        }
                        else // 桩号相等时同时推进
                        {
                            advanceLeft = true;
                            advanceRight = true;
                        }
                    }

                    // 更新索引
                    if (advanceLeft && canAdvanceLeft)
                    {
                        _curidx[0] += 1;
                    }
                    if (advanceRight && canAdvanceRight)
                    {
                        _curidx[1] += 1;
                    }

                    // 显示当前图像
                    ShowImg(_ImgPath[0][_curidx[0]], 0);
                    ShowImg(_ImgPath[1][_curidx[1]], 1);

                }
                else
                {
                    if (_curidx[0] + 1 < _ImgPath[0].Count)
                    {
                        ShowImg(_ImgPath[0][++_curidx[0]], 0);
                    }
                    if (_curidx[1] + 1 < _ImgPath[1].Count)
                    {
                        ShowImg(_ImgPath[1][++_curidx[1]], 1);
                    }
                    if (_curidx[0] + 1 == _ImgPath[0].Count && _curidx[1] + 1 == _ImgPath[1].Count)
                    {
                        if (showTimeCount < 0)
                        {
                            return;
                        }
                        showTimeCount--;
                        MessageBox.Show("景观已经是最后一张图像！");

                    }
                }



            }
            else // 非双侧街道模式保持原逻辑
            {
                if (_curidx[0] + 1 < _ImgPath[0].Count)
                {
                    ShowImg(_ImgPath[0][++_curidx[0]], 0);
                }
                else if (_curidx[0] + 1 == _ImgPath[0].Count)
                {
                    if (showTimeCount < 0)
                    {
                        return;
                    }
                    showTimeCount--;
                    MessageBox.Show("景观已经是最后一张图像！");
                }
            }
        }

        public void ShowLastImg()
        {
            // 清除旧信息
            if (_PicBox[0].ImageLocation != null)
            {
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".txt", _DisRecord,true);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + ".rbd", _DisRecord_RoadBed, true);
                ClearAllDiseaseInfoBox(_PicBox[0].ImageLocation + _UserInfoSubName + ".txt", _UserSignRecord, true);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".txt", _DisRecord,false);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + ".rbd", _DisRecord_RoadBed,false);
                ClearAllDiseaseInfoBox(_PicBox[1].ImageLocation + _UserInfoSubName + ".txt", _UserSignRecord,false);
            }

            if (_ProjectInfo._IsDStreet)
            {
                // 获取拍摄间距（仅用于理解问题，此处逻辑不直接依赖它）
                int left = _ProjectInfo._StreetImgDis_Left;
                int right = _ProjectInfo._StreetImgDis_Right;
                if (left != right)
                {
                    // 假设 _ProjectInfo.IsAscending 表示是否上行（需根据实际情况添加此属性）
                    bool isAscending = _ProjectInfo._Direction == 1; // true 为上行，false 为下行

                    // 如果都已经是第一张，提示用户
                    if (_curidx[0] == 0 && _curidx[1] == 0)
                    {
                        MessageBox.Show("已经是第一张图像！");
                        return;
                    }

                    // 获取当前图像的桩号
                    double currentLeftMile = _ImgPath[0][_curidx[0]].imgmile;
                    double currentRightMile = _ImgPath[1][_curidx[1]].imgmile;

                    // 检查是否可以回退
                    bool canGoBackLeft = _curidx[0] > 0;
                    bool canGoBackRight = _curidx[1] > 0;

                    if (isAscending) // 上行：桩号递增，回退时选择较大的桩号
                    {
                        if (currentLeftMile > currentRightMile && canGoBackLeft)
                        {
                            _curidx[0] -= 1;
                        }
                        else if (currentRightMile > currentLeftMile && canGoBackRight)
                        {
                            _curidx[1] -= 1;
                        }
                        else if (currentLeftMile == currentRightMile && canGoBackLeft && canGoBackRight)
                        {
                            _curidx[0] -= 1;
                            _curidx[1] -= 1;
                        }
                    }
                    else // 下行：桩号递减，回退时选择较小的桩号
                    {
                        if (currentLeftMile < currentRightMile && canGoBackLeft)
                        {
                            _curidx[0] -= 1;
                        }
                        else if (currentRightMile < currentLeftMile && canGoBackRight)
                        {
                            _curidx[1] -= 1;
                        }
                        else if (currentLeftMile == currentRightMile && canGoBackLeft && canGoBackRight)
                        {
                            _curidx[0] -= 1;
                            _curidx[1] -= 1;
                        }
                    }

                    // 显示当前图像
                    ShowImg(_ImgPath[0][_curidx[0]], 0);
                    ShowImg(_ImgPath[1][_curidx[1]], 1);
                }
                else
                {
                    if (_curidx[0] > 0)
                    {
                        ShowImg(_ImgPath[0][--_curidx[0]], 0);
                    }
                    if (_curidx[1] > 0)
                    {
                        ShowImg(_ImgPath[1][--_curidx[1]], 1);
                    }
                    else if (_curidx[0] == 0 && _curidx[1] == 0)
                    {
                        MessageBox.Show("已经是第一张图像！");
                    }
                }
               
            }
            else // 非双侧街道模式保持原逻辑
            {
                if (_curidx[0] > 0)
                {
                    ShowImg(_ImgPath[0][--_curidx[0]], 0);
                }
                else if (_curidx[0] == 0)
                {
                    MessageBox.Show("已经是第一张图像！"); 
                }
            }
        }



        private void WinStreetImg_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Space)
            {
                ShowNextImg();
            }
            else if (e.KeyData == Keys.Escape)
            {
                ShowLastImg();
            }
        }

        private void pictureBox_Img_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            int ri = pictureBox_Img.Height / 8, xi = pictureBox_Img.Width / 8, yi = ri;
            foreach (StreetDisRecord dis in _DisRecord)
            {
                if (dis.isHasRect())
                {

                }
                else
                {
                    g.DrawString(dis.ShowString(), new Font("宋体", 10, FontStyle.Regular), Brushes.Red, xi, yi);
                    if ((yi += 25) >= pictureBox_Img.Height)
                    {
                        xi += pictureBox_Img.Width / 2;
                        yi = ri;
                    }
                }

             
            }

            foreach (StreetDisRecord dis in _DisRecord_RoadBed)
            {
                if (dis.isHasRect())
                {

                }
                else
                {
                    g.DrawString(dis.ShowString(), new Font("宋体", 10, FontStyle.Regular), Brushes.Blue, xi, yi);
                    if ((yi += 25) >= pictureBox_Img.Height)
                    {
                        xi += pictureBox_Img.Width / 2;
                        yi = ri;
                    }
                } 
            }

            drawUserDis(0,g);
            foreach (UserSignMsg dis in _UserSignRecord)
            {
                if (dis.isHasRect())
                {
                 
                }
                else
                {
                    g.DrawString(dis.getDisInfo(), new Font("宋体", 10, FontStyle.Regular), Brushes.Yellow, xi, yi);
                    if ((yi += 25) >= pictureBox_Img.Height)
                    {
                        xi += pictureBox_Img.Width / 2;
                        yi = ri;
                    }
                } 
            }
            // --- 新增：绘制当前正在拖拽的红框 ---
            if (_isDrawing && _currentRect.Width > 0)
            {
                using (Pen p = new Pen(Color.Red, 2))
                {
                    p.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash; // 虚线框
                    g.DrawRectangle(p, _currentRect);
                }
            }
        }

        // 【新增方法】将图片真实坐标 转换回 PictureBox显示坐标
        public Rectangle RectPoint2BoxRect(PictureBox box, Rectangle imgRect, int boxw, int boxh)
        {
            if (firstPicute == null) return imgRect;

            Rectangle boxRect = new Rectangle();

            if (box.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                // 反向计算比例： 控件宽度 / 图片宽度
                double scalew = (double)boxw / firstPicute.Width;
                double scaleh = (double)boxh / firstPicute.Height;

                boxRect.X = (int)(imgRect.X * scalew);
                boxRect.Y = (int)(imgRect.Y * scaleh);
                boxRect.Width = (int)(imgRect.Width * scalew);
                boxRect.Height = (int)(imgRect.Height * scaleh);
            }
            else
            {
                // 如果是其他模式（如Normal），可能不需要转换，视情况而定
                return imgRect;
            }
            return boxRect;
        }

        private void drawUserDis(int side, Graphics g)
        {
            // 获取对应的PictureBox，用于计算比例
            PictureBox targetBox = (side == 0) ? pictureBox_Img : pictureBox_ImgR;
            // 获取对应的当前显示区域大小
            int boxW = (side == 0) ? RoadImgRect.Width : RoadImgRectR.Width;
            int boxH = (side == 0) ? RoadImgRect.Height : RoadImgRectR.Height;
            
            foreach (StreetDisRecord dis in _DisRecord)
            {
              
                // 确保只画当前侧的病害
                if (dis.Side != side) continue;

                if (dis.isHasRect())
                {
                
                    // 【修改点】将存储的 图片坐标 转换为 当前屏幕坐标
                    Rectangle drawRect = RectPoint2BoxRect(targetBox, dis.SignRect, boxW, boxH);
                    int drawY = drawRect.Y - 20;
                    if (drawY<5)
                    {
                        drawY = 5;
                    }
                    // 绘制文字（位置也需要跟随转换后的矩形）
                    g.DrawString(dis.ShowString(), new Font("宋体", 10, FontStyle.Bold), Brushes.Red, drawRect.X, drawY); // 文字画在框上方

                    // 绘制矩形
                    using (Pen p = new Pen(Color.Yellow, 2))
                    {
                        g.DrawRectangle(p, drawRect);
                    }
                }
                else
                {
                    
                }


            }
             
            foreach (StreetDisRecord dis in _DisRecord_RoadBed)
            { 
                if (dis.Side != side) continue;
                if (dis.isHasRect())
                {
                    // 【修改点】将存储的 图片坐标 转换为 当前屏幕坐标
                    Rectangle drawRect = RectPoint2BoxRect(targetBox, dis.SignRect, boxW, boxH);
                    int drawY = drawRect.Y - 20;
                    if (drawY < 5)
                    {
                        drawY = 5;
                    }
                    // 绘制文字（位置也需要跟随转换后的矩形）
                    g.DrawString(dis.ShowString(), new Font("宋体", 10, FontStyle.Bold), Brushes.Blue, drawRect.X, drawY); // 文字画在框上方

                    // 绘制矩形
                    using (Pen p = new Pen(Color.Yellow, 2))
                    {
                        g.DrawRectangle(p, drawRect);
                    }
                }
                else
                {
                    
                }
            }


            foreach (UserSignMsg dis in _UserSignRecord)
            {
                // 确保只画当前侧的病害
                if (dis.Side != side) continue;

                if (dis.isHasRect())
                {
                    // 【修改点】将存储的 图片坐标 转换为 当前屏幕坐标
                    Rectangle drawRect = RectPoint2BoxRect(targetBox, dis.SignRect, boxW, boxH);
                    int drawY = drawRect.Y - 20;
                    if (drawY < 5)
                    {
                        drawY = 5;
                    }
                    // 绘制文字（位置也需要跟随转换后的矩形）
                    g.DrawString(dis.getDisInfo(), new Font("宋体", 10, FontStyle.Bold), Brushes.Yellow, drawRect.X, drawY); // 文字画在框上方

                    // 绘制矩形
                    using (Pen p = new Pen(Color.Yellow, 2))
                    {
                        g.DrawRectangle(p, drawRect);
                    }
                }
                else
                {
                    // 处理没有框的旧数据（保持原有逻辑，或者根据需求调整）
                    // ...
                }
            }
        }


        private void pictureBox_ImgR_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            drawUserDis(1,g);
            // --- 新增：绘制当前正在拖拽的红框 ---
            if (_isDrawing && _currentRect.Width > 0)
            {
                using (Pen p = new Pen(Color.Red, 2))
                {
                    p.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash; // 虚线框
                    g.DrawRectangle(p, _currentRect);
                }
            }
        }

        private void button_Add_Click(object sender, EventArgs e)
        {
            if (!this.ContainsFocus)
            {
                return;
            }
            沿线设施损坏类型 streetpanel = new 沿线设施损坏类型(_mileval, _ProjectInfo._Direction, _ProjectInfo._StartMile, _ProjectInfo._EndMile);
            streetpanel.ShowDialog();
            foreach (StreetDisRecord dis in streetpanel._DisRecord)
            {
                if (!_DisRecord.Contains(dis))
                {
                    _DisRecord.Add(dis);
                }
            }
            foreach (StreetDisRecord dis in streetpanel._DisRecord)
            {
                if (!_AllDisRecord.Contains(dis))
                {
                    _AllDisRecord.Add(dis);
                    EventUpdateDisList(dis, null);
                }
            }            
            pictureBox_Img.Invalidate();
        }

        private void button_jump_Click(object sender, EventArgs e)
        {

            int temp = 0;
            try
            {
               // temp = (int)Math.Ceiling(Mile2DMI(Convert.ToDouble(textBox_mile.Text)));
               // temp = _ProjectInfo.Mile2Dmi((int)Math.Ceiling(Mile2DMI(Convert.ToDouble(textBox_mile.Text))));
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

        private void pictureBox_Img_MouseDoubleClick(object sender, MouseEventArgs e)
        {
             
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                bool flag = false;
                int ri = pictureBox_Img.Height / 8, xi = pictureBox_Img.Width / 8, yi = ri, idx = 0;// yi = ri
                foreach (StreetDisRecord dis in _DisRecord)
                {
                    if (dis.isHasRect() && dis.Side == 0)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_Img, dis.SignRect, RoadImgRect.Width, RoadImgRect.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                    else
                    {
                        Rectangle temp = new Rectangle(xi, yi, pictureBox_Img.Width / 2, 25);
                        if (temp.Contains(e.Location))
                        {
                            flag = true;
                            break;
                        }
                        if ((yi += 25) >= pictureBox_Img.Height)
                        {
                            xi += pictureBox_Img.Width / 2;
                            yi = ri;
                        }
                        idx++;
                    }
                        
                }
                if (flag)
                {
                    EventDeleteDis(_DisRecord[idx], null);
                    _AllDisRecord.Remove(_DisRecord[idx]);
                    _DisRecord.RemoveAt(idx);
                    pictureBox_Img.Invalidate();
                    return;
                }

                flag = false;
                idx = 0;
                foreach (StreetDisRecord dis in _DisRecord_RoadBed)
                {
                    if (dis.isHasRect() && dis.Side == 0)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_Img, dis.SignRect, RoadImgRect.Width, RoadImgRect.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                    else
                    {
                        Rectangle temp = new Rectangle(xi, yi, pictureBox_Img.Width / 2, 25);
                        if (temp.Contains(e.Location))
                        {
                            flag = true;
                            break;
                        }
                        if ((yi += 25) >= pictureBox_Img.Height)
                        {
                            xi += pictureBox_Img.Width / 2;
                            yi = ri;
                        }
                        idx++;
                    }

                        
                }
                if (flag)
                {
                    EventDeleteDis(_DisRecord_RoadBed[idx], null);
                    _AllDisRecord_RoadBed.Remove(_DisRecord_RoadBed[idx]);
                    _DisRecord_RoadBed.RemoveAt(idx);
                    pictureBox_Img.Invalidate();
                }
                flag = false;
                foreach (var dis in _UserSignRecord)
                {
                    if (dis.isHasRect() && dis.Side == 0)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_Img, dis.SignRect, RoadImgRect.Width, RoadImgRect.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                    else
                    {
                        Rectangle temp = new Rectangle(xi, yi, pictureBox_Img.Width / 2, 25);
                        if (temp.Contains(e.Location))
                        {
                            flag = true;
                            break;
                        }
                        if ((yi += 25) >= pictureBox_Img.Height)
                        {
                            xi += pictureBox_Img.Width / 2;
                            yi = ri;
                        }
                        idx++;
                    }
                }
                if (flag)
                {
                    //EventDeleteDis(_DisRecord_RoadBed[idx], null);
                    _UserSignRecord.RemoveAt(idx);
                    pictureBox_Img.Invalidate();
                }
            }
        }
        private void pictureBox_ImgR_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                bool flag = false;
                int ri = pictureBox_ImgR.Height / 8, xi = pictureBox_ImgR.Width / 8, yi = ri, idx = 0;// yi = ri
                foreach (StreetDisRecord dis in _DisRecord)
                {
                    
                    if (dis.isHasRect()&&dis.Side ==1)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_ImgR, dis.SignRect, RoadImgRectR.Width, RoadImgRectR.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                    else
                    {
                        Rectangle temp = new Rectangle(xi, yi, pictureBox_ImgR.Width / 2, 25);
                        if (temp.Contains(e.Location))
                        {
                            flag = true;
                            break;
                        }
                        if ((yi += 25) >= pictureBox_ImgR.Height)
                        {
                            xi += pictureBox_ImgR.Width / 2;
                            yi = ri;
                        }
                        idx++;
                    }

                }
                if (flag)
                {
                    EventDeleteDis(_DisRecord[idx], null);
                    _AllDisRecord.Remove(_DisRecord[idx]);
                    _DisRecord.RemoveAt(idx);
                    pictureBox_ImgR.Invalidate();
                    return;
                }

                flag = false;
                idx = 0;
                foreach (StreetDisRecord dis in _DisRecord_RoadBed)
                {
                    if (dis.isHasRect() && dis.Side == 1)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_ImgR, dis.SignRect, RoadImgRectR.Width, RoadImgRectR.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                    else
                    {
                        Rectangle temp = new Rectangle(xi, yi, pictureBox_ImgR.Width / 2, 25);
                        if (temp.Contains(e.Location))
                        {
                            flag = true;
                            break;
                        }
                        if ((yi += 25) >= pictureBox_ImgR.Height)
                        {
                            xi += pictureBox_ImgR.Width / 2;
                            yi = ri;
                        }
                        idx++;
                    }


                }
                if (flag)
                {
                    EventDeleteDis(_DisRecord_RoadBed[idx], null);
                    _AllDisRecord_RoadBed.Remove(_DisRecord_RoadBed[idx]);
                    _DisRecord_RoadBed.RemoveAt(idx);
                    pictureBox_ImgR.Invalidate();
                }
                flag = false;
                foreach (var dis in _UserSignRecord)
                {
                    if (dis.isHasRect() && dis.Side == 1)
                    {
                        var currect = RectPoint2BoxRect(pictureBox_ImgR, dis.SignRect, RoadImgRectR.Width, RoadImgRectR.Height);
                        if (currect.Contains(e.Location))
                        {
                            flag = true;
                            break;

                        }
                    }
                }
                if (flag)
                {
                    //EventDeleteDis(_DisRecord_RoadBed[idx], null);
                    _UserSignRecord.RemoveAt(idx);
                    pictureBox_ImgR.Invalidate();
                }
            }
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

        private void pictureBox_Img_MouseEnter(object sender, EventArgs e)
        {
            this.pictureBox_Img.Focus();
        }

        private void pictureBox_ImgR_MouseEnter(object sender, EventArgs e)
        {
            this.pictureBox_Img.Focus();
        }

        private void LoadRecInfo(string RecInfoFilename, ref List<StreetDisRecord> disRecord, double curmile, int type)
        {
            disRecord.Clear();
            if (File.Exists(RecInfoFilename))
            {
                FileStream fr = File.OpenRead(RecInfoFilename);
                StreamReader sr = new StreamReader(fr);
                String strline;
                while ((strline = sr.ReadLine()) != null)
                {
                    StreetDisRecord trecord = new StreetDisRecord(strline, (int)curmile, type);
                    if (trecord.isOK)
                    {
                        disRecord.Add(trecord);
                    }
                }
                sr.Close();
                fr.Close();
            }
        }

        private void LoadUserRecInfo(string RecInfoFilename, ref List<UserSignMsg> infos  )
        {
             
            if (File.Exists(RecInfoFilename))
            {
                FileStream fr = File.OpenRead(RecInfoFilename);
                StreamReader sr = new StreamReader(fr);
                String strline;
                while ((strline = sr.ReadLine()) != null)
                {
                    UserSignMsg userSignMsg = new UserSignMsg(strline);
                    if (!infos.Contains(userSignMsg))
                    {
                        infos.Add(userSignMsg);

                    }
                }
                sr.Close();
                fr.Close();
            }
        }
        private void ClearAllDiseaseInfoBox(string recInfoFile, List<StreetDisRecord> disRecord,bool isLeft)
        {
            bool hasDis = false;
            foreach (var dis in disRecord)
            {
                if (isLeft)
                {
                    if ( dis.Side == -1 ||  dis.Side == 0)
                    {
                        hasDis = true;
                        break;
                    }
                }
                else
                {
                    if (dis.Side == 1)
                    {
                        hasDis = true;
                        break;
                    }
                }
               
            }

            if (hasDis)
            {
                FileStream fw = File.Open(recInfoFile, FileMode.Create);
                StreamWriter sw = new StreamWriter(fw);
                for (int i = 0; i < disRecord.Count; ++i)
                {
                    if (isLeft && disRecord[i].Side <1)
                    {
                        sw.WriteLine(disRecord[i].ToString(), Encoding.UTF8);
                    }
                    if (!isLeft && disRecord[i].Side == 1)
                    {
                        sw.WriteLine(disRecord[i].ToString(), Encoding.UTF8);
                    }

                    
                }
                sw.Close();
                fw.Close();
            }
            else
            {
                if (File.Exists(recInfoFile))
                {
                    File.Delete(recInfoFile);
                }
            }
        }

        private void ClearAllDiseaseInfoBox(string recInfoFile, List<UserSignMsg> disRecord, bool isLeft)
        {
            bool hasDis = false;
            foreach (var dis in disRecord)
            {
                if (isLeft)
                {
                    if (dis.Side == -1 || dis.Side == 0)
                    {
                        hasDis = true;
                        break;
                    }
                }
                else
                {
                    if (dis.Side == 1)
                    {
                        hasDis = true;
                        break;
                    }
                }

            }
            if (hasDis)
            {
                FileStream fw = File.Open(recInfoFile, FileMode.Create);
                StreamWriter sw = new StreamWriter(fw);
                for (int i = 0; i < disRecord.Count; ++i)
                {
                    if (isLeft && disRecord[i].Side < 1)
                    {
                        sw.WriteLine(disRecord[i].ToString(), Encoding.UTF8);
                    }
                    if (!isLeft && disRecord[i].Side == 1)
                    {
                        sw.WriteLine(disRecord[i].ToString(), Encoding.UTF8);
                    }
                }
                sw.Close();
                fw.Close();
            }
            else
            {
                if (File.Exists(recInfoFile))
                {
                    File.Delete(recInfoFile);
                }
            }
        }


        private void label_imgpathL_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowDefalutSystemImg(label_imgpathL.Text);
        }
        private void label_imgpathR_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ShowDefalutSystemImg(label_imgpathR.Text);
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

        private void button_AddRoadBed_Click(object sender, EventArgs e)
        {
            路基损坏类型 streetpanel = new 路基损坏类型(_mileval, _ProjectInfo._Direction, _ProjectInfo._StartMile, _ProjectInfo._EndMile);
            streetpanel.ShowDialog();
            foreach (StreetDisRecord dis in streetpanel._DisRecord)
            {
                if (!_DisRecord_RoadBed.Contains(dis))
                {
                    _DisRecord_RoadBed.Add(dis);
                }
            }
            foreach (StreetDisRecord dis in streetpanel._DisRecord)
            {
                if (!_AllDisRecord_RoadBed.Contains(dis))
                {
                    _AllDisRecord_RoadBed.Add(dis);
                    EventUpdateDisList(dis, null);
                }
            }
            pictureBox_Img.Invalidate();
        }
        override public event EventHandler EventUpdateFullImg;
        override public event EventHandler EventUpdateFullPoint;
        private void pictureBox_Img_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Location.X > 0 && e.Location.X < pictureBox_Img.Width && e.Location.Y > 0 && e.Location.Y < pictureBox_Img.Height)
            {
                EventUpdateFullPoint(BoxPoint2RectPoint(pictureBox_Img,e.Location, RoadImgRect.Width, RoadImgRect.Height ), EventArgs.Empty);
            }
            EventUpdateFullImg(pictureBox_Img.Image, EventArgs.Empty);

            // 新增：绘制拉框逻辑
            if (_isDrawing)
            {
                // 计算矩形（支持向左上角拖拽）
                int x = Math.Min(_startPoint.X, e.X);
                int y = Math.Min(_startPoint.Y, e.Y);
                int w = Math.Abs(e.X - _startPoint.X);
                int h = Math.Abs(e.Y - _startPoint.Y);
                _currentRect = new Rectangle(x, y, w, h);

                // 触发重绘
                pictureBox_Img.Invalidate();
            }
        }
        private Rectangle RoadImgRect = new Rectangle();
        private Rectangle RoadImgRectR = new Rectangle();
      
        private void pictureBox_ImgR_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Location.X > 0 && e.Location.X < pictureBox_ImgR.Width && e.Location.Y > 0 && e.Location.Y < pictureBox_ImgR.Height)
            {
                EventUpdateFullPoint(BoxPoint2RectPoint(pictureBox_ImgR,e.Location, RoadImgRectR.Width, RoadImgRectR.Height ), EventArgs.Empty);
            }
            EventUpdateFullImg(pictureBox_ImgR.Image, EventArgs.Empty);
            // 新增：绘制拉框逻辑
            if (_isDrawing)
            {
                // 计算矩形（支持向左上角拖拽）
                int x = Math.Min(_startPoint.X, e.X);
                int y = Math.Min(_startPoint.Y, e.Y);
                int w = Math.Abs(e.X - _startPoint.X);
                int h = Math.Abs(e.Y - _startPoint.Y);
                _currentRect = new Rectangle(x, y, w, h);

                // 触发重绘
                pictureBox_ImgR.Invalidate();
            }
        }
        public Point BoxPoint2RectPoint(PictureBox box, Point boxpoint, int boxw, int boxh)
        { 
            Point imgpoint = new Point();
            //imgpoint.X = boxpoint.X;
            //imgpoint.Y = boxpoint.Y;
            if (firstPicute==null)
            {
                return imgpoint;
            }
            if (box.SizeMode == PictureBoxSizeMode.StretchImage)
            {
                double scalew = (double)firstPicute.Width / boxw;
                double scaleh = (double)firstPicute.Height / boxh;
                imgpoint.X = (int)(boxpoint.X * scalew);
                imgpoint.Y = (int)(boxpoint.Y * scaleh);
            }
            return imgpoint;
        }

        private void pictureBox_Img_Resize(object sender, EventArgs e)
        {
            RoadImgRect.Width = pictureBox_Img.Width;
            RoadImgRect.Height = pictureBox_Img.Height;
        }

        private void pictureBox_ImgR_Resize(object sender, EventArgs e)
        {
            RoadImgRectR.Width = pictureBox_ImgR.Width;
            RoadImgRectR.Height = pictureBox_ImgR.Height;
        }

        private void pictureBox_ImgR_Click(object sender, EventArgs e)
        {

        }

        private void 添加沿线设施损坏病害ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           this.button_Add_Click(sender,e);
        }

        private void 添加路基损坏病害ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.button_AddRoadBed_Click(sender, e);
        }

        /// <summary>
        /// 不带坐标信息的自定义病害添加
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 添加自定义病害ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            string filePath = _PicBox[0].ImageLocation + _UserInfoSubName + ".txt";
           
            //弹出窗口提示用户选择
            CustomStreetInfoForm fs = new CustomStreetInfoForm(_mileval, filePath);
           if( fs.ShowDialog()== DialogResult.OK)
            {
                _UserSignRecord.Add(fs.getUser());

                pictureBox_Img.Invalidate();
            }
 
        
        }
        // 1. 在 WinStreetImg 类中添加以下成员变量
        private bool _isDrawing = false;      // 是否正在绘制
        private Point _startPoint;            // 鼠标按下时的起始点
        private Rectangle _currentRect;       // 当前绘制的矩形（PictureBox坐标系）
                                              // 2. 辅助方法：将PictureBox上的矩形转换为图片的实际矩形坐标
                                              // 这样即使窗口缩放，保存的坐标也是相对于原始图片的
        private Rectangle GetImgRectangle(Rectangle boxRect)
        {
            if (firstPicute == null || pictureBox_Img.Image == null) return boxRect;

            // 利用你现有的 BoxPoint2RectPoint 方法逻辑进行转换
            Point p1 = BoxPoint2RectPoint(pictureBox_Img, new Point(boxRect.X, boxRect.Y), RoadImgRect.Width, RoadImgRect.Height);
            Point p2 = BoxPoint2RectPoint(pictureBox_Img, new Point(boxRect.Right, boxRect.Bottom), RoadImgRect.Width, RoadImgRect.Height);

            return new Rectangle(p1.X, p1.Y, p2.X - p1.X, p2.Y - p1.Y);
        }
        
         
        private void pictureBox_Img_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isDrawing = true;
                _startPoint = e.Location;
                _currentRect = new Rectangle(e.X, e.Y, 0, 0);
            }
        }
        private void pictureBox_ImgR_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _isDrawing = true;
                _startPoint = e.Location;
                _currentRect = new Rectangle(e.X, e.Y, 0, 0);
            }
        }


        private void drawUserStreetDis(int side,MouseEventArgs e)
        {
            if (_isDrawing && e.Button == MouseButtons.Left)
            {
                _isDrawing = false;

                // 如果框太小（例如只是点击了一下），则认为是误操作或普通点击，不触发
                if (_currentRect.Width < 5 || _currentRect.Height < 5)
                {
                    _currentRect = Rectangle.Empty;
                    pictureBox_Img.Invalidate();
                    return;
                }

                //获取用户选择的病害类型
                UserStreetDisModelSelectForm selectForm = new UserStreetDisModelSelectForm();
                selectForm.StartPosition = FormStartPosition.CenterScreen;

                if (selectForm.ShowDialog() == DialogResult.OK&& selectForm.SelectedModelIndex!=-1)
                {
                    
                    // 1. 获取图片实际坐标系的矩形 (用于保存)
                    Rectangle imgRect = GetImgRectangle(_currentRect);
                    if (selectForm.SelectedModelIndex == 1)
                    {
                        if (!this.ContainsFocus)
                        {
                            return;
                        }
                        沿线设施损坏类型 streetpanel = new 沿线设施损坏类型(_mileval, _ProjectInfo._Direction, _ProjectInfo._StartMile, _ProjectInfo._EndMile,side, imgRect);
                        streetpanel.ShowDialog();
                        foreach (StreetDisRecord dis in streetpanel._DisRecord)
                        {
                            if (!_DisRecord.Contains(dis))
                            {
                                _DisRecord.Add(dis);
                            }
                        }
                        foreach (StreetDisRecord dis in streetpanel._DisRecord)
                        {
                            if (!_AllDisRecord.Contains(dis))
                            {
                                _AllDisRecord.Add(dis);
                                EventUpdateDisList(dis, null);
                            }
                        }

                    }
                    if (selectForm.SelectedModelIndex ==2)
                    {
                        路基损坏类型 streetpanel = new 路基损坏类型(_mileval, _ProjectInfo._Direction, _ProjectInfo._StartMile, _ProjectInfo._EndMile,side,imgRect);
                        streetpanel.ShowDialog();
                        foreach (StreetDisRecord dis in streetpanel._DisRecord)
                        {
                            if (!_DisRecord_RoadBed.Contains(dis))
                            {
                                _DisRecord_RoadBed.Add(dis);
                            }
                        }
                        foreach (StreetDisRecord dis in streetpanel._DisRecord)
                        {
                            if (!_AllDisRecord_RoadBed.Contains(dis))
                            {
                                _AllDisRecord_RoadBed.Add(dis);
                                EventUpdateDisList(dis, null);
                            }
                        }
                    }
                    if (selectForm.SelectedModelIndex ==3)
                    {
                        //自定义病害

                        // 2. 准备文件路径
                        string filePath = _PicBox[side].ImageLocation + _UserInfoSubName + ".txt";

                        // 3. 弹出窗口，传入实际坐标
                        // [注意] 这里调用了修改后的构造函数
                        CustomStreetInfoForm fs = new CustomStreetInfoForm(side, _mileval, filePath, imgRect);
                        fs.StartPosition = FormStartPosition.CenterScreen;
                        if (fs.ShowDialog() == DialogResult.OK)
                        {
                            _UserSignRecord.Add(fs.getUser());  
                        }
                    } 

                    // 5. 清除绘制框
                    _currentRect = Rectangle.Empty;

                    if (side == 0)
                    {
                        pictureBox_Img.Invalidate();

                    }
                    else
                    {
                        pictureBox_ImgR.Invalidate();

                    }

                }
               

              
            }
        }
        private void pictureBox_Img_MouseUp(object sender, MouseEventArgs e)
        {
            drawUserStreetDis(0, e);
        }

        private void pictureBox_ImgR_MouseUp(object sender, MouseEventArgs e)
        {
            drawUserStreetDis(1, e);

        }

       
    }
}
