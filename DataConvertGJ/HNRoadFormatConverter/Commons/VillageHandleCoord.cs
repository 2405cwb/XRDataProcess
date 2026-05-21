using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Drawing;

namespace HNRoadFormatConverter.Commons
{
    /// <summary>
    /// 当用户只裁剪即 bool值为false的时候需要   进行矩阵转换
    /// </summary>
     class VillageHandleCoord
    {
        private static int OriImgWidth  =0;
        private static int OriImgHeight =0;
        private static int BigImgWidth  =0;
        private static int BigImgHeight =0;
        private static int CutImgX      =0;
        private static int CutImgY      =0;
        private static int CutImgWidth  =0;
        private static int CutImgHeight =0;
        private static int ImgQuality   =0;
        private static double QxCutImgWidth = 0;
        private static double QxCutImgHeight = 0;
        private static double[,] Mat_U  =null;
        private static double[,] Mat_V  =null;
        private static double widthScale = 0;
        private bool _isjz = true;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_RoadConfig"></param>
        /// <returns></returns>
        //public double getArea(RoadConfig _RoadConfig)
        //{
        //    double areaScale = (Math.Abs(p1.X * p2.Y + p2.X * p3.Y + p3.X * p1.Y - p1.Y * p2.X - p2.Y * p3.X - p3.Y * p1.X) / 2.0
        //       + Math.Abs(p1.X * p4.Y + p4.X * p3.Y + p3.X * p1.Y - p1.Y * p4.X - p4.Y * p3.X - p3.Y * p1.X) / 2.0) / (QxCutImgWidth * QxCutImgHeight);
        //    return areaScale * _RoadConfig.RealHeight * _RoadConfig.RealWidth;

        //    //图片中面积大小
        //}
        public static VillageHandleCoord getInstance(bool isjz, string path)
        {
            //isjz  进行的是矫正变换 
            if (_HightVillageHandleCoord == null)
            {
                _HightVillageHandleCoord = new VillageHandleCoord(isjz, path);

            }

            return _HightVillageHandleCoord;
        }
        /// <summary>
        /// 只允许确定已经有实例的时候使用
        /// </summary>
        /// <returns></returns>
        public static VillageHandleCoord getInstance()
        {
            return _HightVillageHandleCoord;
        }
        /// <summary>
        /// 仅当返回false是代表为仅裁剪  需要进行矩阵变换
        /// </summary>
        public bool Isjz
        {
            get { return _isjz; }
            set { _isjz = value; }
        }
        /// <summary>
        /// 仅且仅当isjz=false  时候使用
        /// </summary>
        /// <param name="r1"></param>
        /// <param name="width"></param>
        /// <param name="hight"></param>
        public void getHandelCoordRect(Rectangle r1, ref double width, ref double hight)
        {
           
            p1 = r1.Location;   //左上角
            p2 = new Point(p1.X + r1.Width, p1.Y);  //右上角
            p3 = new Point(p1.X, p1.Y + r1.Height); //左下角
            p4 = new Point(p1.X + r1.Width, p1.Y + r1.Height);// 右下角
            handle(ref p1);
            handle(ref p2);
            handle(ref p3);
            handle(ref p4);
            width = ((Math.Abs(p2.X - p1.X) + Math.Abs(p4.X - p3.X)) / 2);
            hight = ((Math.Abs(p3.Y - p1.Y) + Math.Abs(p4.Y - p2.Y)) / 2);

        }
        private static Point p1;
        private static Point p2;
        private static Point p3;
        private static Point p4;
      
        private void handle(ref Point p1)
        {
            try
            {
                Point p1Tmep = p1;
                p1.X = (int)Mat_U[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

                p1.Y = (int)Mat_V[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

                while (p1.X == 0 && p1.Y == 0)
                {
                    p1Tmep.X--;
                    p1.X = (int)Mat_U[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

                    p1.Y = (int)Mat_V[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("请检查选择的农村路模块类型与道路数据是否一致！图片大小越界！");
                throw new Exception("请检查选择的农村路模块类型与道路数据是否一致！", ex);
            }

        }


        private void initExtracted(out double[,] Mat_U, out double[,] Mat_V)
        {
            Mat_U = new double[BigImgHeight, BigImgWidth];
            Mat_V = new double[BigImgHeight, BigImgWidth];
        }



        private VillageHandleCoord(bool isjz_,string path)
        {
            this._isjz = isjz_;
            //仅裁剪
            if (!isjz_)
            {
                string iniPath = Path.Combine(path);
                init(iniPath);
            }
        }
      
        private static VillageHandleCoord _HightVillageHandleCoord = null;
      


        /// <summary>
        /// 比较费时，只用执行一次
        /// </summary>
        /// <param name="path"></param>
        private void init(string path)
        {
            using (StreamReader sr = new StreamReader(Path.Combine(path,"CamSetting.ini")))
            {
                while (!sr.EndOfStream)
                {
                    string str = sr.ReadLine();
                    string[] strs = str.Split('=');
                    switch (strs.First())
                    {
                        case "OriImgWidth":
                            OriImgWidth = int.Parse(strs.Last());
                            break;
                        case "OriImgHeight":
                            OriImgHeight = int.Parse(strs.Last());
                            break;
                        case "BigImgWidth":
                            BigImgWidth = int.Parse(strs.Last());
                            break;
                        case "BigImgHeight":
                            BigImgHeight = int.Parse(strs.Last());
                            break;
                        case "CutImgX":
                            CutImgX = int.Parse(strs.Last());
                            break;
                        case "CutImgY":
                            CutImgY = int.Parse(strs.Last());
                            break;
                        case "CutImgWidth":
                            CutImgWidth = int.Parse(strs.Last());
                            break;
                        case "qxCutImgWidth":
                            QxCutImgWidth = double.Parse(strs.Last());
                            break;
                        case "qxCutImgHeight":
                           QxCutImgHeight = double.Parse(strs.Last());
                            break;
                        case "CutImgHeight":
                            CutImgHeight = int.Parse(strs.Last());
                            break;
                        case "ImgQuality":
                            ImgQuality = int.Parse(strs.Last());
                            break;
                    }
                }
            }
      
            initExtracted(out Mat_U, out Mat_V);
            using (var stream = File.Open(Path.Combine(path,"u_d2.bin"), FileMode.Open))
            {
                using (BinaryReader br = new BinaryReader(stream))
                {
                    for (int i = 0; i < BigImgHeight; ++i)
                    {
                        for (int t = 0; t < BigImgWidth; ++t)
                        {
                            Mat_U[i, t] = br.ReadDouble();
                        }
                    }

                }
            }

            using (var stream = File.Open(Path.Combine(path, "v_d2.bin"), FileMode.Open))
            {
                using (BinaryReader br = new BinaryReader(stream))
                {
                    for (int i = 0; i < BigImgHeight; ++i)
                    {
                        for (int t = 0; t < BigImgWidth; ++t)
                        {
                            Mat_V[i, t] = br.ReadDouble();
                        }
                    }

                }
            }
        }

    }
}
