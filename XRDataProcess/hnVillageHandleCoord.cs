//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.IO;
//using System.Drawing;
//using System.Windows.Forms;

//namespace XRDataProcess
//{
//     class hnVillageHandleCoord
//    {

//        public  hnVillageHandleCoord()
//        {
//            if ( Mat_U==null || Mat_V==null)
//            {
//                init();
//            }
           
          
//        }
//        public  void getHandelCoord( ref double x,ref double y)
//        {
//         x = Mat_U[(int)y + CutImgY,(int)x + CutImgX];
//         y = Mat_V[(int)y + CutImgY,(int)x + CutImgX];
//        // x =    matu.at<float>(y+CutImgY,x+CutImgX);
//        // y =    matv.at<float>(y+CutImgY,x+CutImgX);
//        }
//         /// <summary>
//         /// 根据矩形的四个点获得 映射出的  宽高（取平均值）  
//         /// tuple1 宽 ,tuple2 高
//         /// </summary>
//         /// <param name="r1"></param>
//         /// <returns></returns>
//        public  void getHandelCoordRect( Rectangle  r1,ref double width,ref double hight)
//        {

//            p1 = r1.Location;   //左上角
//            p2 = new Point(p1.X + r1.Width, p1.Y);  //右上角
//            p3 = new Point(p1.X, p1.Y + r1.Height); //左下角
//            p4 = new Point(p1.X + r1.Width, p1.Y + r1.Height);// 右下角
//             handle(ref p1);
//             handle(ref p2);
//             handle(ref p3);
//             handle(ref p4);
//             width = ((Math.Abs(p2.X - p1.X) + Math.Abs(p4.X - p3.X)) / 2);
//             hight = ((Math.Abs(p3.Y - p1.Y) + Math.Abs(p4.Y - p2.Y)) / 2);
            
//        }
//        private static  Point p1;
//         private static  Point p2;
//         private static  Point p3;
//         private static  Point p4;
//         //根据四个点求面积，更精确
//         public  double getArea(RoadConfig _RoadConfig)
//         {
//             double areaScale = (Math.Abs(p1.X * p2.Y + p2.X * p3.Y + p3.X * p1.Y - p1.Y * p2.X - p2.Y * p3.X - p3.Y * p1.X) / 2.0
//                + Math.Abs(p1.X * p4.Y + p4.X * p3.Y + p3.X * p1.Y - p1.Y * p4.X - p4.Y * p3.X - p3.Y * p1.X) / 2.0) / (_RoadConfig.ImageWidth * _RoadConfig.ImageHeight);
//             return areaScale * _RoadConfig.RealHeight * _RoadConfig.RealWidth;

//             //图片中面积大小
//         }
//         private void handle(ref Point p1)
//        {

//            try
//            {
//                Point p1Tmep = p1;
//                p1.X = (int)Mat_U[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

//                p1.Y = (int)Mat_V[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

//                while (p1.X == 0 && p1.Y == 0)
//                {
//                    p1Tmep.X--;
//                    p1.X = (int)Mat_U[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];

//                    p1.Y = (int)Mat_V[p1Tmep.Y + CutImgY, p1Tmep.X + CutImgX];
//                }
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("请检查选择的农村路模块类型与道路数据是否一致！图片大小越界！");
//                throw new VillageIndexException("请检查选择的农村路模块类型与道路数据是否一致！", ex);
//            }
                
           
//         }
//        private static int OriImgWidth;
//        private static int OriImgHeight;
//        private static int BigImgWidth;
//        private static int BigImgHeight;
//        private static int CutImgX;
//        private static int CutImgY;
//        private static int CutImgWidth;
//        private static int CutImgHeight;
//        private static int ImgQuality;
//        private static double[,] Mat_U;
//        private static double[,] Mat_V;
//        private static double widthScale=0;

//        public double WidthScale
//        {
//            get { return hnVillageHandleCoord.widthScale; }

//        }
//        private static double hightScale=0;

//        public double HightScale
//        {
//            get { return hnVillageHandleCoord.hightScale; }

//        }

//        private static double detectWidth=0;
//        public double DetectWidth
//        {
//            get { return hnVillageHandleCoord.detectWidth; }

//        }
//        private static void initExtracted(out double[,] Mat_U, out double[,] Mat_V)
//        {
//            Mat_U = new double[BigImgHeight, BigImgWidth];
//            Mat_V = new double[BigImgHeight, BigImgWidth];
//        }

//        public void setRealLength(ref double width, ref double hight)
//        {
//            using (StreamReader sr = new StreamReader("Setting\\lowVillageImgConfig\\ImageParm.cfg"))
//            {
//                while (!sr.EndOfStream)
//                {
//                    string str = sr.ReadLine();
//                    string[] strs = str.Split('=');
//                    switch (strs.First())
//                    {
//                        case "width":
//                            width = double.Parse(strs.Last());
//                            break;
//                        case "height":
//                            hight = double.Parse(strs.Last());
//                            break;
//                    }
//                }
//            }
//        }
//        public  void setShowMaxLength(ref double showMaxW, ref double showMaxH)
//        {
//            using (StreamReader sr = new StreamReader("Setting\\lowVillageImgConfig\\ImageParm.cfg"))
//            {
//                while (!sr.EndOfStream)
//                {
//                    string str = sr.ReadLine();
//                    string[] strs = str.Split('=');
//                    switch (strs.First())
//                    {

//                        case "showMaxW":
//                            showMaxW = double.Parse(strs.Last());
//                            break;
//                        case "showMaxH":
//                            showMaxH = double.Parse(strs.Last());
//                            break;
//                    }
//                }
//            }
//        }
//        public  void init()
//        {
//            using (StreamReader sr = new StreamReader("Setting\\lowVillageImgConfig\\ImageParm.cfg"))
//            {
//                while (!sr.EndOfStream)
//                {
//                    string str = sr.ReadLine();
//                    string[] strs = str.Split('=');
//                    switch (strs.First())
//                    {
//                        case "OriImgWidth":
//                            OriImgWidth = int.Parse(strs.Last());
//                            break;
//                        case "OriImgHeight":
//                            OriImgHeight = int.Parse(strs.Last());
//                            break;
//                        case "BigImgWidth":
//                            BigImgWidth = int.Parse(strs.Last());
//                            break;
//                        case "BigImgHeight":
//                            BigImgHeight = int.Parse(strs.Last());
//                            break;
//                        case "CutImgX":
//                            CutImgX = int.Parse(strs.Last());
//                            break;
//                        case "CutImgY":
//                            CutImgY = int.Parse(strs.Last());
//                            break;
//                        case "CutImgWidth":
//                            CutImgWidth = int.Parse(strs.Last());
//                            break;
//                        case "CutImgHeight":
//                            CutImgHeight = int.Parse(strs.Last());
//                            break;
//                        case "ImgQuality":
//                            ImgQuality = int.Parse(strs.Last());
//                            break;
//                        case "widthScale":
//                            widthScale = double.Parse(strs.Last());
//                            break;
//                        case "hightScale":
//                            hightScale = double.Parse(strs.Last());
//                            break;
//                        case "detectWidth":
//                            detectWidth = double.Parse(strs.Last());
//                            break;
//                    }
//                }
//            }
//           // float[,] matu = new float[BigImgHeight, BigImgWidth];
//           // float[,] matv = new float[BigImgHeight, BigImgWidth];

     
//        initExtracted(out Mat_U, out Mat_V);

//            //byte[] arr1 = File.ReadAllBytes("u_d.bin");
//           // byte[] arr2 = File.ReadAllBytes("v_d.bin");

//        using (var stream = File.Open("Setting\\lowVillageImgConfig\\u_d.bin", FileMode.Open))
//            {
//                using (BinaryReader br = new BinaryReader(stream))
//                {
//                    for (int i = 0; i < BigImgHeight; ++i)
//                    {
//                        for (int t = 0; t < BigImgWidth; ++t)
//                        {
//                            Mat_U[i, t] = br.ReadDouble();
//                        }
//                    }

//                }
//            }

//        using (var stream = File.Open("Setting\\lowVillageImgConfig\\v_d.bin", FileMode.Open))
//            {
//                using (BinaryReader br = new BinaryReader(stream))
//                {
//                    for (int i = 0; i < BigImgHeight; ++i)
//                    {
//                        for (int t = 0; t < BigImgWidth; ++t)
//                        {
//                            Mat_V[i, t] = br.ReadDouble();
//                        }
//                    }

//                }
//            }
//        }
//    }
    
//}
