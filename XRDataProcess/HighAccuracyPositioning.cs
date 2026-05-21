using Framework.Other;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XRDataProcess.Properties;


namespace XRDataProcess
{
    /// <summary>
    /// 汉阳市政  高精度定位系统
    /// </summary>
    public static class HighAccuracyPositioning
    {
        
        [DllImport("hnCalcuMethod.dll", EntryPoint = "calcLatToPicCenter")]
        static extern bool calcLatToPicCenter(
           double dCurLon, double dCurLat, double dCurHeight,
           double dLastLon, double dLastLat, double dLastHeight, double dOffsetX,
           double dOffsetY, double dOffsetZ,
           out double returnLon, out double returnLat, out double returnHeight,
           bool bInverse);

        [DllImport("hnCalcuMethod.dll", EntryPoint = "calcLatToPicPos")]
        static extern bool calcLatToPicPos(bool showGps,
          double dCurPicLon, double dCurPicLat, double dCurPicH,
          double dLastPicLon, double dLastPicLat, double dLastPicH,
          int picX, int picY, int picWidth, int picHeight,
          ref double returnLon, ref double returnLat, ref double returnHeight,int equip,
          bool bInverse = false, double dWidth = 3.75, double dHeight = 2.0);

        /// <summary>
        /// 桩号  经度 维度 高程
        /// </summary>
        static List<(double, double, double, double)> gps2Mile = new List<(double, double, double, double)>();
        public static void getHighAccPosition(bool showGps, int equipType, string _ProjPath, List<MyImgMile> _ImgPath, int curPicuteIndex, int curPosX,
            int curPosY, int pictureW, int pictureH, double roadWidth, double roadLength,
            ref double dDiseaseLon, ref double dDiseaseLat, ref double dDiseaseH
            )
        {
            //20240924 cwb
          //  if (equipType == 1)
            {
            //    curPosX = pictureW - curPosX;

            }

            string mileFile = _ProjPath + "/GPS2Mile.txt";
            if (File.Exists(mileFile))
            {
                List<string> gps2MileFile = File.ReadAllLines(mileFile).ToList();
                gps2Mile.Clear();

                foreach (var msg in gps2MileFile)
                {
                    (double, double, double, double) msgLine;
                    string[] msgSp = msg.Split(' ');
                    double longitudeStr = double.Parse(msgSp[1]);
                    double latitudeStr = double.Parse(msgSp[2]);
                    double altitudeStr = double.Parse(msgSp[3]);
                    double mile = double.Parse(msgSp.Last());
                    gps2Mile.Add((mile, longitudeStr, latitudeStr, altitudeStr));
                }
                if (gps2Mile.Count > 1)
                {

                    double dCenterLon = 0, dCenterLat = 0, dCenterH = 0; //当前图像的经纬高
                    double lastCenterlon = 0, lastCenterlat = 0, lastCenterH = 0;//后一张图像中心的经维高 
                    int curIdx = curPicuteIndex;

                    var soltGps2mileList = gps2Mile
                    .Select((num, index) => new { Number = num.Item1, Index = index }).ToList();
                    int closestIndex = 0;
                    int closestIndex1 = 0;
                    (double, double, double, double) gpsNow;
                    (double, double, double, double) gpsPre;
                    //获取病害经纬度
                    if (curIdx == 0)//在第一张图片 
                    {
                        //根据图片桩号获得 对应经纬度 
                        //第一张图像的经纬度

                        
                        closestIndex = soltGps2mileList
                            .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                            .First()
                            .Index;
                        gpsNow = gps2Mile[closestIndex];

                        closestIndex1 = soltGps2mileList
                          .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx + 1].imgmile))
                          .First()
                          .Index;
                        if (closestIndex == closestIndex1)
                        {
                            closestIndex1++;
                        }
                        gpsPre = gps2Mile[closestIndex1];



                        handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out dCenterLon, out dCenterLat, out dCenterH, equipType, true);



                        closestIndex = soltGps2mileList
                           .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx + 1].imgmile))
                           .First()
                           .Index;
                        gpsNow = gps2Mile[closestIndex];


                        closestIndex1 = soltGps2mileList
                     .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                     .First()
                     .Index;
                        if (closestIndex == closestIndex1)
                        {
                            closestIndex1--;
                        }
                        gpsPre = gps2Mile[closestIndex1];

                        handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out lastCenterlon, out lastCenterlat, out lastCenterH, equipType, false);
                        handelGpsInfoInDisease( equipType, showGps,(0, dCenterLon, dCenterLat, dCenterH),
                            (0, lastCenterlon, lastCenterlat, lastCenterH), roadWidth, roadLength,
                            curPosX, curPosY, pictureW, pictureH,
                            ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH, true);
                    }
                    else
                    {

                        closestIndex = soltGps2mileList
                         .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                         .First()
                         .Index;
                        gpsNow = gps2Mile[closestIndex];

                        closestIndex1 = gps2Mile
                         .Select((num, index) => new { Number = num.Item1, Index = index })
                         .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx - 1].imgmile))
                         .First()
                         .Index;

                        if (closestIndex == closestIndex1)
                        {
                            closestIndex1--;
                        }
                        gpsPre = gps2Mile[closestIndex1];
                        handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out dCenterLon, out dCenterLat, out dCenterH, equipType, false);

                        if (curIdx - 1 == 0)
                        {

                            closestIndex = soltGps2mileList
                           .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx - 1].imgmile))
                           .First()
                           .Index;
                            gpsNow = gps2Mile[closestIndex];


                            closestIndex1 = soltGps2mileList
                         .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                         .First()
                         .Index;
                            if (closestIndex == closestIndex1)
                            {
                                closestIndex1++;
                            }
                            gpsPre = gps2Mile[closestIndex1];
                            handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out lastCenterlon, out lastCenterlat, out lastCenterH, equipType, true);
                        }
                        else
                        {
                            curIdx -= 1;



                            closestIndex = soltGps2mileList
                         .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                         .First()
                         .Index;
                            gpsNow = gps2Mile[closestIndex];


                            closestIndex1 = soltGps2mileList
                     .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx - 1].imgmile))
                     .First()
                     .Index;
                            if (closestIndex == closestIndex1)
                            {
                                closestIndex1--;
                            }
                            gpsPre = gps2Mile[closestIndex1];
                            handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out lastCenterlon, out lastCenterlat, out lastCenterH, equipType, false);
                        }
                        handelGpsInfoInDisease(equipType, showGps,(0, dCenterLon, dCenterLat, dCenterH),
                            (0, lastCenterlon, lastCenterlat, lastCenterH), roadWidth, roadLength,
                             curPosX, curPosY, pictureW, pictureH,
                             ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH, false);
                    }
                }

            }



        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="showGps">是否显示gps</param>
        /// <param name="highGpss">高精度定位信息列表</param>
        /// <param name="equipType"></param>
        /// <param name="_ProjPath"></param>
        /// <param name="nowMile">当前病害桩号</param>
        /// <param name="curPosX">当前病害中心点x坐标</param>
        /// <param name="curPosY"></param>
        /// <param name="pictureW">图像宽度</param>
        /// <param name="pictureH">图像高度</param>
        /// <param name="roadDir"></param>
        /// <param name="roadWidth">道路宽度</param>
        /// <param name="roadLength">道路长度</param>
        /// <param name="dDiseaseLon"></param>
        /// <param name="dDiseaseLat"></param>
        /// <param name="dDiseaseH"></param>

        public static void getHighAccPosition(bool showGps,List<(double, GPSInfo)> highGpss, int equipType, string _ProjPath, double nowMile, int curPosX, int curPosY, int pictureW, int pictureH, int roadDir,
            double roadWidth, double roadLength,
            ref double dDiseaseLon, ref double dDiseaseLat, ref double dDiseaseH)
        {
            //if (equipType == 0)
            //{

            //}
            //else if (equipType == 1)
            //{
            //    curPosX = pictureW - curPosX;
            //}

            //根据桩号找到图片中心gps
            var soltGps2mileList = highGpss
                .Select((num, index) => new { Number = num.Item1, Index = index }).ToList();

            int closestIndex = 0;
            int closestIndex1 = 0;
            double dCenterLon = 0, dCenterLat = 0, dCenterH = 0; //当前图像的经纬高
            double lastCenterlon = 0, lastCenterlat = 0, lastCenterH = 0;//后一张图像中心的经维高   

            closestIndex = soltGps2mileList
                      .OrderBy(item => Math.Abs(item.Number - nowMile))
                      .First()
                      .Index;
            bool isFirst = closestIndex == 0;
            if (isFirst)
            {
                //第一张图片
                closestIndex1 = closestIndex + 1;

            }
            else
            {
                closestIndex1 = closestIndex - 1;

            }
            dCenterLon = highGpss[closestIndex].Item2._longitude;
            dCenterLat = highGpss[closestIndex].Item2._latitude;
            dCenterH = highGpss[closestIndex].Item2._elevation;

            lastCenterlon = highGpss[closestIndex1].Item2._longitude;
            lastCenterlat = highGpss[closestIndex1].Item2._latitude;
            lastCenterH = highGpss[closestIndex1].Item2._elevation;

            if (dCenterLat == lastCenterlat && dCenterLon == lastCenterlon)
            {
                lastCenterlon = highGpss[closestIndex1 - 1].Item2._longitude;
                lastCenterlat = highGpss[closestIndex1 - 1].Item2._latitude;
                lastCenterH = highGpss[closestIndex1 - 1].Item2._elevation;
            }
            handelGpsInfoInDisease(equipType, showGps, (0, dCenterLon, dCenterLat, dCenterH),
                          (0, lastCenterlon, lastCenterlat, lastCenterH), roadWidth, roadLength,
                           curPosX, curPosY, pictureW, pictureH,
                           ref dDiseaseLon, ref dDiseaseLat, ref dDiseaseH, isFirst);
        }
        public static bool writeHighAccPicture(string _ProjPath, int pictureW, int pictureH,int equip
        )
        {
            List<(double, GPSInfo)> allHighAccInfo = new List<(double, GPSInfo)>();
            GetAllImg(_ProjPath + "\\RoadImg\\Camera0");
            gps2Mile.Clear();
            string mileFile = _ProjPath + "/GPS2Mile.txt";
            string outHighGpstxtPath = _ProjPath + "/HighGps2Mile.txt";

            if (File.Exists(mileFile))
            {
                List<string> gps2MileFile = File.ReadAllLines(mileFile).ToList();
                if (gps2Mile.Count == 0)
                {
                    foreach (var msg in gps2MileFile)
                    {
                        (double, double, double, double) msgLine;
                        string[] msgSp = msg.Split(' ');
                        double longitudeStr = double.Parse(msgSp[1]);
                        double latitudeStr = double.Parse(msgSp[2]);
                        double altitudeStr = double.Parse(msgSp[3]);
                        double mile = double.Parse(msgSp.Last());
                        gps2Mile.Add((mile, longitudeStr, latitudeStr, altitudeStr));
                    }
                }
                if (gps2Mile.Count > 1)
                {

                    int imgCount = _ImgPath.Count;
                    if (gps2Mile.Count <= imgCount)
                    {
                        imgCount = gps2Mile.Count;
                    }
                    var soltGps2mileList = gps2Mile
                .Select((num, index) => new { Number = num.Item1, Index = index }).ToList();
                    for (int i = 0; i < imgCount; i++)
                    {
                        int closestIndex = 0;
                        int closestIndex1 = 0;
                        (double, double, double, double) gpsNow;
                        (double, double, double, double) gpsPre;
                        double dCenterLon = 0, dCenterLat = 0, dCenterH = 0; //当前图像的经纬高
                        //double lastCenterlon = 0, lastCenterlat = 0, lastCenterH = 0;//后一张图像中心的经维高 
                        int curIdx = i;
                        //获取病害经纬度
                        if (curIdx == 0)//在第一张图片 
                        {
                            //根据图片桩号获得 对应经纬度 
                            //第一张图像的经纬度 
                            closestIndex = soltGps2mileList
                                .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                                .First()
                                .Index;
                            gpsNow = gps2Mile[closestIndex];
                            closestIndex1 = soltGps2mileList
                                 .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx + 1].imgmile))
                                 .First()
                                 .Index;
                            if (closestIndex == closestIndex1)
                            {
                                closestIndex1++;
                            }
                            gpsPre = gps2Mile[closestIndex1];
                            handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out dCenterLon, out dCenterLat, out dCenterH, equip,true);
                            //gpsNow = gps2Mile.Where(t => t.Item1 == _ImgPath[curIdx + 1].imgmile).First();
                            //gpsNext = gps2Mile.Where(t => t.Item1 == _ImgPath[curIdx].imgmile).First();
                            //handelGpsInfoInPicture(_ProjPath, gpsNow, gpsNext, out lastCenterlon, out lastCenterlat, out lastCenterH, false);
                        }
                        else
                        {
                            closestIndex = soltGps2mileList
                                 .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx].imgmile))
                                 .First()
                                 .Index;
                            gpsNow = gps2Mile[closestIndex];
                            closestIndex1 = soltGps2mileList
                                 .OrderBy(item => Math.Abs(item.Number - _ImgPath[curIdx - 1].imgmile))
                                 .First()
                                 .Index;
                            if (closestIndex == closestIndex1)
                            {
                                closestIndex1--;
                            }
                            gpsPre = gps2Mile[closestIndex1];

                            handelGpsInfoInPicture(_ProjPath, gpsNow, gpsPre, out dCenterLon, out dCenterLat, out dCenterH, equip,  false);
                        }
                        GPSInfo gpsInfo = new GPSInfo();
                        gpsInfo._longitude = dCenterLon;
                        gpsInfo._latitude = dCenterLat;
                        gpsInfo._elevation = dCenterH;
                        //  gpsInfo._longitude = lastCenterlon;
                        //  gpsInfo._latitude = lastCenterlat;
                        //  gpsInfo._elevation = lastCenterH;
                        allHighAccInfo.Add((_ImgPath[i].imgmile, gpsInfo));
                    }

                }

            }
            if (allHighAccInfo.Count <= 0)
            {
                return false;
            }
            else
            {
                using (FileStream fs = new FileStream(outHighGpstxtPath, FileMode.Create))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        foreach (var item in allHighAccInfo)
                        {
                            string line = string.Join(",", item.Item2._longitude, item.Item2._latitude, item.Item2._elevation,
                                item.Item1);
                            sw.WriteLine(line);
                        }
                    }
                }

                return true;
            }

        }


        public static Dictionary<string, ValueTuple<ValueTuple<double, double, double>, string>> March_GPS(string _ProjPath)
        {
            // 将GPS2Mile.txt中 桩号 对应的经纬度 存入字典
            Dictionary<string, ValueTuple<double, double, double>> GPS_Dic = new Dictionary<string, ValueTuple<double, double, double>>();
            string GPSFile = _ProjPath + "/GPS2Mile.txt";
            if (File.Exists(GPSFile))
            {
                var gps2MileFile = File.ReadAllLines(GPSFile);
                foreach (var msg in gps2MileFile)
                {
                    string[] msgSp = msg.Split(' ');
                    double longitudeStr = double.Parse(msgSp[1]);
                    double latitudeStr = double.Parse(msgSp[2]);
                    double altitudeStr = double.Parse(msgSp[3]);
                    string mile = msgSp.Last();
                    if (GPS_Dic.ContainsKey(mile)) MessageBox.Show("GPS2Mile中存在重复桩号");
                    GPS_Dic[mile] = (longitudeStr, latitudeStr, altitudeStr);
                }
            }
            else
            {
                MessageBox.Show("未找到GPS2Mile.txt文件");
            }

            // 将Road2Mile.txt中 桩号 对应的图片名称 存入字典
            Dictionary<string, string> Pic_Dic = new Dictionary<string, string>();
            string PicFile = _ProjPath + "\\RoadImg\\Camera0\\Road2Mile.txt";
            if (File.Exists(PicFile))
            {
                var road2MileFile = File.ReadAllLines(PicFile);
                foreach (var msg in road2MileFile)
                {
                    string[] msgSp = msg.Split(' ');
                    string mile = msgSp[0];
                    if (Pic_Dic.ContainsKey(mile)) MessageBox.Show("Road2Mile中存在重复桩号");
                    if(!GPS_Dic.ContainsKey(mile)) MessageBox.Show("Road2Mile中有桩号在GPS2Mile中找不到");
                    string filename = msgSp[1];
                    Pic_Dic[mile] = filename;
                }

            }
            else
            {
                MessageBox.Show("未找到Road2Mile.txt文件");
            }
            //if(Pic_Dic.Count != GPS_Dic.Count)
            //{
            //    MessageBox.Show("GPS数据文件与图片数据文件 的桩号数量不匹配");
            //}
             var result = new Dictionary<string, ValueTuple<ValueTuple<double, double, double>, string>>();
            foreach (var key in Pic_Dic.Keys)
            {
                result[key] = (GPS_Dic[key], Pic_Dic[key]);
            }
            return result;

            

        }


        private static List<MyImgMile> _ImgPath = new List<MyImgMile>();

        //获取所有图像
        private static void GetAllImg(string path)
        {
            _ImgPath.Clear();


            //
            string[] imgsinfo = File.ReadAllLines(path + "\\Road2Mile.txt");
            foreach (string str in imgsinfo)
            {
                _ImgPath.Add(new MyImgMile(str));
            }
        }
        /// <summary>
        /// 更新静态类的图片列表
        /// </summary>
        /// <param name="path"></param>
        /// <param name="imgs"></param>
        public static void UpdateAllImg(string path)
        {
            _ImgPath.Clear();
            string[] imgsinfo = File.ReadAllLines(path + "\\Road2Mile.txt");
            foreach (string str in imgsinfo)
            {
                _ImgPath.Add(new MyImgMile(str));
            }
        }


        /// <summary>
        /// 20240308 cwb
        /// 处理当前图像的经纬度
        /// 获得了当前图片中心的经纬度
        /// </summary>
        private static void handelGpsInfoInPicture(string projectPath, (double, double, double, double) gps1, (double, double, double, double) gps2, out double retLon, out double retLat, out double retHeight,int equip, bool bInverse = false)
        {
            //获取参数
            string iniPath = projectPath + "/GPSModel/Config.ini";
            double dYOffsetLength = 2.61;
            double dZOffsetLength = 1.8;
            double dXOffsetLength = 0; 
            Dictionary<string, string> iniSettings = OtherHelper.ParseIniFile(iniPath);
            if (iniSettings.Keys.Contains("YOffset"))
            {
                dYOffsetLength = double.Parse(iniSettings["YOffset"]);
            }
            if (iniSettings.Keys.Contains("XOffset"))
            {
                dXOffsetLength = double.Parse(iniSettings["XOffset"]);

            }
            if (iniSettings.Keys.Contains("ZOffset"))
            {
                dZOffsetLength = double.Parse(iniSettings["ZOffset"]);

            }
            calcLatToPicCenter(gps1.Item2, gps1.Item3, gps1.Item4, gps2.Item2, gps2.Item3, gps2.Item4, dXOffsetLength, dYOffsetLength, dZOffsetLength, out retLon, out retLat, out retHeight, bInverse);

        }
        /// <summary>
        /// 获取病害位置的gps
        /// </summary>
        /// <param name="gps1">1图像GPS</param>
        /// <param name="gps2">2图像GPS</param>
        /// <param name="retLon"></param>
        /// <param name="retLat"></param>
        /// <param name="retHeight"></param>
        /// <param name="bInverse">如果1在前则为true,1图像在2图像后则为false</param>
        private static void handelGpsInfoInDisease( int equip,bool showGps,(double, double, double, double) gps1, (double, double, double, double) gps2,
            double roadWidth, double roadLength,
            int curPosX, int curPosY, int pcitureWidth, int pcitureHeight,
            ref double retLon, ref double retLat,
            ref double retHeight, bool bInverse = false)
        {
            calcLatToPicPos(showGps, gps1.Item2, gps1.Item3, gps1.Item4, gps2.Item2, gps2.Item3, gps2.Item4,
                curPosX, curPosY, pcitureWidth, pcitureHeight,
                ref retLon, ref retLat, ref retHeight,equip, bInverse, roadWidth, roadLength);

        }
    }
}
