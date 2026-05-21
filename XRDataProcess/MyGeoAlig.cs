using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using OperateIniFile;
using Framework.Other;

namespace XRDataProcess
{
    /// <summary>
    /// 计算几何线形
    /// 1、数据格式转换，从原始的二进制数据解析转码为明码的文本文件，通过IMUDecode.dll来解析
    /// 2、横坡计算需要用到翻滚角rolling、纵坡计算需要用到海拔altitude、几何线形测试需要用到航向角heading
    /// 3、解析0x2201，从里面解析出GPS时间，从GPS周，转换为UTC日期
    /// 4、解析0x6403，从里面解析出GPS周内秒，转换为UTC日期时间，然后解析出海拔altitude
    /// 5、解析0x6405，从里面解析出GPS周内秒，转换为UTC日期时间，然后解析出翻滚叫rolling和航向角heading
    /// 6、读取车辙的同步文件，车辙同步文件0.5米输出一条，然后时间查找最近的0x6403和0x6405，抽稀或插值每0.5米一条的等距离间隔数据
    /// 7、利用海拔altitude计算横坡
    /// 8、利用航向角heading计算曲率半径
    /// 9、利用车辙横断面数据和rolling角计算横坡
    /// </summary>
    class MyGeoAlig
    {
        [DllImport("IMUDecode.dll", EntryPoint = "DecodeIMUBin")]
        static extern void DecodeIMUBin(string fpath);

        static XRSetting _Setting = XRSetting.GetInstance();

        public static void ComputeGeoalig(string prj, WinProcessBar bar)
        {
            string imuhon_fpath = prj + @"\camera0\imu.hon"; 
            if(!File.Exists(imuhon_fpath))
            {
                MessageBox.Show("缺失\\camera0\\imu.hon文件，请检查数据是否完整！");
                bar.SetGeoAlig(1.0);
                return;
            }

            string ini_fpath = prj + @"\camera0\rutcfg.ini";
            if (!File.Exists(ini_fpath))
            {
                MessageBox.Show("缺失\\camera0\\rutcfg.ini文件，请检查数据是否完整！");
                bar.SetGeoAlig(1.0);
                return;
            }

            string geoalig_fpath = prj + @"\camera0\geoalig_0.5m.txt"; 
            if (File.Exists(geoalig_fpath))
            {
                bar.SetGeoAlig(1.0);
                return;
            }

            string imuhon_fpath_decode = prj + @"\camera0\imu.hon";
            if (File.Exists(imuhon_fpath_decode))
            {
                //从原始的二进制数据解析转码为明码的文本文件
                DecodeIMUBin(imuhon_fpath_decode);
            }

            //读取横坡的系统误差安装初始值
            IniFiles setingini = new IniFiles(ini_fpath);
            double BaseCrossSlope = Convert.ToDouble(setingini.ReadString("GeoAlig", "BaseCrossSlope", "0.0"));
            
            //解析明码的文本文件
            imuhon_fpath_decode = imuhon_fpath_decode + ".csv";
            string[] oristrs = File.ReadAllLines(imuhon_fpath_decode);
            int len = oristrs.Length;
            int gpsweek = 0;
            bool isfind_0x2201 = false;//是否找到第一个有效的0x2201数据
            for (int i = 0; i < len; ++i )
            {
                if (!isfind_0x2201)
                {
                    if (oristrs[i].StartsWith("0x2201"))
                    {
                        string[] tstrs = oristrs[i].Split(',');
                        gpsweek = int.Parse(tstrs[4]);
                        isfind_0x2201 = true;
                        break;
                    }
                }
            }
            if (!isfind_0x2201)
            {
                MessageBox.Show("没有找到有效的GPS时间，外业数据采集时惯导没有对齐成功，无法解算几何线形！");
                bar.SetGeoAlig(1.0);
                return;
            }
            bar.SetGeoAlig(0.05);

            DateTime startdt = new DateTime(1980,1,6);
            startdt = startdt.AddDays(gpsweek * 7);
            startdt = startdt.AddSeconds(-_Setting.GPSJumpTime);   //减去跳秒

            List<Hg0x6403GeodeticPosition> Geo_0x6403_List = new List<Hg0x6403GeodeticPosition>();
            List<Hg0x6405EulerAttitudes> Geo_0x6405_List = new List<Hg0x6405EulerAttitudes>();
            for (int i = 0; i < len; ++i )
            {
                if(oristrs[i].StartsWith("0x6403"))
                {
                    string[] tstrs = oristrs[i].Split(',');
                    Hg0x6403GeodeticPosition tgeo = new Hg0x6403GeodeticPosition(startdt, tstrs[2], tstrs[5], tstrs[6]);
                    if (tgeo.INSMode == 4)
                    {
                        Geo_0x6403_List.Add(tgeo);
                    }
                }
                else if(oristrs[i].StartsWith("0x6405"))
                {
                    string[] tstrs = oristrs[i].Split(',');
                    Hg0x6405EulerAttitudes tgeo = new Hg0x6405EulerAttitudes(startdt, tstrs[2], tstrs[3], tstrs[5], tstrs[9]);
                    if (tgeo.INSMode == 4)
                    {
                        Geo_0x6405_List.Add(tgeo);
                    }
                }
            }

            if (Geo_0x6403_List.Count < 1 || Geo_0x6405_List.Count < 1)
            {
                MessageBox.Show("没有找到有效的位姿数据，外业数据采集时惯导没有对齐成功，无法解算几何线形！");
                bar.SetGeoAlig(1.0);
                return;
            }
            bar.SetGeoAlig(0.1);

            int geolen = 0;

            //航向角，-180和+180是同一个位置，需要调整为连续的
            //数据里面有时候会在-180和180反转的时候，出现中间0值的点，需要先把这种异常点剔除掉，才好判断-180和+180的翻转
            geolen = Geo_0x6405_List.Count;
            float thval = (float)(Math.PI * 0.5f);
            for (int i = 1; i < geolen - 1; ++i)
            {
                //剔除陡变的异常点
                if (Math.Abs(Geo_0x6405_List[i].Heading - Geo_0x6405_List[i - 1].Heading) > thval
                    && Math.Abs(Geo_0x6405_List[i].Heading - Geo_0x6405_List[i + 1].Heading) > thval)
                {
                    Geo_0x6405_List.RemoveAt(i);
                    i = i - 1;
                    --geolen;
                }
            }
            geolen = Geo_0x6405_List.Count;
            thval = (float)(Math.PI * 1.8f);
            float diff = 0.0f;
            for (int i = 1; i < geolen; ++i)
            {
                //-180和+180是同一个位置，反转调整
                diff = Math.Abs(Geo_0x6405_List[i].Heading - Geo_0x6405_List[i - 1].Heading);
                if (diff > thval)
                {
                    if (Geo_0x6405_List[i].Heading > Geo_0x6405_List[i - 1].Heading)
                    {
                        Geo_0x6405_List[i].Heading = Geo_0x6405_List[i].Heading - (float)(2 * Math.PI);
                    }
                    else if (Geo_0x6405_List[i].Heading < Geo_0x6405_List[i - 1].Heading)
                    {
                        Geo_0x6405_List[i].Heading = Geo_0x6405_List[i].Heading + (float)(2 * Math.PI);
                    }
                }
            }
            bar.SetGeoAlig(0.15);

            //int cnt = 0;
            //string[] outstr = null;
            //outstr = new string[Geo_0x6403_List.Count];
            //foreach (Hg0x6403GeodeticPosition tgeo in Geo_0x6403_List)
            //{
            //    outstr[cnt++] = tgeo.ToString();
            //}
            //File.WriteAllLines(imuhon_fpath_decode + ".6403", outstr);

            //cnt = 0;
            //outstr = new string[Geo_0x6405_List.Count];
            //foreach (Hg0x6405EulerAttitudes tgeo in Geo_0x6405_List)
            //{
            //    outstr[cnt++] = tgeo.ToString();
            //}
            //File.WriteAllLines(imuhon_fpath_decode + ".6405", outstr);

            //解析车辙的同步数据
            string trigger_fpath = prj + @"\camera0\trigger.txt";
            if(!File.Exists(trigger_fpath))
            {
                MessageBox.Show("缺失\\camera0\\imu.hon文件，请检查数据是否完整！");
                bar.SetGeoAlig(1.0);
                return;
            }

            startdt = new DateTime(Geo_0x6403_List[0].GpsDateTime.Date.Year,
                Geo_0x6403_List[0].GpsDateTime.Date.Month,
                Geo_0x6403_List[0].GpsDateTime.Date.Day);

            List<SynTrigger> SynTriggerList = new List<SynTrigger>();
            string[] trigstrs = File.ReadAllLines(trigger_fpath);
            int triglen = trigstrs.Length;
            for (int i = 0; i < triglen; ++i )
            {
                if (trigstrs[i].StartsWith("%XRC"))
                {
                    string[] tstr = trigstrs[i].Split(',');
                    if (tstr == null || tstr.Length < 4)
                    {
                        continue;
                    }

                    SynTrigger tsyn = new SynTrigger(startdt, tstr[2], tstr[1], tstr[3]);
                    if (tsyn.IsOK)
                    {
                        SynTriggerList.Add(tsyn);
                    }
                }
            }

            //检查车辙同步数据有没有丢失，如果有丢失的，插值补齐
            triglen = SynTriggerList.Count;
            List<SynTrigger> InsertSynList = new List<SynTrigger>();
            for (int i = 1; i < triglen; ++i )
            {
                if (SynTriggerList[i].DmiVal - SynTriggerList[i - 1].DmiVal > 500)
                {
                    int insertnum = (int)((SynTriggerList[i].DmiVal - SynTriggerList[i - 1].DmiVal) / 500 - 1);
                    double timestep = (SynTriggerList[i].GpsDateTime - SynTriggerList[i - 1].GpsDateTime).TotalSeconds / (insertnum + 1);
                    for (int j = 0; j < insertnum; ++j )
                    {
                        SynTrigger syntrig = new SynTrigger();
                        syntrig.IsOK = true;
                        syntrig.DmiVal = SynTriggerList[i - 1].DmiVal + (j + 1) * 500;
                        syntrig.TrigIdx = SynTriggerList[i - 1].TrigIdx + (j + 1) * 5;
                        syntrig.GpsDateTime = SynTriggerList[i - 1].GpsDateTime.AddSeconds((j + 1) * timestep);
                        InsertSynList.Add(syntrig);
                    }
                }
            }
            SynTriggerList.AddRange(InsertSynList);
            SynTriggerList.Sort(delegate(SynTrigger x, SynTrigger y) { return x.TrigIdx.CompareTo(y.TrigIdx); });
            bar.SetGeoAlig(0.2);

            //cnt = 0;
            //outstr = new string[triglen];
            //foreach (SynTrigger tgeo in SynTriggerList)
            //{
            //    outstr[cnt++] = tgeo.ToString();
            //}
            //File.WriteAllLines(trigger_fpath + ".txt", outstr);


            //根据同步数据的时间，将导航数据抽样成0.5m一条的，导航数据原始频率100Hz等时间间隔0.01s，抽样插值成等距离间隔0.5m
            triglen = SynTriggerList.Count;
            geolen = Geo_0x6403_List.Count;
            List<Hg0x6403GeodeticPosition> Geo_0x6403_List_new = new List<Hg0x6403GeodeticPosition>();
            int lastfind = 1;
            for (int i = 0, j=0; i < triglen; ++i )
            {
                bool isfind = false;
                for (j = lastfind; j < geolen; ++j)
                {
                    if (DateTime.Compare(Geo_0x6403_List[j - 1].GpsDateTime, SynTriggerList[i].GpsDateTime) <= 0
                        && DateTime.Compare(Geo_0x6403_List[j].GpsDateTime, SynTriggerList[i].GpsDateTime) > 0)
                    {
                        Hg0x6403GeodeticPosition tgeo = new Hg0x6403GeodeticPosition(
                            Geo_0x6403_List[j - 1], Geo_0x6403_List[j], SynTriggerList[i].GpsDateTime);
                        Geo_0x6403_List_new.Add(tgeo);
                        isfind = true;
                        lastfind = j;
                        break;
                    }
                }
                if (!isfind)
                {
                    if (DateTime.Compare(Geo_0x6403_List[0].GpsDateTime, SynTriggerList[i].GpsDateTime) > 0)
                    {
                        Hg0x6403GeodeticPosition tgeo = new Hg0x6403GeodeticPosition(
                            Geo_0x6403_List[0], Geo_0x6403_List[1], SynTriggerList[i].GpsDateTime);
                        Geo_0x6403_List_new.Add(tgeo);
                    }

                    if (DateTime.Compare(Geo_0x6403_List[geolen - 1].GpsDateTime, SynTriggerList[i].GpsDateTime) <= 0)
                    {
                        Hg0x6403GeodeticPosition tgeo = new Hg0x6403GeodeticPosition(
                            Geo_0x6403_List[geolen - 2], Geo_0x6403_List[geolen - 1], SynTriggerList[i].GpsDateTime);
                        Geo_0x6403_List_new.Add(tgeo);
                    }
                }
            }
            bar.SetGeoAlig(0.3);
            
            geolen = Geo_0x6405_List.Count;
            List<Hg0x6405EulerAttitudes> Geo_0x6405_List_new = new List<Hg0x6405EulerAttitudes>();
            lastfind = 1;
            for (int i = 0, j = 0; i < triglen; ++i)
            {
                bool isfind = false;
                for (j = lastfind; j < geolen; ++j)
                {
                    if (DateTime.Compare(Geo_0x6405_List[j - 1].GpsDateTime, SynTriggerList[i].GpsDateTime) <= 0
                        && DateTime.Compare(Geo_0x6405_List[j].GpsDateTime, SynTriggerList[i].GpsDateTime) > 0)
                    {
                        Hg0x6405EulerAttitudes tgeo = new Hg0x6405EulerAttitudes(
                            Geo_0x6405_List[j - 1], Geo_0x6405_List[j], SynTriggerList[i].GpsDateTime, SynTriggerList[i].TrigIdx);
                        Geo_0x6405_List_new.Add(tgeo);
                        isfind = true;
                        lastfind = j;
                        break;
                    }
                }
                if (!isfind)
                {
                    if (DateTime.Compare(Geo_0x6405_List[0].GpsDateTime, SynTriggerList[i].GpsDateTime) > 0)
                    {
                        Hg0x6405EulerAttitudes tgeo = new Hg0x6405EulerAttitudes(
                            Geo_0x6405_List[0], Geo_0x6405_List[1], SynTriggerList[i].GpsDateTime, SynTriggerList[i].TrigIdx);
                        Geo_0x6405_List_new.Add(tgeo);
                    }

                    if (DateTime.Compare(Geo_0x6405_List[geolen - 1].GpsDateTime, SynTriggerList[i].GpsDateTime) <= 0)
                    {
                        Hg0x6405EulerAttitudes tgeo = new Hg0x6405EulerAttitudes(
                            Geo_0x6405_List[geolen - 2], Geo_0x6405_List[geolen - 1], SynTriggerList[i].GpsDateTime, SynTriggerList[i].TrigIdx);
                        Geo_0x6405_List_new.Add(tgeo);
                    }
                }
            }
            bar.SetGeoAlig(0.35);

            //cnt = 0;
            //outstr = new string[Geo_0x6403_List_new.Count];
            //foreach (Hg0x6403GeodeticPosition tgeo in Geo_0x6403_List_new)
            //{
            //    outstr[cnt++] = tgeo.ToString();
            //}
            //File.WriteAllLines(imuhon_fpath_decode + ".6403.0.5", outstr);

            //cnt = 0;
            //outstr = new string[Geo_0x6405_List_new.Count];
            //foreach (Hg0x6405EulerAttitudes tgeo in Geo_0x6405_List_new)
            //{
            //    outstr[cnt++] = tgeo.ToString();
            //}
            //File.WriteAllLines(imuhon_fpath_decode + ".6405.0.5", outstr);

            //曲率
            geolen = Geo_0x6405_List_new.Count;
            float dmival = 0.0f;
            float curvature = 0.0f;
            string[] str_vals = new string[geolen - 1];
            for (int i = 1; i < geolen; ++i )
            {
                dmival += 0.5f;
                curvature = (Geo_0x6405_List_new[i].Heading - Geo_0x6405_List_new[i - 1].Heading) / 0.5f;
                str_vals[i - 1] = string.Format("{0},{1},{2},{3}",
                    dmival, curvature, Geo_0x6405_List_new[i].Heading, Geo_0x6405_List_new[i - 1].Heading);
            }
            File.WriteAllLines(prj + @"\camera0\imu.hon.Curvature", str_vals);
            bar.SetGeoAlig(0.4);

            //计算纵坡
            geolen = Geo_0x6403_List.Count;
            dmival = 0.0f;
            float HeightSlope = 0.0f;
            str_vals = new string[geolen - 1];
            for (int i = 1; i < geolen; ++i )
            {
                dmival += 0.5f;
                HeightSlope = (float)(Geo_0x6403_List[i].AltitudeAboveEllipsoid - Geo_0x6403_List[i - 1].AltitudeAboveEllipsoid) / 0.5f;
                str_vals[i - 1] = string.Format("{0},{1},{2},{3}", 
                    dmival, HeightSlope, Geo_0x6403_List[i].AltitudeAboveEllipsoid, Geo_0x6403_List[i - 1].AltitudeAboveEllipsoid);
            }
            File.WriteAllLines(prj + @"\camera0\imu.hon.HeightSlope", str_vals);
            bar.SetGeoAlig(0.45);

            //解析车辙原始数据
            GetRutCrossAngle(prj, ref Geo_0x6405_List_new);
            bar.SetGeoAlig(0.9);

            //计算横坡
            geolen = Geo_0x6405_List_new.Count;
            dmival = 0.0f;
            float CrossSlope = 0.0f;
            str_vals = new string[geolen - 1];
            for (int i = 1; i < geolen; ++i)
            {
                dmival += 0.5f;
                CrossSlope = (float)Math.Tan(Geo_0x6405_List_new[i].Roll - BaseCrossSlope - Geo_0x6405_List_new[i].RutAngle);
                str_vals[i - 1] = string.Format("{0},{1},{2},{3},{4}",
                    dmival, CrossSlope, Geo_0x6405_List_new[i].Roll, Geo_0x6405_List_new[i].RutAngle, BaseCrossSlope);
            }
            File.WriteAllLines(prj + @"\camera0\imu.hon.CrossSlope", str_vals);
            bar.SetGeoAlig(1.0);
        }

        public static bool GetRutCrossAngle(string prj, ref List<Hg0x6405EulerAttitudes> Geo_0x6405_List_new)
        {
            string rutdatapath = prj + @"\camera0\data";
            if (!Directory.Exists(rutdatapath))
            {
                MessageBox.Show("没有找到车辙原始数据！请检查！");
                return false;
            }

            string[] _dats = Directory.GetFiles(rutdatapath, "*.dat");
            Array.Sort(_dats);
            if (_dats.Length <= 0)
            {
                MessageBox.Show("没有找到车辙原始数据！请检查！");
                return false;
            }

            string inifpath = prj + @"\camera0\rutcfg.ini";
            if(!File.Exists(inifpath))
            {
                MessageBox.Show("没有找到车辙配置文件\\camera0\\rutcfg.ini！请检查！");
                return false;
            }

            string matxfapth = prj + @"\camera0\Mat_X.cal";
            string matzfapth = prj + @"\camera0\Mat_Z.cal";
            if (!File.Exists(matxfapth))
            {
                MessageBox.Show("没有找到车辙配置文件\\camera0\\Mat_X.cal！请检查！");
                return false;
            }
            if(!File.Exists(matzfapth))
            {
                MessageBox.Show("没有找到车辙配置文件\\camera0\\Mat_Z.cal！请检查！");
                return false;
            }
            IniFiles rutcfg = new IniFiles(inifpath);
            int hpix = rutcfg.ReadInteger("camera", "hpixel", 2048);
            int vpix = rutcfg.ReadInteger("camera", "calivpix", 3200);
            int linesidx = rutcfg.ReadInteger("camera", "rutastart", 0);
            int lineeidx = rutcfg.ReadInteger("camera", "rutcend", 2048);
            double[,] MatX = new double[vpix, hpix];
            double[,] MatZ = new double[vpix, hpix];
            double[] MatTmp = new double[hpix];
            byte[] rbmat = new byte[hpix * 8];
            byte[] rbarr = new byte[hpix * 2];
            int bidx = hpix * 8;
            short[] profile = new short[hpix];
            LinePoint[] profilept = new LinePoint[hpix];

            for (int i = 0; i < hpix; ++i )
            {
                profilept[i] = new LinePoint();
            }
            
            using (FileStream frstream = new FileStream(matxfapth, FileMode.Open))
            {
                for (int n = 0; n < vpix; ++n)
                {
                    frstream.Read(rbmat, 0, bidx);
                    Buffer.BlockCopy(rbmat, 0, MatTmp, 0, rbmat.Length);
                    for (int k = 0; k < hpix; ++k)
                    {
                        MatX[n, k] = MatTmp[k];
                    }
                }
            }
            using (FileStream frstream = new FileStream(matzfapth, FileMode.Open))
            {
                for (int n = 0; n < vpix; ++n)
                {
                    frstream.Read(rbmat, 0, bidx);
                    Buffer.BlockCopy(rbmat, 0, MatTmp, 0, rbmat.Length);
                    for (int k = 0; k < hpix; ++k)
                    {
                        MatZ[n, k] = MatTmp[k];
                    }
                }
            }

            int cidx = 0;
            int geolen = Geo_0x6405_List_new.Count;
            for (int i = 0, gi = 0; i < _dats.Length && gi < geolen; ++i)
            {
                //读取所有dat文件
                using (FileStream frstream = new FileStream(_dats[i], FileMode.Open))
                {
                    bidx = hpix * 2;
                    while (gi < geolen && frstream.Read(rbarr, 0, bidx) > 0)
                    {
                        if (cidx == Geo_0x6405_List_new[gi].TrigIdx)
                        {
                            Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                            int m = 0, n = 0;
                            for (; m < hpix; ++m)
                            {
                                if (profile[m] != 0)
                                {
                                    profilept[n].X = MatX[profile[m], m];
                                    profilept[n].Y = MatZ[profile[m], m];
                                    ++n;
                                }
                            }

                            int tlineeidx = lineeidx - (m - n);
                            Geo_0x6405_List_new[gi].RutAngle = GetRutAngle(profilept, linesidx, lineeidx);
                            ++gi;
                        }
                        ++cidx;
                    }
                    frstream.Close();
                }
            }
            return true;
        }

        public static double GetRutAngle(LinePoint[] pt, int sidx, int eidx)
        {
            double angle = 0;
            double k = 0, b = 0;

            List<LinePoint> ptlist = new List<LinePoint>();
            ptlist.AddRange(pt);

            FitLine(pt, sidx, eidx, ref k, ref b);

            double sumY = 0;
            for (int i = sidx; i < eidx; ++i )
            {
                pt[i].Y = pt[i].Y - (pt[i].X * k - b);
                sumY += pt[i].Y;
            }

            double meanY = sumY / (eidx - sidx);
            double stdY = 0;
            for (int i = sidx; i < eidx; ++i )
            {
                stdY += (pt[i].Y - meanY) * (pt[i].Y - meanY); 
            }
            stdY = Math.Sqrt(stdY / (eidx - sidx));

            ptlist.RemoveRange(eidx, ptlist.Count - eidx);  //先把尾部的无效点去掉
            double stdThresh = 3.34f * stdY + 2.0f;
            bool isremove = false;
            for (int i = eidx - 1; i >= sidx; --i )
            {
                if (Math.Abs(pt[i].Y - meanY) > stdThresh)
                {
                    ptlist.RemoveAt(i); //从后向前去掉中间的异常点
                    isremove = true;
                }
            }
            ptlist.RemoveRange(0, sidx);  //去掉头部的无效点

            if (isremove)
            {
                FitLine(ptlist.ToArray(), 0, ptlist.Count, ref k, ref b);
            }

            angle = Math.Atan(k);
            
            return angle;
        }

        public static void FitLine(LinePoint[] pt, int sidx, int eidx, ref double k, ref double b)
        {
            int n = eidx - sidx;
            double A = 0, B = 0, C = 0, D = 0, temp = 0;
            for (int i = sidx; i < eidx; i++)
            {
                A += pt[i].X * pt[i].X;
                B += pt[i].X;
                C += pt[i].X * pt[i].Y;
                D += pt[i].Y;
            }

            temp = n * A - B * B;
            if (temp != 0)
            {
                k = (n * C - B * D) / temp;
                b = (A * D - B * C) / temp;
            }
            else
            {
                k = 0;
                b = 0;
            }
        }
    }

    class LinePoint
    {
        public double X;
        public double Y;
    }

    class Hg0x6403GeodeticPosition
	{
        public DateTime GpsDateTime;

        /// <summary>
        /// 海拔高度，m
        /// </summary>
        public double AltitudeAboveEllipsoid; //m 

        /// <summary>
        /// 惯导工作模式，1-待命（默认），2-粗对齐，4-组合导航，15-无效
        /// </summary>
        public int INSMode;

        public Hg0x6403GeodeticPosition()
        { }

        public Hg0x6403GeodeticPosition(DateTime startdate, string strgps, string stralt, string strinsmode)
        {
            double GpsTov = double.Parse(strgps);
            GpsDateTime = startdate.AddSeconds(GpsTov);

            AltitudeAboveEllipsoid = double.Parse(stralt);
            INSMode = int.Parse(strinsmode);
        }

        public override string ToString()
        {
            return string.Format("{0:yyyy-MM-dd,HH:mm:ss.fff},{1},{2}", GpsDateTime, AltitudeAboveEllipsoid, INSMode);
        }
        
        /// <summary>
        /// 插值
        /// </summary>
        /// <param name="SInfo">上一个</param>
        /// <param name="EInfo">下一个</param>
        /// <param name="insertime">插值日期时间</param>
        public Hg0x6403GeodeticPosition(Hg0x6403GeodeticPosition SInfo, Hg0x6403GeodeticPosition EInfo, DateTime insertime)
        {
            System.TimeSpan es_date = EInfo.GpsDateTime - SInfo.GpsDateTime;
            System.TimeSpan cs_date = insertime - SInfo.GpsDateTime;
            double k = cs_date.TotalMilliseconds / es_date.TotalMilliseconds;

            GpsDateTime = insertime;
            AltitudeAboveEllipsoid = k * (EInfo.AltitudeAboveEllipsoid - SInfo.AltitudeAboveEllipsoid) + SInfo.AltitudeAboveEllipsoid;
            INSMode = SInfo.INSMode;
        }
    }

    class Hg0x6405EulerAttitudes
	{
        public DateTime GpsDateTime;

        /// <summary>
        /// 翻滚角，弧度
        /// </summary>
        public float Roll;

        /// <summary>
        /// 航向角，弧度
        /// </summary>
        public float Heading;

        /// <summary>
        /// 惯导工作模式，1-待命（默认），2-粗对齐，4-组合导航，15-无效
        /// </summary>
        public int INSMode;

        /// <summary>
        /// 横断面触发的帧序号
        /// </summary>
        public long TrigIdx;

        /// <summary>
        /// 车辙横断面测出的倾斜角度，弧度
        /// </summary>
        public double RutAngle;

        public Hg0x6405EulerAttitudes()
        { }

        public Hg0x6405EulerAttitudes(DateTime startdate, string strgps, string strroll, string strheading, string strinsmode)
        {
            double GpsTov = double.Parse(strgps);
            GpsDateTime = startdate.AddSeconds(GpsTov);

            Roll = float.Parse(strroll);
            Heading = float.Parse(strheading);
            INSMode = int.Parse(strinsmode);
        }

        public override string ToString()
        {
            return string.Format("{0:yyyy-MM-dd,HH:mm:ss.fff},{1},{2},{3}", GpsDateTime, Roll, Heading, INSMode);
        }

        /// <summary>
        /// 插值
        /// </summary>
        /// <param name="SInfo">上一个</param>
        /// <param name="EInfo">下一个</param>
        /// <param name="insertime">插值日期时间</param>
        public Hg0x6405EulerAttitudes(Hg0x6405EulerAttitudes SInfo, Hg0x6405EulerAttitudes EInfo, DateTime insertime, long trigcnt)
        {
            System.TimeSpan es_date = EInfo.GpsDateTime - SInfo.GpsDateTime;
            System.TimeSpan cs_date = insertime - SInfo.GpsDateTime;
            double k = cs_date.TotalMilliseconds / es_date.TotalMilliseconds;

            GpsDateTime = insertime;
            Roll = (float)(k * (EInfo.Roll - SInfo.Roll) + SInfo.Roll);
            Heading = (float)(k * (EInfo.Heading - SInfo.Heading) + SInfo.Heading);
            INSMode = SInfo.INSMode;

            TrigIdx = trigcnt;
        }
    }

    class SynTrigger
    {
        /// <summary>
        /// 断面同步的时间
        /// </summary>
        public DateTime GpsDateTime;

        /// <summary>
        /// 触发断面的序号
        /// </summary>
        public long TrigIdx;

        /// <summary>
        /// DMI脉冲计数
        /// </summary>
        public long DmiVal;

        /// <summary>
        /// 是否解析正常
        /// </summary>
        public bool IsOK = false;

        public SynTrigger()
        { }

        public SynTrigger(DateTime startdate, string strtime, string strtrigidx, string strdmival)
        {
            if (strtime.Length == 9 && strtrigidx.Length == 8 && strdmival.Length == 8)
            {
                long tmp = 0;
                try { tmp = int.Parse(strtime.Substring(0, 2)); }
                catch (Exception ) { IsOK = false; return; }
                GpsDateTime = startdate.AddHours(tmp);

                try { tmp = int.Parse(strtime.Substring(2, 2)); }
                catch (Exception ) { IsOK = false; return; }
                GpsDateTime = GpsDateTime.AddMinutes(tmp);

                try { tmp = int.Parse(strtime.Substring(4, 2)); }
                catch (Exception ) { IsOK = false; return; }
                GpsDateTime = GpsDateTime.AddSeconds(tmp);

                try { tmp = int.Parse(strtime.Substring(6, 3)); }
                catch (Exception  ) { IsOK = false; return; }
                GpsDateTime = GpsDateTime.AddMilliseconds(tmp);

                try { TrigIdx = Convert.ToInt64(strtrigidx, 16); }
                catch (Exception  ) { IsOK = false; return; }

                try { DmiVal = Convert.ToInt64(strdmival, 16); }
                catch (Exception  ) { IsOK = false; return; }

                IsOK = true;
            }
            else
            {
                IsOK = false;
            }
        }

        public override string ToString()
        {
            return string.Format("{0:yyyy-MM-dd,HH:mm:ss.fff},{1},{2}", GpsDateTime, TrigIdx, DmiVal);
        }
    }
}
