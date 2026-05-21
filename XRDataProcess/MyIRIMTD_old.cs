using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Xml;
using OperateIniFile;
using System.Windows.Forms;

namespace XRDataProcess
{
    public class MyIRIMTD_old
    {
        public static double[][] _LasCali;
        public static double[][] _AccCali;
        public static double[][] _MTDCali;

        public static double[] _LasFilter;
        public static double[] _AccFilter;

        public static int _LasGDoff;
        public static int _AccGDoff;
        public static int _AccLaterLasOff=400;
        private static int _FrameSize=42;
        private static double _IntegraTime=0.000125;//AD采用周期，单位s
        private static int _PluseParm = 50;//编码器分频系数为50
        private static double _LocalGravity = 9.7913;
                
        public static void LoadParm(string iriname)
        {
            _LasCali = new double[2][];
            _AccCali = new double[2][];
            for (int i = 0; i < 2; ++i)
            {
                _LasCali[i] = new double[2];
                _AccCali[i] = new double[2];
            }
            IniFiles iriparm = new IniFiles(iriname);
            _FrameSize = iriparm.ReadInteger("Frame", "FrameSize", 42);
            _AccLaterLasOff = iriparm.ReadInteger("Frame", "LaterLas", 400);
            _LocalGravity = Convert.ToDouble(iriparm.ReadString("Acc", "LocalGravity", "9.7913"));
        }

        public static void LoadParm(DirectoryInfo prjdir)
        {
            _LasCali = new double[2][];
            _AccCali = new double[2][];
            for (int i = 0; i < 2; ++i)
            {
                _LasCali[i] = new double[2];
                _AccCali[i] = new double[2];
                if (Directory.Exists(string.Format(@"{0}\IRIMTD\DAQ{1}", prjdir.FullName, i)))
                {
                    string iris = File.ReadAllText(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prjdir.FullName, i));
                    if(!iris.Contains(Environment.NewLine))
                    {
                        iris = iris.Replace("\n",Environment.NewLine);
                        File.WriteAllText(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prjdir.FullName, i),iris);
                    }

                    IniFiles iriparm = new IniFiles(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prjdir.FullName, i));
                    _LasCali[i][0] = Convert.ToDouble(iriparm.ReadString("CaliLaser", "p1", "1"));
                    _LasCali[i][1] = Convert.ToDouble(iriparm.ReadString("CaliLaser", "p2", "0"));
                    _AccCali[i][0] = Convert.ToDouble(iriparm.ReadString("CaliAcc", "p1", "1"));
                    _AccCali[i][1] = Convert.ToDouble(iriparm.ReadString("CaliAcc", "p2", "0"));
                }
            }
            _MTDCali = new double[2][];
            for (int i = 0; i < 2; i++ )
            {
                _MTDCali[i] = new double[2];
                if (Directory.Exists(string.Format(@"{0}\IRIMTD\Laser{1}", prjdir.FullName, i)))
                {
                    string mtds = File.ReadAllText(string.Format(@"{0}\IRIMTD\Laser{1}\Setting.ini", prjdir.FullName, i));
                    if (!mtds.Contains(Environment.NewLine))
                    {
                        mtds = mtds.Replace("\n", Environment.NewLine);
                        File.WriteAllText(string.Format(@"{0}\IRIMTD\Laser{1}\Setting.ini", prjdir.FullName, i), mtds);
                    }
                    IniFiles mtdparm = new IniFiles(mtds);
                    _MTDCali[i][0] = mtdparm.ReadInteger("Parm", "MTD_k", 1);
                    _MTDCali[i][1] = mtdparm.ReadInteger("Parm", "MTD_b", 0);
                }
            }
            InitMTDX();
        }

        public static void LoadFilter()
        {
            string []sdata = File.ReadAllLines("Laser.filter");
            int len = sdata.Length;
            _LasFilter = new double[len];
            for (int i=0; i<len; i++)
            {
                _LasFilter[i] = double.Parse(sdata[i]);
            }
            _LasGDoff = len / 2;
            int cflen = (int)Math.Pow(2, Math.Ceiling(Math.Log(len, 2)));
            double[] _tlasf = new double[cflen];
            Array.Copy(_LasFilter, _tlasf, len);

            sdata = File.ReadAllLines("Acc.filter");
            len = sdata.Length;
            _AccFilter = new double[len];
            for (int i = 0; i < len; i++)
            {
                _AccFilter[i] = double.Parse(sdata[i]);
            }
            _AccGDoff = len / 2;
            cflen = (int)Math.Pow(2, Math.Ceiling(Math.Log(len, 2)));
            double[] _taccf = new double[cflen];
            Array.Copy(_AccFilter, _taccf, len);
        }

        public static void ComputeIRI(string prj, int side, WinProcessBar bar)
        {
            int fcnt=0;
            List<string> daqfile = new List<string>();
            GetAllFiles(prj, side, "DAQ", "*.daq", ref daqfile);
            bar.SetIRIVal(0.1);            

            string resamplefname = string.Format("{0}\\IRIMTD\\DAQ{1}\\resample.txt", prj, side);
            ushort lastdmi = 0;

            if (!File.Exists(resamplefname))
            {
                foreach (string df in daqfile)
                {
                    lastdmi = AnalysisOneFile_2(df, resamplefname, lastdmi, side);
                    bar.SetIRIVal(0.1 + (++fcnt)*0.6/daqfile.Count);
                }
            }
            //else
            //{
            //    File.Delete(resamplefname);
            //    foreach (string df in daqfile)
            //    {
            //        lastdmi = AnalysisOneFile_2(df, resamplefname, lastdmi, side);
            //    }
            //}
            //先计算1m的平整度，后面再累计计算10m、100m、1000m的平整度
            GenerateIRI(new Tuple<string, int, int>(resamplefname, 10, _PluseParm));
            bar.SetIRIVal(0.8);
            //GenerateIRI(new Tuple<string, int, int>(resamplefname, 100, _PluseParm));
            bar.SetIRIVal(0.9);
            //GenerateIRI(new Tuple<string, int, int>(resamplefname, 1000, _PluseParm));
            bar.SetIRIVal(1);
        }
 
        private static int[] m_laserX;
        private const int SizeLaserData = 24;
        private const int MTDPOINT = 301;
        private static int m_PointNN;

        public static void ComputeMTD_2(string prj, int side, int featurelen, WinProcessBar bar)
        {
            string mtd10fname = string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_10m.txt", prj, side);
            string resfname = string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_{2}m.txt", prj, side, featurelen);

            if (!File.Exists(mtd10fname))
            {
                return;
            }

            int NUM = featurelen / 10;

            FileStream fr = new FileStream(mtd10fname, FileMode.Open);
            StreamReader sr = new StreamReader(fr);

            FileStream fw = new FileStream(resfname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);

            string linestr;
            string[] s;
            int partlinecnt = 0;
            int mtdcnt = 0;
            double sumval = 0;
            while((linestr = sr.ReadLine())!=null)
            {
                s = linestr.Split(' ');
                sumval += double.Parse(s[1]);
                if (++partlinecnt == NUM)
                {
                    partlinecnt = 0;
                    sumval /= NUM;

                    sw.WriteLine(string.Format("{0} {1}", ++mtdcnt, sumval));
                    sumval = 0;
                }
            }

            sw.Close();
            fw.Close();

            sr.Close();
            fr.Close();
        }
    
        /// <summary>
        /// 计算构造深度---计算的是10m的构造深度，其他的由这个的平均值计算
        /// </summary>
        /// <param name="prj"></param>
        /// <param name="side"></param>
        public static void ComputeMTD(string prj, int side, int featurelen, WinProcessBar bar)
        {
            if (File.Exists(string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_{2}m.txt", prj, side, featurelen)))
            {
                return;
            }                        

            List<string> lasfile = new List<string>();
            GetAllFiles(prj, side, "Laser", "*.las", ref lasfile);
            bar.SetMTDVal(0.1);

            bool m_IsMTDFrame0 = true;
            double tempvaly = 0, tempval = 0, ttval = 0, oldttval=0;
            int tempvalx = 0;
            int m_laserYIdx = 0;
            double m_laserYSum = 0;
            double m_laserYYSum = 0;
            double m_laserXYSum = 0;
            double m_laserXXYSum = 0;

            double m_SMTDdSum = 0;
            int m_SMTDdCnt = 0;
            int m_SMTDSubcnt = 0;//不足一个SMTD的点数计数
            int m_SMTDValCnt = 0;//有效的SMTD个数
            int m_featureDis = featurelen;
            int m_SMTDdnum = (int)(m_featureDis * 1000 / (MTDPOINT - 1));//总共的SMTD个数
            int m_SMTDSubnum = m_featureDis * 1000 - m_SMTDdnum * (MTDPOINT - 1);//不足一个SMTD的点数

            double m_CurMTD;
            int m_MTDCnt=0;
            int filecnt = 0;

            ulong framecnt = 0;
            FileStream fmtd = new FileStream(string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_{2}m.txt", prj, side, featurelen), FileMode.Create);
            StreamWriter smtd = new StreamWriter(fmtd);
            foreach (string lf in lasfile)
            {
                FileInfo lasf = new FileInfo(lf);
                long filesize = lasf.Length;
                long framenum = filesize / SizeLaserData;

                FileStream fr = File.OpenRead(lf);
                BinaryReader br = new BinaryReader(fr);

                for (int i = 0; i < framenum; ++i, ++framecnt)
                {
                    fr.Position += 16;
                    ttval = br.ReadDouble();
                    
                    if (m_IsMTDFrame0)
                    {
                        if (++m_SMTDSubcnt >= m_SMTDSubnum)
                        {
                            m_IsMTDFrame0 = false;
                        }
                        oldttval = ttval;
                    }
                    else
                    {
                        if (Math.Abs(oldttval - ttval) > 10)
                        {
                            ttval = oldttval;
                        }

                        tempvaly = ttval;
                        tempvalx = m_laserX[m_laserYIdx];
                        m_laserYSum += tempvaly;
                        m_laserYYSum += tempvaly * tempvaly;
                        tempvaly *= tempvalx;
                        m_laserXYSum += tempvaly;
                        m_laserXXYSum += tempvaly * tempvalx;

                        if (++m_laserYIdx == MTDPOINT)
                        {
                            bar.SetMTDVal(0.1 + 0.9 / lasfile.Count * (filecnt + i / framenum));

                            tempval = (m_PointNN - 1) * m_laserYSum - 12 * m_laserXXYSum;
                            tempval = 5 * tempval * tempval / (4 * (m_PointNN - 4));
                            tempval = (12 * m_laserXYSum * m_laserXYSum + tempval) / (m_PointNN - 1);
                            tempval = (MTDPOINT * m_laserYYSum - m_laserYSum * m_laserYSum - tempval) / m_PointNN;

                            tempval = tempval > 0 ? tempval : 0;
                            m_SMTDdSum += Math.Sqrt(tempval);
                            ++m_SMTDValCnt;

                            if (++m_SMTDdCnt == m_SMTDdnum)
                            {
                                m_CurMTD = m_SMTDdSum / m_SMTDValCnt * _MTDCali[side][0] + _MTDCali[side][1];
                                ++m_MTDCnt;

                                smtd.WriteLine(string.Format("{0} {1}", m_MTDCnt, m_CurMTD));

                                m_SMTDdSum = 0;
                                m_SMTDdCnt = 0;
                                m_SMTDValCnt = 0;
                                m_SMTDSubcnt = 0;
                                m_IsMTDFrame0 = true;
                            }

                            tempvaly = ttval;
                            tempvalx = m_laserX[0];
                            m_laserYSum = tempvaly;
                            m_laserYYSum = tempvaly * tempvaly;
                            tempvaly *= tempvalx;
                            m_laserXYSum = tempvaly;
                            m_laserXXYSum = tempvaly * tempvalx;
                            m_laserYIdx = 1;
                        }
                        oldttval = ttval;
                    }
                }
                br.Close();
                fr.Close();
                filecnt++;
            }

            smtd.Close();
            fmtd.Close();
            bar.SetMTDVal(1.0);
        }

        public static void InitMTDX()
        {
            m_laserX = new int[MTDPOINT];
            for (int i = 0; i < MTDPOINT; i++)
            {
                m_laserX[i] = i - MTDPOINT / 2;
            }
            m_PointNN = MTDPOINT * MTDPOINT;
        }

        private static void GetAllFiles(string prj, int side,string subfolder, string ftype, ref List<string> filepath)
        {
            string prjpath = string.Format("{0}\\IRIMTD\\{1}{2}", prj, subfolder, side);
            DirectoryInfo dir = new DirectoryInfo(prjpath);
            FileInfo[] files = dir.GetFiles(ftype);
            foreach(FileInfo f in files)
            {
                filepath.Add(f.FullName);
            }
        }
        
        /// <summary>
        /// 生成路面平整度，tIRI_距离.txt
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="iridis"></param>
        /// <param name="pluseparm"></param>
        private static void GenerateIRI(Tuple<string, int, int> tuple)
        {
            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(tuple.Item1);
            int len = sdata.Length;

            //250mm抽样
            int qplusenum = Convert.ToInt32(250 / tuple.Item3);//250mm内有多少个编码器脉冲
            string savefname = tuple.Item1.Replace("resample.txt", string.Format("ReSample250.txt", tuple.Item2));            
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i]);
                }
            }
            sw.Close();
            fw.Close();

            //计算IRI
            int plusenum = Convert.ToInt32(tuple.Item2 * 4);//IRI距离内有多少个250mm
            int iricnt = 0;
            double irisum = 0;

            sdata = File.ReadAllLines(savefname);
            len = sdata.Length; double[] oridata = new double[len];
            for (int i = 0; i < len; i++)
            {
                oridata[i] = double.Parse(sdata[i].Split('\t')[2]);
            }

            savefname = tuple.Item1.Replace("resample.txt", string.Format("IRI_{0}m.txt", tuple.Item2));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            for (int i = 1; i < len; ++i)
            {
                if (i % plusenum == 0)
                {
                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irisum / tuple.Item2));
                    irisum = 0;
                }
                else
                {
                    irisum += Math.Abs(oridata[i] - oridata[i - 1]);
                }
            }
            sw.Close();
            fw.Close();
        }

        private static ushort AnalysisOneFile(string fname, string resfname, ushort lastdmi, int side)
        {
            FileInfo fileInfo = new FileInfo(fname);
            long datasize = fileInfo.Length;
            long datalen = datasize / _FrameSize+100;

            ushort[] _dmidata = new ushort[datalen];
            double[] _accdata = new double[datalen];
            double[] _lasdata = new double[datalen];
            double[] _tempdata = new double[datalen];
            double[] _taccdata = new double[datalen];

            //加载数据
            LoadADData(fname, ref _accdata, ref _lasdata, ref _dmidata, ref datalen, side);

            //激光器FIR滤波
            FilterConv(_lasdata, datalen, _LasFilter, ref _tempdata);
            MoveOff(_tempdata, datalen, ref _lasdata, true, _LasGDoff);

            //加速度计积分
            FilterConv(_accdata, datalen, _AccFilter, ref _tempdata);
            MoveOff(_tempdata, datalen, ref _taccdata, true, _AccGDoff);
            ArrSub(_accdata, _taccdata, datalen, ref _tempdata);
            IntegraOnce(_tempdata, datalen, ref _accdata, _IntegraTime, 0);

            FilterConv(_accdata, datalen, _AccFilter, ref _tempdata);
            MoveOff(_tempdata, datalen, ref _taccdata, true, _AccGDoff);
            ArrSub(_accdata, _taccdata, datalen, ref _tempdata);
            IntegraOnce(_tempdata, datalen, ref _accdata, _IntegraTime, 0);

            //加速度计滞后，右移
            MoveOff(_accdata, datalen, ref _taccdata, false, _AccLaterLasOff);

            //编码器抽样
            ReSampleSave(_lasdata, _taccdata, _dmidata, lastdmi, datalen, resfname);
            return _dmidata[datalen-1];
        }

        private static ushort AnalysisOneFile_2(string fname, string resfname, ushort lastdmi, int side)
        {
            FileInfo fileInfo = new FileInfo(fname);
            long datasize = fileInfo.Length;
            long datalen = datasize / _FrameSize+100;

            ushort[] _dmidata = new ushort[datalen];
            double[] _accdata = new double[datalen];
            double[] _lasdata = new double[datalen];
            double[] _tempdata = new double[datalen];
            double[] _taccdata = new double[datalen];

            //加载数据
            LoadADData(fname, ref _accdata, ref _lasdata, ref _dmidata, ref datalen, side);

            //激光器FIR滤波
            FilterConv(_lasdata, datalen, _LasFilter, ref _tempdata);
            MoveOff(_tempdata, datalen, ref _lasdata, true, _LasGDoff);

            //加速度计积分--》速度
            IntegraOnce(_accdata, datalen, ref _tempdata, _IntegraTime, 0);
            //速度积分--》位移
            IntegraOnce(_tempdata, datalen, ref _accdata, _IntegraTime, 0);
            //位移滤波
            FilterConv(_accdata, datalen, _AccFilter, ref _tempdata);
            MoveOff(_tempdata, datalen, ref _taccdata, true, _AccGDoff);
            ArrSub(_accdata, _taccdata, datalen, ref _tempdata);

            //加速度计滞后，右移
            MoveOff(_tempdata, datalen, ref _taccdata, false, _AccLaterLasOff);

            //编码器抽样
            ReSampleSave(_lasdata, _taccdata, _dmidata, lastdmi, datalen, resfname);
            return _dmidata[datalen - 1];
        }

        private static void SaveTempData(double[] data, long dlen, string fname)
        {
            FileStream fw = new FileStream(fname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);

            for (int i = 0; i < dlen; i++)
            {
                sw.WriteLine(data[i]);
            }

            sw.Close();
            fw.Close();
        }

        private static void ReSampleSave(double[] las, double []acc, ushort[]dmi, ushort lastdmi, 
            long dlen, string fname)
        {
            FileStream fw = new FileStream(fname, FileMode.Append);
            StreamWriter sw = new StreamWriter(fw);

            for (int i = 0; i < dlen; ++i )
            {
                if (dmi[i] >= 26173 && lastdmi < 26173)
                {
                    sw.WriteLine(las[i] - acc[i]*1000);
                }
                lastdmi = dmi[i];
            }

            sw.Close();
            fw.Close();
        }

        /// <summary>
        /// 两个数组对应值相减，res=A-B
        /// </summary>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <param name="dlen"></param>
        /// <param name="res"></param>
        private static void ArrSub(double[] A, double[] B, long dlen, ref double[] res)
        {
            for (int i = 0; i < dlen; i++ )
            {
                res[i] = A[i] - B[i];
            }
        }

        /// <summary>
        /// 一次积分
        /// </summary>
        /// <param name="ori">待积分数据</param>
        /// <param name="dlen">积分长度</param>
        /// <param name="res">积分结果</param>
        /// <param name="time">时间间隔，已经*0.5</param>
        /// <param name="res0">结果初值</param>
        private static void IntegraOnce(double []ori, long dlen, ref double[] res, double time, double res0)
        {
            time = time * 0.5;
            res[0] = res0;
            for (int i = 1; i < dlen; ++i )
            {
                res[i] = res[i - 1] + (ori[i] + ori[i - 1]) * time;
            }
        }

        /// <summary>
        /// 数据整体向左/右移动offval个
        /// </summary>
        /// <param name="ori"></param>
        /// <param name="res"></param>
        /// <param name="direction">左移true，右移false</param>
        /// <param name="offval"></param>
        private static void MoveOff(double []ori, long dlen, ref double []res, bool direction, int offval)
        {
            long tlen = dlen - offval;
            //数据左移
            if (direction)
            {
                Array.Copy(ori, offval, res, 0, tlen);
                Array.Copy(ori, tlen, res, tlen, offval);
            }
            //数据右移
            else
            {
                Array.Copy(ori, 0, res, offval, tlen);
                Array.Copy(ori, 0, res, 0, offval);
            }
        }

        //FIR滤波器
        //数字滤波器，计算卷积
        private static void FilterConv(double[] signal, long slen,double[] filter, ref double[] foutdata)
        {
            long flen = filter.Length, minlen = flen > slen? slen : flen;
            long i, j; 
            double tempout = 0;

            for (i = 0; i < minlen; i++)
            {
                tempout = 0;
                for (j = 0; i - j > -1; j++)
                {
                    tempout = tempout + filter[j] * signal[i - j];
                }
                foutdata[i] = tempout;
            }

            for (i = flen; i < slen; i++)
            {
                tempout = 0;
                for (j = 0; j < flen; j++)
                {
                    tempout = tempout + filter[j] * signal[i - j];
                }
                foutdata[i] = tempout;
            }
        }

        //读取原始数据
        private static void LoadADData(string fname, ref double[] accdata, 
            ref double []lasdata, ref ushort []dmidata, ref long dlen, int side)
        {
            dlen = 0;
            int lenstrs = "%AD,094547410,45351,25251,52388,42912,34".Length;
            FileStream fr = File.OpenRead(fname);
            StreamReader sr = new StreamReader(fr);
            string strline;
            while ((strline = sr.ReadLine()) != null)
            {
                if (strline.Length == lenstrs)
                {
                    try
                    {
                        string[] s = strline.Split(',');
                        accdata[dlen] = double.Parse(s[2]) * MyIRIMTD._AccCali[side][0] + MyIRIMTD._AccCali[side][1] - _LocalGravity;
                        lasdata[dlen] = double.Parse(s[3]) * MyIRIMTD._LasCali[side][0] + MyIRIMTD._LasCali[side][1];
                        dmidata[dlen] = ushort.Parse(s[4]);
                        ++dlen;
                    }
                    catch (System.Exception ex)
                    {
                    	
                    }
                }
            }
            sr.Close();
            fr.Close();

            if (dlen == 0)
            {
                accdata = null;
                lasdata = null;
                dmidata = null;
            }
        }
        
        //检测校验
        private static bool Check(string instr)
        {
            if (instr.Substring(0, 4) == "%AD,")
            {
                int strlen = instr.Length - 2;
                int tchar = instr[0];

                for (int i = 1; i < strlen; i++)
                {
                    tchar = tchar ^ instr[i];
                }
                string res = tchar.ToString("X8");
                if (res == instr.Substring(strlen, 2))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

    }
}
