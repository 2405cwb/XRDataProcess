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
    public class MyIRIMTD
    {
        public static double[][] _LasCali;
        public static double[][] _AccCali;
        public static double[][] _MTDCali;

        private static double m_speed_kK;
        private static double m_speed_bK;
        private static double m_speed_kB;
        private static double m_speed_bB;
        private static bool m_speedtype;
        private static int m_Frequency;

        /// <summary>
        /// 加载工程的配置文件
        /// </summary>
        /// <param name="prjdir"></param>
        public static void LoadParm(string prj)
        {
            _LasCali = new double[2][];
            _AccCali = new double[2][];
            for (int i = 0; i < 2; ++i)
            {
                _LasCali[i] = new double[2];
                _AccCali[i] = new double[2];
                if (Directory.Exists(string.Format(@"{0}\IRIMTD\DAQ{1}", prj, i)))
                {
                    string setininame = string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i);
                    if (!File.Exists(setininame))
                    {
                        MessageBox.Show(string.Format("丢失配置文件：\r\n{0}\r\n请从其他工程相同位置拷贝【Setting.ini】至此目录", setininame));
                        System.Environment.Exit(0);
                    }
                    string iris = File.ReadAllText(setininame);
                    if (!iris.Contains(Environment.NewLine))
                    {
                        iris = iris.Replace("\n", Environment.NewLine);
                        File.WriteAllText(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i), iris);
                    }

                    IniFiles iriparm = new IniFiles(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i));
                    _LasCali[i][0] = Convert.ToDouble(iriparm.ReadString("CaliLaser", "p1", "1"));
                    _LasCali[i][1] = Convert.ToDouble(iriparm.ReadString("CaliLaser", "p2", "0"));
                    _AccCali[i][0] = Convert.ToDouble(iriparm.ReadString("CaliAcc", "p1", "1"));
                    _AccCali[i][1] = Convert.ToDouble(iriparm.ReadString("CaliAcc", "p2", "0"));

                    m_speed_kK = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "kofK", "1"));
                    m_speed_bK = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "bofK", "0"));
                    m_speed_kB = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "kofB", "1"));
                    m_speed_bB = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "bofB", "0"));

                    m_speedtype = iriparm.ReadBool("IRISpeedCali", "SpeedType", false);
                    m_Frequency = iriparm.ReadInteger("SampleFrequency", "Frequency", 2000);
                }
            }
            _MTDCali = new double[2][];
            for (int i = 0; i < 2; i++)
            {
                _MTDCali[i] = new double[2];
                string fname = string.Format(@"{0}\IRIMTD\Laser{1}", prj, i);
                if (Directory.Exists(fname))
                {
                    fname += "\\Setting.ini";
                    string mtds = File.ReadAllText(fname);
                    if (!mtds.Contains(Environment.NewLine))
                    {
                        mtds = mtds.Replace("\n", Environment.NewLine);
                        File.WriteAllText(string.Format(@"{0}\IRIMTD\Laser{1}\Setting.ini", prj, i), mtds);
                    }
                    IniFiles mtdparm = new IniFiles(fname);
                    _MTDCali[i][0] = mtdparm.ReadInteger("Parm", "MTD_k", 1);
                    _MTDCali[i][1] = mtdparm.ReadInteger("Parm", "MTD_b", 0);
                }
            }
            InitMTDX();
        }

        public static void ComputeIRI(string prj, int side, WinProcessBar bar)
        {
            List<string> daqfile = new List<string>();
            GetAllFiles(prj, side, "DAQ", "*.daq", ref daqfile);
            bar.SetIRIVal(0.1);

            string resamplefname = string.Format("{0}\\IRIMTD\\DAQ{1}\\resample.txt", prj, side);
            if (!File.Exists(resamplefname))
            {
                MessageBox.Show("遗失平整度数据！");
                return;
            }

            //先计算10m的平整度，后面再累计计算100m、1000m的平整度
            const int dislen = 10;
            GenerateIRI(resamplefname, dislen, "resample.txt");
            AdjustVal(resamplefname.Replace("resample.txt", string.Format("IRI_{0}m.txt", dislen)), MainForm._ErrorIRI, 0.001);
            bar.SetIRIVal(0.8);
            bar.SetIRIVal(0.9);
            bar.SetIRIVal(1);
        }

        /// <summary>
        /// 调整异常值
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="Thrval"></param>
        /// <param name="scale"></param>
        private static void AdjustVal(string fname, double Thrval, double scale)
        {
            if (MainForm._ErrorVal != 1)
                return;

            string[] oristrs = File.ReadAllLines(fname);
            int len = oristrs.Length;
            string[] newstrs = new string[len];
            double[] orival = new double[len];
            for (int i = 0; i < len; ++i)
            {
                orival[i] = double.Parse(oristrs[i].Substring(oristrs[i].LastIndexOf(' ') + 1));
            }
            if (len < 2) return;
            double lastval = orival[0];
            for (int i = 1; i < len; ++i)
            {
                orival[i] = orival[i] > Thrval ? lastval + MainForm.rdval.Next(100) * scale : orival[i];
                lastval = orival[0];
            }

            for (int i = 0; i < len; ++i)
            {
                newstrs[i] = string.Format("{0} {1}", oristrs[i].Split(' ')[0], orival[i]);
            }
            FileStream fw = new FileStream(fname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < len; ++i)
            {
                sw.WriteLine(newstrs[i], Encoding.UTF8);
            }
            sw.Close();
            fw.Close();
        }

        private static int[] m_laserX;
        private const int SizeLaserData = 24;
        private const int MTDPOINT = 151;
        private static int m_PointNN;

        /// <summary>
        /// 计算构造深度---计算的是10m的构造深度，其他的由这个的平均值计算
        /// </summary>
        /// <param name="prj"></param>
        /// <param name="side"></param>
        public static void ComputeMTD(string prj, int side, int featurelen, WinProcessBar bar)
        {
            string fname = string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_{2}m.txt", prj, side, featurelen);
            FileInfo tfile = new FileInfo(fname);
            if (File.Exists(fname) && tfile.Length > 0)
            {
                bar.SetMTDVal(1.0);
                return;
            }

            List<string> lasfile = new List<string>();
            GetAllFiles(prj, side, "Laser", "*.las", ref lasfile);
            bar.SetMTDVal(0.1);

            bool m_IsMTDFrame0 = true;
            double tempvaly = 0, tempval = 0, ttval = 0, oldttval = 0;
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

            int m_laspnum = 300 / (MTDPOINT - 1);
            int m_SMTDdnum = (int)(m_featureDis * 1000 / 300);//总共的SMTD个数
            int m_SMTDSubnum = (m_featureDis * 1000 - m_SMTDdnum * 300) / m_laspnum;//不足一个SMTD的点数

            double m_CurMTD;
            int m_MTDCnt = 0;
            int filecnt = 0;

            ulong framecnt = 0;
            FileStream fmtd = new FileStream(fname, FileMode.Create);
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

                                smtd.WriteLine(string.Format("{0} {1}", m_MTDCnt, m_CurMTD), Encoding.UTF8);

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

            AdjustVal(fname, MainForm._ErrorMTD, 0.0005);
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

        private static void GetAllFiles(string prj, int side, string subfolder, string ftype, ref List<string> filepath)
        {
            string prjpath = string.Format("{0}\\IRIMTD\\{1}{2}", prj, subfolder, side);
            DirectoryInfo dir = new DirectoryInfo(prjpath);
            FileInfo[] files = dir.GetFiles(ftype);
            foreach (FileInfo f in files)
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
        private static void GenerateIRI(string fpath, int vallen, string fname)
        {
            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);
            int len = sdata.Length;

            //250mm抽样
            int qplusenum = Convert.ToInt32(250 / 50);//250mm内有多少个编码器脉冲
            string savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));

            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();

            //计算IRI
            int plusenum = Convert.ToInt32(vallen * 4);//IRI距离内有多少个250mm
            int iricnt = 0;
            double irisum = 0;
            double irival = 0;
            double speedval = 0;
            double stime;
            double etime;

            sdata = File.ReadAllLines(savefname);
            len = sdata.Length;
            double[] oridata = new double[len];
            int[] oritime = new int[len];
            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                oridata[i] = double.Parse(s[0]) - double.Parse(s[1]);
                oritime[i] = int.Parse(s[3]);
            }

            savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            if (m_speedtype)
            {
                stime = oritime[0] * 1.0 / m_Frequency;
            }
            else
            {
                stime = (oritime[0] / 10000000) * 3600
                    + (oritime[0] / 100000) % 100 * 60
                    + (oritime[0] / 1000) % 100
                    + (oritime[0] % 1000) * 0.001;
            }
            for (int i = 1; i < len; ++i)
            {
                if (i % plusenum == 0)
                {
                    if (m_speedtype)
                    {
                        etime = oritime[i] * 1.0 / m_Frequency;
                    }
                    else
                    {
                        etime = (oritime[i] / 10000000) * 3600
                            + (oritime[i] / 100000) % 100 * 60
                            + (oritime[i] / 1000) % 100
                            + (oritime[i] % 1000) * 0.001;
                    }
                    speedval = vallen / (etime - stime) * 3.6;
                    irival = irisum / vallen;
                    irival = irival * (m_speed_kK * speedval + m_speed_bK) + (m_speed_kB * speedval + m_speed_bB);
                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);
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

        //计算路面跳车
        public static void ComputePB(string prj, int side, int vallen)
        {
            string fname = string.Format(@"{0}\IRIMTD\DAQ{1}\Resample.txt", prj, side);
            if (!File.Exists(fname))
            {
                MessageBox.Show("缺少路面纵断面文件：" + fname);
                return;
            }
            string[] datastrs = File.ReadAllLines(fname);

            //原始断面长度 50mm
            const int pluselen = 50;
            //计算跳车断面长度 100mm
            const int baselen = 100;

            int skipnum = baselen / pluselen;
            int valcnt = vallen * 1000 / pluselen;

            int len = datastrs.Length;

            string[] tstrs;
            double[] hval = new double[(len+skipnum-1) / skipnum];
            for (int i = 0; i < len; ++i)
            {
                if (i % skipnum == 0)
                {
                    tstrs = datastrs[i].Split('\t');
                    hval[i / 2] = double.Parse(tstrs[2]);
                }
            }

            len = hval.Length;
            double max = -1000000, min = 1000000;
            string[] rvjval = new string[(len + valcnt - 1) / valcnt];
            for (int i = 0; i < len; ++i)
            {
                max = Math.Max(hval[i], max);
                min = Math.Min(hval[i], min);
                if ((i + 1) % valcnt == 0)
                {
                    rvjval[i / valcnt] = (max - min).ToString();
                    max = -1000000;
                    min = 1000000;
                }
            }
            if (len % valcnt != 0)
            {
                rvjval[rvjval.Length - 1] = (max - min).ToString();
            }

            fname = string.Format(@"{0}\IRIMTD\DAQ{1}\PavementBump.txt", prj, side);
            File.WriteAllLines(fname, rvjval, Encoding.UTF8);
        }
    }
}

