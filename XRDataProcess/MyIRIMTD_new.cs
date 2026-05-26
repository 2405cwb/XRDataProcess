//#define 平整度测试
using DevExpress.Utils.About;
using DevExpress.Utils.CodedUISupport;
using DevExpress.XtraBars.Docking2010.Views.Widget;
using Framework.Log;
using Framework.Other;
using NPOI.SS.Formula.Functions;
using OperateIniFile;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using XRDataProcess.Properties;

namespace XRDataProcess
{
    public class MyIRIMTD
    {
        static XRSetting _Setting = XRSetting.GetInstance();
        private static MyLogger log = new MyLogger(typeof(MyIRIMTD));

        public static double[] SZU = { 0.9966071, 1.091514e-02, -2.083274e-03, 3.190145e-04,
            -0.5563044, 0.9438768, -0.8324718, 5.064701e-02,
            2.153176e-02, 2.126763e-03, 0.7508714, 8.221888e-03,
            3.335013, 0.3376467, -39.12762, 0.4347664};
        public static double[] PZU = { 5.47610e-03, 1.388776, 0.2275968, 35.79262 };



        private static double[] SZU100 =
{
      0.9994014, 0.004442351, 0.0002188854, 5.72179E-05,
      -0.2570548, 0.975036, 0.007966216, 0.02458427,
      0.003960378, 0.0003814527, 0.9548048, 0.004055587,
      1.687312, 0.1638951, -19.34264, 0.7948701
};
        private static double[] PZU100 = new double[4] { 0.0003793992, 0.2490886, 0.04123478, 17.65532 };



        public static double[][] _LasCali;
        public static double[][] _AccCali;
        public static double[][] _MTDCali;
        public static double[][] _MPDCali;
        public static int[] _mmPerPoint;

        /// <summary>
        /// 0-生成路面平整度，国际平整度指数四分之一车轮算法，tIRI_距离.txt
        /// 1-生成路面平整度，相邻算法-高程差，tIRI_距离.txt
        /// 2-生成路面平整度，国际平整度指数四分之一车轮算法，tIRI_距离.txt，用内业软件配置的k、b参数
        /// </summary>
        public static double[] m_IRI_k = new double[] { 0.8969, 0.8969 };
        public static double[] m_IRI_b = new double[] { 0.359, 0.359 };
        private static bool[] m_speedtype;
        private static int[] m_Frequency;

        /// <summary>
        /// 加载工程的配置文件
        /// </summary>
        /// <param name="prjdir"></param>
        public static void LoadParm(string prj)
        {
            
            m_speedtype = new bool[2];
            m_Frequency = new int[2];
            for (int i = 0; i < 2; i++)
            {
                if (Directory.Exists(string.Format(@"{0}\IRIMTD\DAQ{1}", prj, i)))
                {
                    string setininame = string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i);
                    if (!File.Exists(setininame))
                    {
                        MessageBox.Show(string.Format("丢失配置文件：\r\n{0}\r\n请从其他工程相同位置拷贝【Setting.ini】至此目录", setininame));

                    }
                    string iris = File.ReadAllText(setininame);
                    if (!iris.Contains(Environment.NewLine))
                    {
                        iris = iris.Replace("\n", Environment.NewLine);
                        File.WriteAllText(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i), iris);
                    }

                    IniFiles iriparm = new IniFiles(string.Format(@"{0}\IRIMTD\DAQ{1}\Setting.ini", prj, i));
                    m_IRI_k[i] = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "IRIk", "0.8969"));
                    m_IRI_b[i] = Convert.ToDouble(iriparm.ReadString("IRISpeedCali", "IRIb", "0.359"));
                    m_speedtype[i] = iriparm.ReadBool("IRISpeedCali", "SpeedType", false);
                    m_Frequency[i] = iriparm.ReadInteger("SampleFrequency", "Frequency", 2000);
                }

            } 

            _MTDCali = new double[3][];
            _MPDCali = new double[3][];
            _mmPerPoint = new int[3];
            for (int i = 0; i < 3; i++)
            {
                {
                    _MTDCali[i] = new double[2];
                    _MPDCali[i] = new double[2];
                    string fname = string.Format(@"{0}\IRIMTD\Laser{1}", prj, i);
                    if (Directory.Exists(fname))
                    {
                        string mtdfname = fname + "\\MTD_10m.txt";

                        fname += "\\Setting.ini";
                        if (!File.Exists(fname))
                        {
                            throw new Exception($"缺少{fname}文件，请检查数据!");
                        }
                        string mtds = File.ReadAllText(fname);
                        if (!mtds.Contains(Environment.NewLine))
                        {
                            mtds = mtds.Replace("\n", Environment.NewLine);
                            File.WriteAllText(string.Format(@"{0}\IRIMTD\Laser{1}\Setting.ini", prj, i), mtds);
                        }
                        IniFiles mtdparm = new IniFiles(fname);
                        _MTDCali[i][0] = Convert.ToDouble(mtdparm.ReadString("Parm", "MTD_k", "1"));
                        _MTDCali[i][1] = Convert.ToDouble(mtdparm.ReadString("Parm", "MTD_b", "0"));
                        _MPDCali[i][0] = Convert.ToDouble(mtdparm.ReadString("Parm", "MTD_k", "1"));
                        _MPDCali[i][1] = Convert.ToDouble(mtdparm.ReadString("Parm", "MTD_b", "0"));
                        _mmPerPoint[i] = mtdparm.ReadInteger("Parm", "PMode", 2);

                        // 如果存在这个文件，有两种情况：1、旧软件计算的，MPD不加系数，2、新软件计算的，MPD要加系数
                        // 如果不存在这个文件，那么就之间用这个新软件计算，MTD和MPD都会加系数
                        if (File.Exists(mtdfname))
                        {
                            int tmp = mtdparm.ReadInteger("Parm", "Ver", -1);
                            if (tmp == -1)//不存在这个配置，说明用旧软件计算的构造深度，MPD不加系数，如果存在这个配置，说明是用新软件计算的构造深度，MPD加系数
                            {
                                try
                                {
                                    _MPDCali[i][0] = mtdparm.ReadInteger("Parm", "MTD_k", 1);
                                }
                                catch (System.Exception)
                                {
                                    _MPDCali[i][0] = 1;
                                }
                                try
                                {
                                    _MPDCali[i][1] = mtdparm.ReadInteger("Parm", "MTD_b", 0);
                                }
                                catch (System.Exception)
                                {
                                    _MPDCali[i][1] = 0;
                                }
                            }
                        }
                    }
                }
            }

        }

        public static void ComputeIRI(string prj, int side, WinProcessBar bar, ProjectInfo _ProjectInfo = null)
        {
            if (_Setting.IsCheckIRIGPSTime)
            {
                CheckGPSTime(prj, side);
            }

            //先计算10m的平整度，后面再累计计算100m、1000m的平整度
            const int dislen = 10;
            string fname = string.Format("{0}\\IRIMTD\\DAQ{1}\\IRI_{2}m.txt", prj, side, dislen);
            FileInfo tfile = new FileInfo(fname);
            if (File.Exists(fname) && tfile.Length > 0)
            {
                bar.SetIRIVal(1.0);
                return;
            }

            List<string> daqfile = new List<string>();
            GetAllFiles(prj, side, "DAQ", "*.daq", ref daqfile);
            bar.SetIRIVal(0.1);

            string resamplefname = string.Format("{0}\\IRIMTD\\DAQ{1}\\resample.txt", prj, side);
            int oldAdcc_iri = 0; // 默认计算方式
            if (_ProjectInfo._PlusLength != 0)
            {
                //老设备
                _Setting.Acc_IRI = 3;

            }
            else
            {
                _Setting.Acc_IRI = oldAdcc_iri;
            }
            _Setting.WriteData();
            if (!File.Exists(resamplefname) && _Setting.Acc_IRI != 3)
            {

                if (!_ProjectInfo._IsJgAndGd)
                    MessageBox.Show("遗失平整度数据！");
                return;
            }

            bool datasrc = JudgMTDval(prj, side);

            if (_Setting.Acc_IRI == 0)
            {
                 GenerateIRI_NEW(resamplefname, dislen, "resample.txt", datasrc, side, _ProjectInfo); 

               // WorkBankIRIAlgo_withSpeed(resamplefname, "resample.txt", dislen * 4,side); 
                //GenerateIRI(resamplefname, 0.1f, "resample.txt", datasrc, side);
            }
            else if (_Setting.Acc_IRI == 1)
            {
                GenerateIRI_1(resamplefname, dislen, "resample.txt", datasrc,side);
            }
            else if (_Setting.Acc_IRI == 2)
            {
                GenerateIRI_2(resamplefname, dislen, "resample.txt", datasrc,side);
            }
            else if (_Setting.Acc_IRI == 3)
            {
                GenerateIRI_3(resamplefname, 10, "resample.txt", datasrc, _ProjectInfo, bar,side);
            }
            AdjustVal(resamplefname.Replace("resample.txt", string.Format("IRI_{0}m.txt", dislen)), _Setting.ErrorIRI, 0.001);
            bar.SetIRIVal(0.8);
            bar.SetIRIVal(0.9);
            bar.SetIRIVal(1);
        }

        // /LaserX/MTD_100.txt第二列平均值小于0.1，返回true，否则返回false
        private static bool JudgMTDval(string prj, int side)
        {
            string resamplefname = string.Format("{0}\\IRIMTD\\Laser{1}\\MTD_100.txt", prj, side);
            if (!File.Exists(resamplefname))
            {
                return false;
            }
            string[] data = new string[2];
            string[] s;
            data = File.ReadAllLines(resamplefname);
            if (data.Length > 2)
            {
                double tt = 0;
                int num = 0;
                for (int i = 0; i < data.Length; ++i)
                {
                    try
                    {
                        s = data[i].Split('\t');
                        if (s.Length > 1)
                        {
                            tt += Convert.ToDouble(s[1]);
                            num++;
                        }

                    }
                    catch
                    { }
                }
                tt = tt / num;
                if (tt < _Setting.IRI_threshval)
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

        /// <summary>
        /// 调整异常值
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="Thrval"></param>
        /// <param name="scale"></param>
        private static void AdjustVal(string fname, double Thrval, double scale)
        {
            if (_Setting.ErrorVal != 1)
                return;

            if (!File.Exists(fname))
            {
                return;
            }

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
            double sumval = 0;
            // orival[0] = MainForm.rdval.NextDouble() * 10;
            for (int i = 1; i < len; ++i)
            {

                //
                //if (orival[i]>12)
                //{
                //    orival[i] = MainForm.rdval.NextDouble() * 10;
                //}
                // 如果当前值大于异常阈值，则调整当前值为前面所有值的平均值
                orival[i] = orival[i] > Thrval ? lastval + (MainForm.rdval.Next(100) - 50) * scale : orival[i];
                sumval += orival[i];
                lastval = sumval / i;
            }

            for (int i = 0; i < len; ++i)
            {

                orival[i] = _Setting.iriKCorrect * orival[i] + _Setting.iriBCorrect;

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

        /// <summary>
        /// 计算构造深度MPD---计算的是10m的构造深度，其他的由这个的平均值计算
        /// </summary>
        /// <param name="prj"></param>
        /// <param name="side"></param>
        public static void ComputeMPD(string prj, int side, int featurelen, double threshval, WinProcessBar bar)
        {
            string fname = string.Format("{0}\\IRIMTD\\Laser{1}\\MPD_{2}m.txt", prj, side, featurelen);
            FileInfo tfile = new FileInfo(fname);
            if (File.Exists(fname) && tfile.Length > 0)
            {
                bar.SetMPDVal(1.0);
                return;
            }

            List<string> lasfile = new List<string>();
            GetAllFiles(prj, side, "Laser", "*.las", ref lasfile);
            bar.SetMPDVal(0.1);

            const int SizeLaserData = 24;
            double ttval = 0;
            double oldttval = 200;

            int m_featureDis = featurelen;
            double[] ratio;

            ulong MPDPointNum = (ulong)(100 / _mmPerPoint[side]);
            ulong MPDPointNumHalf = MPDPointNum / 2;
            double YMax1 = 0.0;
            double YMax2 = 0.0;
            double YMean = 0.0;
            double curMPDB = 0.0;
            double curMPD = 0.0;
            int MPDBCnt = 0;
            int MPDCnt = 0;
            int MPDBNUM = featurelen * 1000 / 100;

            int ptcnt = 0;
            double filtersum = 0.0;
            double filtermean = 0.0;

            double[] laserX = new double[MPDPointNum];
            for (ulong i = 0; i < MPDPointNum; i++)
            {
                laserX[i] = i;
            }
            double[] laserY = new double[MPDPointNum];

            int filecnt = 0;
            ulong framecnt = 0;
            List<string> lasstrlist = null;

           // string tmpstr = null;
            //if (_Setting.IsOutputLasval)
            //{
            //    lasstrlist = new List<string>();
            //}

            FileStream fmpd = new FileStream(fname, FileMode.Create);
            StreamWriter smpd = new StreamWriter(fmpd);
            List<double> value1M = new List<double>(); //始终记录前25个点  用来获取中间值
           


            foreach (string lf in lasfile)
            {
                FileInfo lasf = new FileInfo(lf);
                long filesize = lasf.Length;
                long framenum = filesize / SizeLaserData;

                FileStream fr = File.OpenRead(lf);
                BinaryReader br = new BinaryReader(fr);

                double tMile = 0;
                for (int i = 0; i < framenum; ++i, ++framecnt)
                {
                    fr.Position += 16;
                    ttval = br.ReadDouble();
                    if (ttval < 201) ttval = 200;
                    else if (ttval > 399) ttval = 400;
                    else
                    {
                        //取25个点，是因为平整度的纵断面原始是50mm的采样间距，构造的激光数字信号是2mm的采样间距，所以取25个点，
                        //先把平整度的纵断面原始激光数据做个去噪，然后得到平整度纵断面数据里面差距最大值作为构造的激光数字信号去噪的差值阈值
                        if (ptcnt < 25)
                        {
                            filtersum += ttval;

                            ++ptcnt;
                        }
                        if (value1M.Count < 500) //记录500-》1m数据 
                        {
                            value1M.Add(ttval);
                        }
                    }
                    if (value1M.Count >= 500)
                    {
                        value1M.RemoveAt(0);
                        value1M.Add(ttval);
                    }
                    //if (_Setting.IsOutputLasval)
                    //{
                    //    tmpstr = ttval.ToString();
                    //}
                    if (ptcnt >= 25)
                    {
                        tMile += 0.002;
                        filtermean = filtersum / ptcnt;
                        if (Math.Abs(filtermean - ttval) > threshval)
                        {
                            ttval = oldttval;
                        }

                        filtersum = filtersum + ttval - filtermean;
                    }

                    //if (_Setting.IsOutputLasval)
                    //{
                    //    tmpstr = tmpstr + "\t" + ttval.ToString();
                    //    lasstrlist.Add(tmpstr);
                    //}

                    List<double> tempSort = new List<double>(value1M);
                    tempSort.Sort();
                    if (value1M.Count > 0)
                    {

                        oldttval = tempSort[tempSort.Count / 2];
                    } 
                    if (framecnt < MPDPointNum)
                    {
                        laserY[framecnt] = ttval;
                    }
                    else
                    {
                        YMax1 = -10000;
                        YMax2 = -10000;
                        YMean = 0;
                        ratio = FittingFunct.TowTimesCurve(laserY, laserX);
                        for (ulong j = 0; j < MPDPointNum; j++)
                        {
                            laserY[j] = laserY[j] - (ratio[0] + ratio[1] * laserX[j] + ratio[2] * laserX[j] * laserX[j]);
                        }
                        for (ulong j = 0; j < MPDPointNumHalf; j++)
                        {
                            YMean += laserY[j];
                            if (laserY[j] > YMax1)
                                YMax1 = laserY[j];
                        }
                        for (ulong j = MPDPointNumHalf; j < MPDPointNum; j++)
                        {
                            YMean += laserY[j];
                            if (laserY[j] > YMax2)
                                YMax2 = laserY[j];
                        }
                        YMean = YMean / MPDPointNum;
                        curMPDB = (YMax1 + YMax2) / 2 - YMean;

                        curMPD = curMPD + curMPDB;
                        ++MPDBCnt;

                        if (MPDBCnt == MPDBNUM)
                        {
                            bar.SetMPDVal(0.1 + 0.9 / lasfile.Count * (filecnt + (double)i / framenum));

                            ++MPDCnt;
                            curMPD = curMPD / MPDBCnt;
                            curMPD = curMPD * _Setting.MPD_K + _Setting.MPD_B;
                            curMPD = curMPD * _MPDCali[side][0] + _MPDCali[side][1];

                            smpd.WriteLine(string.Format("{0} {1}", MPDCnt, curMPD), Encoding.UTF8);

                            MPDBCnt = 0;
                            curMPD = 0;
                        }

                        framecnt = 0;
                        laserY[framecnt] = ttval;
                    }

                }

                br.Close();
                fr.Close();
                filecnt++;
            }

            smpd.Close();
            fmpd.Close();


            AdjustVal(fname, _Setting.ErrorMTD, 0.0005);
            bar.SetMPDVal(1.0);
            //if (_Setting.IsOutputLasval)
            //{
            //    File.WriteAllLines(string.Format("{0}\\IRIMTD\\Laser{1}\\lasval.txt", prj, side), lasstrlist.ToArray());
            //}
        }

        /// <summary>
        /// 计算构造深度---计算的是10m的构造深度，其他的由这个的平均值计算
        /// </summary>
        /// <param name="prj"></param>
        /// <param name="side"></param>
        /// <param name="featurelen"></param>
        /// <param name="threshval">相邻激光点差值的异常判断阈值</param>
        public static void ComputeMTD(string prj, int side, int featurelen, double threshval, WinProcessBar bar)
        {
            if (side == 1)
            {
            }
            //  _Setting.IsOutputLasval = true;
            const int SizeLaserData = 24;
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

            int MTDPOINT = 300 / _mmPerPoint[side] + 1;
            int[] laserX = new int[MTDPOINT];
            for (int i = 0; i < MTDPOINT; i++)
            {
                laserX[i] = i - MTDPOINT / 2;
            }
            int m_PointNN = MTDPOINT * MTDPOINT;

            bool m_IsMTDFrame0 = true;
            double tempvaly = 0, tempval = 0, ttval = 0, oldttval = 200;
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

            int ptcnt = 0;
            double filtersum = 0.0;
            double filtermean = 0.0;

            string tmpstr = null;
            // 修改：移除 lasstrlist List；如果启用输出，直接用 StreamWriter 写入文件
            StreamWriter slasval = null;
            string lasvalPath = null;
            if (_Setting.outMoHaoData)
            {
                lasvalPath = string.Format("{0}\\IRIMTD\\Laser{1}\\lasval.txt", prj, side);
                // 如果文件存在，先删除（或用 FileMode.Append 如果想追加）
                if (File.Exists(lasvalPath)) File.Delete(lasvalPath);
                FileStream flasval = new FileStream(lasvalPath, FileMode.Create, FileAccess.Write);
                slasval = new StreamWriter(flasval, Encoding.UTF8);
            }

            ulong framecnt = 0;
            FileStream fmtd = new FileStream(fname, FileMode.Create);
            StreamWriter smtd = new StreamWriter(fmtd);

            List<double> value1M = new List<double>(); //始终记录前500个点  用来获取中间值
            foreach (string lf in lasfile)
            {
                FileInfo lasf = new FileInfo(lf);
                long filesize = lasf.Length;
                long framenum = filesize / SizeLaserData;

                FileStream fr = File.OpenRead(lf);
                BinaryReader br = new BinaryReader(fr);

                //List<double> jzlb5Value = new List<double>();
                List<string> temps = new List<string>();
                int ceshiInt = 0;
                //bool isJunZhi = false;
                for (int i = 0; i < framenum; ++i, ++framecnt)
                {

                    fr.Position += 16;
                    ttval = br.ReadDouble();
                    ceshiInt++;
                    if (ttval < 201) ttval = 200;
                    else if (ttval > 399) ttval = 400;
                    else
                    {

                        if (ptcnt < 25)
                        {
                            filtersum += ttval;

                            ++ptcnt;
                        }
                        if (value1M.Count < 500) //记录500-》1m数据 
                        {
                            value1M.Add(ttval);
                        }
                    }
                    if (value1M.Count >= 500)
                    {
                        value1M.RemoveAt(0);
                        value1M.Add(ttval);
                    }
                    // 修改：原始 ttval 用于 tmpstr 的第一部分
                    if (_Setting.outMoHaoData)
                    {
                        tmpstr = ttval.ToString();  // 原始值
                    }
                    //取25个点，是因为平整度的纵断面原始是50mm的采样间距，构造的激光数字信号是2mm的采样间距，所以取25个点，
                    //先把平整度的纵断面原始激光数据做个去噪，然后得到平整度纵断面数据里面差距最大值作为构造的激光数字信号去噪的差值阈值
                    if (ptcnt >= 25)
                    {

                        filtermean = filtersum / ptcnt;
                        if (Math.Abs(filtermean - ttval) > threshval)
                        {
                            ttval = oldttval;
                        }
                        filtersum = filtersum + ttval - filtermean;
                    }

                    List<double> tempSort = new List<double>(value1M);
                    tempSort.Sort();
                    if (value1M.Count > 0)
                    {

                        oldttval = tempSort[tempSort.Count / 2];
                    }

                    // 修改：去噪后，构建完整 tmpstr 并直接写入文件（如果启用）
                    if (_Setting.outMoHaoData)
                    {
                        tmpstr += "\t" + ttval.ToString();  // 原始 + 处理后
                        slasval.WriteLine(tmpstr);
                        slasval.Flush();  // 每行后刷新，确保数据不丢失（可选，但大文件时安全）
                    }

                    if (m_IsMTDFrame0)
                    {
                        if (++m_SMTDSubcnt >= m_SMTDSubnum)
                        {
                            m_IsMTDFrame0 = false;
                        }
                    }
                    else
                    {
                        tempvaly = ttval;
                        tempvalx = laserX[m_laserYIdx];
                        m_laserYSum += tempvaly;
                        m_laserYYSum += tempvaly * tempvaly;
                        tempvaly *= tempvalx;
                        m_laserXYSum += tempvaly;
                        m_laserXXYSum += tempvaly * tempvalx;

                        if (++m_laserYIdx == MTDPOINT)
                        {
                            tempval = (m_PointNN - 1) * m_laserYSum - 12 * m_laserXXYSum;
                            tempval = 5 * tempval * tempval / (4 * (m_PointNN - 4));
                            tempval = (12 * m_laserXYSum * m_laserXYSum + tempval) / (m_PointNN - 1);
                            tempval = (MTDPOINT * m_laserYYSum - m_laserYSum * m_laserYSum - tempval) / m_PointNN;

                            tempval = tempval > 0 ? tempval : 0;
                            m_SMTDdSum += Math.Sqrt(tempval);
                            ++m_SMTDValCnt;

                            if (++m_SMTDdCnt == m_SMTDdnum)
                            {
                                bar.SetMTDVal(0.1 + 0.9 / lasfile.Count * (filecnt + (double)i / framenum));

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
                            tempvalx = laserX[0];
                            m_laserYSum = tempvaly;
                            m_laserYYSum = tempvaly * tempvaly;
                            tempvaly *= tempvalx;
                            m_laserXYSum = tempvaly;
                            m_laserXXYSum = tempvaly * tempvalx;
                            m_laserYIdx = 1;
                        }
                    }
                }
                br.Close();
                fr.Close();
                filecnt++;
            }

            smtd.Close();
            fmtd.Close();

            // 修改：关闭 lasval 的 StreamWriter
            if (_Setting.outMoHaoData && slasval != null)
            {
                slasval.Close();
            }

            AdjustVal(fname, _Setting.ErrorMTD, 0.0005);

            string ininame = string.Format("{0}\\IRIMTD\\Laser{1}\\Setting.ini", prj, side);
            IniFiles mtdparm = new IniFiles(ininame);
            mtdparm.WriteInteger("Parm", "Ver", 1);
            bar.SetMTDVal(1.0);

            // 修改：移除 File.WriteAllLines 调用，因为已直接写入
        }

        private static void GetAllFiles(string prj, int side, string subfolder, string ftype, ref List<string> filepath)
        {
            string prjpath = string.Format("{0}\\IRIMTD\\{1}{2}", prj, subfolder, side);
            DirectoryInfo dir = new DirectoryInfo(prjpath);
            FileInfo[] files = dir.GetFiles(ftype);
            //20250331 .las格式导致 文件反序 修改了改bug
            //Array.Sort(files, delegate (FileInfo x, FileInfo y) { return y.CreationTime.CompareTo(x.CreationTime); });
            Array.Sort(files, delegate (FileInfo x, FileInfo y) { return x.CreationTime.CompareTo(y.CreationTime); });

            foreach (FileInfo f in files)
            {
                filepath.Add(f.FullName);
                var time = f.CreationTime;
            }
        }

        public static double GetLaserThresh(string fpath)
        {
            double res = 100;
            string[] sdata = null;
            if (File.Exists(fpath))
            {
                sdata = File.ReadAllLines(fpath);
            }
            else
            {
                return res;
            }

            int len = sdata.Length;
            //不足1米
            if (len < 20)
                return res;

            double[] lasval = new double[len];
            string[] s;
            for (int i = 0; i < len; ++i)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    try
                    {
                        lasval[i] = double.Parse(s[0]);
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            lasval[i] = double.Parse(s[0]);
                        }
                    }
                }
            }

            res = 0;
            double tmpval = 0;
            for (int i = 1; i < len; ++i)
            {
                tmpval = Math.Abs(lasval[i] - lasval[i - 1]);
                if (res < tmpval)
                {
                    res = tmpval;
                }
            }

            return res;
        }

        /// <summary>
        /// 剔除激光测距机里面的异常值，resample.txt文件
        /// </summary>
        /// <param name="fpath"></param>
        public static void FilterLaserData(string fpath, double thresh1 = 5, double thresh2 = 20)
        {
            string[] sdata = null;
            if (File.Exists(fpath + ".bak"))
            {
                sdata = File.ReadAllLines(fpath + ".bak");
            }
            else
            {
                sdata = File.ReadAllLines(fpath);
            }
            int len = sdata.Length;

            //不足1米
            if (len < 20)
                return;

            double[] lasval = new double[len];
            double[] disval = new double[len];
            double[] roadval = new double[len];
            double[] dmival = new double[len];
            double[] diffval = new double[len];

            double sum_lasval = 0;
            string[] s;
            for (int i = 0; i < len; ++i)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    try
                    {
                        lasval[i] = double.Parse(s[0]);
                        disval[i] = double.Parse(s[1]);
                        roadval[i] = double.Parse(s[2]);
                        dmival[i] = double.Parse(s[3]);
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            lasval[i] = double.Parse(s[0]);
                            disval[i] = double.Parse(s[1]);
                            roadval[i] = double.Parse(s[2]);
                            dmival[i] = double.Parse(s[3]);
                        }
                    }
                    sum_lasval += lasval[i];
                    if (i > 0)
                    {
                        diffval[i] = Math.Abs(lasval[i] - lasval[i - 1]);
                    }
                }
            }

            double meanlasval = sum_lasval / len;
            for (int i = 0; i < 3; ++i)
            {
                if (Math.Abs(lasval[i] - meanlasval) > 100)
                {
                    lasval[i] = meanlasval;
                }
            }
            //先根据相邻测距值的差大于指定阈值，将这种孤立的异常值剔除
            for (int i = 2; i < len; ++i)
            {
                if (diffval[i] >= thresh1 && diffval[i - 1] >= thresh1)
                {
                    lasval[i - 1] = lasval[i - 2] - disval[i - 2] + disval[i - 1];
                }
            }
            //然后根据相邻的动态测距差值大于指定阈值，找出第一步成片的异常值的边缘位置，将这种异常值剔除
            for (int i = 1; i < len; ++i)
            {
                if (Math.Abs(lasval[i] - lasval[i - 1]) >= thresh2)
                {
                    lasval[i] = lasval[i - 1] - disval[i - 1] + disval[i];
                }
            }

            for (int i = 0; i < len; ++i)
            {
                roadval[i] = lasval[i] - disval[i];
                sdata[i] = string.Format("{0}\t{1}\t{2}\t{3}", lasval[i], disval[i], roadval[i], dmival[i]);
            }

            if (!File.Exists(fpath + ".bak"))
            {
                File.Move(fpath, fpath + ".bak");
            }

            File.WriteAllLines(fpath, sdata, Encoding.UTF8);
        }

        private static double[,] ST100 = new double[4, 4]
{
    { 0.9994014, 0.004442351, 0.0002188854, 5.72179E-05 },
    { -0.2570548, 0.975036, 0.007966216, 0.02458427 },
    { 0.003960378, 0.0003814527, 0.9548048, 0.004055587 },
    { 1.687312, 0.1638951, -19.34264, 0.7948701 }
};
        private static double[] PR100 = new double[4] { 0.0003793992, 0.2490886, 0.04123478, 17.65532 };



        private static double[,] ST250 = new double[4, 4]
{
    { 0.9966071, 0.01091514, -0.002083274, 0.0003190145 },
    { -0.5563044, 0.9438768, -0.8324718, 0.05064701 },
    { 0.02153176, 0.002126763, 0.7508714, 0.008221888 },
    { 3.335013, 0.3376467, -39.12762, 0.4347564 }
};
        private static double[] PR250 = new double[4] { 0.005476107, 1.388776, 0.2275968, 35.79262 };






        public static void WorkBankIRIAlgo_withSpeed(string fpath, string fname, int Count,int side)
        {

            Count = (int)(Count / _Setting.IRIAlgorithmInterval);
            List<double> Points = new List<double>();
            List<double> SpeedPoints = new List<double>(); // 新增：存储车速数据
            double DeltLen = _Setting.IRIAlgorithmInterval;

            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);
            //Array.Reverse(sdata); // 原地反转数组

            int qplusenum = Convert.ToInt32(DeltLen / 0.05);//250mm内有多少个编码器脉冲
            int len = sdata.Length;

            double[] oridata = new double[len];
            int[] oritime = new int[len]; // 新增：存储时间数据用于计算车速


            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[2]);
                        oritime[i] = int.Parse(s[3]); // 获取时间数据
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];

                        }
                    }
                }
            }
            double[] Src = new double[(len + qplusenum - 1) / qplusenum];
            int[] SrcTime = new int[Src.Length]; // 新增：抽样后的时间数据

            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                Src[j] = oridata[i];
                SrcTime[j] = oritime[i]; // 保存抽样点的时间
            }

            float DX = (float)DeltLen;
            double[,] ST = ST250;
            double[] PR = PR250;

            if ((int)(DX * 1000f) == 100)
            {
                ST = ST100;
                PR = PR100;
            }

            double[] array = new double[27];
            double[] array2 = new double[5];
            double[] array3 = new double[5];
            int num = (int)(0.25f / DX + 0.5f) + 1;
            //int num = (int)(0.25 / (double)DX + 0.5) + 1;
            if (num < 2)
            {
                num = 2;
            }


            float num2 = (float)(num - 1) * DX;
            num--;
            if (Src.Length == 0 || Src.Length <= num)
            {
                return;
            }
            int num3 = (int)Math.Round(11.0 / (double)DX);
            if (num3 >= Src.Length) //  
            {
                num3 = Src.GetLength(0) - 1;
            }
            array[num] = Src[num3];
            array[0] = Src[0];
            array3[0] = (array[num] - array[0]) / 11.0;
            array3[1] = 0.0;
            array3[2] = array3[0];
            array3[3] = 0.0;
            double num4 = 0.0;
            int num5 = 1;
            int num6 = 0;
            int num7 = Count;
            double num8 = 0.0;

            // ===== 修复1：正确的时间解析 =====
            // 时间解析函数 - 根据实际数据格式选择
            double ParseTime(int timeValue)
            {
                // 方法1：脉冲计数转时间（频率方式）
                // return timeValue / frequency; 
                if (m_speedtype[side])
                {
                    return timeValue * 1.0 / m_Frequency[side];
                }
                else
                {
                    return (timeValue / 10000000.0) * 3600 +
                                           (timeValue / 100000 % 100) * 60 +
                                           (timeValue / 1000 % 100) +
                                           (timeValue % 1000) * 0.001;
                }
                // 方法2：HHMMSSmmm格式解析（您实际使用的格式）

            }
            // 初始化时间跟踪
            int lastIndex = 0; // 上一个区间结束点的索引
            double lastTime = ParseTime(SrcTime[0]);

            while (num5 < Src.Length)
            {
                do
                {
                    array[num] = Src[num5];
                    if (num5 < num)
                    {
                        array[num5] = array[num];
                    }
                    num5++;
                }

                while (num5 <= num);
                double num9 = (array[num] - array[0]) / (double)num2;
                for (int i = 1; i <= num; i++)
                {
                    array[i - 1] = array[i];
                }
                for (int i = 0; i <= 3; i++)
                {
                    array2[i] = PR[i] * num9;
                    for (int j = 0; j <= 3; j++)
                    {
                        array2[i] += ST[i, j] * array3[j];
                    }
                }
                for (int i = 0; i <= 3; i++)
                {
                    array3[i] = array2[i];
                }

                num4 += Math.Abs(array2[0] - array2[2]);
                num6++;
                num8 = num4 / (double)num6;
                if (num5 == num7)
                {

                    // 添加IRI值
                    Points.Add(num8);

                    // 计算当前段的车速
                    // 关键：使用区间起点和终点的时间差
                    double currentTime = ParseTime(SrcTime[num5 - 1]); // 当前点时间
                    double timeDiff = currentTime - lastTime;

                    // 实际距离 = 点数 * 0.1米
                    double actualDistance = (num5 - lastIndex) * DX;
                    double speed = (timeDiff > 0) ? (actualDistance / timeDiff * 3.6) : 0;

                    SpeedPoints.Add(speed);

                    // 更新跟踪变量
                    lastTime = currentTime;
                    lastIndex = num5;

                    // 重置计数器
                    num7 += Count;
                    num4 = 0.0;
                    num6 = 0;
                }
            }
            if (num5 < num7)
            {
                Points.Add(num8);
            }
            List<string> texts = new List<string>();
            for (int i = 0; i < Points.Count; i++)
            {
                string line = $"{i + 1} {Points[i]}";
                texts.Add(line);
            }
            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", 10));
            File.WriteAllLines(savefname, texts);
            // 新增：输出车速文件
            List<string> speedTexts = new List<string>();
            for (int i = 0; i < SpeedPoints.Count; i++)
            {
                speedTexts.Add($"{i + 1} {SpeedPoints[i]:F2}");
            }
            string speedSavePath = fpath.Replace(fname, $"Speed_10m.txt");
            File.WriteAllLines(speedSavePath, speedTexts);
        }


        /// <summary>
        /// 生成路面平整度，国际平整度指数四分之一车轮算法，tIRI_距离.txt
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="iridis"></param>
        /// <param name="pluseparm"></param>
        /// <param name="datasrc">true 仅加速度计，false 纵断面</param>
        private static void GenerateIRI(string fpath, int vallen, string fname, bool datasrc,int side)
        {
            bool IsParmFile = false;
            String fparmpath = fpath.Replace("resample.txt", "Coeff.dat");
            if (File.Exists(fparmpath))
            {
                IsParmFile = true;
            }

            string[] parms = null;
            double[] speedparms = null;
            double[] kparms = null;
            double[] bparms = null;
            int parmnum = 0;
            if (IsParmFile)
            {
                int idx = 0;
                parms = File.ReadAllLines(fparmpath);
                try
                {
                    parmnum = int.Parse(parms[idx]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("读取文件出错，请检查！\r\n" + fparmpath);
                    return;
                }
                speedparms = new double[parmnum];
                kparms = new double[parmnum];
                bparms = new double[parmnum];
                for (int i = 0; i < parmnum; ++i)
                {
                    speedparms[i] = double.Parse(parms[++idx]);
                }
                for (int i = 0; i < parmnum; ++i)
                {
                    kparms[i] = double.Parse(parms[++idx]);
                }
                for (int i = 0; i < parmnum; ++i)
                {
                    bparms[i] = double.Parse(parms[++idx]);
                }
            }

            double DeltLen = _Setting.IRIAlgorithmInterval;
            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);

            int qplusenum = Convert.ToInt32(DeltLen / 0.05);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            double[] iridata = new double[(len + qplusenum - 1) / qplusenum];
            int[] oritime = new int[len];
            int[] iritime = new int[(len + qplusenum - 1) / qplusenum];

            //计算IRI
            int plusenum = Convert.ToInt32(vallen / DeltLen);//IRI距离内有多少个250mm
            int iricnt = 0;
            double irisum = 0;
            double irival = 0;

            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[2]);
                        toridata[i] = oridata[i];
                        oritime[i] = int.Parse(s[3]);
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                            toridata[i] = oridata[i];
                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
                        }
                    }
                }
            }
            //取5个点做均值滤波
            for (int i = 2; i < len - 2; ++i)
            {
                oridata[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
            }
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            len = iridata.Length;

            double speedval = 0;
            double stime;
            double etime;
            double YSU = 0;
            double[] ZSU = new double[4];
            double[] oldZSU = new double[4];


            //oldZSU[0] = iridata[1] - iridata[0];
            //oldZSU[1] = iridata[1] - iridata[0];
            //oldZSU[2] = 0;
            //oldZSU[3] = 0;



            if (_Setting.IRIAlgorithmInterval == 0.1)
            {
                oldZSU[0] = (iridata[110] - iridata[0]) / 11;
                oldZSU[2] = (iridata[110] - iridata[0]) / 11;

                oldZSU[1] = 0;
                oldZSU[3] = 0;
            }
            else
            {
                oldZSU[0] = (iridata[44] - iridata[0]) / 11;
                oldZSU[2] = (iridata[44] - iridata[0]) / 11;

                oldZSU[1] = 0;
                oldZSU[3] = 0;
            }

            if (m_speedtype[side])
            {
                stime = iritime[0] * 1.0 / m_Frequency[side];
            }
            else
            {
                stime = (iritime[0] / 10000000) * 3600
                    + (iritime[0] / 100000) % 100 * 60
                    + (iritime[0] / 1000) % 100
                    + (iritime[0] % 1000) * 0.001;
            }

            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            savefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(savefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);
            int count = 0;

            int length = len;
            if (_Setting.IRIAlgorithmInterval == 0.1)
            {
                length = len - 2;
            }
            for (int i = 1; i < length; ++i)
            {

                if (i % plusenum == 0)
                {
                    if (m_speedtype[side])
                    {
                        etime = iritime[i] * 1.0 / m_Frequency[side];
                    }
                    else
                    {
                        etime = (iritime[i] / 10000000) * 3600
                            + (iritime[i] / 100000) % 100 * 60
                            + (iritime[i] / 1000) % 100
                            + (iritime[i] % 1000) * 0.001;
                    }
                    speedval = vallen / (etime - stime) * 3.6;
                    stime = etime;

                    //irival = irisum / plusenum;
                    irival = irisum / count;
                    count = 0;
                    if (datasrc)
                    {
                        irival = irival * _Setting.Acc_IRI_K_1 + _Setting.Acc_IRI_B_1;
                    }
                    else
                    {
                        if (IsParmFile)
                        {
                            //根据车速获取速度系数k、b
                            double kparm = kparms[parmnum - 1];
                            double bparm = bparms[parmnum - 1];
                            for (int pi = 0; pi < parmnum; ++pi)
                            {
                                if (speedval <= speedparms[pi])
                                {
                                    kparm = kparms[pi];
                                    bparm = bparms[pi];
                                    break;
                                }
                            }
                            irival = irival * kparm + bparm;
                        }
                        else
                        {
                            irival = irival * m_IRI_k[side] + m_IRI_b[side];
                        }
                    }

                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);
                    swspeed.WriteLine(string.Format("{0} {1}", iricnt, speedval), Encoding.UTF8);
                    irisum = 0;
                }
                if (_Setting.IRIAlgorithmInterval == 0.1)
                {
                    YSU = (iridata[2 + i] - iridata[i - 1]) / 0.3;

                }
                else
                {
                    YSU = (iridata[i] - iridata[i - 1]) / DeltLen;
                }

                for (int zi = 0; zi < 4; ++zi)
                {
                    ZSU[zi] = 0;
                    for (int zj = 0; zj < 4; ++zj)
                    {
                        if (_Setting.IRIAlgorithmInterval == 0.1)
                        {
                            ZSU[zi] += SZU100[zi * 4 + zj] * oldZSU[zj];
                        }
                        else
                        {
                            ZSU[zi] += SZU[zi * 4 + zj] * oldZSU[zj];
                        }

                    }
                    if (_Setting.IRIAlgorithmInterval == 0.1)
                    {
                        ZSU[zi] += PZU100[zi] * YSU;
                    }
                    else
                    {
                        ZSU[zi] += PZU[zi] * YSU;
                    }

                }
                irisum += Math.Abs(ZSU[0] - ZSU[2]);
                count++;
                for (int zi = 0; zi < 4; ++zi)
                {
                    oldZSU[zi] = ZSU[zi];
                }
            }


            sw.Close();
            fw.Close();
            swspeed.Close();
            fwspeed.Close();

            //250mm抽样
            savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            len = sdata.Length;
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();
        }





        /// <summary>
        /// 修复了原有方法中0.1m计算错误的问题
        /// 生成路面平整度，国际平整度指数（IRI）四分之一车轮算法，输出 tIRI_距离.txt 和 Speed_距离.txt
        /// <param name="fpath">输入文件路径（resample.txt）</param>
        /// <param name="vallen">IRI计算段长度（米，典型为10）</param>
        /// <param name="fname">输入文件名</param>
        /// <param name="datasrc">true: 加速度计数据，false: 纵断面数据</param>
        ///</summary> 
        public static void GenerateIRI_NEW(string fpath, int vallen, string fname, bool datasrc, int side,ProjectInfo projectInfo = null)
        {
             
            // 步骤1: 加载速度修正参数（如果存在）
            var (isParmFile, speedparms, kparms, bparms, parmnum) = LoadParameters(fpath, fname);

            // 步骤2: 加载和预处理原始数据
            double DeltLen = _Setting.IRIAlgorithmInterval; // 采样间隔（0.1m 或 0.25m）
            var (oridata, toridata, oritime, sdata, len) = LoadData(fpath);


            int maxRawLen = (int)Math.Round(projectInfo._EndDmi / 0.05);
            len = Math.Min(len, maxRawLen);
            oridata = oridata.Take(len).ToArray();
            toridata = toridata.Take(len).ToArray();
            oritime = oritime.Take(len).ToArray();
            sdata = sdata.Take(len).ToArray();

            // 步骤3: 应用均值滤波（可选，默认注释，与原始版本一致）
            for (int i = 2; i < len - 2; ++i)
            {
                oridata[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
            }
            string saveSpeedfname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));

            // 步骤3: 抽样到指定间隔（DeltLen）
            int qplusenum = Convert.ToInt32(DeltLen / 0.05); // 每0.1m抽样点数（0.05m原始间隔）
            var (iridata, iritime) = ResampleData(oridata, oritime, len, qplusenum);


            //string textFilePath = "C:\\Users\\cwb\\Desktop\\job\\平整度验证20251223\\右高程记录表300-400%25厘米.txt";
            //DeltLen = 0.25;
            //string [] textFileTxts =  File.ReadAllLines(textFilePath);
            //List<double> testDatas = new List<double>();
            //for (int i = 0; i < textFileTxts.Length; i++)
            //{
            //    if (double.TryParse(textFileTxts[i], out double value))
            //    {
            //        testDatas.Add(value*1000);
            //    } 
            //}
            //iridata = testDatas.ToArray();




            //先计算速度文件
            // 提前计算并输出速度文件
            string speedSavefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(speedSavefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);

            int plusenum = Convert.ToInt32(vallen / DeltLen); // 每段点数（e.g., 10m / 0.1m = 100）
            int start_i = (DeltLen == 0.1) ? 3 : 1; // 0.1m时从i=3开始以支持尾随YSU
            double stime = ParseTime(iritime[0], side);
            int iricnt = 0;

            // 预先计算所有速度点
            for (int i = 0; i < iritime.Length; i++)
            {
                if (i % plusenum == 0 && i > 0)
                {
                    double etime = ParseTime(iritime[i], side);
                    double speedval = (etime - stime) > 0 ? (vallen / (etime - stime) * 3.6) : 0;

                    swspeed.WriteLine(string.Format("{0} {1}", ++iricnt, speedval));
                    stime = etime;
                }
            }

            // 处理最后一个不完整段的速度
            if (iritime.Length > 0)
            {
                double etime = ParseTime(iritime[iritime.Length - 1], side);
                double partial_distance = (iritime.Length % plusenum) * DeltLen;
                double speedval = (etime - stime) > 0 ? (partial_distance / (etime - stime) * 3.6) : 0;
                swspeed.WriteLine(string.Format("{0} {1}", ++iricnt, speedval));
            }

            swspeed.Close();
            fwspeed.Close();
            int count = 0;
#if 平整度测试

            if (side !=0)
            {
                return;
            }
            List<string> values = new List<string>(); 
            List<string> newValues = new List<string>();

            double d_CurrentB = 0.0;    // 当前k值
            double d_BeforeB = 0.0;     // 前一次k值

            if (File.Exists(speedSavefname))
            {
                string[] speedTexts = File.ReadAllLines(speedSavefname);
                double[] speedValues = new double[speedTexts.Length];
                for (int i = 0; i < speedTexts.Length; i++)
                {
                    string[] parts = speedTexts[i].Split(' ');
                    if (parts.Length == 2) speedValues[i] = double.Parse(parts[1]);
                }

                int pointsPerSegment = plusenum;

                for (int i = 0; i < iridata.Length; i++)
                {
                    int segmentIndex = i / pointsPerSegment;
                    if (segmentIndex >= speedValues.Length) segmentIndex = speedValues.Length - 1;

                    double speedval = speedValues[segmentIndex];
                    double kparm = kparms[parmnum - 1];


                    kparm =  kparms.Average();

                    for (int pi = 0; pi < parmnum; ++pi)
                    {
                        if (speedval <= speedparms[pi])
                        {
                            kparm = kparms[pi];
                            break;
                        }
                    }

                    if(segmentIndex == 1)
                    {
                        d_CurrentB = kparm;
                        d_BeforeB = d_CurrentB;
                    }
                    else
                    {
                        if()
                        {
                        }
                    }

                        newValues[i] = iridata[i] * kparm;
                        if (i >=1100 && i < 1199)
                        {
                            values.Add(iridata[i].ToString());
                            newValues.Add((iridata[i] * kparm).ToString());
                        }

                  iridata[i] = iridata[i] * kparm;
                }
            }
               File.WriteAllLines("E:\\aaa\\oridata.txt",values);
            File.WriteAllLines("E:\\aaa\\oridataNew.txt", newValues);
#endif
            len = iridata.Length;

            // 步骤4: 初始化状态变量
            double[] oldZSU = InitializeState(iridata, DeltLen, len);

            // 步骤5: 初始化时间和输出文件
            stime = ParseTime(iritime[0], side);
            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);

            // 步骤6: 主循环计算IRI
            iricnt = 0;
            double irisum = 0;
            count = 0;
            
            double[] ZSU = new double[4];

            for (int i = start_i; i < len; ++i)
            {
                // 每段结束时处理IRI
                if (i % plusenum == 0)
                {
                    double etime = ParseTime(iritime[i], side);
                    double speedval = (etime - stime) > 0 ? (vallen / (etime - stime) * 3.6) : 0;
                    double irival = count > 0 ? irisum / count : 0;

                    // 应用IRI修正（加速度或速度因子）
                    irival = ApplyCorrection(irival, datasrc, speedval, isParmFile, speedparms, kparms, bparms, parmnum);

                    // 写入输出
                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);

                    // 重置
                    irisum = 0;
                    count = 0;
                    stime = etime;
                }

                // 计算YSU（输入坡度）
                double YSU = ComputeYSU(iridata, i, DeltLen);

                // 更新状态
                UpdateState(ZSU, oldZSU, YSU, DeltLen);

                // 累加IRI
                irisum += Math.Abs(ZSU[0] - ZSU[2]);
                count++;

                // 更新oldZSU
                Array.Copy(ZSU, oldZSU, 4);
            }

            // 步骤7: 处理最后一个不完整段
            if (count > 0)
            {
                double etime = ParseTime(iritime[len - 1], side);
                double partial_distance = count * DeltLen;
                double speedval = (etime - stime) > 0 ? (partial_distance / (etime - stime) * 3.6) : 0;
                double irival = irisum / count;

                // 应用IRI修正
                irival = ApplyCorrection(irival, datasrc, speedval, isParmFile, speedparms, kparms, bparms, parmnum);

                // 写入输出
                sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);
            }

            // 步骤8: 关闭输出文件
            sw.Close();
            fw.Close();

            // 步骤9: 生成250mm抽样文件
            savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            for (int i = 0; i < sdata.Length; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();
        }



        /// <summary>
        /// 加载速度修正参数从 Coeff.dat
        /// </summary>
        private static (bool isParmFile, double[] speedparms, double[] kparms, double[] bparms, int parmnum) LoadParameters(string fpath, string fname)
        {
            bool isParmFile = false;
            string fparmpath = fpath.Replace("resample.txt", "Coeff.dat");
            double[] speedparms = null, kparms = null, bparms = null;
            int parmnum = 0;

            if (File.Exists(fparmpath))
            {
                isParmFile = true;
                string[] parms = File.ReadAllLines(fparmpath);
                try
                {
                    parmnum = int.Parse(parms[0]);
                    speedparms = new double[parmnum];
                    kparms = new double[parmnum];
                    bparms = new double[parmnum];
                    int idx = 1;
                    for (int i = 0; i < parmnum; ++i) speedparms[i] = double.Parse(parms[idx++]);
                    for (int i = 0; i < parmnum; ++i) kparms[i] = double.Parse(parms[idx++]);
                    for (int i = 0; i < parmnum; ++i) bparms[i] = double.Parse(parms[idx++]);
                }
                catch
                {
                    MessageBox.Show("读取文件出错，请检查！\r\n" + fparmpath);
                }
            }
            return (isParmFile, speedparms, kparms, bparms, parmnum);
        }

        /// <summary>
        /// 加载和解析 resample.txt 数据
        /// </summary>
        private static (double[] oridata, double[] toridata, int[] oritime, string[] sdata, int len) LoadData(string fpath)
        {
            string[] sdata = File.ReadAllLines(fpath);
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            int[] oritime = new int[len];

            for (int i = 0; i < len; i++)
            {
                string[] s = sdata[i].Split('\t');
                if (s.Length <= 1)
                {
                    s = sdata[i].Split(' ');
                }
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[2]);
                        toridata[i] = oridata[i];
                        oritime[i] = int.Parse(s[3]);
                    }
                    catch
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                            toridata[i] = oridata[i];
                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
                        }
                    }
                }
            }
            return (oridata, toridata, oritime, sdata, len);
        }


        /// <summary>
        /// 测试方法
        /// </summary>
        /// <param name="fpath"></param>
        /// <returns></returns>
        private static (double[] oridata, double[] toridata, int[] oritime, string[] sdata, int len) LoadDataFormExcel(string fpath)
        {
            string[] sdata = File.ReadAllLines(fpath);
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            int[] oritime = new int[len];

            for (int i = 0; i < len; i++)
            {
                string[] s = sdata[i].Split('\t');
                if (s.Length <= 1)
                {
                    s = sdata[i].Split(' ');
                }
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[2]);
                        toridata[i] = oridata[i];
                        oritime[i] = int.Parse(s[3]);
                    }
                    catch
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                            toridata[i] = oridata[i];
                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
                        }
                    }
                }
            }
            return (oridata, toridata, oritime, sdata, len);
        }

        /// <summary>
        /// 抽样数据到指定间隔（0.1m 或 0.25m）
        /// </summary>
        private static (double[] iridata, int[] iritime) ResampleData(double[] oridata, int[] oritime, int len, int qplusenum)
        {
            double[] iridata = new double[(len + qplusenum - 1) / qplusenum];
            int[] iritime = new int[iridata.Length];
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            return (iridata, iritime);
        }

        /// <summary>
        /// 初始化状态变量 oldZSU
        /// </summary>
        private static double[] InitializeState(double[] iridata, double DeltLen, int len)
        {
            double[] oldZSU = new double[4];
            int index = (DeltLen == 0.1) ? Math.Min(110, len - 1) : Math.Min(44, len - 1);
            oldZSU[0] = (iridata[index] - iridata[0]) / 11;
            oldZSU[2] = (iridata[index] - iridata[0]) / 11;
            oldZSU[1] = 0;
            oldZSU[3] = 0;
            return oldZSU;
        }

        /// <summary>
        /// 解析时间戳为秒（支持脉冲计数或 HHMMSSmmm 格式）
        /// </summary>
        private static double ParseTime(int timeValue,int side)
        {
            if (m_speedtype[side])
            {
                return timeValue * 1.0 / m_Frequency[side];
            }
            else
            {
                return (timeValue / 10000000) * 3600 +
                       (timeValue / 100000 % 100) * 60 +
                       (timeValue / 1000 % 100) +
                       (timeValue % 1000) * 0.001;
            }
        }

        /// <summary>
        /// 计算 YSU（输入坡度）
        /// </summary>
        private static double ComputeYSU(double[] iridata, int i, double DeltLen)
        {
            if (DeltLen == 0.1)
            {
                return (iridata[i] - iridata[i - 3]) / 0.3;
            }
            else
            {
                return (iridata[i] - iridata[i - 1]) / DeltLen;
            }
        }

        /// <summary>
        /// 更新状态向量 ZSU
        /// </summary>
        private static void UpdateState(double[] ZSU, double[] oldZSU, double YSU, double DeltLen)
        {
            double[] szu = (DeltLen == 0.1) ? SZU100 : SZU;
            double[] pzu = (DeltLen == 0.1) ? PZU100 : PZU;
            for (int zi = 0; zi < 4; ++zi)
            {
                ZSU[zi] = 0;
                for (int zj = 0; zj < 4; ++zj)
                {
                    ZSU[zi] += szu[zi * 4 + zj] * oldZSU[zj];
                }
                ZSU[zi] += pzu[zi] * YSU;
            }
        }

        /// <summary>
        /// 应用 IRI 修正（加速度或速度因子）
        /// </summary>
        private static double ApplyCorrection(double irival, bool datasrc, double speedval, bool isParmFile, double[] speedparms, double[] kparms, double[] bparms, int parmnum)
        {
            if (kparms==null)
            {
                return irival;
            }
            double kparm = kparms[parmnum - 1];
            double bparm = bparms[parmnum - 1];
            for (int pi = 0; pi < parmnum; ++pi)
            {
                if (speedval <= speedparms[pi])
                {
                    kparm = kparms[pi];
                    bparm = bparms[pi];
                    break;
                }
            }
            kparm = kparms.Average();
            irival = irival * kparm;
             //irival = irival * kparm + bparm;
             
            return irival;
        }





        private static void GenerateIRI(string fpath, float vallen, string fname, bool datasrc,int side)
        {
            bool IsParmFile = false;
            String fparmpath = fpath.Replace("resample.txt", "Coeff.dat");
            if (File.Exists(fparmpath))
            {
                IsParmFile = true;
            }

            string[] parms = null;
            double[] speedparms = null;
            double[] kparms = null;
            double[] bparms = null;
            int parmnum = 0;
            if (IsParmFile)
            {
                int idx = 0;
                parms = File.ReadAllLines(fparmpath);
                try
                {
                    parmnum = int.Parse(parms[idx]);
                }
                catch (System.Exception)
                {
                    MessageBox.Show("读取文件出错，请检查！\r\n" + fparmpath);
                    return;
                }
                speedparms = new double[parmnum];
                kparms = new double[parmnum];
                bparms = new double[parmnum];
                for (int i = 0; i < parmnum; ++i)
                {
                    speedparms[i] = double.Parse(parms[++idx]);
                }
                for (int i = 0; i < parmnum; ++i)
                {
                    kparms[i] = double.Parse(parms[++idx]);
                }
                for (int i = 0; i < parmnum; ++i)
                {
                    bparms[i] = double.Parse(parms[++idx]);
                }
            }

            const double DeltLen = 0.25;
            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);

            int qplusenum = Convert.ToInt32(DeltLen / 0.05);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            double[] iridata = new double[len];
            int[] oritime = new int[len];
            int[] iritime = new int[len];

            //计算IRI
            int plusenum = Convert.ToInt32(vallen / DeltLen);//IRI距离内有多少个250mm
            int iricnt = 0;


            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[2]);
                        toridata[i] = oridata[i];
                        oritime[i] = int.Parse(s[3]);
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                            toridata[i] = oridata[i];
                            oritime[i] = oritime[i - 1] * 2 - oritime[i - 2];
                        }
                    }
                }
            }
            //取5个点做均值滤波
            for (int i = 2; i < len - 2; ++i)
            {
                oridata[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
            }
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            //len = iridata.Length;

            double speedval = 0;
            double stime;
            double etime;

            double YSU = 0;
            double[] ZSU = new double[4];
            double[] oldZSU = new double[4];
            oldZSU[0] = iridata[1] - iridata[0];
            oldZSU[1] = iridata[1] - iridata[0];
            oldZSU[2] = 0;
            oldZSU[3] = 0;

            if (m_speedtype[side])
            {
                stime = iritime[0] * 1.0 / m_Frequency[side];
            }
            else
            {
                stime = (iritime[0] / 10000000) * 3600
                    + (iritime[0] / 100000) % 100 * 60
                    + (iritime[0] / 1000) % 100
                    + (iritime[0] % 1000) * 0.001;
            }
            string savefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(savefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);
            for (int i = 1; i < len; ++i)
            {
                if (m_speedtype[side])
                {
                    etime = iritime[i] * 1.0 / m_Frequency[side];
                }
                else
                {
                    etime = (iritime[i] / 10000000) * 3600
                        + (iritime[i] / 100000) % 100 * 60
                        + (iritime[i] / 1000) % 100
                        + (iritime[i] % 1000) * 0.001;
                }
                speedval = vallen / (etime - stime) * 3.6;
                stime = etime;
                if (datasrc)
                {

                }
                else
                {
                    if (IsParmFile)
                    {
                        //根据车速获取速度系数k、b
                        double kparm = kparms[parmnum - 1];
                        double bparm = bparms[parmnum - 1];
                        for (int pi = 0; pi < parmnum; ++pi)
                        {
                            if (speedval <= speedparms[pi])
                            {
                                kparm = kparms[pi];
                                bparm = bparms[pi];
                                break;
                            }

                        }
                        

                    }
                    else
                    {

                    }
                }
                swspeed.WriteLine(string.Format("{0} {1}", ++iricnt, speedval), Encoding.UTF8);
            }

            swspeed.Close();
            fwspeed.Close();
        }
        /// <summary>
        /// 生成路面平整度，相邻算法-高程差，tIRI_距离.txt
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="iridis"></param>
        /// <param name="pluseparm"></param>
        /// <param name="datasrc">true 仅加速度计，false 纵断面</param>
        /// 



        private static void GenerateIRI_1(string fpath, int vallen, string fname, bool datasrc,int side)
        {
            const double DeltLen = 0.25;

            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);

            int qplusenum = Convert.ToInt32(DeltLen / 0.05);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            double[] iridata = new double[(len + qplusenum - 1) / qplusenum];
            int[] oritime = new int[len];
            int[] iritime = new int[(len + qplusenum - 1) / qplusenum];

            //计算IRI
            int plusenum = Convert.ToInt32(vallen / DeltLen);//IRI距离内有多少个250mm
            int iricnt = 0;
            double irisum = 0;
            double irival = 0;

            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    oridata[i] = double.Parse(s[2]);
                    toridata[i] = oridata[i];
                    oritime[i] = int.Parse(s[3]);
                }
            }
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            len = iridata.Length;

            double speedval = 0;
            double stime;
            double etime;
            if (m_speedtype[side])
            {
                stime = iritime[0] * 1.0 / m_Frequency[side];
            }
            else
            {
                stime = (iritime[0] / 10000000) * 3600
                    + (iritime[0] / 100000) % 100 * 60
                    + (iritime[0] / 1000) % 100
                    + (iritime[0] % 1000) * 0.001;
            }

            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            savefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(savefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);
            for (int i = 1; i < len; ++i)
            {
                if (i % plusenum == 0)
                {
                    if (m_speedtype[side])
                    {
                        etime = iritime[i] * 1.0 / m_Frequency[side];
                    }
                    else
                    {
                        etime = (iritime[i] / 10000000) * 3600
                            + (iritime[i] / 100000) % 100 * 60
                            + (iritime[i] / 1000) % 100
                            + (iritime[i] % 1000) * 0.001;
                    }
                    speedval = vallen / (etime - stime) * 3.6;
                    stime = etime;

                    irival = irisum / vallen;
                    if (datasrc)
                    {
                        irival = irival * _Setting.Acc_IRI_K_1 + _Setting.Acc_IRI_B_1;
                    }
                    else
                    {
                        irival = irival * _Setting.IRIk + _Setting.IRIb;
                    }
                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);
                    swspeed.WriteLine(string.Format("{0} {1}", iricnt, speedval), Encoding.UTF8);
                    irisum = 0;
                }
                irisum += Math.Abs(iridata[i] - iridata[i - 1]);
            }
            sw.Close();
            fw.Close();
            swspeed.Close();
            fwspeed.Close();

            //250mm抽样
            savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            len = sdata.Length;
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();
        }
        private static float[] LaserFilter(float[] InputArray)
        {
            //从xml文件中获取高斯滤波器系数
            String filepath = System.Windows.Forms.Application.StartupPath + @"\Coeff.xml";
            XmlDocument xml = new XmlDocument();//初始化一个xml实例
            xml.Load(filepath);//导入指定xml文件
            XmlNodeList nodelist = xml.SelectNodes(rootcoeffname + "/激光器滤波模板");
            String tempcoeff = nodelist[0].InnerText;
            String[] s = System.Text.RegularExpressions.Regex.Split(tempcoeff, @"\r\n");
            int coefflen = s.Length;
            float[] LaserFilterCoeff = new float[coefflen];
            for (int i = 0; i < coefflen; ++i)
            {
                LaserFilterCoeff[i] = float.Parse(s[i]);
            }

            nodelist = xml.SelectNodes(rootcoeffname + "/激光器滤波群延时");
            tempcoeff = nodelist[0].InnerText;
            int gdoff = int.Parse(tempcoeff);

            //进行滤波
            int n = InputArray.Length;
            float tempzero = InputArray[0];
            float[] FilteredArray = new float[n];
            for (int i = 0; i < n; i++)
            {
                FilteredArray[i] = InputArray[i] - tempzero;
            }
            FilteredArray = FilterConv(FilteredArray, LaserFilterCoeff);
            //for (int i = gdoff; i < n; ++i)
            //{
            //    FilteredArray[i - gdoff] = FilteredArray[i];
            //}
            return FilteredArray;
        }
        //数字滤波器，计算卷积
        private static float[] FilterConv(float[] signal, float[] filter)
        {
            int len = signal.Length;
            int flen = filter.Length;
            float[] output = new float[len];
            int i, j;

            for (i = 0; i < len; i++)
            {
                float tempout = 0;
                for (j = 0; j < flen && i - j > -1; j++)
                {
                    tempout = tempout + filter[j] * signal[i - j];
                }
                output[i] = tempout;
            }
            return output;
        }   //将加速度推算成车的颠簸值
        private static float[] Acc2Dis(float[] Laser, float[] Acc, string[] time, float DAQSampleT)
        {
            int cnt = Acc.Length;
            float[] Dis = new float[cnt];
            int gdoff;
            int LaserAccdisOff;

            float Aver = Acc.Average();
            for (int i = 0; i < cnt; ++i)
            {
                Acc[i] = Acc[i] - Aver;
            }

            //从xml文件中获取滤波器系数
            String filepath = System.Windows.Forms.Application.StartupPath + @"\Coeff.xml";
            XmlDocument xml = new XmlDocument();//初始化一个xml实例
            xml.Load(filepath);//导入指定xml文件
            XmlNodeList nodelist = xml.SelectNodes(rootcoeffname + "/加速度计滤波模板");
            String tempcoeff = nodelist[0].InnerText;
            String[] s = System.Text.RegularExpressions.Regex.Split(tempcoeff, @"\r\n");
            int coefflen = s.Length;
            float[] FilterCoeff = new float[coefflen];
            for (int i = 0; i < coefflen; ++i)
            {
                FilterCoeff[i] = float.Parse(s[i]);
            }
            nodelist = xml.SelectNodes(rootcoeffname + "/加速度计滤波群延时");
            tempcoeff = nodelist[0].InnerText;
            gdoff = int.Parse(tempcoeff);

            nodelist = xml.SelectNodes(rootcoeffname + "/颠簸移位数目");
            tempcoeff = nodelist[0].InnerText;
            LaserAccdisOff = int.Parse(tempcoeff);
            GC.Collect();

            //加速度滤波
            Dis = FilterConv(Acc, FilterCoeff);
            GC.Collect();
            for (int i = gdoff; i < cnt; ++i)
            {
                Dis[i - gdoff] = Dis[i];
            }
            for (int i = 0; i < cnt; ++i)
            {
                Acc[i] = Acc[i] - Dis[i];
            }
            //积分
            Acc = OnceIntergation(Acc, DAQSampleT);
            GC.Collect();


            //速度滤波
            Dis = FilterConv(Acc, FilterCoeff);
            GC.Collect();
            for (int i = gdoff; i < cnt; ++i)
            {
                Dis[i - gdoff] = Dis[i];
            }
            for (int i = 0; i < cnt; ++i)
            {
                Acc[i] = Acc[i] - Dis[i];
            }
            Acc = OnceIntergation(Acc, DAQSampleT);
            GC.Collect();


            //颠簸滞后于激光
            for (int i = cnt - LaserAccdisOff - 1; i > -1; --i)
            {
                Acc[i + LaserAccdisOff] = Acc[i];
            }

            //激光-颠簸
            for (int i = 0; i < cnt; ++i)
            {
                Dis[i] = Laser[i] - Acc[i] * 1000;
            }

            return Dis;
        }    //速度推算位移
        private static float[] OnceIntergation(float[] input, float time)
        {
            int len = input.Length;
            float[] output = new float[len];
            output[0] = 0;
            output[1] = 0;
            for (int i = 2; i < len; ++i)
            {
                output[i] = output[i - 1] + (input[i] + 4 * input[i - 1] + input[i - 2]) / 6 * time;
            }
            return output;
        }
        //采集卡采样率
        private static void Dis2IRI(float[] Dis, float[] DMI, float PluseDis, float PartLength, string[] Time, string outPath)
        {
            int DMIPluseCnt = 0;
            int Dislen = Dis.Length;
            ArrayList SampleDis = new ArrayList();
            ArrayList SampleTime = new ArrayList();
            //用编码器脉冲抽样
            for (int i = 1; i < Dislen; ++i)
            {
                if (DMI[i - 1] < 1.65 && DMI[i] >= 1.65)
                {
                    SampleDis.Add(Dis[i]);
                    SampleTime.Add(Time[i]);
                    ++DMIPluseCnt;
                }
            }

            //求整个里程的脉冲距离，为什么数采集卡的脉冲数和拍照数量对不上？？？
            //PluseDis = ProjectTotalMile / DMIPluseCnt;
            //PulseLength = PluseDis;

            int PartDMICnt = (int)(PartLength / (PluseDis * 0.001));//每段中DMI脉冲的个数

            String filename = outPath;
            filename = filename + "\\DAQ0\\resample.txt";
            FileStream fr2 = File.Open(filename, FileMode.Create);
            StreamWriter sr2 = new StreamWriter(fr2);
            //用编码器脉冲抽样
            double valueThree = 0;
            for (int i = 0; i < SampleDis.Count; ++i)
            {
                if (i == 0)
                {
                    sr2.WriteLine("{1}\t0\t{1}\t{0}", valueThree, SampleDis[i]);
                }
                else
                {
                    valueThree += double.Parse(SampleTime[i].ToString()) - double.Parse(SampleTime[i - 1].ToString());
                    sr2.WriteLine("{1}\t0\t{1}\t{0}", valueThree, SampleDis[i]);
                }


            }
            sr2.Close();
            fr2.Close();

            //计算再次抽样，计算IRI
            //   ArrayList IRI = new ArrayList();
            // IRI = Dis2IRI(SampleTime, SampleDis, PulseLength,10, false);
            //  return IRI;
        }
        private static float PulseLength = 0.06591143f;
        /// <summary>
        /// 生成路面平整度，相邻算法-高程差，tIRI_距离.txt  
        /// cwb 适配旧设备
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="iridis"></param>
        /// <param name="pluseparm"></param>
        /// <param name="datasrc">true 仅加速度计，false 纵断面</param>
        private static void GenerateIRI_3(string fpath, int vallen, string fname, bool datasrc, ProjectInfo _ProjectInfo, WinProcessBar bar,int side)
        {
            const double DeltLen = 0.25;

            if (_ProjectInfo._PlusLength != 0)
            {
                PulseLength = _ProjectInfo._PlusLength / 1000;
            }
            String basePath = Directory.GetParent(fpath).Parent.FullName;
            if (!File.Exists(fpath))
            {
                float DAQSampleT = 1.0f / _ProjectInfo.DAQSampleFrequency;

                const string IRMstandard = "%AD,015018769,13193,07298,00000,00000,97";
                //获取DAQ.txt文件的行数

                string filename = basePath + "\\DAQ0\\daq_0.daq";
                FileStream fr = File.OpenRead(filename);
                StreamReader sr = new StreamReader(fr);
                String strline;
                int DAQtotallineCnt = 0;

                string[] LocalTimeChanged;
                float[] EcoderChanged;
                float[] LaserRangeChanged;
                float[] AccChanged;

                while ((strline = sr.ReadLine()) != null)
                {
                    if (strline.Length == IRMstandard.Length && strline.StartsWith("%AD"))
                    {
                        DAQtotallineCnt++;
                    }
                }
                sr.Close();
                fr.Close();
                GC.Collect();



                LocalTimeChanged = new string[DAQtotallineCnt];
                EcoderChanged = new float[DAQtotallineCnt];
                LaserRangeChanged = new float[DAQtotallineCnt];
                AccChanged = new float[DAQtotallineCnt];
                //打开DAQ.txt文件并解析
                fr = File.OpenRead(filename);
                sr = new StreamReader(fr);
                int i = 0;

                float tempscale = 4.096f / 65565.0f;
                while ((strline = sr.ReadLine()) != null)
                {
                    if (strline.Length == IRMstandard.Length && strline.StartsWith("%AD"))
                    {
                        String[] record = System.Text.RegularExpressions.Regex.Split(strline, @"\,");
                        try
                        {
                            LocalTimeChanged[i] = record[1];
                            EcoderChanged[i] = float.Parse(record[4]) * tempscale;
                            LaserRangeChanged[i] = float.Parse(record[3]) * tempscale;
                            AccChanged[i] = float.Parse(record[2]) * tempscale;
                            i++;
                        }
                        catch (System.Exception)
                        { }
                    }
                }
                bar.SetIRIVal(0.3);
                sr.Close();
                fr.Close();
                GC.Collect();

                LaserRangeChange(ref LaserRangeChanged);//将原始激光电压值转换为激光测距值
                GC.Collect();
                bar.SetIRIVal(0.4);
                CalibrationAcc(ref AccChanged);//将原始的Acc的电压值转化为真实的Acc的值
                GC.Collect();
                bar.SetIRIVal(0.5);
                LaserRangeChanged = LaserFilter(LaserRangeChanged);//激光器滤波 
                GC.Collect();

                AccChanged = Acc2Dis(LaserRangeChanged, AccChanged, LocalTimeChanged, DAQSampleT);
                GC.Collect();


                //计算平整度
                Dis2IRI(AccChanged, EcoderChanged, PulseLength, 10, LocalTimeChanged, basePath);
                GC.Collect();

            }


            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);

            int qplusenum = Convert.ToInt32(DeltLen / PulseLength);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            int DMIPluseCnt = len;
            ArrayList IRIVal = new ArrayList();
            ArrayList DMICnt = new ArrayList();
            ArrayList myIRILable = new ArrayList();
            ArrayList ReSampleDis = new ArrayList();
            ArrayList SampleDis = new ArrayList();
            GetIRIStack(myIRILable, DMICnt, DMIPluseCnt, _ProjectInfo, PulseLength, vallen);
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            double[] iridata = new double[(len + qplusenum - 1) / qplusenum];
            int[] oritime = new int[len];
            int[] iritime = new int[(len + qplusenum - 1) / qplusenum];

            //计算IRI
            int plusenum = Convert.ToInt32(vallen / 0.25);//IRI距离内有多少个250mm
                                                          //   int plusenum =37;//IRI距离内有多少个250mm
                                                          //   int plusenum = Convert.ToInt32((vallen * PulseLength)/0.25);//IRI距离内有多少个250mm
                                                          // double d = vallen * PulseLength/ 0.25;


            int iricnt = 0;
            double irisum = 0;

            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    oridata[i] = double.Parse(s[2]);
                    toridata[i] = oridata[i];
                    oritime[i] = int.Parse(s[3]);
                    SampleDis.Add(s[2]);
                }
            }
            float temp = 0;

            {
                int i, j, k;
                for (i = 0, k = 0, j = 0; i < DMIPluseCnt - 1 && j < DMICnt.Count; ++i)
                {
                    ReSampleDis.Add(float.Parse(SampleDis[i].ToString()));

                    if (k == int.Parse(DMICnt[j].ToString()))
                    {
                        temp = float.Parse(DMICnt[j].ToString()) * PulseLength;
                        IRIVal.Add(OneIRI(ReSampleDis, temp, PulseLength));

                        ReSampleDis.Clear();
                        j++;
                        k = 0;
                        i--;
                    }
                    else
                    {
                        ++k;
                    }
                }
                if (j < DMICnt.Count)
                {
                    temp = float.Parse(DMICnt[j].ToString()) * PulseLength;
                    IRIVal.Add(OneIRI(ReSampleDis, temp, PulseLength));
                }
                IRIVal.Add(OneIRI(ReSampleDis, temp, PulseLength));
                DMICnt.Add(DMICnt[DMICnt.Count - 1]);

                String filename = basePath;
                String StackDMICnt = filename + "\\DAQ0\\StackDMICnt.txt";
                FileStream fr1 = File.Open(StackDMICnt, FileMode.Create);
                StreamWriter sr1 = new StreamWriter(fr1);
                //用编码器脉冲抽样
                for (i = 0; i < myIRILable.Count && i < DMICnt.Count; ++i)
                {
                    sr1.WriteLine(string.Format("{0} {1}", myIRILable[i], DMICnt[i]));
                }
                sr1.Close();
                fr1.Close();

            }
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            len = iridata.Length;

            double speedval = 0;
            double stime;
            double etime;
            if (m_speedtype[side])
            {
                stime = iritime[0] * 1.0 / m_Frequency[side];
            }
            else
            {
                stime = (iritime[0] / 10000000) * 3600
                    + (iritime[0] / 100000) % 100 * 60
                    + (iritime[0] / 1000) % 100
                    + (iritime[0] % 1000) * 0.001;
            }

            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            savefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(savefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);
            for (int i = 1; i < len; ++i)
            {
                if (i % plusenum == 0)
                {
                    if (m_speedtype[side])
                    {
                        etime = iritime[i] * 1.0 / m_Frequency[side];
                    }
                    else
                    {
                        etime = (iritime[i] / 10000000) * 3600
                            + (iritime[i] / 100000) % 100 * 60
                            + (iritime[i] / 1000) % 100
                            + (iritime[i] % 1000) * 0.001;
                    }
                    speedval = vallen / (etime - stime) * 3.6;
                    stime = etime;





                    swspeed.WriteLine(string.Format("{0} {1}", ++iricnt, speedval), Encoding.UTF8);
                    irisum = 0;
                }


            }
            iricnt = 0;
            for (int i = 0; i < IRIVal.Count; i++)
            {
                sw.WriteLine(string.Format("{0} {1}", ++iricnt, IRIVal[i]), Encoding.UTF8);
            }
            sw.Close();
            fw.Close();
            swspeed.Close();
            fwspeed.Close();

            //250mm抽样
            savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            len = sdata.Length;
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();
        }
        private const string rootcoeffname = "/CoeffVal/IRM";
        //将原始激光电压值转换为激光测距值
        private static void LaserRangeChange(ref float[] LaserRangeOrigin)
        {
            //从xml文件中获取激光器标定系数
            String filepath = System.Windows.Forms.Application.StartupPath + @"\Coeff.xml";
            XmlDocument xml = new XmlDocument();//初始化一个xml实例
            xml.Load(filepath);//导入指定xml文件
            XmlNodeList nodelist = xml.SelectNodes(rootcoeffname + "/激光器标定系数");
            String tempcoeff = nodelist[0].InnerText;
            String[] s = System.Text.RegularExpressions.Regex.Split(tempcoeff, @"\r\n");
            int coefflen = s.Length;
            float[] LaserCoeff = new float[coefflen];
            for (int i = 0; i < coefflen; ++i)
            {
                LaserCoeff[i] = float.Parse(s[i]);
            }

            //标定激光器
            int n = LaserRangeOrigin.Length;
            for (int ii = 0; ii < n; ii++)
            {
                LaserRangeOrigin[ii] = LaserCoeff[0] * LaserRangeOrigin[ii] + LaserCoeff[1];
            }
        }   //将原始的Acc的电压值转化为真实的Acc的值
        private static void CalibrationAcc(ref float[] Acc)
        {
            //从xml文件中获取激光器标定系数
            String filepath = System.Windows.Forms.Application.StartupPath + @"\Coeff.xml";
            XmlDocument xml = new XmlDocument();//初始化一个xml实例
            xml.Load(filepath);//导入指定xml文件
            XmlNodeList nodelist = xml.SelectNodes(rootcoeffname + "/加速度计标定系数");
            String tempcoeff = nodelist[0].InnerText;
            String[] s = System.Text.RegularExpressions.Regex.Split(tempcoeff, @"\r\n");
            int coefflen = s.Length;
            float[] AccCoeff = new float[coefflen];
            for (int i = 0; i < coefflen; ++i)
            {
                AccCoeff[i] = float.Parse(s[i]);
            }

            //float tempscale = 4.096f / 65565.0f;
            //AccCoeff[0] = AccCoeff[0] * tempscale;
            //标定加速度计
            int cnt = Acc.Length;
            for (int ii = 0; ii < cnt; ii++)
            {
                Acc[ii] = AccCoeff[0] * Acc[ii] + AccCoeff[1];
            }
        }
        private static float OneIRI(ArrayList ReSampleDis, float PartLength, float PulseLength)
        {
            //计算单个平整度
            PartLength = (float)Math.Round(PartLength);
            int reresamplenum = Convert.ToInt16(Math.Round(0.25 / PulseLength));
            float IRI = 0;
            int len = ReSampleDis.Count / reresamplenum; //250mm中有多少个编码器脉冲
            for (int i = 1; i < len; ++i)
            {
                IRI = IRI + Math.Abs(float.Parse(ReSampleDis[(i - 1) * reresamplenum].ToString()) - float.Parse(ReSampleDis[i * reresamplenum].ToString()));
            }
            IRI = IRI / PartLength;

            ////根据速度对平整度进行缩放
            //DateTime start = DateTime.ParseExact(ReSampleTime[0].ToString(), "hhmmssfff", null);
            //DateTime end = DateTime.ParseExact(ReSampleTime[ReSampleTime.Count - 1].ToString(), "hhmmssfff", null);
            //float time = (end - start).TotalSeconds;
            //float Vel = PartLength / time / 3.6;

            //float scale = IRIScaleCoeff[0];
            //len = IRIScaleCoeff.Length;
            //for (int i = 1; i < len; i++)
            //{
            //    scale = Vel * scale + IRIScaleCoeff[i];
            //}
            //IRI = IRI * scale;
            return IRI;
        }
        private static void GetIRIStack(ArrayList Stack, ArrayList DMICnt, int DMIPluseCnt, ProjectInfo _ProjectInfo, float PulseLength, int My_gap)
        {
            int currentDistance = _ProjectInfo._StartMile;
            int endDmiVal = _ProjectInfo._EndMile;
            int Direction = _ProjectInfo._Direction;

            //生成加入桩号信息的record文件
            String filename = _Setting.DefaultPath;
            String MileStoneCaliInfo = filename + "\\MileStoneCaliInfo.txt";
            Stack.Clear();
            DMICnt.Clear();


            int oldOrigStact = currentDistance;
            int curOrigStact = 0;
            int oldLaterStack = currentDistance;
            int curLaterStack = 0;
            int oldDMIVal = 0;
            int curDMIVal = 0;
            float DMIScale = 1;
            int oldcurrentDistance = currentDistance;
            int tempcnt = 0;
            if (File.Exists(MileStoneCaliInfo))
            {
                Stack.Add(currentDistance.ToString("K0000+000"));
                FileStream fr = File.OpenRead(MileStoneCaliInfo);
                StreamReader sr = new StreamReader(fr);
                String strline;
                while ((strline = sr.ReadLine()) != null && currentDistance >= 0)
                {
                    String[] s = strline.Split(' ');
                    curOrigStact = int.Parse(s[0].Substring(1, 4)) * 1000 + int.Parse(s[0].Substring(6, 3));
                    curLaterStack = int.Parse(s[1].Substring(1, 4)) * 1000 + int.Parse(s[1].Substring(6, 3));
                    curDMIVal = int.Parse(s[3]);

                    DMIScale = (curDMIVal - oldDMIVal) * 1.0f / (curLaterStack - oldLaterStack);
                    while ((Direction > 0 && currentDistance <= curLaterStack)
                        || (Direction < 0 && currentDistance >= curLaterStack))
                    {
                        if (currentDistance % My_gap == 0 && currentDistance != oldcurrentDistance && currentDistance >= 0)
                        {
                            Stack.Add(currentDistance.ToString("K0000+000"));
                            tempcnt = Convert.ToInt32(Math.Floor((currentDistance - oldcurrentDistance) * DMIScale / PulseLength));
                            DMIPluseCnt = DMIPluseCnt - tempcnt;
                            DMICnt.Add(tempcnt);
                            oldcurrentDistance = currentDistance;
                        }
                        currentDistance = currentDistance + Direction;
                    }
                    currentDistance = currentDistance - Direction;

                    oldOrigStact = curOrigStact;
                    oldLaterStack = curLaterStack;
                    oldDMIVal = curDMIVal;
                }
                sr.Close();
                fr.Close();

                tempcnt = Convert.ToInt32(Math.Round(My_gap / PulseLength));
                while (DMIPluseCnt > tempcnt)
                {
                    currentDistance = currentDistance + My_gap * Direction;
                    if (currentDistance >= 0)
                    {
                        Stack.Add(currentDistance.ToString("K0000+000"));
                        DMIPluseCnt = DMIPluseCnt - tempcnt;
                        DMICnt.Add(tempcnt);
                    }
                    else
                    {
                        break;
                    }
                }
                currentDistance = Convert.ToInt32(Math.Round(currentDistance + DMIPluseCnt * PulseLength * Direction));
                if (currentDistance >= 0)
                {
                    Stack.Add(currentDistance.ToString("K0000+000"));
                    DMICnt.Add(DMIPluseCnt);
                }
            }

            else
            {
                if (Direction > 0)
                {
                    tempcnt = Convert.ToInt32(Math.Round(Math.Min(My_gap - currentDistance % My_gap, endDmiVal - currentDistance % My_gap) / PulseLength));
                }
                else
                {
                    tempcnt = Convert.ToInt32(Math.Round((currentDistance % My_gap) / PulseLength));
                }

                if (tempcnt != 0)
                {
                    Stack.Add(currentDistance.ToString("K0000+000"));
                }
                currentDistance = Convert.ToInt32(Math.Round(currentDistance + tempcnt * PulseLength * Direction));
                Stack.Add(currentDistance.ToString("K0000+000"));
                DMIPluseCnt = DMIPluseCnt - tempcnt;
                if (tempcnt != 0)
                {
                    DMICnt.Add(tempcnt);
                }

                tempcnt = Convert.ToInt32(Math.Round(My_gap / PulseLength));
                while (DMIPluseCnt > tempcnt)
                {
                    currentDistance = currentDistance + My_gap * Direction;
                    if (currentDistance >= 0)
                    {
                        Stack.Add(currentDistance.ToString("K0000+000"));
                        DMIPluseCnt = DMIPluseCnt - tempcnt;
                        DMICnt.Add(tempcnt);
                    }
                    else
                    {
                        break;
                    }
                }
                currentDistance = Convert.ToInt32(Math.Round(currentDistance + DMIPluseCnt * PulseLength * Direction));
                if (currentDistance >= 0)
                {
                    Stack.Add(currentDistance.ToString("K0000+000"));
                    DMICnt.Add(DMIPluseCnt);
                }
            }
        }


        /// <summary>
        /// 生成路面平整度，国际平整度指数四分之一车轮算法，tIRI_距离.txt，用内业软件配置的k、b参数
        /// </summary>
        /// <param name="fname"></param>
        /// <param name="iridis"></param>
        /// <param name="pluseparm"></param>
        /// <param name="datasrc">true 仅加速度计，false 纵断面</param>
        private static void GenerateIRI_2(string fpath, int vallen, string fname, bool datasrc,int side)
        {
            const double DeltLen = 0.25;
            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);

            int qplusenum = Convert.ToInt32(DeltLen / 0.05);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            double[] oridata = new double[len];
            double[] toridata = new double[len];
            double[] iridata = new double[(len + qplusenum - 1) / qplusenum];
            int[] oritime = new int[len];
            int[] iritime = new int[(len + qplusenum - 1) / qplusenum];

            //计算IRI
            int plusenum = Convert.ToInt32(vallen / DeltLen);//IRI距离内有多少个250mm
            int iricnt = 0;
            double irisum = 0;
            double irival = 0;

            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split('\t');
                if (s.Length > 3)
                {
                    oridata[i] = double.Parse(s[2]);
                    toridata[i] = oridata[i];
                    oritime[i] = int.Parse(s[3]);
                }
            }
            //取5个点做均值滤波
            for (int i = 2; i < len - 2; ++i)
            {
                oridata[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
            }
            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                iridata[j] = oridata[i];
                iritime[j] = oritime[i];
            }
            len = iridata.Length;

            double speedval = 0;
            double stime;
            double etime;

            double YSU = 0;
            double[] ZSU = new double[4];
            double[] oldZSU = new double[4];
            oldZSU[0] = iridata[1] - iridata[0];
            oldZSU[1] = iridata[1] - iridata[0];
            oldZSU[2] = 0;
            oldZSU[3] = 0;

            if (m_speedtype[side] )
            {
                stime = iritime[0] * 1.0 / m_Frequency[side];
            }
            else
            {
                stime = (iritime[0] / 10000000) * 3600
                    + (iritime[0] / 100000) % 100 * 60
                    + (iritime[0] / 1000) % 100
                    + (iritime[0] % 1000) * 0.001;
            }

            string savefname = fpath.Replace(fname, string.Format("IRI_{0}m.txt", vallen));
            FileStream fw = new FileStream(savefname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            savefname = fpath.Replace(fname, string.Format("Speed_{0}m.txt", vallen));
            FileStream fwspeed = new FileStream(savefname, FileMode.Create);
            StreamWriter swspeed = new StreamWriter(fwspeed);
            for (int i = 1; i < len; ++i)
            {
                if (i % plusenum == 0)
                {
                    if (m_speedtype[side])
                    {
                        etime = iritime[i] * 1.0 / m_Frequency[side];
                    }
                    else
                    {
                        etime = (iritime[i] / 10000000) * 3600
                            + (iritime[i] / 100000) % 100 * 60
                            + (iritime[i] / 1000) % 100
                            + (iritime[i] % 1000) * 0.001;
                    }
                    speedval = vallen / (etime - stime) * 3.6;
                    stime = etime;

                    irival = irisum / plusenum;
                    if (datasrc)
                    {
                        irival = irival * _Setting.Acc_IRI_K_1 + _Setting.Acc_IRI_B_1;
                    }
                    else
                    {
                        irival = irival * _Setting.IRIk + _Setting.IRIb;
                    }
                    sw.WriteLine(string.Format("{0} {1}", ++iricnt, irival), Encoding.UTF8);
                    swspeed.WriteLine(string.Format("{0} {1}", iricnt, speedval), Encoding.UTF8);
                    irisum = 0;
                }

                YSU = (iridata[i] - iridata[i - 1]) / DeltLen;
                for (int zi = 0; zi < 4; ++zi)
                {
                    ZSU[zi] = 0;
                    for (int zj = 0; zj < 4; ++zj)
                    {
                        ZSU[zi] += SZU[zi * 4 + zj] * oldZSU[zj];
                    }
                    ZSU[zi] += PZU[zi] * YSU;
                }
                irisum += Math.Abs(ZSU[0] - ZSU[2]);

                for (int zi = 0; zi < 4; ++zi)
                {
                    oldZSU[zi] = ZSU[zi];
                }
            }
            sw.Close();
            fw.Close();
            swspeed.Close();
            fwspeed.Close();

            //250mm抽样
            savefname = fpath.Replace(fname, string.Format("ReSample250.txt", vallen));
            fw = new FileStream(savefname, FileMode.Create);
            sw = new StreamWriter(fw);
            len = sdata.Length;
            for (int i = 0; i < len; ++i)
            {
                if (i % qplusenum == 0)
                {
                    sw.WriteLine(sdata[i], Encoding.UTF8);
                }
            }
            sw.Close();
            fw.Close();
        }

        //计算路面跳车
        public static void ComputePB(string prj, int side, int vallen, ProjectInfo _ProjectInfo = null)
        {
            string fname = string.Format(@"{0}\IRIMTD\DAQ{1}\PavementBump.txt", prj, side);
            if (File.Exists(fname))
            {
                return;
            }

            fname = string.Format(@"{0}\IRIMTD\DAQ{1}\Resample.txt", prj, side);
            if (!File.Exists(fname))
            {
                if (!_ProjectInfo._IsJgAndGd)
                    MessageBox.Show("缺少路面纵断面文件：" + fname);
                return;
            }
            string[] datastrs = File.ReadAllLines(fname);

            //原始断面长度 50mm
            const int pluselen = 50;
            //计算跳车断面长度 100mm  0.1m计一个高程
            const int baselen = 100;

            int skipnum = baselen / pluselen;
            int valcnt = vallen * 1000 / baselen;

            int len = datastrs.Length;

            string[] tstrs;
            double[] hval = new double[(len + skipnum - 1) / skipnum];
            double[] toridata = new double[(len + skipnum - 1) / skipnum];
            for (int i = 0; i < len; ++i)
            {
                if (i % skipnum == 0)
                {
                    tstrs = datastrs[i].Split('\t');
                    if (tstrs.Length > 3)
                    {
                        hval[i / 2] = double.Parse(tstrs[2]);
                        toridata[i / 2] = double.Parse(tstrs[2]);

                    }
                }
            }

            len = hval.Length;
            //取5个点做均值滤波

            for (int i = 2; i < len - 2; ++i)
            {
                hval[i] = (toridata[i - 2] + toridata[i - 1] + toridata[i] + toridata[i + 1] + toridata[i + 2]) / 5;
            }

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

        /// <summary>
        /// 同步板采集的AD数据里面有时间跳秒，CheckGPSTime把跳秒的时间调整过来
        /// </summary>
        /// <param name="prj"></param>
        /// <param name="side"></param>
        private static void CheckGPSTime(string prj, int side)
        {
            string resamplefname = string.Format("{0}\\IRIMTD\\DAQ{1}\\resample.txt", prj, side);
            if (!File.Exists(resamplefname))
            {
                MessageBox.Show("遗失平整度数据！");
                return;
            }

            if (!File.Exists(resamplefname + ".bak"))
            {
                File.Copy(resamplefname, resamplefname + ".bak", false);
            }

            string[] oristr = null;
            if (File.Exists(resamplefname + ".bak"))
            {
                oristr = File.ReadAllLines(resamplefname + ".bak");
            }
            else
            {
                oristr = File.ReadAllLines(resamplefname);
            }

            int len = oristr.Length;
            if (len <= 2)
            {
                MessageBox.Show("平整度数据不合法，请检查是否是有效工程！");
                return;
            }

            string[] strs = oristr[0].Split('\t');
            UInt64 frameidx_old = UInt64.Parse(strs[3]);
            DateTime time_old = new DateTime(2021, 1, 1,
                int.Parse(strs[4].Substring(0, 2)),
                int.Parse(strs[4].Substring(2, 2)),
                int.Parse(strs[4].Substring(4, 2)),
                int.Parse(strs[4].Substring(6, 3)));

            UInt64 frameidx_new = 0;
            for (int i = 1; i < len; ++i)
            {
                strs = oristr[i].Split('\t');
                frameidx_new = UInt64.Parse(strs[3]);
                DateTime time_new = new DateTime(2021, 1, 1,
                    int.Parse(strs[4].Substring(0, 2)),
                    int.Parse(strs[4].Substring(2, 2)),
                    int.Parse(strs[4].Substring(4, 2)),
                    int.Parse(strs[4].Substring(6, 3)));

                TimeSpan usetime = time_new - time_old;
                double usetimeframe = (frameidx_new - frameidx_old) * 0.0005;
                double diff = usetime.TotalSeconds - usetimeframe;
                if (Math.Abs(diff) > 0.5)
                {
                    //出现跳秒了
                    if (time_new < time_old)
                    {
                        time_old = time_old.AddMilliseconds(usetimeframe);
                        usetime = time_old - time_new;
                        time_new = time_new.AddSeconds((long)Math.Round(usetime.TotalSeconds));
                        oristr[i] = string.Format("{0}\t{1}\t{2}\t{3}\t{4:HHmmssfff}",
                            strs[0], strs[1], strs[2], strs[3], time_new);
                    }
                    //丢数据了
                    else { }
                }

                time_old = time_new;
                frameidx_old = frameidx_new;
            }

            File.WriteAllLines(resamplefname, oristr, Encoding.UTF8);
        }
    }
}

