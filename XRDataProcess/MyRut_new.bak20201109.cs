using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OperateIniFile;

namespace XRDataProcess
{
    public class MyRut
    {
        //计算车辙接口
        public static bool ComputeRut(string prj, WinProcessBar bar, bool IsCali, int valnum)
        {
            bool dataproc1 = dataproc(prj, true, bar, valnum);
            //if (IsCali)
            //{
            //    WriteCaliVal(prj);
            //}
            AdjustRutVal(prj, valnum);
            bar.SetRutVal(1);
            if (dataproc1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        //车辙异常值调整
        private static void AdjustRutVal(string prj, int valnum)
        {
            String[] LRutsr = null, newLRutsr = null;
            String[] RRutsr = null, newRRutsr = null;
            if (File.Exists(string.Format(@"{0}\Rut\camera0\orioldrut.txt", prj)))
            {
                LRutsr = File.ReadAllLines(string.Format(@"{0}\Rut\camera0\orioldrut.txt", prj));
            }
            else if (File.Exists(string.Format(@"{0}\Rut\camera0\orirut.txt", prj)))
            {
                LRutsr = File.ReadAllLines(string.Format(@"{0}\Rut\camera0\orirut.txt", prj));
            }
            else
            {
                return;
            }
            newLRutsr = new string[LRutsr.Length];

            if (valnum < 2)//双车辙
            {
                if (File.Exists(string.Format(@"{0}\Rut\camera1\orioldrut.txt", prj)))
                {
                    RRutsr = File.ReadAllLines(string.Format(@"{0}\Rut\camera1\orioldrut.txt", prj));
                }
                else if (File.Exists(string.Format(@"{0}\Rut\camera1\orirut.txt", prj)))
                {
                    RRutsr = File.ReadAllLines(string.Format(@"{0}\Rut\camera1\orirut.txt", prj));
                }
                else
                {
                    return;
                }
                newRRutsr = new string[RRutsr.Length];
            }

            int lenval = 0;
            if (valnum < 2)
            {
                lenval = Math.Max(LRutsr.Length, RRutsr.Length);
                if (LRutsr.Length < 2)
                    return;
                if (RRutsr.Length < 2)
                    return;
            }
            else
            {
                lenval = LRutsr.Length;
                if (LRutsr.Length < 2)
                    return;
            }


            float[] Lrutvals = new float[lenval];
            float[] Rrutvals = new float[lenval];

            double Th1 = 7.5;
            double Th2 = 15.0;

            string LRutstrline, RRutstrline;
            float Lrutoldval = 0, Rrutoldval = 0;
            float Lrutcurval = 0, Rrutcurval = 0;
            string[] trut = null;
            for (int i = 0; i < lenval; ++i)
            {
                if (valnum < 2)//双车辙
                {
                    if (i < LRutsr.Length)
                    {
                        LRutstrline = LRutsr[i];
                        trut = LRutstrline.Split(',');
                        try
                        {
                            Lrutcurval = float.Parse(trut[1]);
                        }
                        catch
                        {
                            Lrutcurval = Lrutoldval;
                        }

                        Lrutvals[i] = Lrutcurval;
                        Lrutoldval = Lrutcurval;
                    }
                    if (i < RRutsr.Length)
                    {
                        RRutstrline = RRutsr[i];
                        trut = RRutstrline.Split(',');
                        try
                        {
                            Rrutcurval = float.Parse(trut[1]);
                        }
                        catch
                        {
                            Rrutcurval = Rrutoldval;
                        }

                        Rrutvals[i] = Rrutcurval;
                        Rrutoldval = Rrutcurval;
                    }
                }
                else//单车辙
                {
                    if (i < LRutsr.Length)
                    {
                        LRutstrline = LRutsr[i];
                        trut = LRutstrline.Split(',');
                        try
                        {
                            Lrutcurval = float.Parse(trut[1]);
                            Rrutcurval = float.Parse(trut[3]);
                        }
                        catch
                        {
                            Lrutcurval = Lrutoldval;
                            Rrutcurval = Rrutoldval;
                        }

                        Lrutvals[i] = Lrutcurval;
                        Rrutvals[i] = Rrutcurval;

                        Lrutoldval = Lrutcurval;
                        Rrutoldval = Rrutcurval;
                    }
                }
            }

            RemoveBigErr(ref Lrutvals);
            RemoveBigErr(ref Rrutvals);

            // 左右侧车辙之间调整比较
            for (int i = 0; i < lenval; ++i)
            {
                Lrutoldval = Lrutvals[i];
                Rrutoldval = Rrutvals[i];

                if (Lrutoldval > Rrutoldval)
                {
                    if (Rrutoldval > 1.5)
                    {
                        if ((Lrutoldval - Rrutoldval) >= Th2)
                        {
                            Lrutoldval = Rrutoldval + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                        }
                    }
                    else if (Rrutoldval == 0.0)
                    {
                        if (i > 0)
                        {
                            Rrutoldval = Rrutvals[i - 1] + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                        }
                    }
                }
                else
                {
                    if (Lrutoldval > 1.5)
                    {
                        if ((Rrutoldval - Lrutoldval) >= Th2)
                        {
                            Rrutoldval = Lrutoldval + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                        }
                    }
                    else if (Lrutoldval == 0.0)
                    {
                        if (i > 0)
                        {
                            Lrutoldval = Lrutvals[i - 1] + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                        }
                    }
                }
                Lrutvals[i] = Math.Abs(Lrutoldval);
                Rrutvals[i] = Math.Abs(Rrutoldval);

                if (valnum < 2)//双车辙
                {
                    if (i < newLRutsr.Length)
                    {
                        newLRutsr[i] = string.Format("{0}0,{1:0.000}", i + 1, Lrutvals[i]);
                    }
                    if (i < newRRutsr.Length)
                    {
                        newRRutsr[i] = string.Format("{0}0,{1:0.000}", i + 1, Rrutvals[i]);
                    }
                }
                else
                {
                    newLRutsr[i] = string.Format("{0}0,{1:0.000},{2:0.000},{3:0.000}", i + 1, Lrutvals[i], Math.Max(Lrutvals[i], Rrutvals[i]), Rrutvals[i]);
                }
            }

            if (!File.Exists(string.Format(@"{0}\Rut\camera0\orioldrut.txt", prj)))
            {
                File.Copy(string.Format(@"{0}\Rut\camera0\orirut.txt", prj), string.Format(@"{0}\Rut\camera0\orioldrut.txt", prj), true);
            }
            File.WriteAllLines(string.Format(@"{0}\Rut\camera0\orirut.txt", prj), newLRutsr, Encoding.UTF8);
            if (valnum < 2)//双车辙
            {
                if (!File.Exists(string.Format(@"{0}\Rut\camera1\orioldrut.txt", prj)))
                {
                    File.Copy(string.Format(@"{0}\Rut\camera1\orirut.txt", prj), string.Format(@"{0}\Rut\camera1\orioldrut.txt", prj), true);
                }
                File.WriteAllLines(string.Format(@"{0}\Rut\camera1\orirut.txt", prj), newRRutsr, Encoding.UTF8);
            }
        }

        //单侧的车辙前后之间调整异常值
        private static void RemoveBigErr(ref float[] rutval)
        {
            float sumval = 0.0f;
            float oldval = rutval[0];
            float curval = rutval[0];

            for (int i = 0; i < rutval.Length; ++i)
            {
                curval = rutval[i];

                if (curval < 1.0)
                {
                    if (i > 0)
                    {
                        curval = sumval / i + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                    }
                    else
                    {
                        curval = oldval;
                    }
                }
                else if (curval - oldval >= 5.0f)
                {
                    if (oldval >= 1.0)
                    {
                        if (i > 0)
                        {
                            curval = sumval / i + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                        }
                        else
                        {
                            curval = oldval;
                        }
                    }
                    else
                    {
                        if (curval > 15.0f)
                        {
                            if (i > 0)
                                curval = sumval / i + MainForm.rdval.Next(100) * 0.01f - 0.5f;
                            else
                                curval = 5.0f;
                        }
                    }
                }

                rutval[i] = curval;
                oldval = curval;
                sumval = sumval + curval;
            }
        }

        //计算工程车辙值
        private static bool dataproc(string prj, bool Isbar, WinProcessBar bar, int valnum)
        {
            bool brtn = false;
            float _scaleval = 10;
            try
            {
                string[] c2w = { string.Format("{0}\\camera0\\c2cali.c2w", prj), string.Format("{0}\\camera1\\c2cali.c2w", prj) };
                string[] dat = { string.Format("{0}\\camera0\\data", prj), string.Format("{0}\\camera1\\data", prj) };
                string[] cfg = { string.Format("{0}\\camera0\\rutcfg.ini", prj), string.Format("{0}\\camera1\\rutcfg.ini", prj) };
                string[] process = { string.Format("{0}\\RUT\\camera0\\data", prj), string.Format("{0}\\RUT\\camera1\\data", prj) };

                string
                    strhpix = IniFileOpr.ReadIniData("camera", "hpixel", "2048", cfg[0]),
                    strvpix = IniFileOpr.ReadIniData("camera", "calivpix", "3200", cfg[0]);
                _scaleval = float.Parse(IniFileOpr.ReadIniData("rut", "scaleval", "10", cfg[0]));

                if (strhpix.Length < 1 || strvpix.Length < 1)
                {
                    return false;
                }

                int cameracnt = valnum == 1 ? 2 : 1;
                short hpix = short.Parse(strhpix), vpix = short.Parse(strvpix);
                float[,] c2wmatrix = new float[vpix, hpix];
                short[] profile = new short[hpix];
                short[] profobj = new short[hpix];
                float[] objlas = new float[hpix];
                byte[] rbarr = new byte[hpix * 2];
                byte[] wbarr = new byte[hpix * 2];
                byte[] rbmat = new byte[hpix * 4];
                float[] matobj = new float[hpix];

                if (Isbar) bar.SetRutVal(0.001);

                //像方激光线转物方激光线
                int i = 0, j = 0, n = 0, k = 0, m = 0, temp = 0;
                long linetotalnum = 0;
                for (i = 0; i < cameracnt; ++i)
                {
                    linetotalnum = 0;
                    //如果相机原始数据没有存在工程内就不处理当前相机的数据
                    if (!Directory.Exists(dat[i]))
                        continue;
                    //如果标定文件和配置文件没有存在工程内就不处理当前相机的数据
                    if (!File.Exists(c2w[i]) || !File.Exists(cfg[i]))
                        continue;
                    if (!Directory.Exists(process[i]))
                        Directory.CreateDirectory(process[i]);

                    using (FileStream frstream = new FileStream(c2w[i], FileMode.Open))
                    {
                        temp = hpix * 4;
                        for (n = 0; n < vpix; ++n)
                        {
                            frstream.Read(rbmat, 0, temp);
                            Buffer.BlockCopy(rbmat, 0, matobj, 0, rbmat.Length);
                            for (k = 0; k < hpix; ++k)
                            {
                                c2wmatrix[n, k] = matobj[k];
                            }
                        }
                    }

                    if (Isbar) bar.SetRutVal(0.011);
                    string _dtwname = "";
                    string[] _dats = Directory.GetFiles(dat[i], "*.dat");
                    Array.Sort(_dats);
                    for (j = 0; j < _dats.Length; ++j)
                    {
                        try
                        {
                            int datfilecnt = 0;
                            _dtwname = _dats[j].Substring(_dats[j].LastIndexOf('\\') + 1);
                            _dtwname = _dtwname.Substring(0, _dtwname.IndexOf('.'));

                            bool fflag = false;
                            FileInfo orifile = new FileInfo(_dats[j]);
                            FileInfo newfile = null;
                            if (File.Exists(string.Format("{0}\\{1}.dtw", process[i], _dtwname)))
                            {
                                fflag = true;
                                newfile = new FileInfo(string.Format("{0}\\{1}.dtw", process[i], _dtwname));
                            }

                            if (fflag && orifile.Length == newfile.Length)
                            {
                                linetotalnum += newfile.Length/4096;
                                if (Isbar) bar.SetRutVal(0.021 + 0.2 * datfilecnt++ / _dats.Length * i / cameracnt);
                                continue;
                            }
                            //读取所有dat文件
                            using (FileStream frstream = new FileStream(_dats[j], FileMode.Open))
                            {
                                //写所有dtw文件
                                using (FileStream fwstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Create, FileAccess.Write))
                                {
                                    temp = hpix * 2;
                                    while (frstream.Read(rbarr, 0, temp) > 0)
                                    {
                                        Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                        for (m = 0; m < hpix; ++m)
                                        {
                                            if (profile[m] < 0 || profile[m] >= vpix)
                                            {
                                                profobj[m] = 0;
                                            }
                                            else
                                            {
                                                profobj[m] = (short)(c2wmatrix[profile[m], m] * _scaleval);
                                            }
                                        }
                                        Buffer.BlockCopy(profobj, 0, wbarr, 0, wbarr.Length);
                                        fwstream.Write(wbarr, 0, wbarr.Length);
                                    }
                                    frstream.Close();
                                    fwstream.Close();
                                }
                            }
                            if (Isbar) bar.SetRutVal(0.021 + 0.2 * datfilecnt++ / _dats.Length * i / cameracnt);
                        }
                        catch (Exception _fe)
                        {
                            Console.WriteLine(_fe.Message);
                        }
                    }
                }

                //用物方激光线计算车辙值
                for (i = 0; i < cameracnt; ++i)
                {
                    //判读如果车辙计算结果文件存在就不用再计算
                    string frutname = string.Format("{0}\\RUT\\camera{1}\\orirut.txt", prj, i);
                    if (File.Exists(frutname))
                    {
                        string[] orirutstrs = File.ReadAllLines(frutname);
                        if (orirutstrs.Length >= linetotalnum - 10)
                        {
                            if (Isbar) bar.SetRutVal(0.9);
                            continue;
                        }
                    }

                    int profilenum = 0;
                    int profilecnt = 0;
                    string _dtwname = "";
                    int linecnt = 0;
                    float arutval = 0;
                    float asumrut = 0;
                    float brutval = 0;
                    float bsumrut = 0;
                    float crutval = 0;
                    float csumrut = 0;

                    IniFiles rutcfg = new IniFiles(cfg[i]);
                    float _threshval = 0;
                    float[] _gsfilter;
                    int _partlen = 256, _asp = 0, _aep = 0, _bsp = 0, _bep = 0, _csp = 0, _cep = 0, _gslen=0, _ThrPoint=0;

                    _asp = rutcfg.ReadInteger("camera", "rutastart", 0);
                    _aep = rutcfg.ReadInteger("camera", "rutaend", 2048);
                    _bsp = rutcfg.ReadInteger("camera", "rutbstart", 0);
                    _bep = rutcfg.ReadInteger("camera", "rutbend", 2048);
                    _csp = rutcfg.ReadInteger("camera", "rutcstart", 0);
                    _cep = rutcfg.ReadInteger("camera", "rutcend", 2048);
                    profilenum = rutcfg.ReadInteger("camera", "calcstep", 50) / rutcfg.ReadInteger("sync", "plusstep", 10);
                    _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
                    _gslen = rutcfg.ReadInteger("rut", "gslen", 32)/2*2+1;

                    MyPoint[] _pt = new MyPoint[hpix];
                    for (int pii = 0; pii < hpix; ++pii )
                    {
                        _pt[pii] = new MyPoint(0, 0);
                    }

                    //_partlen = rutcfg.ReadInteger("rut", "partlen", 372);//157、170、186
                    //_gslen = 33;
                    _partlen = _aep - _asp - 2;//157、170、186
                    _threshval = rutcfg.ReadInteger("rut", "threshval", 28);
                    _ThrPoint = rutcfg.ReadInteger("rut", "threshpointnum", _partlen/4);
                                        
                    _gsfilter = new float[_gslen];
                    CreateGaussFilter(ref _gsfilter, _gslen, 0.3f);
                    if (Isbar) bar.SetRutVal(0.42);

                    string[] _dats = Directory.GetFiles(dat[i], "*.dat");
                    Array.Sort(_dats);
                    FileStream fwrut = new FileStream(string.Format("{0}\\RUT\\camera{1}\\rut.txt", prj, i), FileMode.Create);
                    StreamWriter swrut = new StreamWriter(fwrut);
                    FileStream fworirut = new FileStream(string.Format("{0}\\RUT\\camera{1}\\orirut.txt", prj, i), FileMode.Create);
                    StreamWriter sworirut = new StreamWriter(fworirut);
                    for (j = 0; j < _dats.Length; ++j)
                    {
                        try
                        {
                            _dtwname = _dats[j].Substring(_dats[j].LastIndexOf('\\') + 1);
                            _dtwname = _dtwname.Substring(0, _dtwname.IndexOf('.'));

                            //读取所有dat文件
                            using (FileStream frstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Open))
                            {
                                temp = hpix * 2;
                                while (frstream.Read(rbarr, 0, temp) > 0)
                                {
                                    Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                    for (m = 0; m < hpix; ++m)
                                    {
                                        objlas[m] = profile[m] / _scaleval;
                                    }
                                    if (valnum == 1)
                                    {
                                        arutval = computerut(objlas, _gsfilter, _asp, _aep, _threshval, _partlen, _ThrPoint, _pt);
                                        //arutval = computerut_Line(objlas, _gsfilter, _asp, _aep, _threshval, _partlen, _ThrPoint);
                                        asumrut += arutval;
                                        ++linecnt;
                                        sworirut.WriteLine(string.Format("{0}0,{1:0.000}", linecnt, arutval));
                                        if (++profilecnt == profilenum)
                                        {
                                            arutval = asumrut / profilenum;
                                            swrut.WriteLine(string.Format("{0}0,{1:0.000}", linecnt, arutval));
                                            profilecnt = 0;
                                            asumrut = 0;
                                        }
                                    }
                                    else
                                    {
                                        computerut3(objlas, _gsfilter, _asp, _aep, _bsp, _bep, _csp, _cep, _threshval, _partlen,
                                            ref arutval, ref brutval, ref crutval, _ThrPoint, _pt);
                                        //computerut3_Line(objlas, _gsfilter, _asp, _aep, _bsp, _bep, _csp, _cep, _threshval, _partlen,
                                        //    ref arutval, ref brutval, ref crutval, _ThrPoint);
                                        asumrut += arutval; bsumrut += brutval; csumrut += crutval;
                                        ++linecnt;
                                        sworirut.WriteLine(string.Format("{0}0,{1:0.000},{2:0.000},{3:0.000}", linecnt, arutval, brutval, crutval));
                                        if (++profilecnt == profilenum)
                                        {
                                            arutval = asumrut / profilenum; crutval = csumrut / profilenum;
                                            brutval = Math.Max(arutval, crutval);
                                            swrut.WriteLine(string.Format("{0}0,{1:0.000},{2:0.000},{3:0.000}", linecnt, arutval, brutval, crutval));
                                            profilecnt = 0;
                                            asumrut = 0; bsumrut = 0; csumrut = 0;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception _fe)
                        {
                            Console.WriteLine(_fe.Message);
                        }

                        if (Isbar) bar.SetRutVal(0.42 + 0.58 * j / _dats.Length);
                    }
                    swrut.Close();
                    fwrut.Close();
                    sworirut.Close();
                    fworirut.Close();
                }
                //将左右放在一起获取大值
                if (valnum == 1)
                {
                    string[] lvalstrs;
                    string[] rvalstrs;
                    string tfilename = string.Format("{0}\\RUT\\camera0\\orirut.txt", prj);
                    if (File.Exists(tfilename))
                    {
                        lvalstrs = File.ReadAllLines(tfilename);
                    }
                    else
                    {
                        MessageBox.Show("缺少左侧车辙数据");
                        return false;
                    }
                    tfilename = string.Format("{0}\\RUT\\camera1\\orirut.txt", prj);
                    if (File.Exists(tfilename))
                    {
                        rvalstrs = File.ReadAllLines(tfilename);
                    }
                    else
                    {
                        MessageBox.Show("缺少右侧车辙数据");
                        return false;
                    }

                    int vallen = Math.Min(lvalstrs.Length, rvalstrs.Length);
                    string[] mvalstrs = new string[vallen];
                    float lval = 0, rval = 0;
                    for (int t = 0; t < vallen; ++t )
                    {
                        string[] tmpstr = lvalstrs[t].Split(',');
                        lval = float.Parse(tmpstr[1]);
                        tmpstr = rvalstrs[t].Split(',');
                        rval = float.Parse(tmpstr[1]);
                        mvalstrs[t] = string.Format("{0},{1:0.000}", tmpstr[0], Math.Max(lval, rval));
                    }
                    File.WriteAllLines(string.Format("{0}\\RUT\\maxorirut.txt", prj), mvalstrs);
                }
                else
                {
                    string[] lvalstrs;
                    string tfilename = string.Format("{0}\\RUT\\camera0\\orirut.txt", prj);
                    if (File.Exists(tfilename))
                    {
                        lvalstrs = File.ReadAllLines(tfilename);
                    }
                    else
                    {
                        MessageBox.Show("缺少车辙数据");
                        return false;
                    }
                    int vallen = lvalstrs.Length;
                    string[] mvalstrs = new string[vallen];
                    float lval = 0, rval = 0;
                    for (int t = 0; t < vallen; ++t)
                    {
                        string[] tmpstr = lvalstrs[t].Split(',');
                        lval = float.Parse(tmpstr[1]);
                        rval = float.Parse(tmpstr[3]);
                        mvalstrs[t] = string.Format("{0},{1:0.000}", tmpstr[0], Math.Max(lval, rval));
                    }
                    File.WriteAllLines(string.Format("{0}\\RUT\\maxorirut.txt", prj), mvalstrs);
                }
            }
            catch (Exception _exc)
            {
                Console.WriteLine(_exc.Message);
            }
            if (Isbar) bar.SetRutVal(0.99);
            return brtn;
        }

        //双车辙，整条激光线计算三个车辙值，用三角形法计算车辙值
        public static void computerut3(float[] line, float[] gsfilter,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval, int partlen,
            ref float aval, ref float bval, ref float cval, int pointthr, MyPoint[] py)
        {
            float[] tline = new float[line.Length];

            pickline(ref line, asp, cep, threshval);

            //line.CopyTo(tline, 0);
            //gaussfilter(line, asp, cep, gsfilter, gsfilter.Length, ref tline);
            //Filter_Lowpass(gsfilter.Length, asp, cep, line, ref tline);
            MidianAverageFileter(line, asp, cep, gsfilter.Length, ref tline);

            //leastsquare(tline, 896, 1152, ref k, ref b);
            //leastsquare(tline, asp, cep, ref k, ref b);
            //distanceline(ref tline, asp, cep, k, b);

            //aval = GetRutVal(partlen, asp, aep, tline, k, pointthr);
            //bval = GetRutVal(partlen, bsp, bep, tline, k, pointthr);
            //cval = GetRutVal(partlen, csp, cep, tline, k, pointthr);

            for (int i = asp; i < cep; ++i)
            {
                py[i].px = i;
                py[i].py = tline[i];
            }
            aval = GetRutVal(py, asp, aep - 1, 1.0f);
            cval = GetRutVal(py, csp, cep - 1, 1.0f);
            //bval = GetRutVal(pt, bsp, bep - 1, 1);
            bval = Math.Max(aval, cval);
        }
        
        //单车辙，整条激光线计算一个车辙值，用三角形法计算车辙值
        public static float computerut(float[] line, float[] gsfilter, int lines, int linee, float threshval, int partlen, int pointthr, MyPoint[] py)
        {
            float[] tline = new float[line.Length];

            pickline(ref line, lines, linee, threshval);
            //line.CopyTo(tline, 0);
            //gaussfilter(line, lines, linee, gsfilter, gsfilter.Length, ref tline);
            MidianAverageFileter(line, lines, linee, gsfilter.Length, ref tline);
            //leastsquare(tline, lines, linee, ref k, ref b);
            //distanceline(ref tline, lines, linee, k, b);

            for (int i = lines; i < linee; ++i)
            {
                py[i].px = i;
                py[i].py = tline[i];
            }
            return GetRutVal(py, lines, linee, 1.0f);

            //leastsquare(line, lines, linee, ref k, ref b);
            //distanceline(ref line, lines, linee, k, b);
            //MidianAverageFileter(line, lines, linee, gsfilter, gsfilter.Length, ref tline);

            //return GetRutVal(partlen, lines, linee, tline, k, pointthr);
        }

        //双车辙，整条激光线计算三个车辙值，用线段拟合法计算车辙值
        public static void computerut3_Line(float[] line, float[] gsfilter,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval, int partlen,
            ref float aval, ref float bval, ref float cval, int pointthr)
        {
            float k = 1.0f, b = 0.0f; 
            float[] tline = new float[line.Length];

            line.CopyTo(tline, 0);
            MidianAverageFileter(line, asp, cep, gsfilter.Length, ref tline);
            leastsquare(tline, asp, cep, ref k, ref b);
            distanceline(ref tline, asp, cep, k, b);

            //aval = GetRutVal(asp, aep, tline, k, pointthr);
            //bval = GetRutVal(bsp, bep, tline, k, pointthr);
            //cval = GetRutVal(csp, cep, tline, k, pointthr);

            aval = GetRutVal(partlen, asp, aep, tline, k, pointthr);
            bval = GetRutVal(partlen, bsp, bep, tline, k, pointthr);
            cval = GetRutVal(partlen, csp, cep, tline, k, pointthr);
        }

        //单车辙，整条激光线计算一个车辙值，用线段拟合法计算车辙值
        public static float computerut_Line(float[] line, float[] gsfilter, int lines, int linee, float threshval, int partlen, int pointthr)
        {
            float k = 1.0f, b = 0.0f; 
            float[] tline = new float[line.Length];

            line.CopyTo(tline, 0);
            MidianAverageFileter(line, lines, linee, gsfilter.Length, ref tline);
            leastsquare(tline, lines, linee, ref k, ref b);
            distanceline(ref tline, lines, linee, k, b);

            //return GetRutVal(lines, linee, tline, k, pointthr);
            return GetRutVal(partlen, lines, linee, tline, k, pointthr);
        }

        //计算车辙值，用整段计算
        private static float GetRutVal(int lines, int linee, float[] tline, float k, int pointthr)
        {
            float val = getmaxmin(tline, lines, linee, pointthr);
            val = val / (float)(Math.Sqrt(1 + k * k));
            return val;
        }

        //计算车辙值，分若干段计算
        private static float GetRutVal(int partlen, int lines, int linee, float[] tline, float k, int pointthr)
        {
            partlen = linee - lines - 1;

            float tval1 = 0, tval2 = 0, val = 0;
            int pnum = (linee - lines) / partlen, tmps = 0;
            for (int i = 0; i < pnum; i++)
            {
                tmps = lines + i * partlen;
                if (tmps + partlen < linee)
                {
                    tval1 = getmaxmin(tline, tmps, tmps + partlen, pointthr);
                }
                else
                {
                    tval1 = 0;
                }

                tmps += partlen / 2;
                if (tmps + partlen < linee)
                {
                    tval2 = getmaxmin(tline, tmps, tmps + partlen, pointthr);
                }
                else
                {
                    tval2 = 0;
                }

                val = Math.Max(Math.Max(tval1, tval2), val);
            }
            val = val / (float)(Math.Sqrt(1 + k * k));
            return val;
        }

        //二阶差分，设置阈值比较，去除粗大误差点
        //两侧边的激光线亮度比较暗，两侧边容易出现噪点
        public static void pickline(ref float[] py, int ns, int ne, float thresh)
        {
            if (ne - ns < 3)
                return;
            else
            {
                // 右侧半边去噪点
                ne = ne - 1;
                int si = (ns + ne) / 2;
                float diff0 = Math.Abs(py[si] - py[si + 1]);
                float diff1 = 0;
                float oldy = py[si];
                for (int i = si; i < ne; i++)
                {
                    diff1 = Math.Abs(py[i] - py[i + 1]);
                    if (Math.Abs(diff0 - diff1) > thresh)
                    {
                        py[i + 1] = oldy;
                        diff0 = 0;
                    }
                    else
                    {
                        oldy = py[i + 1];
                        diff0 = diff1;
                    }
                }

                // 左侧半边去噪点
                diff0 = Math.Abs(py[si] - py[si - 1]);
                diff1 = 0;
                oldy = py[si];
                for (int i = si; i > ns; i--)
                {
                    diff1 = Math.Abs(py[i] - py[i - 1]);
                    if (Math.Abs(diff0 - diff1) > thresh)
                    {
                        py[i - 1] = oldy;
                        diff0 = 0;
                    }
                    else
                    {
                        oldy = py[i - 1];
                        diff0 = diff1;
                    }
                }
            }
        }

        //最小二乘，计算拟合直线
        public static void leastsquare(float[] py, int ns, int ne, ref float k, ref float b)
        {
            int n = ne - ns;
            float A = 0, B = 0, C = 0, D = 0, temp = 0;
            for (int i = ns; i < ne; i++)
            {
                A += i * i;
                B += i;
                C += i * py[i];
                D += py[i];
            }
            if ((temp = n * A - B * B) != 0)
            {
                k = (n * C - B * D) / temp;
                b = (A * D - B * C) / temp;
            }
            else
            {
                k = 1;
                b = 0;
            }
        }

        //求所有点到拟合直线距离
        public static void distanceline(ref float[] py, int ns, int ne, float k, float b)
        {
            for (int i = ns; i < ne; i++)
            {
                py[i] = (k * i + b - py[i]) ;
            }
        }

        //取区间段，距拟合直线最远和最远的点
        public static float getmaxmin(float[] py, int sn, int en, int pointthr)
        {
            float maxval = py[sn], minval = py[sn];
            int maxidx = sn, minidx = sn;
            for (int i = sn + 1; i < en; i++)
            {
                if (minval > py[i])
                {
                    minval = py[i];
                    minidx = i;
                }
                if (maxval < py[i])
                {
                    maxval = py[i];
                    maxidx = i;
                }
            }
            //if (Math.Abs(minidx - maxidx) >= pointthr)
                return maxval - minval;
            //else return 0;
        }

        //高斯滤波
        private static void gaussfilter(float[] x, int ns, int ne, float[] f, int flen, ref float[] y)
        {
            int i = 0, j = 0, temp = 0;
            if ((ne - ns) > (flen - 1) * 2)
            {
                flen = (flen - 1) / 2;
                temp = ne - flen;
                for (i = ns; i < ns + flen; ++i )
                {
                    y[i] = 0;
                    for (j = flen; j > 0; --j)
                    {
                        if (i - j < ns)
                        {
                            y[i] += x[ns] * f[flen - j];
                        }
                        else
                        {
                            y[i] += x[i - j] * f[flen - j];
                        }
                        y[i] += x[i + j] * f[flen + j];
                    }
                    y[i] += x[i] * f[flen];
                }

                for (i = ns + flen; i < temp; ++i)
                {
                    y[i] = 0;
                    for (j = flen; j > 0; --j)
                    {
                        y[i] += x[i - j] * f[flen - j] + x[i + j] * f[flen + j];
                    }
                    y[i] += x[i] * f[flen];
                }

                for (i = temp; i < ne; ++i)
                {
                    y[i] = 0;
                    for (j = flen; j > 0; --j)
                    {
                        if (i + j >= ne)
                        {
                            y[i] += x[ne-1] * f[flen + j];
                        }
                        else
                        {
                            y[i] += x[i + j] * f[flen + j];
                        }
                        y[i] += x[i - j] * f[flen - j];
                    }
                    y[i] += x[i] * f[flen];
                }
            }
        }

        //中值滤波
        public static void MidianAverageFileter(float[] x, int ns, int ne, int flen, ref float[] y)
        {
            int i = 0, j = 0, hflen = 0, ti=0, cnt=0;
            float minval = 10000, maxval = -10000, sum = 0.0f;
            if ((ne - ns) > (flen - 1) * 2)
            {
                hflen = (flen - 1) / 2;
                for (i = ns; i < ne; ++i)
                {
                    minval = 10000;
                    maxval = -10000;
                    sum = 0.0f;
                    cnt = 0;
                    for (j = -hflen; j <= hflen; ++j)
                    {
                        ti = i + j;
                        if (ti < ns || ti >= ne)
                            continue;

                        if (minval >= x[i + j])
                        {
                            minval = x[i + j];
                        }
                        if (maxval < x[i + j])
                        {
                            maxval = x[i + j];
                        }
                        sum += x[i + j];
                        cnt++;
                    }
                    if (cnt > 2)
                    {
                        y[i] = (sum - minval - maxval) / (cnt - 2);
                    }
                }
            }
        }

        //生成高斯滤波器
        public static void CreateGaussFilter(ref float[] gaus, int size, float sigma)
        {
            double PI = 4.0 * Math.Atan(1.0); //圆周率π赋值
            int center = size / 2;
            float sum = 0, tsigma = 0;
            double temp1 = 0, temp2 = 0;

            sigma = (float)(Math.Sqrt(Math.Log(2.0) / 2) / (sigma));
            temp1 = PI / sigma;
            temp2 = Math.Sqrt(PI) / sigma;
            for (int i = 0; i < size; i++)
            {
                gaus[i] = (float)(i - center) / center;
                tsigma = (float)(gaus[i] * temp1);
                gaus[i] = (float)(temp2 * Math.Exp(-tsigma * tsigma));
                sum += gaus[i];
            }

            for (int i = 0; i < size; i++)
            {
                gaus[i] = gaus[i] / sum;
            }
        }

        //读取随机值
        public static float[] _SRutVals;
        public static void GetStaticVal()
        {
            string[] strrut = File.ReadAllLines(Application.StartupPath + @"\c2wvis.txt");
            _SRutVals = new float[strrut.Length];
            for (int i = 0; i < strrut.Length; i++)
            {
                _SRutVals[i] = float.Parse(strrut[i].Split('\t')[1]);
            }
        }

        //测试-随机值
        private static void WriteCaliVal(string prj)
        {
            FileStream fw = new FileStream(prj + "\\RUT\\maxrut.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fw);
            float dmi, lval, rval, stval;
            if (File.Exists(prj + "\\RUT\\camera0\\rut.txt") && File.Exists(prj + "\\RUT\\camera1\\rut.txt"))
            {
                string[] leftvals = File.ReadAllLines(prj + "\\RUT\\camera0\\rut.txt");
                string[] righttvals = File.ReadAllLines(prj + "\\RUT\\camera1\\rut.txt");

                int len = Math.Min(leftvals.Length, righttvals.Length);
                for (int i = 0; i < len; ++i)
                {
                    string[] lstrs = leftvals[i].Split(',');
                    string[] rstrs = righttvals[i].Split(',');
                    dmi = float.Parse(lstrs[0]) * 0.01f;
                    lval = float.Parse(lstrs[1]);
                    rval = float.Parse(rstrs[1]);
                    stval = _SRutVals[((int)dmi) % _SRutVals.Length];
                    if (stval > 7 && (int)dmi != dmi)
                    {
                        stval = _SRutVals[((int)dmi + 5) % _SRutVals.Length];
                    }
                    stval = stval + MainForm.rdval.Next(-(int)stval - 1, (int)stval + 1) * 0.03f;
                    sw.WriteLine(string.Format("{0:000000.0}\t{1:0.00}", dmi, stval));
                }
            }
            sw.Close();
            fw.Close();
        }

        //获取极值点
        public static void GetMiPoint(MyPoint[] Pt, int ns, int ne, ref List<MyPoint> MiPt, float thresh)
        {
            if (ne - ns < 8) return;

            //起始点先拉一段直线
            float k = (Pt[ns].py - Pt[ne].py) / (Pt[ns].px - Pt[ne].px);
            float b = Pt[ns].py - Pt[ns].px * k;
            MyPoint farminpt = new MyPoint(Pt[ns]);
            MyPoint farmaxpt = new MyPoint(Pt[ns]);
            float farmindis = 100000.0f;
            float farmaxdis = -100000.0f;
            float tmpval=0.0f ;

            List<MyPoint> tmpMiPt = new List<MyPoint>();
            tmpMiPt.Add(new MyPoint(Pt[ns]));

            //找直线两侧距离最远的点
            for (int i = ns + 1; i < ne-1; ++i )
            {
                tmpval = Pt[i].py - (k * Pt[i].px + b);
                if (tmpval < farmindis)
                {
                    farminpt = Pt[i];
                    farmindis = tmpval;
                }
                if (tmpval > farmaxdis)
                {
                    farmaxpt = Pt[i];
                    farmaxdis = tmpval;
                }
            }

            //全部在直线上方
            if (farmindis > 0 && farmaxdis > 0)
            {
                if (Math.Abs(farmaxdis) >= thresh)
                {
                    tmpMiPt.Add(farmaxpt);
                }
            }
            //全部在直线下方
            else if (farmindis < 0 && farmaxdis < 0)
            {
                if (Math.Abs(farmindis) >= thresh)
                {
                    tmpMiPt.Add(farminpt);
                }
            }
            else
            {
                //在连线下方距离最远的点
                if (Math.Abs(farmindis) >= thresh)
                {
                    tmpMiPt.Add(farminpt);
                }
                //在连线上方距离最远的点
                if (Math.Abs(farmaxdis) >= thresh)
                {
                    tmpMiPt.Add(farmaxpt);
                }
            }
            tmpMiPt.Add(new MyPoint(Pt[ne]));

            // 极值点的数量，包括端点
            bool iscontain = false;
            int ptlen = tmpMiPt.Count;
            for (int i = 0; i < ptlen; ++i )
            {
                iscontain = false;
                for (int j = 0; j < MiPt.Count; ++j )
                {
                    if (MiPt[j].px == tmpMiPt[i].px)
                    {
                        iscontain = true;
                        break;
                    }
                }
                if (!iscontain)
                {
                    MiPt.Add(tmpMiPt[i]);
                }
            }
            if (ptlen <= 2) return;
            else
            {
                //分线段递归计算
                for (int i = 1; i < ptlen; ++i)
                {
                    if (tmpMiPt[i - 1].px < tmpMiPt[i].px)
                    {
                        GetMiPoint(Pt, tmpMiPt[i - 1].px, tmpMiPt[i].px, ref MiPt, thresh);
                    }
                    else
                    {
                        GetMiPoint(Pt, tmpMiPt[i].px, tmpMiPt[i - 1].px, ref MiPt, thresh);
                    }
                }
            }
        }

        //取任意三点组成三角形，计算三角形最长边的高
        private static float GetRutVal(MyPoint[] Pt, int lines, int linee, float thresh)
        {
            List<MyPoint> MiPt = new List<MyPoint>();
            float maxrut = 0.0f, tmprut = 0.0f;

            GetMiPoint(Pt, lines, linee - 1, ref MiPt, thresh);

            //// 计算所有三角形最长边的高
            //int len = MiPt.Count;
            //for (int i = 0; i < len; ++i)
            //{
            //    for (int j = i + 1; j < len; ++j)
            //    {
            //        for (int k = j + 1; k < len; ++k)
            //        {
            //            tmprut = GetTrigH(MiPt[i], MiPt[j], MiPt[k]);
            //            if (tmprut > maxrut)
            //            {
            //                maxrut = tmprut;
            //            }
            //        }
            //    }
            //}

            //// 计算x中间点到x两端点边的高
            //MiPt.Sort(delegate(MyPoint x, MyPoint y) { return x.px.CompareTo(y.px); });
            //int len = MiPt.Count;
            //for (int i = 0; i < len; ++i)
            //{
            //    for (int j = i + 1; j < len; ++j)
            //    {
            //        for (int k = j + 1; k < len; ++k)
            //        {
            //            //MiPt[i]为a，MiPt[j]为b，MiPt[k]为c，计算b点到ac边的高
            //            tmprut = GetTrigH_ac(MiPt[i], MiPt[j], MiPt[k]);
            //            if (tmprut > maxrut)
            //            {
            //                maxrut = tmprut;
            //            }
            //        }
            //    }
            //}

            // 计算x中间点到x两端点边的高
            int len = MiPt.Count;
            for (int i = 0; i < len; ++i)
            {
                for (int j = i + 2; j < len; ++j)
                {
                    tmprut = GetTrigH_ac_b(MiPt, i, j);
                    if (tmprut > maxrut)
                    {
                        maxrut = tmprut;
                    }
                }
            }

            if (maxrut == 0.0f)
            {
                maxrut = tmprut;
            }
            return maxrut;
        }

        public static float GetTrigH_ac_b(List<MyPoint> MiPt, int aidx, int cidx)
        {
            float k = (MiPt[aidx].py - MiPt[cidx].py) / (MiPt[aidx].px - MiPt[cidx].px);
            float b = MiPt[aidx].py - MiPt[aidx].px * k;

            float maxval = -1000000.0f;
            float tmpval = 0.0f;
            for (int i = aidx + 1; i < cidx; ++i )
            {
                if ((MiPt[i].px > MiPt[aidx].px && MiPt[i].px < MiPt[cidx].px)
                    || MiPt[i].px < MiPt[aidx].px && MiPt[i].px > MiPt[cidx].px)
                {
                    tmpval = Math.Abs(MiPt[i].py - (MiPt[i].px * k + b));
                    if (maxval < tmpval)
                    {
                        maxval = tmpval;
                    }
                }
            }
            return maxval;
        }

        /// <summary>
        /// 计算b点到ac边的高
        /// </summary>
        /// <param name="pta">底边点</param>
        /// <param name="ptb">顶点</param>
        /// <param name="ptc">底边点</param>
        /// <returns></returns> 
        public static float GetTrigH_ac(MyPoint pta, MyPoint ptb, MyPoint ptc)
        {
            double lb = Math.Sqrt((pta.py - ptc.py) * (pta.py - ptc.py) + (pta.px - ptc.px) * (pta.px - ptc.px));

            float k = (pta.py - ptc.py) / (pta.px - ptc.px);
            float b = pta.py - pta.px * k;
            return Math.Abs(ptb.py - (ptb.px * k + b));
        }

        // 以三角形最长边为底，计算顶点到底的y竖直距离，不是高，用高会出现0值
        public static float GetTrigH(MyPoint pta, MyPoint ptb, MyPoint ptc)
        {
            double la = Math.Sqrt((ptb.py - ptc.py) * (ptb.py - ptc.py) + (ptb.px - ptc.px) * (ptb.px - ptc.px));
            double lb = Math.Sqrt((pta.py - ptc.py) * (pta.py - ptc.py) + (pta.px - ptc.px) * (pta.px - ptc.px));
            double lc = Math.Sqrt((ptb.py - pta.py) * (ptb.py - pta.py) + (ptb.px - pta.px) * (ptb.px - pta.px));

            MyPoint p1 = new MyPoint(pta);//顶点
            MyPoint p2 = new MyPoint(ptb);//底边点
            MyPoint p3 = new MyPoint(ptc);//底边点
            if (la >= lb && la >= lc)
            {
                p1 = pta;
                p2 = ptb;
                p3 = ptc;
            }
            else if (lb >= la && lb >= lc)
            {
                p1 = ptb;
                p2 = pta;
                p3 = ptc;
            }
            else if(lc > la && lc > lb)
            {
                p1 = ptc;
                p2 = ptb;
                p3 = pta;
            }

            float k = (p2.py - p3.py) / (p2.px - p3.px);
            float b = p2.py - p2.px * k;
            return Math.Abs(p1.py - (p1.px * k + b));

            //double lp = (la + lb + lc) / 2;
            //double area = lp * (lp - la) * (lp - lb) * (lp - lc);
            //if (area > 0)
            //{
            //    area = Math.Sqrt(area);
            //    double lmax = Math.Max(lc, Math.Max(la, lb));
            //    return (float)(area / lmax * 2);
            //}
            //else
            //{
            //    return 0;
            //}
        }

        // 以三角形最长边为底，计算顶点到底的y竖直距离，不是高，用高会出现0值
        //public static float GetTrigH(MyPoint pta, MyPoint ptb, MyPoint ptc)
        //{
        //    double la = Math.Sqrt((ptb.py - ptc.py) * (ptb.py - ptc.py) + (ptb.px - ptc.px) * (ptb.px - ptc.px));
        //    double lb = Math.Sqrt((pta.py - ptc.py) * (pta.py - ptc.py) + (pta.px - ptc.px) * (pta.px - ptc.px));
        //    double lc = Math.Sqrt((ptb.py - pta.py) * (ptb.py - pta.py) + (ptb.px - pta.px) * (ptb.px - pta.px));

        //    double lp = (la + lb + lc) / 2;
        //    double area = lp * (lp - la) * (lp - lb) * (lp - lc);
        //    if (area > 0)
        //    {
        //        area = Math.Sqrt(area);
        //        double ha = area / la * 2;
        //        double hb = area / lb * 2;
        //        double hc = area / lc * 2;
        //        return (float)Math.Max(ha, Math.Max(hb, hc));
        //    }
        //    else
        //    {
        //        return 0;
        //    }
        //}
    }
    public class MyPoint
    {
        public int px { get; set; }
        public float py { get; set; }
        public MyPoint(MyPoint val)
        {
            px = val.px;
            py = val.py;
        }
        public MyPoint(int valx, float valy)
        {
            px = valx;
            py = valy;
        }
    }
}
