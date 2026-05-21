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

    public class MyRut
    {
        static XRSetting _Setting = XRSetting.GetInstance();

        //计算车辙接口
        public static bool ComputeRut(string prj, WinProcessBar bar, int valnum, int rutmode)
        {
            bool dataproc1 = dataproc(prj, true, bar, valnum, rutmode);
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

                        if (_Setting.IsThresholdRut)
                        {
                            if (Lrutcurval > 0)
                            {
                                Lrutcurval = Lrutcurval / (int)Math.Ceiling(Lrutcurval / _Setting.ErrorRut);
                            }
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

                        if (_Setting.IsThresholdRut)
                        {
                            if (Rrutcurval > 0)
                            {
                                Rrutcurval = Rrutcurval / (int)Math.Ceiling(Rrutcurval / _Setting.ErrorRut);
                            }
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

                        if (_Setting.IsThresholdRut)
                        {
                            if (Lrutcurval > 0)
                            {
                                Lrutcurval = Lrutcurval / (int)Math.Ceiling(Lrutcurval / _Setting.ErrorRut);
                            }
                            if (Rrutcurval > 0)
                            {
                                Rrutcurval = Rrutcurval / (int)Math.Ceiling(Rrutcurval / _Setting.ErrorRut);
                            }
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
                    if (Rrutoldval > 1.0)
                    {
                        if ((Lrutoldval - Rrutoldval) >= _Setting.ErrorRutTh1)
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
                    if (Lrutoldval > 1.0)
                    {
                        if ((Rrutoldval - Lrutoldval) >= _Setting.ErrorRutTh1)
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

                if (curval < 0.1)
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
                else if (curval - oldval >= _Setting.ErrorRutTh2)
                {
                    if (oldval >= 0.1)
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
                        if (curval > _Setting.ErrorRutTh2)
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

        private static bool isouttxt = false;

        //计算工程车辙值
        private static bool dataproc(string prj, bool Isbar, WinProcessBar bar, int valnum, int rutmode)
        {
            float[,] c2wmatrix = null;
            float[] matobj = null;
            double[] matobj3d = null;
            byte[] rbmat = null;

            int[] MP_idx = new int[20];
            float[] MP_val = new float[20];
            float lineK = 0.0f;

            try
            {
                string[] c2w = { string.Format("{0}\\camera0\\c2cali.c2w", prj), string.Format("{0}\\camera1\\c2cali.c2w", prj) };
                string[] dat = { string.Format("{0}\\camera0\\data", prj), string.Format("{0}\\camera1\\data", prj) };
                string[] cfg = { string.Format("{0}\\camera0\\rutcfg.ini", prj), string.Format("{0}\\camera1\\rutcfg.ini", prj) };
                string[] process = { string.Format("{0}\\RUT\\camera0\\data", prj), string.Format("{0}\\RUT\\camera1\\data", prj) };

                short hpix = short.Parse(IniFileOpr.ReadIniData("camera", "hpixel", "2048", cfg[0]));
                short vpix = short.Parse(IniFileOpr.ReadIniData("camera", "calivpix", "3200", cfg[0]));
                float _scaleval = float.Parse(IniFileOpr.ReadIniData("rut", "scaleval", "10", cfg[0]));
                short basezval = short.Parse(IniFileOpr.ReadIniData("rut", "basezval", "0", cfg[0]));

                int cameracnt = valnum == 1 ? 2 : 1;
                if (rutmode == 2)
                {
                    matobj3d = new double[hpix];
                    rbmat = new byte[hpix * 8];
                }
                else
                {
                    matobj = new float[hpix];
                    rbmat = new byte[hpix * 4];
                }
                c2wmatrix = new float[vpix, hpix];

                short[] profile = new short[hpix];
                short[] profobj = new short[hpix];
                float[] objlas = new float[hpix];
                float[] tobjlas = new float[hpix];
                byte[] rbarr = new byte[hpix * 2];
                byte[] wbarr = new byte[hpix * 2];

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
                        if (rutmode == 2)
                        {
                            temp = hpix * 8;
                            for (n = 0; n < vpix; ++n)
                            {
                                frstream.Read(rbmat, 0, temp);
                                Buffer.BlockCopy(rbmat, 0, matobj3d, 0, rbmat.Length);
                                for (k = 0; k < hpix; ++k)
                                {
                                    c2wmatrix[n, k] = (float)matobj3d[k];
                                }
                            }
                        }
                        else
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
                                linetotalnum += newfile.Length / (hpix * 2);
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
                                            if (profile[m] <= 0 || profile[m] >= vpix)
                                            {
                                                profobj[m] = 0x7fff;
                                            }
                                            else
                                            {
                                                profobj[m] = (short)((c2wmatrix[profile[m], m] - basezval) * _scaleval);
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
                    string _dtwname = "";
                    int linecnt = 0;
                    float arutval = 0;
                    float brutval = 0;
                    float crutval = 0;

                    IniFiles rutcfg = new IniFiles(cfg[i]);
                    float _threshval = 0;
                    float[] _gsfilter;
                    int _partlen = 256, _asp = 0, _aep = 0, _bsp = 0, _bep = 0, _csp = 0, _cep = 0, _gslen = 0, _ThrPoint = 0;

                    _asp = rutcfg.ReadInteger("camera", "rutastart", 0);
                    _aep = rutcfg.ReadInteger("camera", "rutaend", 2048);
                    _bsp = rutcfg.ReadInteger("camera", "rutbstart", 0);
                    _bep = rutcfg.ReadInteger("camera", "rutbend", 2048);
                    _csp = rutcfg.ReadInteger("camera", "rutcstart", 0);
                    _cep = rutcfg.ReadInteger("camera", "rutcend", 2048);
                    profilenum = rutcfg.ReadInteger("camera", "calcstep", 50) / rutcfg.ReadInteger("sync", "plusstep", 10);
                    _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
                    _gslen = rutcfg.ReadInteger("rut", "gslen", 32) / 2 * 2 + 1;

                    MyPoint[] _pt = new MyPoint[hpix];
                    for (int pii = 0; pii < hpix; ++pii)
                    {
                        _pt[pii] = new MyPoint(0, 0);
                    }

                    _partlen = _aep - _asp - 2;//157、170、186
                    _threshval = rutcfg.ReadInteger("rut", "threshval", 28);
                    _ThrPoint = rutcfg.ReadInteger("rut", "threshpointnum", _partlen / 4);

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

                            double fsbar2 = 0.58 * j / _dats.Length;
                            double fsbar = 0.58 / _dats.Length;
                            //读取所有dat文件
                            using (FileStream frstream = new FileStream(string.Format("{0}\\{1}.dtw", process[i], _dtwname), FileMode.Open))
                            {
                                fsbar = fsbar / frstream.Length;
                                temp = hpix * 2;
                                while (frstream.Read(rbarr, 0, temp) > 0)
                                {
                                    if (linecnt == 1507)
                                    {
                                        isouttxt = true;
                                    }
                                    else
                                    {
                                        isouttxt = false;
                                    }

                                    Buffer.BlockCopy(rbarr, 0, profile, 0, rbarr.Length);
                                    for (m = 0; m < hpix; ++m)
                                    {
                                        objlas[m] = profile[m] / _scaleval;
                                        tobjlas[m] = objlas[m];
                                    }
                                    if (valnum == 1)
                                    {
                                        arutval = computerut(objlas, _gsfilter, _asp, _aep, _threshval, _partlen, _ThrPoint, _pt, tobjlas);

                                        //arutval = computerut_Line(objlas, _gsfilter, _asp, _aep, _threshval, _partlen, _ThrPoint, tobjlas);
                                        //arutval = computerut3(objlas, tobjlas, _asp, _aep, ref arutval, ref crutval, ref lineK, _gslen, MP_idx, MP_val);

                                        ++linecnt;
                                        sworirut.WriteLine(string.Format("{0}0,{1:0.000}", linecnt, arutval));
                                        swrut.WriteLine(string.Format("{0}0,{1:0.000}", linecnt, arutval));
                                    }
                                    else
                                    {
                                        //computerut3_allline(objlas, _gsfilter, _asp, _aep, _bsp, _bep, _csp, _cep, _threshval, _partlen,
                                        //    ref arutval, ref brutval, ref crutval, _ThrPoint, _pt, tobjlas);

                                        //computerut3(objlas, _gsfilter, _asp, _aep, _bsp, _bep, _csp, _cep, _threshval, _partlen,
                                        //    ref arutval, ref brutval, ref crutval, _ThrPoint, _pt, tobjlas);

                                        //computerut3_Line(objlas, _gsfilter, _asp, _aep, _bsp, _bep, _csp, _cep, _threshval, _partlen,
                                        //    ref arutval, ref brutval, ref crutval, _ThrPoint, tobjlas);

                                        brutval = computerut3(objlas, tobjlas, _asp, _cep, ref arutval, ref crutval, ref lineK, _gslen, MP_idx, MP_val);

                                        ++linecnt;
                                        sworirut.WriteLine(string.Format("{0}0,{1:0.000},{2:0.000},{3:0.000}", linecnt, arutval, brutval, crutval));
                                        swrut.WriteLine(string.Format("{0}0,{1:0.000},{2:0.000},{3:0.000}", linecnt, arutval, brutval, crutval));
                                    }

                                    if (linecnt % 1000 == 0)
                                    {
                                        if (Isbar) bar.SetRutVal(0.42 + fsbar2 + frstream.Position * fsbar);
                                    }
                                }
                            }
                        }
                        catch (Exception _fe)
                        {
                            Console.WriteLine(_fe.Message);
                        }
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
                    for (int t = 0; t < vallen; ++t)
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
            return false;
        }

        /// <summary>
        /// 单车辙，整条激光线去噪均值滤波，分左右两边各自用极点三角形法计算车辙值
        /// </summary>
        /// <param name="line"></param>
        /// <param name="gsfilter"></param>
        /// <param name="asp"></param>
        /// <param name="aep"></param>
        /// <param name="bsp"></param>
        /// <param name="bep"></param>
        /// <param name="csp"></param>
        /// <param name="cep"></param>
        /// <param name="threshval"></param>
        /// <param name="partlen"></param>
        /// <param name="aval"></param>
        /// <param name="bval"></param>
        /// <param name="cval"></param>
        /// <param name="pointthr"></param>
        /// <param name="py"></param>
        /// <param name="tline"></param>
        public static void computerut3(float[] line, float[] gsfilter,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval, int partlen,
            ref float aval, ref float bval, ref float cval, int pointthr, MyPoint[] py, float[] tline)
        {
            aval = 0.0f;
            bval = 0.0f;
            cval = 0.0f;

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
            aval = GetRutVal(py, asp, aep, 0.1f);
            cval = GetRutVal(py, csp, cep, 0.1f);
            //bval = GetRutVal(pt, bsp, bep - 1, 1);
            bval = Math.Max(aval, cval);
        }

        /// <summary>
        /// 单车辙，整条激光线去噪均值滤波，用极点三角形法计算车辙值，全线取点，任意两点在同侧为某侧车辙
        /// </summary>
        /// <param name="line"></param>
        /// <param name="gsfilter"></param>
        /// <param name="asp"></param>
        /// <param name="aep"></param>
        /// <param name="bsp"></param>
        /// <param name="bep"></param>
        /// <param name="csp"></param>
        /// <param name="cep"></param>
        /// <param name="threshval"></param>
        /// <param name="partlen"></param>
        /// <param name="aval"></param>
        /// <param name="bval"></param>
        /// <param name="cval"></param>
        /// <param name="pointthr"></param>
        /// <param name="py"></param>
        /// <param name="tline"></param>
        private static void computerut3_allline(float[] line, float[] gsfilter,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval, int partlen,
            ref float aval, ref float bval, ref float cval, int pointthr, MyPoint[] py, float[] tline)
        {
            aval = 0.0f;
            bval = 0.0f;
            cval = 0.0f;

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

            List<MyPoint> MiPt = new List<MyPoint>();
            float tmprut = 0.0f;
            GetMiPoint(py, asp, cep - 1, ref MiPt, 0.1f);
            MiPt.Sort(delegate(MyPoint x, MyPoint y) { return x.px.CompareTo(y.px); });

            int len = MiPt.Count;
            for (int i = 0; i < len; ++i)
            {
                for (int j = i + 1; j < len; ++j)
                {
                    for (int k = j + 1; k < len; ++k)
                    {
                        tmprut = GetTrigH_ac(MiPt[i], MiPt[j], MiPt[k]);
                        if (MiPt[j].px <= aep)
                        {
                            if (aval < tmprut)
                            {
                                aval = tmprut;
                            }
                        }
                        else
                        {
                            if (cval < tmprut)
                            {
                                cval = tmprut;
                            }
                        }
                    }
                }
            }

            bval = Math.Max(aval, cval);
        }

        /// <summary>
        /// 双车辙，整条激光线计算一个车辙值，用极点三角形法计算车辙值
        /// </summary>
        /// <param name="line"></param>
        /// <param name="gsfilter"></param>
        /// <param name="lines"></param>
        /// <param name="linee"></param>
        /// <param name="threshval"></param>
        /// <param name="partlen"></param>
        /// <param name="pointthr"></param>
        /// <param name="py"></param>
        /// <param name="tline"></param>
        /// <returns></returns>
        public static float computerut(float[] line, float[] gsfilter, int lines, int linee,
            float threshval, int partlen, int pointthr, MyPoint[] py, float[] tline)
        {
            float m_k = 0.0f;
            float m_b = 0.0f;
            LSLineFit(line, lines, linee, ref m_k, ref m_b);
            Height2Distance(ref line, lines, linee, m_k, m_b);
            eraseOutliers(ref line, lines, linee);

            //pickline(ref line, lines, linee, threshval);
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
            return GetRutVal(py, lines, linee, 0.1f);

            //leastsquare(line, lines, linee, ref k, ref b);
            //distanceline(ref line, lines, linee, k, b);
            //MidianAverageFileter(line, lines, linee, gsfilter, gsfilter.Length, ref tline);

            //return GetRutVal(partlen, lines, linee, tline, k, pointthr);
        }

        /// <summary>
        /// 单车辙，整条激光线去噪滤波，左右两边各自用线段拟合法计算车辙值
        /// </summary>
        /// <param name="line"></param>
        /// <param name="gsfilter"></param>
        /// <param name="asp"></param>
        /// <param name="aep"></param>
        /// <param name="bsp"></param>
        /// <param name="bep"></param>
        /// <param name="csp"></param>
        /// <param name="cep"></param>
        /// <param name="threshval"></param>
        /// <param name="partlen"></param>
        /// <param name="aval"></param>
        /// <param name="bval"></param>
        /// <param name="cval"></param>
        /// <param name="pointthr"></param>
        /// <param name="tline"></param>
        public static void computerut3_Line(float[] line, float[] gsfilter,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval, int partlen,
            ref float aval, ref float bval, ref float cval, int pointthr, float[] tline)
        {
            aval = 0.0f;
            bval = 0.0f;
            cval = 0.0f;

            float k = 1.0f, b = 0.0f;

            //line.CopyTo(tline, 0);
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

        /// <summary>
        /// 双车辙，整条激光线去噪滤波，用线段拟合法计算车辙值
        /// </summary>
        /// <param name="line"></param>
        /// <param name="gsfilter"></param>
        /// <param name="lines"></param>
        /// <param name="linee"></param>
        /// <param name="threshval"></param>
        /// <param name="partlen"></param>
        /// <param name="pointthr"></param>
        /// <param name="tline"></param>
        /// <returns></returns>
        public static float computerut_Line(float[] line, float[] gsfilter, int lines, int linee,
            float threshval, int partlen, int pointthr, float[] tline)
        {
            float k = 1.0f, b = 0.0f;

            //line.CopyTo(tline, 0);
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
                py[i] = (k * i + b - py[i]);
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
                for (i = ns; i < ns + flen; ++i)
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
                            y[i] += x[ne - 1] * f[flen + j];
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
            int i = 0, j = 0, hflen = 0, ti = 0, cnt = 0;
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
                    else
                    {
                        y[i] = x[i];
                    }
                }
            }
        }

        //中值滤波
        public static void MidianAverageFileter(List<MyPoint> x, int flen, ref List<MyPoint> y)
        {
            int i = 0, j = 0, hflen = 0, ti = 0, cnt = 0;
            float minval = 10000, maxval = -10000, sum = 0.0f;
            int len = x.Count;
            if (len > (flen - 1) * 2)
            {
                hflen = (flen - 1) / 2;
                for (i = 0; i < len; ++i)
                {
                    minval = 10000;
                    maxval = -10000;
                    sum = 0.0f;
                    cnt = 0;
                    for (j = -hflen; j <= hflen; ++j)
                    {
                        ti = i + j;
                        if (ti < 0 || ti >= len)
                            continue;

                        if (minval >= x[i + j].py)
                        {
                            minval = x[i + j].py;
                        }
                        if (maxval < x[i + j].py)
                        {
                            maxval = x[i + j].py;
                        }
                        sum += x[i + j].py;
                        cnt++;
                    }
                    if (cnt > 2)
                    {
                        y[i].py = (sum - minval - maxval) / (cnt - 2);
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

        //获取极值点
        public static void GetMiPoint(MyPoint[] Pt, int ns, int ne, ref List<MyPoint> MiPt, float thresh)
        {
            if (ne - ns < 127) return;

            //起始点先拉一段直线
            float k = (Pt[ns].py - Pt[ne].py) / (Pt[ns].px - Pt[ne].px);
            float b = Pt[ns].py - Pt[ns].px * k;
            MyPoint farminpt = new MyPoint(Pt[ns]);
            MyPoint farmaxpt = new MyPoint(Pt[ns]);
            float farmindis = 100000.0f;
            float farmaxdis = -100000.0f;
            float tmpval = 0.0f;

            List<MyPoint> tmpMiPt = new List<MyPoint>();
            tmpMiPt.Add(new MyPoint(Pt[ns]));

            //找直线两侧距离最远的点
            for (int i = ns + 1; i < ne - 1; ++i)
            {
                tmpval = Pt[i].py - (k * Pt[i].px + b);
                if (tmpval < farmindis)
                {
                    farminpt.px = Pt[i].px;
                    farminpt.py = Pt[i].py;
                    farmindis = tmpval;
                }
                if (tmpval > farmaxdis)
                {
                    farmaxpt.px = Pt[i].px;
                    farmaxpt.py = Pt[i].py;
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
            tmpMiPt.Sort(delegate(MyPoint x, MyPoint y) { return x.px.CompareTo(y.px); });

            // 极值点的数量，包括端点
            bool iscontain = false;
            int ptlen = tmpMiPt.Count;
            for (int i = 0; i < ptlen; ++i)
            {
                iscontain = false;
                for (int j = 0; j < MiPt.Count; ++j)
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
            MiPt.Sort(delegate(MyPoint x, MyPoint y) { return x.px.CompareTo(y.px); });

            //// 计算x中间点到x两端点边的高
            //int len = MiPt.Count;
            //for (int i = 0; i < len; ++i)
            //{
            //    for (int j = i + 2; j < len; ++j)
            //    {
            //        tmprut = GetTrigH_ac_b(MiPt, i, j);
            //        if (tmprut > maxrut)
            //        {
            //            maxrut = tmprut;
            //        }
            //    }
            //}
            int len = MiPt.Count;
            for (int i = 0; i < len; ++i)
            {
                for (int j = i + 1; j < len; ++j)
                {
                    for (int k = j + 1; k < len; ++k)
                    {
                        tmprut = GetTrigH_ac(MiPt[i], MiPt[j], MiPt[k]);
                        if (tmprut > maxrut)
                        {
                            maxrut = tmprut;
                        }
                    }
                }
            }
            return maxrut;
        }

        public static float GetTrigH_ac_b(List<MyPoint> MiPt, int aidx, int cidx)
        {
            float k = (MiPt[aidx].py - MiPt[cidx].py) / (MiPt[aidx].px - MiPt[cidx].px);
            float b = MiPt[aidx].py - MiPt[aidx].px * k;

            float maxval = -1000000.0f;
            float tmpval = 0.0f;
            for (int i = aidx + 1; i < cidx; ++i)
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
            //double lb = Math.Sqrt((pta.py - ptc.py) * (pta.py - ptc.py) + (pta.px - ptc.px) * (pta.px - ptc.px));

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
            else if (lc > la && lc > lb)
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

        //双车辙，以中间凸起点为界，左右各计算一个车辙深度
        public static float computerut3(float[] ArrayHeight, float[] ArrayDistance,
            int nStart, int nEnd, ref float RD_left, ref float RD_right, ref float m_k, int flen,
            int[] MP_idx, float[] MP_val)
        {
            List<MyPoint> pt = new List<MyPoint>();
            for (int i = nStart; i < nEnd; ++i)
            {
                if (ArrayHeight[i] < 300)
                {
                    pt.Add(new MyPoint(i, ArrayHeight[i]));
                }
            }

            RD_left = 0.0f;
            RD_right = 0.0f;

            int type_envelope = -1;
            bool isErase = false;

            if (isouttxt)
            {
                string[] hstr = new string[ArrayHeight.Length];
                for (int i = 0; i < ArrayHeight.Length; ++i)
                {
                    hstr[i] = ArrayHeight[i].ToString();
                }
                File.WriteAllLines("G:\\oriH.txt", hstr);
            }

            m_k = 0.0f;
            float m_b = 0.0f;
            LSLineFit(pt, ref m_k, ref m_b);
            Height2Distance(ref pt, m_k, m_b);
            isErase = eraseOutliers(ref pt);
            if (isErase)
            {
                m_k = 0.0f;
                m_b = 0.0f;

                LSLineFit(pt, ref m_k, ref m_b);
                Height2Distance(ref pt, m_k, m_b);
                eraseOutliers(ref pt);
            }

            nEnd = pt.Count;
            if (nEnd > 5)
            {
                nStart = 0;
                for (int i = 0; i < nEnd; ++i)
                {
                    ArrayHeight[i] = pt[i].py;
                }

                MidianAverageFileter(ArrayHeight, nStart, nEnd, flen, ref ArrayDistance);

                if (isouttxt)
                {
                    string[] hstr = new string[ArrayDistance.Length];
                    for (int i = 0; i < ArrayDistance.Length; ++i)
                    {
                        hstr[i] = ArrayDistance[i].ToString();
                    }
                    File.WriteAllLines("G:\\oriH1.txt", hstr);
                }

                findMaximumPiont(ArrayDistance, nStart, nEnd, ref MP_idx, ref MP_val, ref type_envelope);
                getRD(ArrayDistance, MP_idx, MP_val, type_envelope, ref RD_left, ref RD_right);
            }

            return Math.Max(RD_left, RD_right);
        }

        public static void Average_near(float[] ArrayHeight, int sidx, int eidx, int idx_center, ref float value_average)
        {
            int idx_left = idx_center - 10;
            int idx_right = idx_center + 10;
            if (idx_left < sidx)
            {
                idx_left = sidx;
            }
            if (idx_right > eidx - 1)
            {
                idx_right = eidx - 1;
            }

            value_average = 0.0f;
            for (int i = idx_left; i <= idx_right; i++)
            {
                value_average += ArrayHeight[i];
            }
            value_average = value_average / (idx_right - idx_left + 1);
        }

        /***********************************************************************************/
        /*       w0 \     w2 /\     w4 /                                                   */
        /* 			 \      /  \      /                                                    */
        /*			  \    /    \    /                                                     */
        /*			   \  /      \  /                                                      */
        /*			 w1 \/     w3 \/               车辙W模型，算法中各凹凸点示意           */
        /***********************************************************************************/
        //寻找车辙的中间凸起点，通过高程数据[500,1500]内的数据的极大值确定,并确定包络线类型
        //VecHeights:为路面高程数据
        //num_start:计算数据的左端点序号，小于500  num_end:计算数据的右端点序号，大于1500
        //RP_idx:凸起点序号      RP_val：凸起点值
        //type_envelope:0或者1，分别对应端点连线或端点中间凸起点折线作为包络线
        public static void findMaximumPiont(float[] ArrayHeight, int num_start, int num_end, ref int[] MP_idx, ref float[] MP_val, ref int type_envelope)
        {
            MP_val[2] = -10000.0f;
            //for (int i = (num_start + num_end) / 2 - 300; i < (num_start + num_end) / 2 + 300; i++)
            int tsidx = (num_start + num_end) / 3;
            int teidx = (num_start + num_end) * 2 / 3;
            //int tsidx = (num_start + num_end) / 2 - 300;
            //int teidx = (num_start + num_end) * 2 + 300;
            int meidx = (tsidx + teidx) / 2;
            for (int i = tsidx; i < meidx; i++)
            {
                if (ArrayHeight[i] >= MP_val[2])
                {
                    MP_val[2] = ArrayHeight[i];
                    MP_idx[2] = i;
                }
            }
            for (int i = meidx; i < teidx; i++)
            {
                if (ArrayHeight[i] > MP_val[2])
                {
                    MP_val[2] = ArrayHeight[i];
                    MP_idx[2] = i;
                }
            }

            Average_near(ArrayHeight, num_start, num_end, MP_idx[2], ref MP_val[2]);

            //寻找左侧凹点，W1
            //findPit(ArrayHeight, num_start+300, MP_idx[2], ref MP_idx[1], ref MP_val[1]); 
            findPit(ArrayHeight, num_start+10, MP_idx[2], ref MP_idx[1], ref MP_val[1], 0);

            //寻找左侧端点，W0
            findedge(ArrayHeight, num_start, MP_idx[1], ref MP_idx[0], ref MP_val[0], 0);

            //寻找右侧凹点，W3
            //findPit(ArrayHeight, MP_idx[2], num_end-300, ref MP_idx[3], ref MP_val[3]);
            findPit(ArrayHeight, MP_idx[2], num_end-10, ref MP_idx[3], ref MP_val[3], 1);

            //寻找右侧端点，W4
            findedge(ArrayHeight, MP_idx[3], num_end, ref MP_idx[4], ref MP_val[4], 1);

            //判断包络线类型，凸W：type_envelope =1 凹W：type_envelope =0
            double k = (MP_val[4] - MP_val[0]) / (MP_idx[4] - MP_idx[0]);
            double b = MP_val[0] - k * MP_idx[0];
            double dis = MP_val[2] - (k * MP_idx[2] + b);

            //中间凸点在端点连线下方，凹W
            if (dis <= 0)
            {
                type_envelope = 0;
            }
            //中间凸点在端点连线上方，凸W
            else
            {
                type_envelope = 1;
            }
        }

        //计算左右车辙深度
        //VecHeights:为路面高程数据
        //RP_idx:凸起点序号      RP_val：凸起点值
        //type_envelope:0或者1，分别对应端点连线或端点中间凸起点折线作为包络线
        //RD_left:左车辙值             RD_right：右车辙值
        public static void getRD(float[] ArrayHeight, int[] MP_idx, float[] MP_val, int type_envelope, ref float RD_left, ref float RD_right)
        {
            //根据包络线类型计算左右车辙
            if (type_envelope == 0)
            {
                //端点连线包络线
                //左车辙
                calculateRD(ArrayHeight, MP_idx, MP_val, 0, 4, 1, ref RD_left);
                //右车辙
                calculateRD(ArrayHeight, MP_idx, MP_val, 0, 4, 3, ref RD_right);
            }
            else if (type_envelope == 1)
            {
                //端点凸点折线包络线
                //左车辙
                calculateRD(ArrayHeight, MP_idx, MP_val, 0, 2, 1, ref RD_left);
                //右车辙
                calculateRD(ArrayHeight, MP_idx, MP_val, 2, 4, 3, ref RD_right);
            }
        }

        //寻找[idx_start,idx_end]内的极小点作为凹点，记录凹点的序号和对应的值
        public static void findPit(float[] ArrayHeight, int idx_start, int idx_end, ref int Pit_idx, ref float Pit_val, int side)
        {
            Pit_val = 10000.0f;
            for (int i = idx_start; i < idx_end; i++)
            {
                if (side == 0)
                {
                    if (ArrayHeight[i] <= Pit_val)
                    {
                        Pit_val = ArrayHeight[i];
                        Pit_idx = i;
                    }
                }
                else
                {
                    if (ArrayHeight[i] < Pit_val)
                    {
                        Pit_val = ArrayHeight[i];
                        Pit_idx = i;
                    }
                }
            }
            Average_near(ArrayHeight, idx_start, idx_end, Pit_idx, ref Pit_val);
        }

        //寻找车辙的左右端点
        public static void findedge(float[] ArrayHeight, int idx_start, int idx_end, ref int edge_idx, ref float edge_val, int side)
        {
            edge_val = -10000.0f;
            for (int i = idx_start; i < idx_end; i++)
            {
                if (side == 1)
                {
                    if (ArrayHeight[i] >= edge_val)
                    {
                        edge_val = ArrayHeight[i];
                        edge_idx = i;
                    }
                }
                else
                {
                    if (ArrayHeight[i] > edge_val)
                    {
                        edge_val = ArrayHeight[i];
                        edge_idx = i;
                    }
                }
            }
            Average_near(ArrayHeight, idx_start, idx_end, edge_idx, ref edge_val);
        }

        //根据包络线和凹点数据计算车辙深度
        public static void calculateRD(float[] ArrayHeight, int[] MP_idx, float[] MP_val, int MPleft_idx, int MPright_idx, int MPpoint_idx, ref float RD)
        {
            if (MP_idx[MPleft_idx] != MP_idx[MPright_idx])
            {
                float k = (MP_val[MPleft_idx] - MP_val[MPright_idx]) / (MP_idx[MPleft_idx] - MP_idx[MPright_idx]);
                float b = MP_val[MPleft_idx] - k * MP_idx[MPleft_idx];

                float tempRD = 0.0f;
                for (int i = MP_idx[MPpoint_idx - 1]; i <= MP_idx[MPpoint_idx + 1]; i++)
                {
                    tempRD = k * i + b - ArrayHeight[i];
                    if (tempRD > RD)
                    {
                        RD = tempRD;
                    }
                }
            }
            else
            {
                RD = 0;
            }
        }

        //拟合直线，对高度进行转换
        public static void Height2Distance(ref float[] Arrayheight, int numstart, int numend, float m_k, float m_b)
        {
            for (int i = numstart; i < numend; i++)
            {
                //计算点到线的距离
                Arrayheight[i] = Arrayheight[i] - (m_k * i + m_b);
            }
        }

        //拟合直线，对高度进行转换
        public static void Height2Distance(ref List<MyPoint> Arrayheight, float m_k, float m_b)
        {
            int len = Arrayheight.Count;
            for (int i = 0; i < len; i++)
            {
                //计算点到线的距离
                Arrayheight[i].py = Arrayheight[i].py - (m_k * Arrayheight[i].px + m_b);
            }
        }

        //计算数组的平均值和方差
        public static void mean_std(float[] ArrayHeight, int nStart, int nEnd, ref float mean, ref float std)
        {
            mean = 0.0f;
            for (int i = nStart; i < nEnd; i++)
            {
                mean += ArrayHeight[i];
            }
            mean = mean / (nEnd - nStart);

            std = 0.0f;
            for (int i = nStart; i < nEnd; i++)
            {
                std += (ArrayHeight[i] - mean) * (ArrayHeight[i] - mean);
            }
            std = (float)Math.Sqrt(std / (nEnd - nStart));
        }

        //计算数组的平均值和方差
        public static void mean_std(List<MyPoint> ArrayHeight, ref float mean, ref float std)
        {
            int len = ArrayHeight.Count;
            mean = 0.0f;
            for (int i = 0; i < len; i++)
            {
                mean += ArrayHeight[i].py;
            }
            mean = mean / len;

            std = 0.0f;
            for (int i = 0; i < len; i++)
            {
                std += (ArrayHeight[i].py - mean) * (ArrayHeight[i].py - mean);
            }
            std = (float)Math.Sqrt(std / len);
        }

        //通过拉依达法则去除异常值
        public static bool eraseOutliers(ref float[] ArrayHeight, int nStart, int nEnd)
        {
            bool isErase = false;

            float mean = 0.0f, std = 0.0f;
            mean_std(ArrayHeight, nStart, nEnd, ref mean, ref std);
            float stdThresh = 3.34f * std + 2.0f;
            for (int i = nStart; i < nEnd; i++)
            {
                if (Math.Abs(ArrayHeight[i] - mean) > stdThresh)
                {
                    ArrayHeight[i] = mean;
                    isErase = true;
                }
            }

            return isErase;
        }

        //通过拉依达法则去除异常值
        public static bool eraseOutliers(ref List<MyPoint> ArrayHeight)
        {
            int len = ArrayHeight.Count;
            bool isErase = false;

            float mean = 0.0f, std = 0.0f;
            mean_std(ArrayHeight, ref mean, ref std);
            float stdThresh = 3.34f * std + 2.0f;
            for (int i = len - 1; i >= 0; --i)
            {
                if (Math.Abs(ArrayHeight[i].py - mean) > stdThresh)
                {
                    ArrayHeight.RemoveAt(i);
                    isErase = true;
                }
            }

            return isErase;
        }

        //最小二乘法拟合直线
        public static void LSLineFit(float[] ArrayHeight, int nStart, int nEnd, ref float m_k, ref float m_b)
        {
            float sumX2 = 0.0f;
            float sumX = 0.0f;
            float sumXY = 0.0f;
            float sumY = 0.0f;

            for (int i = nStart; i < nEnd; i++)
            {
                sumX2 += i * i;
                sumX += i;
                sumXY += i * ArrayHeight[i];
                sumY += ArrayHeight[i];
            }

            //计算斜率和截距
            int num_point = nEnd - nStart;
            float tempDenominator = num_point * sumX2 - sumX * sumX;

            if (tempDenominator != 0)
            {
                m_k = (num_point * sumXY - sumX * sumY) / tempDenominator;
                m_b = (sumX2 * sumY - sumX * sumXY) / tempDenominator;
            }
            else
            {
                m_k = 1.0f;
                m_b = 0.0f;
            }
        }

        //最小二乘法拟合直线
        public static void LSLineFit(List<MyPoint> ArrayHeight, ref float m_k, ref float m_b)
        {
            float sumX2 = 0.0f;
            float sumX = 0.0f;
            float sumXY = 0.0f;
            float sumY = 0.0f;

            int len = ArrayHeight.Count;
            for (int i = 0; i < len; i++)
            {
                sumX2 += ArrayHeight[i].px * ArrayHeight[i].px;
                sumX += ArrayHeight[i].px;
                sumXY += ArrayHeight[i].px * ArrayHeight[i].py;
                sumY += ArrayHeight[i].py;
            }

            //计算斜率和截距
            float tempDenominator = len * sumX2 - sumX * sumX;
            if (tempDenominator != 0)
            {
                m_k = (len * sumXY - sumX * sumY) / tempDenominator;
                m_b = (sumX2 * sumY - sumX * sumXY) / tempDenominator;
            }
            else
            {
                m_k = 1.0f;
                m_b = 0.0f;
            }
        }
    }
}
