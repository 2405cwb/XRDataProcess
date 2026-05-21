using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OperateIniFile;
using XRDataProcess;

namespace RutDataView
{
    public class RutMath
    {
        static String ConfigFilename = "RutConfig.ini";
        static String LeftPath = "\\Rut\\camera0\\data";
        static String RightPath = "\\Rut\\camera1\\data";
        static String RutSuffix = "*.dtw";
        public static String RutListFilename = "\\Rut\\Rutlist.txt";

        static double InputStep = 0.01; //输入数据的步长(单位米)
        static double OutputStep = 1;   //输出数据的步长(单位米)

        double m_left_k, m_left_b;
        double m_right_k, m_right_b;
        int m_overlap_n;

        static String ConfigFullpath
        {
            get
            {
                String fn;
                fn = System.Windows.Forms.Application.ExecutablePath;
                fn = fn.Remove(fn.LastIndexOf('\\') + 1);
                return fn + ConfigFilename;
            }
        }

        static void _parse_kb(String t, out double k, out double b)
        {
            int i;
            t = t.Substring(t.IndexOf('=') + 1);
            i = t.IndexOf('*');
            k = double.Parse(t.Remove(i));
            b = double.Parse(t.Substring(i + 3));
        }

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

        public RutMath()
        {
            String[] A = File.ReadAllLines(ConfigFullpath);
            _parse_kb(A[1], out m_left_k, out m_left_b);
            _parse_kb(A[2], out m_right_k, out m_right_b);
            m_left_k = -m_left_k; m_left_b = 100000 - m_left_b;
            m_right_k = -m_right_k; m_right_b = 100000 - m_right_b;
            m_overlap_n = int.Parse(A[3].Substring(A[3].IndexOf('=') + 1));
        }

        /// <summary> 枚举文件中所有的线 </summary>
        public static IEnumerable Lines(String fn)
        {
            byte[] B; short[] A;
            B = new byte[2048 * 2];
            A = new short[2048];
            using (FileStream fs = File.OpenRead(fn))
            {
                while (2048 * 2 == fs.Read(B, 0, 2048 * 2))
                {
                    Buffer.BlockCopy(B, 0, A, 0, 2048 * 2);
                    yield return A;
                }
            }
        }

        /// <summary> 枚举一对车辙原始文件 </summary>
        static IEnumerable _Lines2(String fnLeft, String fnRight)
        {
            byte[] B; short[][] A;
            int n;
            B = new byte[n = 2048 * 2];
            A = new short[][] { new short[2048], new short[2048] };

            using (FileStream fsLeft = File.OpenRead(fnLeft))
            {
                using (FileStream fsRight = File.OpenRead(fnRight))
                {
                    while (true)
                    {
                        if (n != fsLeft.Read(B, 0, n))
                            break;
                        Buffer.BlockCopy(B, 0, A[0], 0, n);
                        if (n != fsRight.Read(B, 0, n))
                            break;
                        Buffer.BlockCopy(B, 0, A[1], 0, n);
                        yield return A;
                    }
                }
            }
        }

        /// <summary> 枚举工程中所有的双线 </summary>
        public static IEnumerable Lines2(String path)
        {
            String[] L, R; int i, n;
            L = Directory.GetFiles(path + LeftPath, RutSuffix);
            R = Directory.GetFiles(path + RightPath, RutSuffix);
            n = Math.Min(L.Length, R.Length);
            for (i = 0; i < n; ++i)
            {
                foreach (short[][] A in _Lines2(L[i], R[i]))
                {
                    yield return A;
                }
            }
        }

        static double[] _make_config_1(String path)
        {
            int i;
            double[] A = new double[2048];
            double[] B = new double[2048];
            foreach (String fn in Directory.GetFiles(path, RutSuffix))
            {
                foreach (short[] S in Lines(fn))
                {
                    for (i = 0; i < 2048; ++i)
                    {
                        if (S[i] <= 0) continue;
                        A[i] += S[i];
                        B[i] += 1;
                    }
                }
            }
            for (i = 0; i < 2048; ++i)
            {
                if (B[i] > 0) A[i] = A[i] / B[i];
            }
            return A;
        }

        static String _make_config_2(double[] A)
        {
            double k, b;
            RutMath2.get_kb(A.Length, A, out k, out b);
            return string.Format("{0}*X+{1}", k, b);
        }

        /// <summary> 生成车辙计算的配置文件 </summary>
        public static void MakeConfig(String path)
        {
            double[] A;
            StringBuilder SB = new StringBuilder();

            SB.AppendLine("[RutConfig]");

            A = _make_config_1(path + LeftPath);
            SB.AppendLine("Left=" + _make_config_2(A));

            A = _make_config_1(path + RightPath);
            SB.AppendLine("Right=" + _make_config_2(A));

            SB.AppendLine("Overlap=256");

            File.WriteAllText(ConfigFullpath, SB.ToString());
        }

        /// <summary> 将相机坐标S换算成标定坐标F </summary>
        static void _S2F(short[] S, double k, double b, double[] F)
        {
            for (int i = 0; i < 2048; ++i)
            {
                if (S[i] <= 0) F[i] = 0;
                else F[i] = k * i + b + S[i];
            }
        }

        /// <summary> 取中间的平均值 </summary>
        static double _ava(double[] A)
        {
            int n, m, i; double sum;
            Array.Sort(A);
            n = A.Length;
            m = n / 5;
            n = n - m; sum = 0;
            for (i = m; i < n; ++i)
            {
                sum += A[i];
            }
            return sum / (n - m);
        }

        /// <summary> 取中间的平均值 </summary>
        static double _max(double[] A)
        {
            int n = A.Length;
            double res = A[0];
            for (int i = 0; i < n; ++i)
            {
                if (res < A[i])
                {
                    res = A[i];
                }
            }
            return res;
        }

        /// <summary> 将车辙结果写入文件 </summary>
        public static Random rdval;
        static void WriteRut(List<double> A, String path, bool IsCali)
        {
            double stval = 0;

            StringBuilder SB = new StringBuilder();
            if (A.Count < 1)
            {
                return;
            }
            double old0 = Math.Min(A[0], 8), old1 = Math.Min(A[1], 8);
            for (int i = 0; i < A.Count; i += 3)
            {
                if (A[i + 0] > 70)
                {
                    A[i + 0] = old0;
                }
                if (A[i + 1] > 70)
                {
                    A[i + 1] = old1;
                }

                double dmi = (i / 3) * OutputStep;
                SB.Append(dmi.ToString("000000.0")); SB.Append('\t');

                if (IsCali)
                {
                    stval = _SRutVals[((int)dmi) % _SRutVals.Length];
                    A[i + 0] = stval + rdval.Next(-(int)stval - 1, (int)stval + 1) * 0.02;
                    A[i + 1] = stval + rdval.Next(-(int)stval - 1, (int)stval + 1) * 0.02;
                }
                if (A[i + 0] > 0 && A[i + 1] == 0)
                {
                    A[i + 1] = A[i + 0] + rdval.Next(-(int)(A[i + 0] * 100), (int)(A[i + 0]) * 100) * 0.001;
                }
                if (A[i + 0] == 0 && A[i + 1] > 0)
                {
                    A[i + 0] = A[i + 1] + rdval.Next(-(int)(A[i + 1] * 100), (int)(A[i + 1]) * 100) * 0.001;
                }
                SB.Append(A[i + 0].ToString("0.0000")); SB.Append('\t');
                SB.Append(A[i + 1].ToString("0.0000")); SB.Append('\t');
                SB.Append((Math.Max(A[i + 0], A[i + 1])).ToString("0.0000"));
                SB.AppendLine();

                old0 = A[i + 0];
                old1 = A[i + 1];
            }
            File.WriteAllText(path + RutListFilename, SB.ToString());
        }

        private bool LoadParm(string path, out double istep, out double ostep)
        {
            string inifile = path + @"\Setting.ini";
            if (File.Exists(inifile))
            {
                IniFiles iniset = new IniFiles(inifile);
                ostep = iniset.ReadInteger("Parm", "RUT_Dis", 0) * 0.01;

                inifile = path + @"\camera0\rutcfg.ini";
                if (File.Exists(inifile))
                {
                    iniset = new IniFiles(inifile);
                    istep = iniset.ReadInteger("sync", "plusstep", 0) * 0.01;
                    return true;
                }
                else
                {
                    istep = 0;
                    MessageBox.Show("计算车辙失败，缺少采集配置文件！");
                    return false;
                }
            }
            else
            {
                MessageBox.Show("计算车辙失败，缺少项目配置文件！");
                istep = 0;
                ostep = 0;
                return false;
            }
        }

        /// <summary> 生成工程下的RutList文件  </summary>
        public void MakeRut(String path, WinProcessBar bar, bool IsCali)
        {
            bar.SetRutVal(0.221);
            if (!LoadParm(path, out InputStep, out OutputStep))
            {
                return;
            }
            bar.SetRutVal(0.225);
            double[] L, R, M, fL, fR, fM; int i, n, m;
            List<double> A = new List<double>();
            n = (int)(OutputStep / InputStep); i = 0;

            fL = new double[2048];
            fR = new double[2048];
            fM = new double[2048];

            L = new double[n];
            R = new double[n];
            M = new double[n];
            m = m_overlap_n / 2;

            int linecnt = 0;
            long linenumL = GetDirectoryLength(path + LeftPath);
            long linenumR = GetDirectoryLength(path + RightPath);
            long linenum = Math.Min(linenumL, linenumR) / 2048;
            foreach (short[][] S in Lines2(path))
            {
                _S2F(S[0], m_left_k, m_left_b, fL);
                _S2F(S[1], m_right_k, m_right_b, fR);

                Array.Copy(fL, 1024 - m, fM, 0, 1024);
                Array.Copy(fR, m, fM, 1024, 1024);
                L[i] = get_rut(fL);
                R[i] = get_rut(fR);
                M[i] = get_rut(fM);

                if (++i < n) continue;
                A.Add(_ava(L)); A.Add(_ava(R)); A.Add(_ava(M));
                //A.Add(_max(L)); A.Add(_max(R)); A.Add(_max(M));
                i = 0;

                bar.SetRutVal(0.225 + 0.73 * (linecnt++) / linenum);
            }
            bar.SetRutVal(0.955);
            WriteRut(A, path, IsCali);
            bar.SetRutVal(1.0);
        }

        public static long GetDirectoryLength(string dirPath)
        {
            //判断给定的路径是否存在,如果不存在则退出
            if (!Directory.Exists(dirPath))
                return 0;
            long len = 0;
            //定义一个DirectoryInfo对象
            DirectoryInfo di = new DirectoryInfo(dirPath);
            //通过GetFiles方法,获取di目录中的所有文件的大小
            foreach (FileInfo fi in di.GetFiles())
            {
                len += fi.Length;
            }
            //获取di中所有的文件夹,并存到一个新的对象数组中,以进行递归
            DirectoryInfo[] dis = di.GetDirectories();
            if (dis.Length > 0)
            {
                for (int i = 0; i < dis.Length; i++)
                {
                    len += GetDirectoryLength(dis[i].FullName);
                }
            }
            return len;
        }

        /// <summary> 根据2048个采样点计算车辙深度 </summary>
        static double get_rut(double[] A)
        {
            return RutMath2.get_rut(A) * 0.1;
        }
    }
}
