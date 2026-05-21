using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OperateIniFile;
using System.Windows.Forms;

namespace poly2tin
{
    public class RutStep2
    {
        static String ConfigFilename = "RutConfig.ini";
        static String LeftPath = "\\Rut\\Camera0\\data";
        static String RightPath = "\\Rut\\Camera1\\data";
        static String RutSuffix = "*.dtw";
        static String RutListFilename = "\\Rut\\Rutlist.txt";

        static double InputStep = 0.01; //输入数据的步长(单位米)，就是采集的时候多少米一个激光线
        static double OutputStep = 10;   //输出数据的步长(单位米)，就是多少米出一个车辙值，输出步长必须是输入的整数倍

        double m_left_k, m_left_b;
        double m_right_k, m_right_b;
        int m_overlap_n;
        
        static String ConfigFullpath
        {
            get
            {
                //String fn;
                //fn = System.Windows.Forms.Application.ExecutablePath;
                //fn = fn.Remove(fn.LastIndexOf('\\') + 1);
                //return fn + ConfigFilename;

                return System.Windows.Forms.Application.StartupPath +"\\"+ ConfigFilename;
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

        public RutStep2()
        {
            String[] A = File.ReadAllLines(ConfigFullpath);
            _parse_kb(A[1], out m_left_k, out m_left_b);
            _parse_kb(A[2], out m_right_k, out m_right_b);
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
                    istep = iniset.ReadInteger("sync", "plusstep", 0)*0.01;
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
        /// 干活儿时调用这个
        public void MakeRut(String path)
        {
            if (!LoadParm(path,out InputStep,out OutputStep))
            {
                return;
            }
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
                i = 0;
            }
            WriteRut(A, path);
        }

        /// <summary> 生成车辙计算的配置文件 </summary>
        /// 出厂的时候生成配置文件
        /// 激光点要求全部打在型材上
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
            double x, y, xx, yy, v; int i, n;
            xx = yy = x = y = 0;
            for (i = n = 0; i < 2048; ++i)
            {
                if (A[i] <= 0) continue;
                x += i; y += A[i]; ++n;
            }
            x = x / n; y = y / n;
            for (i = n = 0; i < 2048; ++i)
            {
                if (A[i] <= 0) continue;
                v = i - x; xx += v * v;
                yy += v * (A[i]-y);
            }
            v = yy / xx;
            y = 10000+y - v * x;
            return v.ToString() + "*X+" + y.ToString();
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
            int n=0, m=0, i=0; 
            double sum;
            Array.Sort(A);
            n = A.Length;
            m = n / 5;
            n = n - m; sum = 0;
            for (i = m; i < n; ++i)
            {
                sum += A[i];
            }
            return sum / (n - m);
            //return (A[n - 3] + A[n - 2] + A[n - 1]) / 3;
            //return A[n - 1];
        }

        /// <summary> 将车辙结果写入文件 </summary>
        static void WriteRut(List<double> A, String path)
        {
            StringBuilder SB = new StringBuilder();            
            SB.AppendLine("DMI(M)  \tLeft\tRight\tMiddle");
            for (int i = 0; i < A.Count; i += 3)
            {
                double dmi = (i / 3) * OutputStep;
                SB.Append(dmi.ToString("000000.0")); SB.Append('\t');
                SB.Append(A[i + 0].ToString("0.0000")); SB.Append('\t');
                SB.Append(A[i + 1].ToString("0.0000")); SB.Append('\t');
                SB.Append(A[i + 2].ToString("0.0000")); SB.AppendLine();
            }
            File.WriteAllText(path + RutListFilename, SB.ToString());
        }

        /// <summary> 根据2048个采样点计算车辙深度 </summary>
        static double get_rut(double[] A)
        {
            return RutMath2.get_rut(A)*0.1;
        }
    }
}
