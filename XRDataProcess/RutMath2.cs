using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RutDataView
{
    static class RutMath2
    {
        static void Filter_Lowpass(int r, int n, double[] P, double[] Q)
        {
            int i, m; double sum;

            sum = m = 0;
            for (i = 0; i < r; ++i)
            {
                sum += P[i];
                m++;
            }

            for (i = 0; i < n; ++i)
            {
                if (i + r < n)
                {
                    sum += P[i + r];
                    ++m;
                }
                if (i - r > 0)
                {
                    sum -= P[i - r - 1];
                    --m;
                }
                if (m == 0) Q[i] = -1;
                else Q[i] = sum / m;
            }
        }

        static int max_i(double[] A, int f, int l)
        {
            int i = f;
            for (; ++f < l; )
            {
                if (A[f] > A[i])
                    i = f;
            }
            return i;
        }

        static int min_i(double[] A, int f, int l)
        {
            int i = f;
            for (; ++f < l; )
            {
                if (A[f] < A[i])
                    i = f;
            }
            return i;
        }

        static double dis(double[] A, int i, int j)
        {
            double x, y;
            x = i - j;
            y = A[i] - A[j];
            return Math.Sqrt(x * x + y * y);
        }

        //描述: 计算车辙深度
        //参数: P是断面的高程值
        //      i是候选的车辙最深点的位置
        //返回: 车辙深度
        static double _rut(double[] P, int u, int v, int w)
        {
            double a, b, c, s;
            a = dis(P, u, v);
            b = dis(P, v, w);
            c = dis(P, u, w);
            s = (a + b + c) / 2;
            s *= (s - a) * (s - b) * (s - c);//海伦公式求面积
            if (s > 0 && c != 0)
            {
                s = 2 * Math.Sqrt(s) / c;
            }
            else
            {
                s = 0;
            }
            return s;
        }

        //描述: 计算车辙深度
        //参数: [P,P+n)是断面的高程值
        //返回: 车辙深度
        static double _get_rut(int n, double[] A)
        {
            double s = 0, maxs, s1;
            int i, m, a, b, c, a1, b1, c1;
            maxs = 0; m = n / 10;
            for (i = 1; i < 9; ++i)
            {
                b = i * m;
                b = min_i(A, b, b + m);
                a = max_i(A, 0, b);
                c = max_i(A, b, n);
                s = _rut(A, a, b, c);
                //if (s > maxs) maxs = s;

                //b1 = i * m;
                //b1 = max_i(A, b1, b1 + m);
                //a1 = min_i(A, 0, b1);
                //c1 = min_i(A, b1, n);
                //s1 = _rut(A, a1, b1, c1);
                //if (s1 > maxs) maxs = s1;
            }
            //return maxs;
            return s;
        }

        public static void get_kb(int m, double[] A, out double k, out double b)
        {
            double x, y, xx, yy, v; int i, n;
            xx = yy = x = y = 0;
            for (i = n = 0; i < m; ++i)
            {
                if (A[i] <= 0) continue;
                x += i; y += A[i]; ++n;
            }
            x = x / n; y = y / n;
            for (i = n = 0; i < m; ++i)
            {
                if (A[i] <= 0) continue;
                v = i - x; xx += v * v;
                yy += v * (A[i] - y);
            }
            k = yy / xx;
            b = y - k * x;
        }

        //描述: 计算车辙深度
        //参数: A是断面的高程值
        //返回: 车辙深度
        public static double get_rut(double[] A)
        {
            double[] B; int i, n; double k, b;
            get_kb(A.Length, A, out k, out b);
            B = new double[A.Length];
            double oldval = 0;
            if (A.Length > 0)
            {
                oldval = A[0];
            }
            for (i = n = 0; i < A.Length; ++i)
            {
                if (A[i] > 0)
                {
                    if (Math.Abs(oldval - A[i]) > 100)
                    {
                        A[i] = oldval;
                    }
                    //B[n++] = A[i] - i * k - b;
                    B[n++] = A[i];
                }
                oldval = A[i];
            }
            if (n < 100) return 0;

            A = new double[n];
            Filter_Lowpass(20, n, B, A);
            return _get_rut(n, B);
        }

    }
}
