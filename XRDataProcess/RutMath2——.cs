using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace poly2tin
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
        static double _rut(double[] P, int u,int v,int w)
        {
            double a, b, c, s;
            a = dis(P, u, v);
            b = dis(P, v, w);
            c = dis(P, u, w);
            s = (a + b + c) / 2;
            s *= (s - a) * (s - b) * (s - c);
            s = 0.2 * Math.Sqrt(s) / c;
            return s;
        }

        //描述: 计算车辙深度
        //参数: [P,P+n)是断面的高程值
        //返回: 车辙深度
        static double _get_rut(int n,double[] A)
        {
            double s, maxs;
            int i, m, a, b, c;
            maxs = 0; m = n / 10;
            for (i = 1; i < 9; ++i)
            {
                b = i * m;
                b = min_i(A, b, b + m);
                a = max_i(A, 0, b);
                c = max_i(A, b, n);
                s = _rut(A, a, b, c);
                if (s > maxs) maxs = s;
            }
            return maxs;
        }

        //描述: 计算车辙深度
        //参数: A是断面的高程值
        //返回: 车辙深度
        public static double get_rut(double[] A)
        {
            double[] B; int i,n;
            B = new double[A.Length];
            for (i = n = 0; i < A.Length; ++i)
            {
                if (A[i] > 0)
                    B[n++] = A[i];
            }
            if (n < 100) return 0;

            A = new double[n];
            Filter_Lowpass(10, n, B, A);
            return _get_rut(n, A);
        }
    }
}
