using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pylon
{
    class Program
    {
        static void Main(string[] args)
        {
            /* double[,] xArray = new double[,]
             {
                
                     { 2.000000 ,-1.000000 , 3.000000,  1.000000},
                     { 4.000000 , 2.000000 , 5.000000,  4.000000},
                     { 1.000000 , 2.000000 , 0.000000 , 7.000000}
             };*/

            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            double[] y = new double[] { 29152.3, 47025.3, 86852.3, 132450.6, 200302.3, 284688.1, 396988.3 };
            double[] x = new double[] { 1.24, 2.37, 5.12, 8.12, 12.19, 17.97, 24.99 };

            // double[,] xArray;
            double[] ratio;
            double[] yy = new double[y.Length];

            Console.WriteLine("一次拟合：");
            sw.Start();
            ratio = FittingFunct.Linear(y, x);
            sw.Stop();

            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[0] + ratio[1] * x[i];
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy) + "\r\n");
            //Console.WriteLine("一次拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.WriteLine("一次拟合(截距为0，即强制过原点)：");
            sw.Start();
            ratio = FittingFunct.LinearInterceptZero(y, x);
            sw.Stop();

            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[0] * x[i];
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy) + "\r\n");
            //Console.WriteLine("一次拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.WriteLine("二次拟合：");
            sw.Start();
            ratio = FittingFunct.TowTimesCurve(y, x);
            sw.Stop();

            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[0] + ratio[1] * x[i] + ratio[2] * x[i] * x[i];
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy) + "\r\n");
            //Console.WriteLine("二次拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.WriteLine("对数拟合计算时间：");
            sw.Start();
            ratio = FittingFunct.LOGEST(y, x);
            sw.Stop();

            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[1] * Math.Log10(x[i]) + ratio[0];
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy) + "\r\n");
            //Console.WriteLine("对数拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.WriteLine("幂级数拟合：");
            sw.Start();
            ratio = FittingFunct.PowEST(y, x);
            sw.Stop();

            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[0] * Math.Pow(x[i], ratio[1]);
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy) + "\r\n");
            //Console.WriteLine("幂级数拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.WriteLine("指数函数拟合：");
            sw.Start();
            ratio = FittingFunct.IndexEST(y, x);
            sw.Stop();
            foreach (double num in ratio)
            {
                Console.WriteLine(num);
            }
            for (int i = 0; i < x.Length; i++)
            {
                yy[i] = ratio[0] * Math.Exp(x[i] * ratio[1]);
            }
            Console.WriteLine("R²=: " + FittingFunct.Pearson(y, yy));
            //Console.WriteLine("指数函数拟合计算时间：");
            //Console.WriteLine(sw.ElapsedMilliseconds);

            Console.ReadKey();
        }
    }
}