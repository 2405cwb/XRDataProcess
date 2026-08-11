using HNRoadFormatConverter.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Animation;

namespace HNRoadFormatConverter
{
    public static class IRM_Algorithm
    {
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fpath"></param>
        /// <param name="fname"></param>
        /// <param name="side">0 左 1右</param>
        /// <param name="Count"></param>
        /// <param name="spacing"></param>
        public static List<double> WorkBankIRIAlgo_withSpeed(string fpath, string outPath,int side, int Count,double spacing)
        {
             
            string sideStr = side == 0 ? "L" : "R";
            int iriSideIndex = side == 0 ? 1 : 2;
            int length = Count;

            Count = (int)(Count / spacing);
            List<double> Points = new List<double>();
 
            double DeltLen = spacing;

            //加载抽样后的路面纵断面值
            string[] sdata = File.ReadAllLines(fpath);
            sdata = sdata.Skip(1).ToArray(); // 移除第一行并重新赋值 

            int qplusenum = Convert.ToInt32(DeltLen / spacing);//250mm内有多少个编码器脉冲
            int len = sdata.Length; 
            double[] oridata = new double[len];
           


            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split(',');
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[iriSideIndex]);
                         
                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];
                           

                        }
                    }
                }
            }
            double[] Src = new double[(len + qplusenum - 1) / qplusenum];
            int[] SrcTime = new int[Src.Length]; // 新增：抽样后的时间数据

            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                Src[j] = oridata[i];
              
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
            // 原算法（保留用于复核）：
            // int num = (int)(0.25f / DX + 0.5f) + 1;
            // float num2 = (float)(num - 1) * DX;
            // num--;
            // 当 DX=0.1 时，原逻辑等价于 (Src[i]-Src[i-3])/0.3，
            // 但后续仍使用 ST100/PR100 做 0.1m 状态更新，输入步长与矩阵不匹配。
            // 现统一为相邻 DX 采样点坡度：(Src[i]-Src[i-1])/DX。
            int num = 1;
            float num2 = DX;
            if (Src.Length == 0 || Src.Length <= num)
            {
                return Points;
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
            double num8 = 0.0;

           
           
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
                // 本采样点的状态已累计完成后，按状态步数输出完整 IRI 段。
                if (num6 == Count)
                {

                    // 添加IRI值
                    Points.Add(num8);
 
                  

                    // 重置计数器
                    num4 = 0.0;
                    num6 = 0;
                }
            }
            if (num6 > 0)
            {
                Points.Add(num4 / (double)num6);
            }
            
            return Points;


        }



        public static List<double> WorkBankIRIAlgo_withSpeed(List<string> datas, int side, int Count, double spacing)
        {

            string sideStr = side == 0 ? "L" : "R";
            int iriSideIndex = side == 0 ? 1 : 2;
            int length = Count;

            Count = (int)(Count / spacing);
            List<double> Points = new List<double>();

            double DeltLen = spacing;

            //加载抽样后的路面纵断面值
            string[] sdata = datas.ToArray();
           

            int qplusenum = Convert.ToInt32(DeltLen / 0.1);//250mm内有多少个编码器脉冲
            int len = sdata.Length;
            double[] oridata = new double[len];



            string[] s;
            for (int i = 0; i < len; i++)
            {
                s = sdata[i].Split(',');
                if (s.Length > 3)
                {
                    try
                    {
                        oridata[i] = double.Parse(s[iriSideIndex]);

                    }
                    catch (System.Exception)
                    {
                        if (i > 1)
                        {
                            oridata[i] = oridata[i - 1] * 2 - oridata[i - 2];


                        }
                    }
                }
            }
            double[] Src = new double[(len + qplusenum - 1) / qplusenum];
            int[] SrcTime = new int[Src.Length]; // 新增：抽样后的时间数据

            for (int i = 0, j = 0; i < len; i += qplusenum, ++j)
            {
                Src[j] = oridata[i];

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
            // 原算法（保留用于复核）：
            // int num = (int)(0.25f / DX + 0.5f) + 1;
            // float num2 = (float)(num - 1) * DX;
            // num--;
            // 当 DX=0.1 时，原逻辑等价于 (Src[i]-Src[i-3])/0.3，
            // 但后续仍使用 ST100/PR100 做 0.1m 状态更新，输入步长与矩阵不匹配。
            // 现统一为相邻 DX 采样点坡度：(Src[i]-Src[i-1])/DX。
            int num = 1;
            float num2 = DX;
            if (Src.Length == 0 || Src.Length <= num)
            {
                return Points;
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
            double num8 = 0.0;



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
                // 本采样点的状态已累计完成后，按状态步数输出完整 IRI 段。
                if (num6 == Count)
                {

                    // 添加IRI值
                    Points.Add(num8);



                    // 重置计数器
                    num4 = 0.0;
                    num6 = 0;
                }
            }
            if (num6 > 0)
            {
                Points.Add(num4 / (double)num6);
            }

            return Points;


        }
    }
}
