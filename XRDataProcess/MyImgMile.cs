using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XRDataProcess
{
    public class MyImgMile
    {
        /// <summary>
        /// 图像关联桩号
        /// </summary>
        public double imgmile;

        /// <summary>
        /// 图像绝对路径
        /// </summary>
        public string imgpath;

        public MyImgMile(string info)
        {
            string[] s = info.Split(' ');
            imgmile = double.Parse(s[0]);
            imgpath = s[1];
        }
    }
}
