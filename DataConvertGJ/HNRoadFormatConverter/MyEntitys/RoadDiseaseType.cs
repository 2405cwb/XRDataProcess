using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.MyEntitys
{
    public class RoadDiseaseType
    {
        public RoadDiseaseType(string sr, string sd, string sdg, string we, string uw, string uh, string ua, string ct, string sw, string sc)
        {
            roadtype = sr;
            disname = sd;
            degree = sdg;
            weight = double.Parse(we);
            usewidth = float.Parse(uw);
            uselength = float.Parse(uh);
            usearea = float.Parse(ua);
            computetype = int.Parse(ct);
            shortcut = sc;

            if (degree != null)
            {
                fulltype = string.Format("{0}.{1}.{2}", roadtype, disname, degree);
            }
            else
            {
                fulltype = string.Format("{0}.{1}", roadtype, disname);
            }

            if (sw == "0")
            {
                isshow = false;
            }
            else if (sw == "1")
            {
                isshow = true;
            }
        }

        //路面类型，病害名称，权重，显示
        public RoadDiseaseType(string sr, string sd, string we, string sw, string sc)
        {
            roadtype = sr;
            disname = sd;
            shortcut = sc;

            weight = double.Parse(we);
            if (degree != null)
            {
                fulltype = string.Format("{0}.{1}.{2}", roadtype, disname, degree);
            }
            else
            {
                fulltype = string.Format("{0}.{1}", roadtype, disname);
            }

            if (sw == "0")
            {
                isshow = false;
            }
            else if (sw == "1")
            {
                isshow = true;
            }
        }

        //2001版本水泥的A系数和B系数
        public RoadDiseaseType(string sr, string sd, double para_a, double para_b, string sw, string sc)
        {
            roadtype = sr;
            disname = sd;
            para_A = para_a;
            para_B = para_b;
            shortcut = sc;

            if (sw == "0")
            {
                isshow = false;
            }
            else if (sw == "1")
            {
                isshow = true;
            }
            if (degree != null)
            {
                fulltype = string.Format("{0}.{1}.{2}", roadtype, disname, degree);
            }
            else
            {
                fulltype = string.Format("{0}.{1}", roadtype, disname);
            }
        }

        override public string ToString()
        {
            if (degree != null)
            {
                return string.Format("{0}.{1}.{2}", roadtype, disname, degree);
            }
            else
            {
                return string.Format("{0}.{1}", roadtype, disname);
            }
        }
        /// <summary>
        /// 2001等级公路的计算单项扣分值的AB系数
        /// </summary>
        public double para_A;
        public double para_B;
        /// <summary>
        /// 板块数
        /// </summary>
        public double platenum;
        /// <summary>
        /// 病害所属路面类型 路面类型.病害名.病害程度
        /// </summary>
        public string fulltype;

        /// <summary>
        /// 病害所属路面类型
        /// </summary>
        public string roadtype;

        /// <summary>
        /// 病害名
        /// </summary>
        public string disname;

        /// <summary>
        /// 病害程度
        /// </summary>
        public string degree;

        /// <summary>
        /// 权重
        /// </summary>
        public double weight;

        /// <summary>
        /// 区间段内病害总面积
        /// </summary>
        public double totalarea = 0;

        /// <summary>
        /// 病害个数
        /// </summary>
        public int count = 0;

        /// <summary>
        /// 区间段内病害总长度
        /// </summary>
        public double totallength = 0;

        /// <summary>
        /// 影响宽度
        /// </summary>
        public float usewidth = 0;

        /// <summary>
        /// 有效长度
        /// </summary>
        public float uselength = 0;

        /// <summary>
        /// 有效面积
        /// </summary>
        public float usearea = 0;

        /// <summary>
        /// 面积计算公式
        /// </summary>
        public int computetype = 0;

        /// <summary>
        /// 是否显示
        /// </summary>
        public bool isshow = true;

        /// <summary>
        /// 快捷键
        /// </summary>
        public string shortcut = null;
    }
}
