using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.MyEntitys
{
    public class MilePart
    { 
        public int dmi = 0;
        public int mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行

        /// <summary>
        /// 第i到i+1个里程区间内的路面材质， 0-沥青，1-水泥，2-砂石
        /// </summary>
        public int roadtype = -1;

        /// <summary>
        /// 第i到i+1个里程区间内的公路等级，0-高速、一级，1-二三四级 或者 0-快速路，1-主干路次干路，2-支路
        /// </summary>
        public int roaddegree = -1;
        /// <summary>
        /// 道路等级字符串
        /// </summary>
        public string degreestr;

        public int roadcross = -1; //路口单元 1-路口

        public string roadCondition = ""; //路面情况

        public List<RoadTypePart> roadtypelist = new List<RoadTypePart>();
        //合肥报表用的道路宽度 从资产表获得
        public double width;
        //合肥报表用
        public string unit;
        //合肥报表用
        public bool isPub;
        //是否是从资产表插入的
        public bool isZC;
    }
    public class RoadTypePart
    {
        public int dmi = 0;
        public int mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行
        public int roadtype = -1;//第i到i+1个里程区间内的路面材质，沥青/水泥，0-沥青，1-水泥
    }
}
