using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace XRDataProcess
{
    
    public class MilePart
    {

        [JsonInclude] public int dmi = 0;
        [JsonInclude] public int mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行

        /// <summary>
        /// 第i到i+1个里程区间内的路面材质， 0-沥青，1-水泥，2-砂石
        /// </summary>
        [JsonInclude] public int roadtype = -1;

        /// <summary>
        /// 第i到i+1个里程区间内的公路等级，0-高速、一级，1-二三四级 或者 0-快速路，1-主干路次干路，2-支路
        /// </summary>
        [JsonInclude] public int roaddegree = -1;
        /// <summary>
        /// 道路等级字符串
        /// </summary>
        [JsonInclude] public string degreestr;

        [JsonInclude] public int roadcross = -1; //路口单元 1-路口

        [JsonInclude] public string roadCondition = ""; //路面情况

        [JsonInclude] public List<RoadTypePart> roadtypelist = new List<RoadTypePart>();
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
        [JsonInclude] public int dmi = 0;
        [JsonInclude] public int mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行
        [JsonInclude] public int roadtype = -1;//第i到i+1个里程区间内的路面材质，沥青/水泥，0-沥青，1-水泥
    }
    /// <summary>
    /// 为了服务   国检转换中TP表  桩号间隔需要为0.1
    /// </summary>
    public class MilePartD
    {
        public double dmi = 0;
        public double mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行

        /// <summary>
        /// 第i到i+1个里程区间内的路面材质， 0-沥青，1-水泥，2-砂石
        /// </summary>
        public int roadtype = -1;

        /// <summary>
        /// 第i到i+1个里程区间内的公路等级，0-高速、一级，1-二三四级 或者 0-快速路，1-主干路次干路，2-支路
        /// </summary>
        public int roaddegree = -1;
        public string degreestr;
        public int roadcross = -1; //路口单元 1-路口

        public List<RoadTypePartF> roadtypelist = new List<RoadTypePartF>();
    }
    public class RoadTypePartF
    {
        public float dmi = 0;
        public float mile = -1;//里程区间，从第一个到第n+1个 --> 报表应该有n行
        public int roadtype = -1;//第i到i+1个里程区间内的路面材质，沥青/水泥，0-沥青，1-水泥
    }
}
