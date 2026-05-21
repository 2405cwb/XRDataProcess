using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNDtos
{
    [SugarTable("T_hefei2")]//当和数据库名称不一样可以设置表别名 指定表明
    public class HeFeiEntity:EntityBase
    {
      

        public string Unit { get; set; }
        [SugarColumn(IsNullable = false)]
        public double StartMile { get; set; }
        [SugarColumn(IsNullable = false)]
        public double EndMile { get; set; }

        public string Grad { get; set; }

        public string IsPub { get; set; }
        /// <summary>
        /// 共有路段编号
        /// </summary>
        public string PubRoadNum { get; set; }

        public string Width { get; set; }

        public string RoadType { get; set; }
    }
    /// <summary>
    /// 一行数据
    /// </summary>
    public class RowRoad
    {

        /// <summary>
        /// 道路编号
        /// </summary>
        public string RoadNum { get; set; }
        public string Unit { get; set; }
        public int StartMile { get; set; }
        public int EndMile { get; set; }
        public int Grad { get; set; }
        public string RoadType { get; set; }
        public double RoadWid { get; set; }

        /// <summary>
        /// 是否是共用路段  区分标准 是共用同时 所用路段代码G S
        /// </summary>
        public bool IsPub { get; set; }
    }
}
