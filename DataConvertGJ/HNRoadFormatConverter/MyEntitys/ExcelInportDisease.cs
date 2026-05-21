using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.MyEntitys
{
    /// <summary>
    /// 道路病害信息实体类
    /// </summary>
    public class ExcelInportDisease
    {
        /// <summary>
        /// 桩号（米）
        /// </summary>
        public int Mile { get; set; }

        /// <summary>
        /// 道路车道信息
        /// </summary>
        public string RoadLane { get; set; }

        /// <summary>
        /// 病害类型
        /// </summary>
        public string DiseaseType { get; set; }

        /// <summary>
        /// 病害等级
        /// </summary>
        public string DiseaseLevel { get; set; }

        /// <summary>
        /// 病害长度（米）
        /// </summary>
        public double DiseaseLength { get; set; }

        /// <summary>
        /// 病害宽度（米）
        /// </summary>
        public double DiseaseWidth { get; set; }

        /// <summary>
        /// 病害位置（距离道路中心线的偏移）
        /// </summary>
        public double DiseaseLocation { get; set; }

        /// <summary>
        /// 病害面积（平方米）
        /// </summary>
        public double DiseaseArea { get; set; }

        /// <summary>
        /// 计算用长度（米）
        /// </summary>
        public double CalculateLength { get; set; }

        /// <summary>
        /// 计算用宽度（米）
        /// </summary>
        public double CalculateWidth { get; set; }

        /// <summary>
        /// 起始桩号（米）
        /// </summary>
        public int Smile { get; set; }

        /// <summary>
        /// 结束桩号（米）
        /// </summary>
        public int EMile { get; set; }

        /// <summary>
        /// 病害深度（毫米）
        /// </summary>
        public double DiseaseDepth { get; set; }

        /// <summary>
        /// 面积计算公式类型
        /// </summary>
        public int AreaFormula { get; set; }

        /// <summary>
        /// 病害权重系数
        /// </summary>
        public double Weight { get; set; }

        /// <summary>
        /// 道路图片名称
        /// </summary>
        public string RoadPictureName { get; set; }

        /// <summary>
        /// 道路图片路径
        /// </summary>
        public string RoadPicturePath { get; set; }

        /// <summary>
        /// 道路类型（如沥青路面、水泥路面等）
        /// </summary>
        public string RoadType { get; set; }

        /// <summary>
        /// 备注信息
        /// </summary>
        public string Remark { get; set; }

        /// <summary>
        /// 经度
        /// </summary>
        public double Longitude { get; set; }

        /// <summary>
        /// 纬度
        /// </summary>
        public double Latitude { get; set; }

        /// <summary>
        /// 海拔高度（米）
        /// </summary>
        public double Elevation { get; set; }
    }
}
