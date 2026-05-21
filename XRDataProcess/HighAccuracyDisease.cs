using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace XRDataProcess
{
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi/*,Pack=1*/)]//注意此处对齐方式,不能用1字节对齐
    public class HighAccuracyDisease
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
        public string DiseaseName;
        public int Mile;
        /// <summary>
        /// 左上
        /// </summary>
        public HighAccureacyGps HighAccureacyGpsP0;
        /// <summary>
        /// 右上
        /// </summary>
        public HighAccureacyGps HighAccureacyGpsP1;
        /// <summary>
        /// 左下
        /// </summary>
        public HighAccureacyGps HighAccureacyGpsP2;
        /// <summary>
        /// 右下
        /// </summary>
        public HighAccureacyGps HighAccureacyGpsP3;
        public HighAccureacyGps HighAccureacyGpsCenter;

    }
    public struct HighAccureacyGps
    {
        public HighAccureacyGps(double lon, double lat, double height)
        {
            this.DiseaseLon = lon;
            this.DiseaseLat = lat;
            this.DiseaseHeight = height;
        }
        /// <summary>
        /// 经度
        /// </summary>
        public double DiseaseLon { get; set; }
        /// <summary>
        /// 纬度
        /// </summary>
        public double DiseaseLat { get; set; }
        public double DiseaseHeight { get; set; }
    }
}
