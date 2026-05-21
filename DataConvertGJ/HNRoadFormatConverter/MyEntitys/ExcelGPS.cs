using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.MyEntitys
{
    public class ExcelGPS
    {
        public double _actualMile; // 新增：基于经纬度计算的累计距离

        /// <summary>
        /// utc时间
        /// </summary>
        public string _utctime;

        /// <summary>
        /// 纬度
        /// </summary>
        public string _latitude;

        /// <summary>
        /// 经度
        /// </summary>
        public string _longitude;

        /// <summary>
        /// 高程
        /// </summary>
        public string _elevation;

        /// <summary>
        /// 桩号
        /// </summary>
        public int _mile;
        public ExcelGPS()
        { }

        public ExcelGPS(string str)
        {
            string[] strs = str.Split(' ');
            _utctime = strs[0];
            _longitude = strs[1];
            _latitude = strs[2];
            _elevation = strs[3];
            _mile = int.Parse(strs[5]);

        }
        // 构造时计算实际里程（假设已有前一点的实际里程）
        public ExcelGPS(string data, double prevActualMile, ExcelGPS prevPoint)
        {
            string[] parts = data.Split(' ');
            _utctime = parts[0];
            _longitude = parts[1];
            _latitude = parts[2];
            _elevation = parts[3];
            _mile = int.Parse(parts[5]);
            // 计算实际空间距离（新增）
            if (prevPoint != null)
            {
                double dist = CalculateDistance(
                    double.Parse(prevPoint._longitude),
                    double.Parse(prevPoint._latitude),
                    double.Parse(this._longitude),
                    double.Parse(this._latitude)
                );
                _actualMile = prevActualMile + dist;
            }
            else
            {
                _actualMile = 0;
            }
        }
        private static double CalculateDistance(double lon1, double lat1, double lon2, double lat2)
        {
            const double R = 6371000; // 地球半径(米)
            var dLat = (lat2 - lat1) * Math.PI / 180;
            var dLon = (lon2 - lon1) * Math.PI / 180;
            var a = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) +
                    Math.Cos(lat1 * Math.PI / 180) * Math.Cos(lat2 * Math.PI / 180) *
                    Math.Sin(dLon / 2) * Math.Sin(dLon / 2);
            return R * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a));
        }
    }
}
