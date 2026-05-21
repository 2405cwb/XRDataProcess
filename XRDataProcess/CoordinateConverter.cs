using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XRDataProcess
{
    public class CoordinateConverter
    {
        const double x_pi = 3.14159265358979324 * 3000.0 / 180.0;
        const double pi = 3.1415926535897932384626;
        const double a = 6378245.0;
        const double ee = 0.00669342162296594323;


        /***
 * wgs84 84年提出，大地坐标，也是原始坐标。
 * gcj02 02年提出，火星坐标，经过加密算法。大多数非百度中国地图厂商基本都是使用的火星坐标：高德，腾讯，谷歌中国cn
 * bd09  09年提出，百度坐标，经过火星坐标再次加密，相当于对大地坐标经过了二次加密。百度自己使用
 * 一般的算法，没有直接bd09->wgs84或者wgs84->bd09，都需要借助wgs84->gcj02或者gcj02->wgs84算法推导。
***/

      public static bool OutOfChina(double lng, double lat)
        {
            if (lng < 72.004 || lng > 137.8347)
            {
                return true;
            }

            if (lat < 0.8293 || lat > 55.8271)
            {
                return true;
            }

            return false;
        }


        public static double TransformLat(double lng, double lat)
        {
            
            double ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + 0.1 * lng * lat + 0.2 * Math.Sqrt(Math.Abs(lng));
            ret += (20.0 * Math.Sin (6.0 * lng * pi) + 20.0 * Math.Sin(2.0 * lng * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(lat * pi) + 40.0 * Math.Sin(lat / 3.0 * pi)) * 2.0 / 3.0;
            ret += (160.0 * Math.Sin(lat / 12.0 * pi) + 320 * Math.Sin(lat * pi / 30.0)) * 2.0 / 3.0;
            return ret;
        }

        public static double TransformLng(double lng, double lat)
        {
            double ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + 0.1 * lng * lat + 0.1 * Math.Sqrt(Math.Abs(lng));
            ret += (20.0 * Math.Sin(6.0 * lng * pi) + 20.0 * Math.Sin(2.0 * lng * pi)) * 2.0 / 3.0;
            ret += (20.0 * Math.Sin(lng * pi) + 40.0 * Math.Sin(lng / 3.0 * pi)) * 2.0 / 3.0;
            ret += (150.0 * Math.Sin(lng / 12.0 * pi) + 300.0 * Math.Sin(lng / 30.0 * pi)) * 2.0 / 3.0;
            return ret;
        }
        /// <summary>
        ///  WGS-84 坐标系（全球定位系统标准坐标系）转换为 GCJ-02 坐标系（中国大陆使用的“火星坐标系”）
        /// </summary>
        /// <param name="lng"></param>
        /// <param name="lat"></param>
        /// <param name="mglng"></param>
        /// <param name="mglat"></param>
        public static void Wgs84ToGcj02(double lng, double lat, out double mglng, out double mglat)
        {
            if (OutOfChina(lng, lat))
            {
                mglng = lng;
                mglat = lat;
                return;
            }

            double dlat = TransformLat(lng - 105.0, lat - 35.0);
            double dlng = TransformLng(lng - 105.0, lat - 35.0);
            double radlat = lat / 180.0 * pi;
            double magic = Math.Sin(radlat);
            magic = 1 - ee * magic * magic;
            double sqrtmagic = Math.Sqrt(magic);
            dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * pi);
            dlng = (dlng * 180.0) / (a / sqrtmagic * Math.Cos(radlat) * pi);
            mglat = lat + dlat;
            mglng = lng + dlng;
        }

        public static void gcj02towgs84(double lng, double lat, out double mglng, out double mglat)
        {
            if (OutOfChina(lng, lat))
            {
                mglng = lng;
                mglat = lat;
                return;
            }

            double dlat = TransformLat(lng - 105.0, lat - 35.0);
            double dlng = TransformLng(lng - 105.0, lat - 35.0);
            double radlat = lat / 180.0 * pi;
            double magic = Math.Sin(radlat);
            magic = 1 - ee * magic * magic;
            double sqrtmagic = Math.Sqrt(magic);
            dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * pi);
            dlng = (dlng * 180.0) / (a / sqrtmagic * Math.Cos(radlat) * pi);
            mglat = lat + dlat;
            mglng = lng + dlng;

            mglng = lng * 2 - mglng;
            mglat = lat * 2 - mglat;
        }
        /*
         将 GCJ-02 坐标（中国大陆使用的"火星坐标系"）转换回 WGS-84 坐标（全球定位系统标准坐标）
        // GCJ02这个接口反算更准确一点;
         */
        public static void gcj02towgs84Extra(double lon, double lat,out double mglng,out double mglat)
        {
            if (OutOfChina(lon, lat))
            {
                Wgs84ToGcj02(lon, lat,out mglng,out mglat);

            }

            double initDelta = 0.01;
            double threshold = 0.000001;
            double dLat = initDelta, dLon = initDelta;
            double mLat = lat - dLat, mLon = lon - dLon;
            double pLat = lat + dLat, pLon = lon + dLon;
            double tmpLat, tmpLon;
            double wgsLat = 0;
            double wgsLon = 0;
            int i = 0;

            do
            {
                wgsLat = (mLat + pLat) / 2;
                wgsLon = (mLon + pLon) / 2;
                Wgs84ToGcj02(wgsLon, wgsLat, out tmpLon, out tmpLat);
                dLat = tmpLat - lat;
                dLon = tmpLon - lon;

                if (Math.Abs(dLat) < threshold && Math.Abs(dLon) < threshold)
                    break;

                if (dLat > 0) pLat = wgsLat; else mLat = wgsLat;
                if (dLon > 0) pLon = wgsLon; else mLon = wgsLon;

            } while (++i <= 1000);
            mglng = wgsLon;
            mglat = wgsLat;
        }

        /// <summary>
        /// 将 GCJ-02 坐标（中国大陆地图服务使用的坐标系，又称"火星坐标系"）转换为 BD-09 坐标（百度地图使用的坐标系）
        /// </summary>
        /// <param name="ggLon"></param>
        /// <param name="ggLat"></param>
        /// <param name="bdLon"></param>
        /// <param name="bdLat"></param>
        public static void gcj02ToBd09(double ggLon, double ggLat, out double bdLon, out double bdLat)
        {
            double x = ggLon, y = ggLat;
            double z = Math.Sqrt(x * x + y * y) + 0.00002 * Math.Sin(y * x_pi);
            double theta = Math.Atan2(y, x) + 0.000003 * Math.Cos(x * x_pi);
            bdLon = z * Math.Cos(theta) + 0.0065;
            bdLat = z * Math.Sin(theta) + 0.006;
        }

        /// <summary>
        ///  BD-09 坐标系（百度地图使用的坐标系）转换为 GCJ-02 坐标系（中国大陆地图服务使用的“火星坐标系”）
        /// </summary>
        /// <param name="bd_lon"></param>
        /// <param name="bd_lat"></param>
        /// <param name="gcjLon"></param>
        /// <param name="gcjLat"></param>
        public static void bd09ToGcj02(double bd_lon, double bd_lat, out double gcjLon, out double gcjLat)
        {
            double x = bd_lon - 0.0065, y = bd_lat - 0.006;
            double z = Math.Sqrt(x * x + y * y) - 0.00002 * Math.Sin(y * x_pi);
            double theta = Math.Atan2(y, x) - 0.000003 * Math.Cos(x * x_pi);
            double gg_lon = z *Math.Cos(theta);
            double gg_lat = z * Math.Sin(theta);

            gcjLon = gg_lon;
            gcjLat = gg_lat;
           
        }

    }
}
