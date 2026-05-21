using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace XRDataProcess
{
    public class Disease
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();
        VillageHandleCoord _villageHandleCoord = VillageHandleCoord.getInstance();

        /// <summary>
        /// 病害框在原始图像中的像素位置及长宽
        /// </summary>
        public Rectangle rect;

        /// <summary>
        /// 病害类型，不带路面材质
        /// </summary>
        public String RoadDisType;

        /// <summary>
        /// 路面类型
        /// </summary>
        public String RoadType;

        /// <summary>
        /// 病害面积
        /// </summary>
        public double Area;

        /// <summary>
        /// 病害桩号
        /// </summary>
        public int m_mile;

        /// <summary>
        /// 病害程度
        /// </summary>
        public string degree;

        /// <summary>
        /// 病害深度
        /// </summary>
        public double depth = 0;

        /// <summary>
        /// 病害框的宽度
        /// </summary>
        public double realwidth;

        /// <summary>
        /// 病害框的长度
        /// </summary>
        public double realheight;

        /// <summary>
        /// 计算面积宽度
        /// </summary>
        public double calcwidth;

        /// <summary>
        /// 计算面积长度
        /// </summary>
        public double calcheight;

        /// <summary>
        /// 路面图像名
        /// </summary>
        public string imgname;

        /// <summary>
        /// 面积计算方式
        /// </summary>
        public int computetype = 0;

        /// <summary>
        /// 病害备注 
        /// </summary>
        public string remarks = "";

        /// <summary>
        /// 路面图像相对路径
        /// </summary>
        public string imgpath;

        public Disease(String line, int mile)
        {
            m_mile = mile;
            SetDisInfoValFromTXT(line);
        }

        public Disease()
        {
            rect = new Rectangle(0, 0, 0, 0);
            RoadType = null;
            imgname = null;
            imgpath = null;
        }

        public void SetDisInfoValFromTXT(String line)
        {
            String[] s = line.Split(' ');
            String[] t;
           
            rect = new Rectangle(int.Parse(s[0]), int.Parse(s[1]), int.Parse(s[2]), int.Parse(s[3]));
            RoadType = s[6];
            RoadDisType = s[4];

            if (s.Length > 7)
            {
                remarks = s[7];
            }
            else
            {
                remarks = "";
            }

            string fulltype = string.Format("{0}.{1}", RoadType, RoadDisType);

            t = RoadDisType.Split('.');
            if (t.Length > 1)
            {
                degree = t[1];
            }
            else
            {
                degree = "无";
            }
            double width_ = 0;
            double hight_ = 0;
            //低配版农村路
            //矫正后面积
            double correctArea = 0;
            bool needCorrect = false;
            if (_villageHandleCoord!=null)
            {
                if (!_villageHandleCoord.Isjz)
                {
                    needCorrect = true;
                    _villageHandleCoord.getHandelCoordRect(rect, ref width_, ref hight_);
                    realwidth = width_ * _RoadConfig.WidthScale;
                    realheight = hight_ * _RoadConfig.HeightScale;
                    correctArea = _villageHandleCoord.getArea(_RoadConfig);
                }
                else
                {
                    needCorrect = false;
                    realwidth = rect.Width * _RoadConfig.WidthScale;
                    realheight = rect.Height * _RoadConfig.HeightScale;
                }
            }
            else
            {
                needCorrect = false;
                realwidth = rect.Width * _RoadConfig.WidthScale;
                realheight = rect.Height * _RoadConfig.HeightScale;
            }


            //城镇道路，客户提出要把修补单独列出来
            if (_Setting.IsRepair = true && _Setting.ParmStyle == StandardParmType.CityRoad && RoadDisType == "修补")
            {
                calcwidth = realwidth;
                calcheight = realheight;
                Area = realwidth * realheight;
            }
            else
            {
                int typeidx = 0;
                bool res = RoadDiseaseTypes.DiseaseTypeDict[RoadDiseaseTypes.roadtypedict[RoadType]].TryGetValue(fulltype, out typeidx);
                if (res)
                {
                    RoadDiseaseType type = RoadDiseaseTypes.roaddis[RoadDiseaseTypes.roadtypedict[RoadType]][typeidx];
                    computetype = type.computetype;

                    ///面积公式：
                    ///0.框的面积长X宽，
                    ///1.框的对角线X影响宽度，
                    ///2.一个框1m2，
                    ///3.框的长边X影响宽度，
                    ///4.框沿路的方向的边长X影响宽度
                    ///5.板块长度X板块宽度
                    switch (computetype)
                    {
                        case 0: //直接就是框的面积
                            {
                                if (needCorrect)
                                {
                                    //高低等级农村路 直接根据四个点算出面积
                                    calcwidth = realwidth;
                                    calcheight = realheight;
                                    Area = correctArea;
                                }
                                else
                                {
                                    calcwidth = realwidth;
                                    calcheight = realheight;
                                    Area = calcwidth * calcheight;
                                    if (type.usearea != 0)
                                    {
                                        if (Area < type.usearea)
                                        {
                                            calcwidth = 0;
                                            calcheight = 0;
                                            Area = calcwidth * calcheight;
                                        }
                                    }
                                }
                                break;
                            }
                        case 1: //框的对角线 X 影响宽度
                            {
                                double tlen = Math.Sqrt(realwidth * realwidth + realheight * realheight);
                                if (type.uselength != 0)
                                {
                                    if (tlen < Convert.ToDouble(type.uselength))
                                    {
                                        calcwidth = 0;
                                        calcheight = 0;
                                    }
                                    else
                                    {
                                        calcwidth = Convert.ToDouble(type.usewidth);
                                        calcheight = tlen;
                                    }
                                }
                                else
                                {
                                    calcwidth = Convert.ToDouble(type.usewidth);
                                    calcheight = tlen;
                                }
                                Area = calcwidth * calcheight;
                                break;
                            }
                        case 2://框的个数
                            {
                                calcwidth = 1.0;
                                calcheight = Convert.ToDouble(type.usewidth);
                                Area = calcwidth * calcheight;
                                break;
                            }
                        case 3://框的长边X影响宽度
                            {
                                calcwidth = Convert.ToDouble(type.usewidth);
                                calcheight = Math.Max(realwidth, realheight);
                                Area = calcwidth * calcheight;
                                break;
                            }
                        case 4://框沿路的方向的边长X影响宽度
                            {
                                calcwidth = Convert.ToDouble(type.usewidth);
                                calcheight = realheight;
                                Area = calcwidth * calcheight;
                                break;
                            }
                        case 5://板块长度X宽度
                            {
                                if (_Setting.BrokenPlatetype == 0)
                                {
                                    calcwidth = realwidth;
                                    calcheight = realheight;
                                }
                                else if (_Setting.BrokenPlatetype == 1)
                                {
                                    calcwidth = _Setting.PlateWidth;
                                    calcheight = _Setting.PlateLength;
                                }
                                Area = calcwidth * calcheight;
                                break;
                            }
                    }
                }
                else
                {
                    Area = 0;
                }
            }
        }

        public String GetDisInfoStr()
        {
            return string.Format("{0} {1} {2} {3} {4} 桩号:{5:K0+000} {6} {7}",
                rect.Location.X,
                rect.Location.Y,
                rect.Width,
                rect.Height,
                RoadDisType,
                m_mile,
                RoadType,
                remarks);
        }

        public String GetRectInfoStr()
        {
            double width_ = 0;
            double hight_ = 0;
            double withval = rect.Height * _RoadConfig.HeightScale;
            double heightval = rect.Width * _RoadConfig.WidthScale;
            if (_villageHandleCoord != null)
            {
                if (!_villageHandleCoord.Isjz)
                {
                    _villageHandleCoord.getHandelCoordRect(rect, ref width_, ref hight_);
                    realwidth = width_ * _RoadConfig.WidthScale;
                    withval = width_ * _RoadConfig.WidthScale;
                    heightval = hight_ * _RoadConfig.HeightScale;
                }
            }
           
            return string.Format("{0}\n桩号:{1:K0+000}\n病害框长度:{2:0.000}m\n病害框宽度:{3:0.000}m\n{4}",
                RoadDisType,
                m_mile,
                Math.Max(withval, heightval),
                Math.Min(withval, heightval),
                remarks);
        }
    }
}
