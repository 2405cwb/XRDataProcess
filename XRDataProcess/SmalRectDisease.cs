using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
namespace XRDataProcess
{
    class SmalRectDisease
    {
        XRSetting _Setting = XRSetting.GetInstance();
        RoadConfig _RoadConfig = RoadConfig.GetInstance();
        VillageHandleCoord _VillageHandleCoord = VillageHandleCoord.getInstance();
        //记录所有病害的矩形框信息
        private static List<PartRectInfo> PartRectInfos = new List<PartRectInfo>();
        /// <summary>
        /// 病害的位置，多个小矩形的编号
        /// </summary>
        public string dispos;
        public double m_DistanceToRight { get; set; } = 0;
        /// <summary>
        /// 病害的第一个矩形编号
        /// </summary>
        public int FirstRectNum;

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
        /// 路面图像名
        /// </summary>
        public string imgname;

        /// <summary>
        /// 路面图像相对路径
        /// </summary>
        public string imgpath;

        /// <summary>
        /// 小方格编号数组
        /// </summary>
        public List<int> rectarry = new List<int>();

        public bool selectfg = false;

        /// <summary>
        /// 从txt解析病害数据是否成功，有时候自动识别软件输出的txt信息不全导致解析失败异常
        /// </summary>
        public bool isDiseaseOK = true;

        public SmalRectDisease()
        {
            if (_VillageHandleCoord!=null)
            {
                if (!_VillageHandleCoord.Isjz)
                {
                    if (PartRectInfos.Count <= 0)
                    {
                        initPartRectInfo();
                    }
                }
            }
               
                
            
        }
       
            
        public SmalRectDisease(String line, int mile)
        {
            
            isDiseaseOK = true;
            m_mile = mile;
            if (_VillageHandleCoord!=null)
            {
                if (!_VillageHandleCoord.Isjz)
                {
                    if (PartRectInfos.Count <= 0)
                    {
                        initPartRectInfo();
                    }
                }
            }
          
            SetDisInfoValFromTXT(line);

            
        }

        private void initPartRectInfo()
        {
            int len = _RoadConfig.PartWidthNum * _RoadConfig.PartHeightNum;
            for (int i = 0; i < len; ++i)
            {
                PartRectInfos.Add(new PartRectInfo(i, new Rectangle()));
            }
            //添加rect
            _RoadConfig.PartImgWidth = (int)(_RoadConfig.ImageWidth * 1.0 / _RoadConfig.PartWidthNum);
            _RoadConfig.PartImgHeight = (int)(_RoadConfig.ImageHeight * 1.0 / _RoadConfig.PartHeightNum);

            for (int i = 0; i < _RoadConfig.PartHeightNum; ++i)
            {
                for (int j = 0; j < _RoadConfig.PartWidthNum; ++j)
                {
                    Rectangle timgrect = new Rectangle(j * _RoadConfig.PartImgWidth, i * _RoadConfig.PartImgHeight, _RoadConfig.PartImgWidth, _RoadConfig.PartImgHeight);
                    //从图像坐标系转换到控件坐标系
                    //Rectangle tpicrect = Img2Box(timgrect, RoadImgRect.Width, RoadImgRect.Height, _RoadConfig.ImageWidth, _RoadConfig.ImageHeight);
                    PartRectInfos[i * _RoadConfig.PartWidthNum + j].SetRect(timgrect);
                }
            }


        }
        public void SetDisInfoValFromTXT(String line)
        {
            String[] s = line.Split(' ');
            char[] csplit = { '-' };

            RoadDisType = s[0];
            RoadType = s[2];
            dispos = s[3];
            string[] t = s[3].Split(csplit, StringSplitOptions.RemoveEmptyEntries);
            t = t.Select(oneStr => oneStr.Split('.').First()).ToArray();
            rectarry.AddRange(Array.ConvertAll<string, int>(t, int.Parse));
            rectarry.Sort(delegate(int x, int y) { return x.CompareTo(y); });
            if (rectarry.Count == 0)
            {
                isDiseaseOK = false;
                return;
            }

            FirstRectNum = rectarry[0];
            string fulltype = string.Format("{0}.{1}", RoadType, RoadDisType);

            int typeidx = 0;
            bool res = RoadDiseaseTypes.DiseaseTypeDict[RoadDiseaseTypes.roadtypedict[RoadType]].TryGetValue(fulltype, out typeidx);
            if (res)
            {
                RoadDiseaseType type = RoadDiseaseTypes.roaddis[RoadDiseaseTypes.roadtypedict[RoadType]][typeidx];
                if (_Setting.BrokenPlatetype == 1 && type.fulltype == "水泥.破碎板")
                {
                    Area = _Setting.PlateWidth * _Setting.PlateLength;
                }

                else if (_VillageHandleCoord!=null)
                {
                    if (!_VillageHandleCoord.Isjz)
                    {
                        double tempArea = 0;
                        foreach (int idx in rectarry)
                        {
                            Rectangle rect = PartRectInfos[idx].GetRect();
                            double width_ = 0;
                            double hight_ = 0;
                            _VillageHandleCoord.getHandelCoordRect(rect, ref width_, ref hight_);
                            // tempArea += width_ * _RoadConfig.WidthScale * hight_ * _RoadConfig.HeightScale;
                            tempArea += _VillageHandleCoord.getArea(_RoadConfig);
                        }
                        Area = Math.Round(tempArea, 3);
                    }
                    else
                    {
                        Area = rectarry.Count * 0.1 * 0.1;
                    }
                    
                }
                else
                {
                    Area = rectarry.Count * 0.1 * 0.1;
                }
                
                m_DistanceToRight = (_RoadConfig.PartWidthNum - GetHorCetreRectWidth()) * 0.1;
            }
        }
        /// <summary>
        /// 获得病害水平方向中心矩形框
        /// </summary>
        /// <returns></returns>
        private double GetHorCetreRectWidth()
        {
            double min = 0;
            double max = 0;
            if (rectarry!=null &&rectarry.Count>0)
            {
                min = rectarry[0] % _RoadConfig.PartWidthNum;
            }
          
            for (int i = 0; i < rectarry.Count; i++)
            {
                if (rectarry[i] % _RoadConfig.PartWidthNum>=max)
                {
                   max = rectarry[i] % _RoadConfig.PartWidthNum;
                }
                if (rectarry[i] % _RoadConfig.PartWidthNum <= min)
                {
                   min = rectarry[i] % _RoadConfig.PartWidthNum;
                }
            }
            return Math.Abs(max - min)/2+min;
        }
        //返回病害信息 病害类型，桩号，路面类型，病害位置
        public String GetDisInfoStr()
        {
            return string.Format("{0} 桩号:{1:K0+000} {2} {3}", RoadDisType, m_mile, RoadType, dispos);
        }

        // 返回 病害类型,桩号,方格个数,面积
        public String GetRectInfoStr()
        {
            return string.Format("{0}\n桩号:{1:K0+000}\n个数:{2}\n面积{3}", RoadDisType, m_mile, rectarry.Count, Area);
        }

        public void Update()
        {
            rectarry.Sort(delegate(int x, int y) { return x.CompareTo(y); });
            FirstRectNum = rectarry[0];
            Area = rectarry.Count * 0.1 * 0.1;

            dispos = string.Empty;
            foreach(int tidx in rectarry)
            {
                dispos = string.Format("{0}{1}-", dispos, tidx);
            }
        }
    }
}
