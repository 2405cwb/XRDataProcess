using DevExpress.Data.Svg;
using DevExpress.Utils.About;
using DevExpress.Utils.CodedUISupport;
using NPOI.OpenXmlFormats.Encryption;
using Org.BouncyCastle.Ocsp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Xml;

namespace XRDataProcess
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

    public class StreetDiseaseType
    {
        /// <summary>
        /// 病害名，包含病害类型和程度
        /// </summary>
        public string disname;

        /// <summary>
        /// 损坏类型
        /// </summary>
        public string distype;

        /// <summary>
        /// 单位扣分
        /// </summary>
        public double unitscore;

        /// <summary>
        /// 权重
        /// </summary>
        public double weight;

        /// <summary>
        /// 影响计量
        /// </summary>
        public int unitval;

        /// <summary>
        /// 描述
        /// </summary>
        public string description;





        /// <summary>
        /// 该类病害的总处/长度
        /// </summary>
        public int sumval = 0;

        /// <summary>
        /// 快捷键
        /// </summary>
        public string shortcut;

        public StreetDiseaseType(string sname, string stype, string suc, string sw, string suv, string des, string sc)
        {
            disname = sname;
            distype = stype;
            unitscore = double.Parse(suc);
            weight = double.Parse(sw);
            unitval = int.Parse(suv);
            description = des;
            shortcut = sc;
        }
        public override string ToString()
        {
            return disname;
        }
    }

    public class UserSignMsg
    {
        public UserSignMsg( string msg,Rectangle _rect,int side)
        {
            SignRect = _rect;
            this.Side = side;
            string[] sp = msg.Split(' ');

            if (sp.Length == 3)
            {
                Mile = sp[0];
                DisName = sp[1];
                Info = sp[2];
            }
            if (sp.Length < 4)
            {
                return;
            }

            Mile = sp[0];
            DisName = sp[1];
            DisCnt = int.Parse(sp[2]);
            Info = sp[3];
            
        }
        public UserSignMsg(string msg)
        {
             string[] sp =  msg.Split(' ',',');
            if (sp.Length == 9)
            {
                Side = int.Parse(sp[0]);
                //矩形
                SignRect = new Rectangle(int.Parse(sp[1]), int.Parse(sp[2]), int.Parse(sp[3]), int.Parse(sp[4]));
                Mile = sp[5];
                DisName = sp[6];
                DisCnt = int.Parse(sp[7]);
                Info = sp[8];
            }

            if (sp.Length==3)
            {
                Mile = sp[0];
                DisName = sp[1]; 
                Info = sp[2];
            } 
            if (sp.Length == 4)
            {
                Mile = sp[0];
                DisName = sp[1];
                DisCnt = int.Parse(sp[2]);
                Info = sp[3];
            } 
          
        }

        public int Side = 0; 

        public Rectangle SignRect = Rectangle.Empty;

        /// <summary>
        /// 桩号
        /// </summary>
        public string Mile { get; set; } = "";
        /// <summary>
        /// 病害名
        /// </summary>
        public string DisName { get; set; } = "";

        public int DisCnt { get; set; } = 0;

        /// <summary>
        /// 备注
        /// </summary>
        public string Info { get; set; } = "";

        public bool isHasRect()
        {
            if (SignRect != Rectangle.Empty)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string getDisInfo()
        {
            return string.Join(",", Mile, DisName, DisCnt, Info);
        }
        public override string ToString()
        {

            if (SignRect !=  Rectangle.Empty)
            { 
                string recordLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}",Side,
                SignRect.X, SignRect.Y, SignRect.Width, SignRect.Height, Mile,DisName,DisCnt,Info);
                return recordLine;
            }
            else
            {
                return string.Join(",", Mile, DisName, DisCnt, Info);

            }
        }
        public override bool Equals(object obj)
        {
            if (obj is UserSignMsg  d)
            {
                if (d.isHasRect() && this.isHasRect())
                {
                    if (d.ToString() == this.ToString())
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
    public class StreetDisRecord
    {
        /// <summary>
        /// 桩号
        /// </summary>
        public string _mile = null;
        /// <summary>
        /// 病害名
        /// </summary>
        public string _disname = null;
        /// <summary>
        /// 扣分值
        /// </summary>
        public double _score = 0;
        /// <summary>
        /// 扣分处
        /// </summary>
        public string _disnum = null;
        /// <summary>
        /// 扣分长度
        /// </summary>
        public string _dislen = null;

        public int _nmile = 0;
        public int _ndisnum = 0;
        public int _ndislen = 0;

        /// <summary>
        /// 是否是有效病害
        /// </summary>
        public bool isOK = false;


        public int Side = -1;
        public Rectangle SignRect = Rectangle.Empty;

        public StreetDisRecord()
        {
        }

        public bool isHasRect()
        {
            if (SignRect != Rectangle.Empty)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public StreetDisRecord(string mile, string name, double score, string num, string len,int side = -1, Rectangle rect = default)
        {
            _mile = mile;
            _disname = name;
            _score = score;
            _disnum = num;
            _dislen = len;
            SignRect = rect;
            Side = side;
            TranStr2Int();
        }

        public StreetDisRecord(string info, int curmile, int type )
        {
            string[] s = info.Split(',');
            _disname = s[0];

            if (type == 0)
            {
                try
                {
                    int typeidx = DiseaseTypes.streetdisIdx[_disname];
                    isOK = true;
                }
                catch (Exception ex)
                {
                    isOK = false;
                    return;
                }
            }
            else if(type == 1)
            {
                try
                {
                    int typeidx = DiseaseTypes.roadbeddisIdx[_disname];
                    isOK = true;
                }
                catch (Exception ex)
                {
                    isOK = false;
                    return;
                }
            }

            _mile = string.Format("{0:K0000+000}", curmile);
            _score = double.Parse(s[2]);
            _disnum = s[3];
            _dislen = s[4];
            TranStr2Int();

            if (s.Length ==10)
            {
                Side = int.Parse(s[5]);
                //矩形
                SignRect = new Rectangle(int.Parse(s[6]), int.Parse(s[7]), int.Parse(s[8]), int.Parse(s[9])); 
            }

        }

        private void TranStr2Int()
        {
            _nmile = Convert.ToInt32(_mile.Replace("K", "").Replace("+", ""));
            _ndisnum = Convert.ToInt32(_disnum);
            _ndislen = Convert.ToInt32(_dislen);
        }

        public override string ToString()
        {
            if (SignRect != Rectangle.Empty)
            {
                return string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", _disname, _mile, _score, _disnum, _dislen,Side, SignRect.X, SignRect.Y, SignRect.Width, SignRect.Height);
               
            }
            else
            {
                return string.Format("{0},{1},{2},{3},{4}", _disname, _mile, _score, _disnum, _dislen);
            }
        
        }

        public string ShowString()
        {
            return string.Format("{0},{1},扣分：{2}", _mile, _disname, _score);
        }
    }



    class MyStreetDisRecord
    {
        public string RoadCode { get; set; }
        public string Direction { get; set; }
        public string RoadNum { get; set; }
        public int StartMile { get; set; }

        public int EndMile { get; set; }

        public string DisName { get; set; }

        public string DisGrad { get; set; }

        public int Area { get; set; }


        public string Mark { get; set; }

        public double Latitude { get; set; }

        public double Longitude { get; set; }

        public double Height { get; set; }
    };

    public class MyStreetMile2Path
    {
        public int Mile { get; set; }

        public string FilePath { get; set; }
    }

    public class MyStreetMile2DisInfo
    {


        public System.Drawing.Rectangle Rect { get; set; }

        public string DisInfo { get; set; }

        public int Mile { get; set; }

    }
}
