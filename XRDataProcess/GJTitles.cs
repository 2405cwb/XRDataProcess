using DevExpress.XtraMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 

namespace XRDataProcess
{
    class GJTitles
    {
        static GJTitles()
        {

        }
      
        public static void  SetPara(List<string> txt)
        {
            LbiTitle = txt[1];
            IriTitle = txt[3];
            RIFileJgTitle = txt[5];
            RIFileGdTitle = txt[7];
            LQ_2018_BIG = txt[9];
            LQ_2018_SMALL = txt[11];
            SN_2018_BIG = txt[13];
            SN_2018_SMALL = txt[15];
            LQ_NC_BIG = txt[17];
            LQ_NC_SMALL = txt[19];
            SN_NC_BIG = txt[21];
            SN_NC_SMALL = txt[23];
            SS_NC_BIG = txt[25];
            SS_NC_SMALL = txt[27];
          
            
        }
        public static string LbiTitle { get; set; }
        public  static string  IriTitle { get; set; }
        public static string RIFileJgTitle { get; set; }
        public static string RIFileGdTitle { get; set; }
        //沥青 2018  大框
        public static string LQ_2018_BIG { get; set; }
        public static string LQ_2018_SMALL { get; set; }
        public static string SN_2018_BIG { get; set; }
        public static string SN_2018_SMALL { get; set; }
        public static string LQ_NC_BIG { get; set; }
        
        public static string LQ_NC_SMALL { get; set; }

        public static string SN_NC_BIG { get; set; }
        //水泥  农村路 小框
        public static string SN_NC_SMALL { get; set; }
        public static string SS_NC_BIG { get; set; }
        public static string SS_NC_SMALL { get; set; }
    }
}
