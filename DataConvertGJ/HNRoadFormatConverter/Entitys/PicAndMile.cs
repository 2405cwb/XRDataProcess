using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter.Entitys
{
    public struct PicAndMile
    {
        /// <summary>
        /// 原始图片
        /// </summary>
        public string PicPath { get; set; }

        /// <summary>
        /// 结果图片名称
        /// </summary>
        public string ResultPicName { get; set; }


        public void updateResultPicName(string name)
        {
            this.ResultPicName = name;
        }
        /// <summary>
        /// 原始对应文本
        /// </summary>
        public string sourceTxt { get; set; }

        /// <summary>
        /// 转换后文本
        /// </summary>
        public string ResultTxt { get; set; } 
        public int Mile { get; set; }

        /// <summary>
        /// 校桩前桩号，按工程起点和图像采集间隔累计。
        /// </summary>
        public int BeforeCalibrationMile { get; set; }

        /// <summary>
        /// 校桩后桩号，从 Road2Mile.txt 或 Street2Mile.txt 读取。
        /// </summary>
        public int AfterCalibrationMile { get; set; }

    }
}
