using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace XRDataProcess
{
    class PartRectInfo
    {        /// <summary>
        /// 小方格在PictureBox上绘制的矩形
        /// </summary>
        private Rectangle _Rect = new Rectangle();

        /// <summary>
        /// 当前小方格在PictureBox上的索引，从0开始，与病害记录txt的索引对应
        /// </summary>
        private int _RectIdx { get; set; }

        /// <summary>
        /// 标志当前小方格是否已经绘制为病害，false非病害，true病害
        /// </summary>
        private bool _IsCheck = false;

        /// <summary>
        /// 初始化函数
        /// </summary>
        /// <param name="idx">小方格索引</param>
        /// <param name="rectx">小方格左上角X</param>
        /// <param name="recty">小方格左上角Y</param>
        /// <param name="rectw">小方格宽度</param>
        /// <param name="recth">小方格高度</param>
        public PartRectInfo(int idx, Rectangle rect)
        {
            _RectIdx = idx;
            _Rect = rect;
            _IsCheck = false;
        }

        /// <summary>
        /// 设置小方格病害记录状态
        /// </summary>
        /// <param name="ischeck">病害记录状态</param>
        public void SetCheck(bool ischeck)
        {
            _IsCheck = ischeck;
        }

        public bool GetChek()
        {
            return _IsCheck;
        }

        public Rectangle GetRect()
        {
            return _Rect;
        }

        public int GetRectIdx()
        {
            return _RectIdx;
        }

        public void SetRect(Rectangle rect)
        {
            _Rect = rect;
        }
    }
}
