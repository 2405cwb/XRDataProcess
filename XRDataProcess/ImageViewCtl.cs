using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace XRDataProcess
{
    public partial class ImageViewCtl : PictureBox
    {
        [DllImport("YuGuang.dll")]
        static extern int YG_PaintImg(IntPtr hmain);

        public ImageViewCtl()
        {
            InitializeComponent();
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            //base.OnPaint(pe);

            YG_PaintImg(this.Handle);
        }
    }
}
