using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XRDataProcess
{
    public partial class WinImgFull : Form
    {
      
        private Bitmap _curimg = null;
        private Point _curpoint = new Point(0,0);
        public WinImgFull()
        {
            InitializeComponent();
        }

        public void UpdateImg(object newimg)
        {
            _curimg = (Bitmap)newimg;
            pictureBox_Img.Invalidate();
        }

        public void UpdateShowImg(object centerpoint)
        {
            _curpoint = (Point)centerpoint;
            pictureBox_Img.Invalidate();
        }

        private void pictureBox_Img_Paint(object sender, PaintEventArgs e)
        {
            if (_curimg != null)
            {
                Pen p = new Pen(Color.Red, 1);
                Point p1 = new Point(0, 0);
                Point p2 = new Point(0, 0);
                int linh = 20, linw = 20;
                int srcx = 0, srcy = 0;
                int srcw = pictureBox_Img.Width, srch = pictureBox_Img.Height;
                int srcw2 = srcw / 2, srch2 = srch / 2;

                if (_curpoint.X < srcw2)
                {
                    srcx = 0;
                }
                else if (_curpoint.X > _curimg.Width - srcw2)
                {
                    srcx = _curimg.Width - srcw;
                }
                else
                {
                    srcx = _curpoint.X - srcw2;
                }

                if (_curpoint.Y < srch2)
                {
                    srcy = 0;
                }
                else if (_curpoint.Y > _curimg.Height - srch2)
                {
                    srcy = _curimg.Height - srch;
                }
                else
                {
                    srcy = _curpoint.Y - srch2;
                }
                
                Graphics g = e.Graphics;
            
                g.DrawImage(_curimg, new Rectangle(0, 0, srcw, srch), new Rectangle(srcx, srcy, srcw, srch), GraphicsUnit.Pixel);

                //linh = pictureBox_Img.Height;
                //linw = pictureBox_Img.Width;
                p1.X = srcw2 - linw/2;
                p1.Y = srch2;
                p2.X = srcw2 + linw / 2;
                p2.Y = srch2;
                g.DrawLine(p, p1, p2); //画竖线

                p1.X = srcw2;
                p1.Y = srch2 - linh / 2;
                p2.X = srcw2;
                p2.Y = srch2 + linh / 2;
                g.DrawLine(p, p1, p2);
            }
        }
    }

}
