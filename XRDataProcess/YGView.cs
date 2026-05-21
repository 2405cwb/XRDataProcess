using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace XRDataProcess
{
    public partial class YGView : UserControl
    {
        public event EventHandler EventUpdateImg;
        [DllImport("YuGuang.dll", EntryPoint = "YG_SetPara")]
        static extern void YG_SetPara(int ld, int dbd, int width);

        [DllImport("YuGuang.dll", EntryPoint = "YG_Mode")]
        static extern int YG_Mode(int mode);

        [DllImport("YuGuang.dll", EntryPoint = "YG_ZmGet")]
        static extern void YG_ZmGet(ref float zm, int width);

        //[DllImport("YuGuang.dll", EntryPoint = "YG_ZmGen")]
        //static extern void YG_ZmGen(IntPtr B, int width, int height);
        [DllImport("YuGuang.dll", EntryPoint = "YG_ZmGen")]
        static extern void YG_ZmGen();

        [DllImport("YuGuang.dll", EntryPoint = "YG_ZmSet")]
        static extern void YG_ZmSet(ref float zm, int width);

        [DllImport("YuGuang.dll", EntryPoint = "YG_RH_Set")]
        static extern int YG_RH_Set(int rh, ref int r, ref int d);

        //[DllImport("YuGuang.dll", EntryPoint = "YG_Adjust")]
        //static extern void YG_Adjust(int w, int h, [In, Out] IntPtr bits);

        [DllImport("YuGuang.dll", EntryPoint = "YG_Ruihua2")]
        static extern void YG_Ruihua2();

        [DllImport("YuGuang.dll", EntryPoint = "YG_Adjust2")]
        static extern void YG_Adjust2();

        [DllImport("YuGuang.dll", EntryPoint = "YG_Recover")]
        static extern void YG_Recover();

        //[DllImport("YuGuang.dll", EntryPoint = "YG_Copy")]
        //static extern void YG_Copy(int w, int h, IntPtr srcp, [In, Out] IntPtr destp);

        private const int m_imgwidth = 4096;

        public YGView()
        {
            InitializeComponent();
        }

        void DP2LP(ref double x, ref double y)
        {
            Rectangle rc = m_pic.ClientRectangle;
            double sa = (double)((m_imgwidth + 3) / 4 * 4) / rc.Width;
            x *= sa;
            y = (y - rc.Height * 0.5) / rc.Height;
            y = -2 * y;
        }

        private Point sp, ep;
        private void m_pic_MouseDown(object sender, MouseEventArgs e)
        {
            sp = e.Location;
        }

        private void m_pic_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
                return;

            int tmpwidth = (m_imgwidth + 3) / 4 * 4;
            double y0, x1, y1, t;
            int i, j;
            ep = e.Location;

            double x0 = sp.X;
            y0 = sp.Y;
            x1 = ep.X;
            y1 = ep.Y;

            DP2LP(ref x0, ref y0);
            DP2LP(ref x1, ref y1);
            if (x0 > x1) { t = x0; x0 = x1; x1 = t; }
            if (y0 > y1) { t = y0; y0 = y1; y1 = t; }

            float[] zm = new float[tmpwidth];
            YG_ZmGet(ref zm[0], tmpwidth);
            for (i = tmpwidth - 1; --i > 0; )
            {
                if (i > x1 || i < x0)
                    continue;
                if (zm[i] > y1 || zm[i] < y0)
                    continue;
                zm[i] = -2;
            }
            for (i = tmpwidth - 1; --i > 0; )
            {
                if (zm[i] != -2)
                    continue;
                j = i;
                while (zm[--j] == -2) ;
                zm[i] = (zm[i + 1] * (i - j) + zm[j]) / (i - j + 1);
            }
            YG_ZmSet(ref zm[0], tmpwidth);
            ParaValueChanged(null, null);
            m_pic.Invalidate();
        }

        void m_pic_Paint(object sender, PaintEventArgs e)
        {
            int i, j; double k, b, sa;
            Rectangle rc = new Rectangle(0, 0, m_pic.Width, m_pic.Height);
            Point[] ptA = new Point[rc.Width];
            int tmpwidth = (m_imgwidth + 3) / 4 * 4;

            k = -rc.Height / 2;
            b = rc.Height / 2;
            sa = (double)tmpwidth / rc.Width;

            float[] A = new float[tmpwidth];
            YG_ZmGet(ref A[0], tmpwidth);
            for (i = rc.Width; --i >= 0; )
            {
                ptA[i].X = i;
                j = (int)(i * sa);
                ptA[i].Y = (int)(A[j] * k + b);
            }
            e.Graphics.FillRectangle(Brushes.White, rc);
            e.Graphics.DrawLines(Pens.Black, ptA);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            m_pic.Invalidate();
        }

        //对比度或亮度变化消息响应函数
        private void ParaValueChanged(object sender, EventArgs e)
        {
            int ld, dbd;
             ld = (int)m_n_ld.Value;
            dbd = (int)m_n_dbd.Value;
            YG_SetPara(ld, dbd, (m_imgwidth + 3) / 4 * 4);

            YG_Recover();
            bool IsOriImg = true;
            if (m_b_YG.Checked || m_b_zm.Checked)
            {
                IsOriImg = false;
                YG_Adjust2();
            }
            if (锐化.Checked)
            {
                IsOriImg = false;
                YG_Ruihua2();
            }
            EventUpdateImg(IsOriImg, null);
        }

        // 使用照明参数设置
        private void ModeChanged(object sender, EventArgs e)
        {
            int mode = 0;
            if (m_b_YG.Checked)
                mode = 1;
            if (m_b_zm.Checked)
                mode = 3;
            YG_Mode(mode);
            ParaValueChanged(null, null);
        }

        public void InitNewImg()
        {
            if (m_b_zm.Checked)
            {
                按钮_照明分析_Click(null, EventArgs.Empty);
            }
            else
            {
                ParaValueChanged(null, null);
            }
        }

        private void 按钮_照明分析_Click(object sender, EventArgs e)
        {
            try
            {
                YG_ZmGen();
                ParaValueChanged(null, null);
                m_pic.Invalidate();
            }
            catch (System.Exception ex)
            {
                Console.Write(ex.ToString());
            }
        }

        private void 锐化_ValueChanged(object sender, EventArgs e)
        {
            if (锐化.Checked)
            {
                int d;
                int r = (int)锐化半径.Value;
                d = (int)锐化强度.Value;
                int rh = 锐化.Checked ? 1 : 0;
                YG_RH_Set(rh, ref r, ref d);
            }
            ParaValueChanged(null, null);
        }
    }
}
