using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using OpenTK;
using OpenTK.Input;
using System.Diagnostics;
using System.Windows.Forms;

namespace XRDataProcess
{
    enum State
    {
        NONE,
        ROTATE,
        PAN
    };
    
    public class TrackBall
    {
	    private const double SQRT1_2 = 0.7071067811865476;
        private Rectangle m_screen;
        private State m_state;

        //旋转
        private Vector3 m_rotateStart, m_rotateEnd;
        public Vector3 m_rotateAxis;
        public float m_rotateAngle;

        //平移
        private Vector2 m_panStart, m_panEnd;
        public Vector3 m_panVal;
        
        public TrackBall(Rectangle screenSize)
        {
            m_screen = screenSize;            
            m_state = State.NONE;
            m_rotateStart = Vector3.Zero;
            m_rotateEnd = Vector3.Zero;
            
            m_panStart = Vector2.Zero;
            m_panEnd = Vector2.Zero;
            m_panVal = Vector3.Zero;
        }

        public void UpdataScreen(Rectangle screenSize)
        {
            m_screen = screenSize;  
        }

        private Vector3 GetMouseProjectionOnBall(int clientX, int clientY)
        {
            //计算平面鼠标点在trackball上的空间坐标
            Vector3 mouseOnBall = new Vector3(
              ((float)clientX - (float)m_screen.Width * 0.5f) / (float)(m_screen.Width * 0.5f),
               ((float)clientY - (float)m_screen.Height * 0.5f) / (float)(m_screen.Height * 0.5f),
               0.0f
            );

            double length = mouseOnBall.Length;

            if (length < SQRT1_2)
            {
                mouseOnBall.Z = (float)Math.Sqrt(1.0 - length * length);
            }
            else if (length > 1.0)
            {
                mouseOnBall.Normalize();
            }
            else
            {
                mouseOnBall.Z = (float)(0.5 / length);
            }
            return mouseOnBall;
        }

        private void RotateCamera()
        {
            //计算从起点到终点的坐标轴旋转矢量和角度
            m_rotateAngle = (float)Math.Acos(Vector3.Dot(m_rotateStart, m_rotateEnd) / m_rotateStart.Length / m_rotateEnd.Length);
            if (!float.IsNaN(m_rotateAngle) && m_rotateAngle != 0.0f)
            {
                m_rotateAxis = Vector3.Cross(m_rotateStart, m_rotateEnd); //m_rotateStart.cross(m_rotateEnd).normalize();
                m_rotateAxis.Normalize();
                if (float.IsNaN(m_rotateAxis.X))
                {
                    m_rotateAxis = Vector3.Zero; // a hack,sometimes NAN comes from "axis" and fucks up everything. Zeroing of it resolves the issue.
                }
            }
        }
        private void PanCamera()
        {
            Vector2 mouseChange = (m_panEnd - m_panStart)*0.1f;
            m_panVal = new Vector3(mouseChange);
        }

        public void Update()
        {
            if (m_state == State.ROTATE)
            {
                RotateCamera();
            }
            else if (m_state == State.PAN)
            {
                PanCamera();
            }
        }

        ///////////////////event listeners///////////////////////////////////////
        public void OpenGLTess_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                m_state = State.PAN;
                m_panStart = new Vector2((float)e.X,(float)e.Y);
                m_panEnd = m_panStart;
            }
            else if (e.Button == MouseButtons.Right)
            {
                m_state = State.ROTATE;
                m_rotateStart = GetMouseProjectionOnBall(e.X, e.Y);
                m_rotateEnd = m_rotateStart;
            }
        }
        public void OpenGLTess_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            m_state = State.NONE;
        }
        public void OpenGLTess_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (m_state == State.ROTATE)
            {
                m_rotateEnd = GetMouseProjectionOnBall(e.X, e.Y);
            }
            else if (m_state == State.PAN)
            {
                m_panEnd = new Vector2((float)e.X, (float)e.Y);
            }
            Update();
        }
    }
}
