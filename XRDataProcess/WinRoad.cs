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
    public partial class WinRoad : Form
    {
        virtual public event EventHandler EventUpdateDis;
        virtual public event EventHandler EventChangeType;
        virtual public event EventHandler EventUpdateMile;
        virtual public event EventHandler EventUpdateDmi;
        virtual public event EventHandler EventUpdateYG;
        virtual public event EventHandler EventUpdateFullImg;
        virtual public event EventHandler EventUpdateFullPoint;

        public bool _IsInitLoad = false;
        public bool _IsActivated = false;
        
        public WinRoad()
        {
            InitializeComponent();
        }

        virtual public void ShowJumpImg(double jval) 
        {
        }

        virtual public void SaveDisease()
        {
        }

        virtual public void UpdateDisType(object updateinfo)
        {
        }

        virtual public void GetTypeMilePart(string projectpath, int direction)
        { }
        
        virtual public void UpdateYG(object YGStatu)
        {

        }
    }
}
