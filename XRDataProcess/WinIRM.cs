using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OperateIniFile;
using DevExpress.XtraCharts;
using OpenTK;
using OpenTK.Graphics;
using OpenTK.Graphics.OpenGL;
using Framework.Other;

namespace XRDataProcess
{
    public partial class WinIRM : Form
    {
        XRSetting _Setting = XRSetting.GetInstance();

        private bool IsShowRutAOI = false;
        
        public ProjectInfo _ProjInfo;
        public string _ProjPath;
        public List<MyIRMValMile> _LIRIList = null;
        public List<MyIRMValMile> _RIRIList = null;

        public List<MyIRMValMile> _LMTDList = null;
        public List<MyIRMValMile> _RMTDList = null;
        public List<MyIRMValMile> _CMTDList = null;

        public List<MyIRMValMile> _LRutList = null;
        public List<MyIRMValMile> _RRutList = null;

        private int _curidxiri = 0;
        private int _curidxmtd = 0;
        private int _curidxrut = 0;

        private int MAXPNUM = 10;
        //折线
        private Series _liriline, _ririline;
        private Series _cmtdline, _lmtdline, _rmtdline;
        private Series _lrutline, _rrutline;

        public List<MilePart> _RoadPartIRIMTD = null;
        public List<MilePart> _RoadPartRut = null;

        public static string[] _RoadGradeStr;
        public Dictionary<string, int> _RoadGradeDict;
        
        private GLControl glc;//opengl显示3D点云的空间
        private List<Vector3d> glpoints;
        private List<Vector3d> glpoints_ROI;
        private List<Vector3d> glpoints_Rut;

        private List<MyPoint> ori_ROI;
        private List<MyPoint> ori_Rut;

        private TrackBall trackball;
        private Vector3[] CoordinateArr;//坐标轴

        string[] rutfilepaths_L = {};
        string[] rutfilepaths_R = {};

        long[] rutfilepath_L_linenum = { };
        long[] rutfilepath_R_linenum = { };

        private RutParm rutparm_L = null;
        private RutParm rutparm_R = null;
        private byte[] _rbarr = null;
        private short[] _profile = null;
        private float[] _profileZ = null;

        public WinIRM(ProjectInfo prjinfo, string ppath)
        {
            InitializeComponent();

            _ProjInfo = prjinfo;
            _ProjPath = ppath;

            _liriline = chartIRI.Series[0];
            _ririline = chartIRI.Series[1];

            _lmtdline = chartMTD.Series[0];
            _rmtdline = chartMTD.Series[1];
            _cmtdline = chartMTD.Series[2];

            _lrutline = chartRut.Series[0];
            _rrutline = chartRut.Series[1];

            _LIRIList = new List<MyIRMValMile>();
            _RIRIList = new List<MyIRMValMile>();

            _LMTDList = new List<MyIRMValMile>();
            _RMTDList = new List<MyIRMValMile>();
            _CMTDList = new List<MyIRMValMile>();

            _LRutList = new List<MyIRMValMile>();
            _RRutList = new List<MyIRMValMile>();


            _RoadPartIRIMTD = new List<MilePart>();
            _RoadPartRut = new List<MilePart>();

            if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
            {
                _RoadGradeStr = new string[4];
                _RoadGradeStr[0] = "快速路";
                _RoadGradeStr[1] = "主干路";
                _RoadGradeStr[2] = "次干路";
                _RoadGradeStr[3] = "支路";

                _RoadGradeDict = new Dictionary<string, int>();
                for (int i = 0; i < _RoadGradeStr.Length; ++i)
                {
                    _RoadGradeDict.Add(_RoadGradeStr[i], i);
                }
            }
            else
            {
                _RoadGradeStr = new string[5];
                _RoadGradeStr[0] = "高速公路";
                _RoadGradeStr[1] = "一级公路";
                _RoadGradeStr[2] = "二级公路";
                _RoadGradeStr[3] = "三级公路";
                _RoadGradeStr[4] = "四级公路";

                _RoadGradeDict = new Dictionary<string, int>();
                for (int i = 0; i < _RoadGradeStr.Length; ++i)
                {
                    _RoadGradeDict.Add(_RoadGradeStr[i], i);
                }
            }

            //OpenGL相关
            glc = new GLControl();
            glc.Load += new EventHandler(glc_Load);
            glc.Paint += new System.Windows.Forms.PaintEventHandler(glc_Paint);
            glc.MouseDown += new System.Windows.Forms.MouseEventHandler(glc_MouseDown);
            glc.MouseMove += new System.Windows.Forms.MouseEventHandler(glc_MouseMove);
            glc.MouseWheel += new System.Windows.Forms.MouseEventHandler(glc_MouseWheel);

            panel_RutPoints.Controls.Add(glc);
            glc.Dock = DockStyle.Fill;
            glpoints = new List<Vector3d>();
            glpoints_ROI = new List<Vector3d>();
            glpoints_Rut = new List<Vector3d>();
            ori_ROI = new List<MyPoint>();
            ori_Rut = new List<MyPoint>();

            trackball = new TrackBall(new Rectangle(0, 0, glc.Width, glc.Height));

            CoordinateArr = new Vector3[4]{new Vector3(0.0f,0.0f,0.0f),
                                            new Vector3(50.0f,0.0f,0.0f),
                                            new Vector3(0.0f,50.0f,0.0f),
                                            new Vector3(0.0f,0.0f,50.0f)};
        }

        private void GetIRMVal()
        {
            int disval = 10;
            DirectoryInfo prjdir = new DirectoryInfo(_ProjPath);
            List<MilePart> _RoadPart = new List<MilePart>();
            string[] _MarkVal = null;
            double[] _LIRIMeanVal = null;
            double[] _RIRIMeanVal = null;
            double[] _LMTDMeanVal = null;
            double[] _RMTDMeanVal = null;
            double[] _CMTDMeanVal = null;
            double[] _LRutMeanVal = null;
            double[] _RRutMeanVal = null;
            double[] _SRutMeanVal = null;

            MilePart spart = null;
            try
            {
                spart = new MilePart()
                { dmi = 0, roadtype = _ProjInfo._RoadType, mile = _ProjInfo._StartMile, 
                    roaddegree = _RoadGradeDict[_ProjInfo._RoadGrade], degreestr = _ProjInfo._RoadGrade 
                };
            }
            catch
            {
                if (_Setting.ParmStyle == StandardParmType.CityRoad || _Setting.ParmStyle == StandardParmType.CityRoadShanghai)
                {
                    MessageBox.Show(string.Format("【城镇道路】不包含【{0}】请检查工程数据！", _ProjInfo._RoadGrade));
                }
                else
                {
                    MessageBox.Show(string.Format("【等级公路】不包含【{0}】请检查工程数据！", _ProjInfo._RoadGrade));
                }
                System.Environment.Exit(0);
            }
            _RoadPart.Add(spart);
            GlobalExcel.GetAllMilePart(prjdir.FullName, _ProjInfo, disval, _ProjInfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
            GlobalExcel.GetMarkInfo(_ProjInfo, prjdir, _RoadPart, ref _MarkVal);

            if (_ProjInfo._IsIRIMTD)
            {
                GlobalExcel.GetIRIMeanVal(_ProjInfo, prjdir, _RoadPart, ref _LIRIMeanVal, ref _RIRIMeanVal, true);
                GlobalExcel.GetMTDMeanVal(_ProjInfo, prjdir, _RoadPart, ref _LMTDMeanVal, ref _RMTDMeanVal, ref _CMTDMeanVal, true);
            }

            if (_ProjInfo._IsRut)
            {
                GlobalExcel.GetRutMeanVal(_ProjInfo, prjdir, _RoadPart, ref _LRutMeanVal, ref _RRutMeanVal, ref _SRutMeanVal, true);
            }

            int len = _RoadPart.Count-1;
            for (int i = 0; i < len; ++i )
            {
                if (_LIRIMeanVal != null && i <_LIRIMeanVal.Length)
                    _LIRIList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _LIRIMeanVal[i]));
                if (_RIRIMeanVal != null && i < _RIRIMeanVal.Length)
                    _RIRIList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _RIRIMeanVal[i]));

                if (_LMTDMeanVal != null && i < _LMTDMeanVal.Length)
                    _LMTDList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _LMTDMeanVal[i]));
                if (_RMTDMeanVal != null && i < _RMTDMeanVal.Length)
                    _RMTDList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _RMTDMeanVal[i]));
                if (_CMTDMeanVal != null && i < _CMTDMeanVal.Length)
                    _CMTDList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _CMTDMeanVal[i]));

                if (_LRutMeanVal != null && i < _LRutMeanVal.Length)
                    _LRutList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _LRutMeanVal[i]));
                if (_RRutMeanVal != null && i < _RRutMeanVal.Length)
                    _RRutList.Add(new MyIRMValMile(_RoadPart[i].mile, _RoadPart[i + 1].mile, _RRutMeanVal[i]));
            }
        }


        private double tempjval = 0;
        private double curjval = 0;
        public void ShowVal(double jval)
        {

            if (!_IsInited)
            {
                WinIRM_Load(null, null);
            }

            //更新IRM的曲线
            try
            {
                curjval = jval;
                if (jval % 10 == 0)
                {
                    if (Math.Abs(jval - tempjval) > 30)
                    {
                        _liriline.Points.Clear();
                        _ririline.Points.Clear();

                        _lmtdline.Points.Clear();
                        _rmtdline.Points.Clear();

                        _lrutline.Points.Clear();
                        _rrutline.Points.Clear();
                    }
                    tempjval = jval;
                    _curidxiri = Convert.ToInt32(jval * 0.1);
          
                    if (_ProjInfo._Direction > 0 && _curidxiri < _LIRIList.Count - 1 /*&& jval >= _LIRIList[_curidxiri]._smile*/
                        || _ProjInfo._Direction < 0 && _curidxiri < _LIRIList.Count - 1 /*&& jval <= _LIRIList[_curidxiri]._smile*/)
                    {
                        _liriline.Points.Add(new SeriesPoint(_LIRIList[_curidxiri]._smile, _LIRIList[_curidxiri]._val));
                        _ririline.Points.Add(new SeriesPoint(_RIRIList[_curidxiri]._smile, _RIRIList[_curidxiri]._val));
                        if (_liriline.Points.Count > MAXPNUM)
                        {
                            _liriline.Points.RemoveAt(0);
                            _ririline.Points.RemoveAt(0);
                        }
                    }

                    _curidxmtd = Convert.ToInt32(jval * 0.1);
                   
                    if (_ProjInfo._Direction > 0 && _curidxmtd < _LMTDList.Count - 1 /*&& jval >= _LMTDList[_curidxmtd]._smile*/
                        || _ProjInfo._Direction < 0 && _curidxmtd < _LMTDList.Count - 1 /*&& jval <= _LMTDList[_curidxmtd]._smile*/)
                    {

                        _cmtdline.Points.Add(new SeriesPoint(_CMTDList[_curidxmtd]._smile, _CMTDList[_curidxmtd]._val));
                        _lmtdline.Points.Add(new SeriesPoint(_LMTDList[_curidxmtd]._smile, _LMTDList[_curidxmtd]._val));
                        _rmtdline.Points.Add(new SeriesPoint(_RMTDList[_curidxmtd]._smile, _RMTDList[_curidxmtd]._val));
                        if (_lmtdline.Points.Count > MAXPNUM)
                        {
                            _cmtdline.Points.RemoveAt(0);
                            _lmtdline.Points.RemoveAt(0);
                            _rmtdline.Points.RemoveAt(0);
                        }
                    }

                    _curidxrut = Convert.ToInt32(jval * 0.1);
                    if (_ProjInfo._Direction > 0 && _curidxrut < _LRutList.Count - 1 /*&& jval >= _LRutList[_curidxrut]._smile*/
                        || _ProjInfo._Direction < 0 && _curidxrut < _LRutList.Count - 1 /*&& jval <= _LRutList[_curidxrut]._smile*/)
                    {
                        _lrutline.Points.Add(new SeriesPoint(_LRutList[_curidxrut]._smile, _LRutList[_curidxrut]._val));
                        _rrutline.Points.Add(new SeriesPoint(_RRutList[_curidxrut]._smile, _RRutList[_curidxrut]._val));
                        if (_lrutline.Points.Count > MAXPNUM)
                        {
                            _lrutline.Points.RemoveAt(0);
                            _rrutline.Points.RemoveAt(0);
                        }
                    }
                }
            }
            catch(Exception e)
            {}

            //更新加载的车辙点云
            LoadRutPoints(jval);
        }

        /// <summary>
        /// 是否初始化，加载过IRM的数据
        /// </summary>
        private bool _IsInited = false;

        //  中线MTD的值 暂时取左右的平均值，后续修改
        private void WinIRM_Load(object sender, EventArgs e)
        {
            if (!_IsInited)
            {
                _IsInited = true;

                checkBox_IsShow.Visible = _Setting.IsShowAnalysis;

                GetIRMVal();
                LoadRutData();

                if (_ProjInfo._IsIRIMTD || _ProjInfo._IsRut)
                {
                    ShowVal(0);
                }   
            }     
        }

        private void LoadRutData()
        {
            if (_ProjInfo._IsRut && Directory.Exists(_ProjPath + "\\RUT\\camera0\\data"))
            {
                if (!File.Exists(_ProjPath + "\\camera0\\rutcfg.ini"))
                {
                    MessageBox.Show("丢失车辙配置文件：..\\camera0\\rutcfg.ini！");
                    return;
                }
                rutparm_L = new RutParm(_ProjPath + "\\camera0\\rutcfg.ini");
                if (_ProjInfo._RutMode == 1)
                {
                    if (!File.Exists(_ProjPath + "\\camera1\\rutcfg.ini"))
                    {
                        MessageBox.Show("丢失车辙配置文件：..\\camera1\\rutcfg.ini！");
                        return;
                    }
                    rutparm_R = new RutParm(_ProjPath + "\\camera1\\rutcfg.ini");
                }
                else
                {
                    rutparm_R = rutparm_L;
                }

                int linebytenum = rutparm_L._hpixel * rutparm_L._pixsize;
                DirectoryInfo dir = new DirectoryInfo(_ProjPath + "\\RUT\\camera0\\data");
                FileInfo[] files = dir.GetFiles("*.dtw");
                int len = files.Length;
                rutfilepaths_L = new string[len];
                rutfilepath_L_linenum = new long[len];
                long sumline = 0;
                for (int i = 0; i < len; ++i)
                {
                    rutfilepaths_L[i] = files[i].FullName;
                    sumline = sumline + files[i].Length / linebytenum;
                    rutfilepath_L_linenum[i] = sumline;
                }

                if (_ProjInfo._RutMode == 1)
                {
                    if (Directory.Exists(_ProjPath + "\\RUT\\camera1\\data"))
                    {
                        dir = new DirectoryInfo(_ProjPath + "\\RUT\\camera1\\data");
                        files = dir.GetFiles("*.dtw");
                        len = files.Length;
                        rutfilepaths_R = new string[len];
                        rutfilepath_R_linenum = new long[len];
                        sumline = 0;
                        for (int i = 0; i < len; ++i)
                        {
                            rutfilepaths_R[i] = files[i].FullName;
                            sumline = sumline + files[i].Length / linebytenum;
                            rutfilepath_R_linenum[i] = sumline;
                        }
                    }
                }

                _rbarr = new byte[linebytenum];
                _profile = new short[rutparm_L._hpixel];
                _profileZ = new float[rutparm_L._hpixel];
            }
        }

        private double _minZ = 10000;
        private double _maxZ = -10000;            
        private void LoadRutPoints(double dmival)
        {
            if (!_ProjInfo._IsRut)
            {
                return;
            }

            if (rutfilepath_L_linenum == null || rutfilepath_L_linenum.Length < 1)
            {
                return;
            }

            int fileidx = 0;
            int lineidx = 0;
            int dmiline = (int)(dmival * 10) + Convert.ToInt32(numericUpDown1.Value);
            for (int i = 0; i < rutfilepath_L_linenum.Length; ++i )
            {
                if (dmiline <= rutfilepath_L_linenum[i])
                {
                    fileidx = i;
                    if (i > 0)
                    {
                        lineidx = (int)(dmiline - rutfilepath_L_linenum[i - 1]);
                    }
                    else
                    {
                        lineidx = (int)dmival;
                    }
                    break;
                }
            }

            if (lineidx < 0)
            {
                lineidx = 0;
            }

            if (lineidx >= rutfilepath_L_linenum[rutfilepath_L_linenum.Length - 1])
            {
                lineidx = (int)(rutfilepath_L_linenum[rutfilepath_L_linenum.Length - 1] - rutfilepath_L_linenum[rutfilepath_L_linenum.Length - 2]) - 20;
            }


            if (fileidx >= rutfilepaths_L.Length)
            {
                return;
            }

            glpoints.Clear();

            if (IsShowRutAOI)
            {
                glpoints_Rut.Clear();
            }
            
            _minZ = 10000;
            _maxZ = -10000;

            float arutval = 0.0f, brutval = 0.0f, crutval = 0.0f;
            double yscale = 100.0f;
            try
            {
                yscale = double.Parse(textBox_Y.Text);
            }
            catch { }

            int lineshownum = Convert.ToInt32(numericUpDown2.Value);

            int linebytes = rutparm_L._hpixel * rutparm_L._pixsize;
            if (_ProjInfo._RutMode == 0)//单车辙
            {
                using (FileStream frstream = new FileStream(rutfilepaths_L[fileidx], FileMode.Open))
                {
                    int tmp1 = (rutparm_L._cep + rutparm_L._asp) / 2;
                    double tmp2 = 3500 / (rutparm_L._cep - rutparm_L._asp);
                    frstream.Seek(lineidx * linebytes, SeekOrigin.Begin);

                    for (int i = 0; i < lineshownum; ++i)
                    {
                        if (frstream.Read(_rbarr, 0, linebytes) > 0)
                        {
                            Buffer.BlockCopy(_rbarr, 0, _profile, 0, linebytes);
                            for (int j = 0; j < rutparm_L._hpixel; ++j)
                            {
                                _profileZ[j] = _profile[j] / rutparm_L._scaleval;
                            }

                            if (IsShowRutAOI)
                            {
                                computerut3(_profileZ, rutparm_L._gslen, rutparm_L._asp, rutparm_L._aep, rutparm_L._bsp, rutparm_L._bep,
                                    rutparm_L._csp, rutparm_L._cep, rutparm_L._threshval, ref arutval, ref brutval, ref crutval);
                            }
                           
                            for (int j = rutparm_L._asp; j < rutparm_L._cep; ++j)
                            {
                                Vector3d tpt = new Vector3d() { X = (j - tmp1) * tmp2, Y = (i - 10) * yscale, Z = _profileZ[j] };
                                glpoints.Add(tpt);
                                if (tpt.Z < _minZ)
                                {
                                    _minZ = tpt.Z;
                                }
                                if (tpt.Z > _maxZ)
                                {
                                    _maxZ = tpt.Z;
                                }
                            }

                            if (IsShowRutAOI)
                            {
                                foreach (MyPoint pt in ori_Rut)
                                {
                                    Vector3d tpt = new Vector3d() { X = (pt.px - tmp1) * tmp2, Y = (i - 10) * yscale, Z = pt.py };
                                    glpoints_Rut.Add(tpt);
                                }
                            }
                        }
                    }
                }
            }
            else if (_ProjInfo._RutMode == 1)//双车辙
            {
                int tmp1 = (rutparm_L._aep + rutparm_L._asp) / 2;
                int tmp2 = (rutparm_R._aep + rutparm_R._asp) / 2;
                double tmp3 = 4700 / (rutparm_L._aep - rutparm_L._asp + rutparm_R._aep - rutparm_R._asp);
                using (FileStream frstream = new FileStream(rutfilepaths_L[fileidx], FileMode.Open))
                {
                    frstream.Seek(lineidx * linebytes, SeekOrigin.Begin);
                    for (int i = 0; i < lineshownum; ++i)
                    {
                        if (frstream.Read(_rbarr, 0, linebytes) > 0)
                        {
                            Buffer.BlockCopy(_rbarr, 0, _profile, 0, linebytes);
                            for (int j = 0; j < rutparm_L._hpixel; ++j)
                            {
                                _profileZ[j] = _profile[j] / rutparm_L._scaleval;
                            }

                            if (IsShowRutAOI)
                            {
                                computerut(_profileZ, rutparm_L._gslen, rutparm_L._asp, rutparm_L._aep, rutparm_L._threshval);
                            }
          
                            for (int j = rutparm_L._asp; j < rutparm_L._aep; ++j)
                            {
                                Vector3d tpt = new Vector3d() { X = (j - tmp1) * tmp3 - 900, Y = (i - 10) * yscale, Z = _profileZ[j] };
                                glpoints.Add(tpt);
                                if (tpt.Z < _minZ)
                                {
                                    _minZ = tpt.Z;
                                }
                                if (tpt.Z > _maxZ)
                                {
                                    _maxZ = tpt.Z;
                                }
                            }

                            if (IsShowRutAOI)
                            {
                                foreach (MyPoint pt in ori_Rut)
                                {
                                    Vector3d tpt = new Vector3d() { X = (pt.px - tmp1) * tmp3 - 900, Y = (i - 10) * yscale, Z = pt.py };
                                    glpoints_Rut.Add(tpt);
                                }
                            }
                        }
                    }
                }

                if (fileidx >= rutfilepaths_R.Length)
                {
                    return;
                }
                using (FileStream frstream = new FileStream(rutfilepaths_R[fileidx], FileMode.Open))
                {
                    frstream.Seek(lineidx * linebytes, SeekOrigin.Begin);
                    for (int i = 0; i < lineshownum; ++i)
                    {
                        if (frstream.Read(_rbarr, 0, linebytes) > 0)
                        {
                            Buffer.BlockCopy(_rbarr, 0, _profile, 0, linebytes);
                            for (int j = 0; j < rutparm_L._hpixel; ++j)
                            {
                                _profileZ[j] = _profile[j] / rutparm_R._scaleval;
                            }

                            if (IsShowRutAOI)
                            {
                                computerut(_profileZ, rutparm_R._gslen, rutparm_R._asp, rutparm_R._aep, rutparm_R._threshval);
                            }

                            for (int j = rutparm_R._asp; j < rutparm_R._aep; ++j)
                            {
                                Vector3d tpt = new Vector3d() { X = (j - tmp2) * tmp3 + 900, Y = (i - 10) * yscale, Z = _profileZ[j] };
                                glpoints.Add(tpt);
                                if (tpt.Z < _minZ)
                                {
                                    _minZ = tpt.Z;
                                }
                                if (tpt.Z > _maxZ)
                                {
                                    _maxZ = tpt.Z;
                                }
                            }

                            if (IsShowRutAOI)
                            {
                                foreach (MyPoint pt in ori_ROI)
                                {
                                    Vector3d tpt = new Vector3d() { X = (pt.px - tmp2) * tmp3 + 900, Y = (i - 10) * yscale, Z = pt.py };
                                    glpoints_ROI.Add(tpt);
                                }

                                foreach (MyPoint pt in ori_Rut)
                                {
                                    Vector3d tpt = new Vector3d() { X = (pt.px - tmp2) * tmp3 + 900, Y = (i - 10) * yscale, Z = pt.py };
                                    glpoints_Rut.Add(tpt);
                                }
                            }
                        }
                    }
                }
            }
            else if (_ProjInfo._RutMode == 2)//单车辙
            {
                using (FileStream frstream = new FileStream(rutfilepaths_L[fileidx], FileMode.Open))
                {
                    int tmp1 = (rutparm_L._cep + rutparm_L._asp) / 2;
                    double tmp2 = 3500 / (rutparm_L._cep - rutparm_L._asp);
                    frstream.Seek(lineidx * linebytes, SeekOrigin.Begin);

                    for (int i = 0; i < lineshownum; ++i)
                    {
                        if (frstream.Read(_rbarr, 0, linebytes) > 0)
                        {
                            Buffer.BlockCopy(_rbarr, 0, _profile, 0, linebytes);
                            for (int j = 0; j < rutparm_L._hpixel; ++j)
                            {
                                _profileZ[j] = _profile[j] / rutparm_L._scaleval;
                            }

                            if (IsShowRutAOI)
                            {
                                computerut3(_profileZ, rutparm_L._gslen, rutparm_L._asp, rutparm_L._aep, rutparm_L._bsp, rutparm_L._bep,
                                    rutparm_L._csp, rutparm_L._cep, rutparm_L._threshval, ref arutval, ref brutval, ref crutval);
                            }

                            for (int j = rutparm_L._asp; j < rutparm_L._cep; ++j)
                            {
                                Vector3d tpt = new Vector3d() { X = (j - tmp1) * tmp2, Y = (i - 10) * yscale, Z = _profileZ[j] };
                                glpoints.Add(tpt);
                                if (tpt.Z < _minZ)
                                {
                                    _minZ = tpt.Z;
                                }
                                if (tpt.Z > _maxZ)
                                {
                                    _maxZ = tpt.Z;
                                }
                            }

                            if (IsShowRutAOI)
                            {
                                foreach (MyPoint pt in ori_Rut)
                                {
                                    Vector3d tpt = new Vector3d() { X = (pt.px - tmp1) * tmp2, Y = (i - 10) * yscale, Z = pt.py };
                                    glpoints_Rut.Add(tpt);
                                }
                            }
                        }
                    }
                }
            }
            glc.Invalidate();
        }

        #region OpenGL显示点云相关函数
        void glc_Load(object sender, EventArgs e)
        {
            // 设置背景色
            GL.ClearColor(System.Drawing.Color.Black);

            int w = glc.Width;
            int h = glc.Height;

            // 设置初始状态
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.Ortho(-2000, 2000, 2000, -2000, -40000, 40000);
            GL.Viewport(0, 0, w, h);
            GL.Rotate(-90, 1.0f, 0.0f, 0.0f);
        }
        void glc_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            GL.Clear(ClearBufferMask.ColorBufferBit | ClearBufferMask.DepthBufferBit);

            //GL.ClearDepth(1.0f);
            GL.Enable(EnableCap.DepthTest);
            GL.DepthFunc(DepthFunction.Lequal);

            //坐标轴
            GL.LineWidth(2);
            GL.Begin(BeginMode.Lines);
            GL.Color3(1.0f, 0.0f, 0.0f); GL.Vertex3(CoordinateArr[0]); GL.Vertex3(CoordinateArr[1]);
            GL.Color3(0.0f, 1.0f, 0.0f); GL.Vertex3(CoordinateArr[0]); GL.Vertex3(CoordinateArr[2]);
            GL.Color3(0.0f, 0.0f, 1.0f); GL.Vertex3(CoordinateArr[0]); GL.Vertex3(CoordinateArr[3]);
            GL.End();

            DrawFrameGLPoints();
            glc.SwapBuffers();
        }

        void DrawFrameGLPoints()
        {
            if (IsShowRutAOI)
            {
                //GL.LineWidth(1);
                //GL.Begin(BeginMode.Lines);
                //GL.Color3(1.0f, 0.0f, 0.0f);
                //foreach (Vector3d points in glpoints_ROI)
                //{
                //    GL.Vertex3(points);
                //}
                //GL.End();

                GL.Begin(BeginMode.Triangles);
                GL.Color3(0.0f, 1.0f, 0.0f);
                foreach (Vector3d points in glpoints_Rut)
                {
                    GL.Vertex3(points);
                }
                GL.End();
            }

            GL.PointSize(2);
            GL.Begin(BeginMode.Points);
            double cl = 0.0f;
            double scale = 1.0f / (_maxZ - _minZ + 2);
            foreach (Vector3d points in glpoints)
            {
                cl = (points.Z - _minZ + 1) * scale;
                GL.Color3(cl, 0, 1 - cl);
                GL.Vertex3(points);
            }
            GL.End();
        }

        private void glc_MouseWheel(object sender, MouseEventArgs e)
        {
            float delta = 1 + e.Delta * 0.001f;
            GL.Scale(delta, delta, delta);
            glc.Invalidate();
        }
        private void glc_MouseDown(object sender, MouseEventArgs e)
        {
            trackball.OpenGLTess_MouseDown(sender, e);
        }
        private void glc_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                trackball.OpenGLTess_MouseMove(sender, e);
                GL.Translate(trackball.m_panVal);
                glc.Invalidate();
            }
            else if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                trackball.OpenGLTess_MouseMove(sender, e);
                GL.Rotate(trackball.m_rotateAngle, trackball.m_rotateAxis);
                glc.Invalidate();
            }
        }
        #endregion

        private void pictureBox_Up_Click(object sender, EventArgs e)
        {
            int w = glc.Width;
            int h = glc.Height;
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.Ortho(-2000, 2000, 2000, -2000, -40000, 40000);
            GL.Viewport(0, 0, w, h);
            GL.Rotate(-90, 1.0f, 0.0f, 0.0f);
            glc.Invalidate();
        }

        private void pictureBox_Front_Click(object sender, EventArgs e)
        {
            int w = glc.Width;
            int h = glc.Height;
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.Ortho(-2000, 2000, 2000, -2000, -40000, 40000);
            GL.Viewport(0, 0, w, h);
            glc.Invalidate();
        }

        private void pictureBox_Right_Click(object sender, EventArgs e)
        {
            int w = glc.Width;
            int h = glc.Height;
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.Ortho(-2000, 2000, 2000, -2000, -40000, 40000);
            GL.Viewport(0, 0, w, h);
            GL.Rotate(-90, 0.0f, 1.0f, 0.0f);
            glc.Invalidate();
        }

        private void pictureBox_45_Click(object sender, EventArgs e)
        {
            int w = glc.Width;
            int h = glc.Height;
            GL.MatrixMode(MatrixMode.Projection);
            GL.LoadIdentity();
            GL.Ortho(-2000, 2000, 2000, -2000, -40000, 40000);
            GL.Viewport(0, 0, w, h);
            GL.Rotate(-30, 1.0f, 1.0f, 0.0f);
            glc.Invalidate();
        }

        private void panel_RutPoints_Resize(object sender, EventArgs e)
        {
            if (trackball == null || glc == null || glc.IsDisposed)
                return;

            Size size = glc.ClientSize;
            if (size.Width <= 0 || size.Height <= 0)
                return;

            trackball.UpdataScreen(new Rectangle(Point.Empty, size));
        }
        
        //双车辙，整条激光线计算三个车辙值
        private void computerut3(float[] line, int flen,
            int asp, int aep, int bsp, int bep, int csp, int cep, float threshval,
            ref float aval, ref float bval, ref float cval)
        {
            float[] tline = new float[line.Length];

            MyRut.pickline(ref line, asp, cep, threshval);
            MyRut.MidianAverageFileter(line, asp, cep, flen, ref tline);

            int llen = line.Length;
            MyPoint[] pt = new MyPoint[llen];
            for (int i = asp; i < cep; ++i)
            {
                pt[i] = new MyPoint(i, tline[i]);
            }
            aval = GetRutVal(pt, asp, aep, 0.1f);
            cval = GetRutVal(pt, csp, cep, 0.1f);
            //bval = GetRutVal(pt, bsp, bep - 1, 1);
            bval = Math.Max(aval, cval);
        }

        //单车辙，整条激光线计算一个车辙值
        private float computerut(float[] line, int flen, int lines, int linee, float threshval)
        {
            float[] tline = new float[line.Length];

            MyRut.pickline(ref line, lines, linee, threshval);
            MyRut.MidianAverageFileter(line, lines, linee, flen, ref tline);
            
            int llen = line.Length;
            MyPoint[] pt = new MyPoint[llen];
            for (int i = lines; i < linee; ++i)
            {
                pt[i] = new MyPoint(i, tline[i]);
            }
            return GetRutVal(pt, lines, linee - 1, 0.1f);
        }

        //取任意三点组成三角形，计算三角形最长边的高
        private float GetRutVal(MyPoint[] Pt, int lines, int linee, float thresh)
        {
            ori_ROI.Clear();
            ori_Rut.Clear();

            List<MyPoint> MiPt = new List<MyPoint>();
            float maxrut = 0.0f, tmprut = 0.0f;

            MyRut.GetMiPoint(Pt, lines, linee - 1, ref MiPt, thresh);
            MiPt.Sort(delegate(MyPoint x, MyPoint y) { return x.px.CompareTo(y.px); });

            ori_ROI.AddRange(MiPt.ToArray());

            int len = MiPt.Count;
            for (int i = 0; i < len; ++i)
            {
                for (int j = i + 1; j < len; ++j)
                {
                    for (int k = j + 1; k < len; ++k)
                    {
                        tmprut = MyRut.GetTrigH(MiPt[i], MiPt[j], MiPt[k]);
                        if (tmprut > maxrut)
                        {
                            maxrut = tmprut;
                            ori_Rut.Clear();
                            ori_Rut.Add(MiPt[i]);
                            ori_Rut.Add(MiPt[j]);
                            ori_Rut.Add(MiPt[k]);
                        }
                    }
                }
            }
            if (maxrut == 0.0f)
            {
                maxrut = tmprut;
            }
            return maxrut;
        }

        private void checkBox_IsShow_CheckedChanged(object sender, EventArgs e)
        {
            IsShowRutAOI = checkBox_IsShow.Checked;
            ShowVal(curjval);
        }

        private void button_Update_Click(object sender, EventArgs e)
        {
            ShowVal(curjval);
        }
    }

    public class RutParm
    {
        public int _partlen = 256;
        public int _asp = 0;
        public int _aep = 0;
        public int _bsp = 0;
        public int _bep = 0;
        public int _csp = 0;
        public int _cep = 0;
        public int _gslen = 0;
        public int _ThrPoint = 0;
        public float _scaleval = 10;
        public float _threshval = 0;
        public int _hpixel = 2048;
        public int _pixsize = 2;

        public RutParm()
        {

        }

        public RutParm(string inifile)
        {
            IniFiles rutcfg = new IniFiles(inifile);

            _asp = rutcfg.ReadInteger("camera", "rutastart", 0);
            _aep = rutcfg.ReadInteger("camera", "rutaend", 2048);
            _bsp = rutcfg.ReadInteger("camera", "rutbstart", 0);
            _bep = rutcfg.ReadInteger("camera", "rutbend", 2048);
            _csp = rutcfg.ReadInteger("camera", "rutcstart", 0);
            _cep = rutcfg.ReadInteger("camera", "rutcend", 2048);
            _hpixel = rutcfg.ReadInteger("camera", "hpixel", 2048);
            _pixsize = rutcfg.ReadInteger("camera", "pixsize", 2);

            _gslen = rutcfg.ReadInteger("rut", "gslen", 32) / 2 * 2 + 1;
            _scaleval = rutcfg.ReadInteger("rut", "scaleval", 10);
            _partlen = _aep - _asp - 2;//157、170、186
            _threshval = rutcfg.ReadInteger("rut", "threshval", 28);
            _ThrPoint = rutcfg.ReadInteger("rut", "threshpointnum", _partlen / 4);
        }
    }
}
