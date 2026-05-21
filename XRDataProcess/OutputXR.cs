using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms; 
using MSExcel = Microsoft.Office.Interop.Excel;
namespace XRDataProcess
{
    class OutputXR
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fpath"></param>
        /// <param name="diseaseNum"></param>
        /// <param name="diseasePtr"></param>
        /// <param name="gridPtr"></param>
        /// <param name="direction"></param>
        /// <returns></returns>
        [DllImport("hnDxfIO.dll", EntryPoint = "OutputXRDxf")]
        private static extern bool OutputXRDxf(string fpath, int diseaseNum, IntPtr diseasePtr, IntPtr gridPtr, int direction);
        [DllImport("hnDxfIO.dll", EntryPoint = "OutputXRDxfCityRoad")]
        private static extern bool OutputXRDxfCityRoad(string fpath, int diseaseNum, IntPtr diseasePtr, IntPtr gridPtr, int direction);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fpath"></param>
        /// <param name="diseaseNum"></param>
        /// <param name="diseasePtr"></param>
        /// <param name="gridPtr"></param>
        /// <param name="direction"></param>
        /// <returns></returns>
        [DllImport("hnDxfIO.dll", EntryPoint = "OutputXRDxfProvinceRoad")]
        private static extern bool OutputXRDxfProvinceRoad(string fpath, int diseaseNum, IntPtr diseasePtr, IntPtr gridPtr, int direction);

        [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi/*,Pack=1*/)]//注意此处对齐方式,不能用1字节对齐
        public struct diseaseInfo
        {
            public int mile;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string roadNum;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string diseaseType;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string diseaseDegree;
            public double rectHeight;
            public double rectWidth;
            public double distToCenter;
            public double diseaseArea;
            public double calcHeight;
            public double calcWidth;
            public bool bOnRoad;
        }
        //public struct diseaseInfo
        //{
        //    public int rect_top; /// 病害框在原始图像中的像素位置及长宽
        //    public int rect_left;
        //    public int rect_width;
        //    public int rect_height;
        //    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        //    public byte[] roadDisType; /// 病害类型，不带路面材质
        //    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        //    public byte[] roadType; /// 路面类型
        //    public double area;  /// 病害面积
        //    public int mile; /// 病害桩号
        //    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        //    public byte[] degree; /// 病害程度
        //    public double depth;  /// 病害深度
        //    public double realWidth;  /// 病害框的宽度
        //    public double realHeight; /// 病害框的长度
        //    public double calcwidth;
        //    public double calcheight;
        //    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        //    public byte[] imgname;
        //    public int computetype;
        //    [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        //    public byte[] remarks;
        //};

        [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public struct GridDiseaseInfo
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string strName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string strBegMile;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string strEndMile;
            public double dBegMileage;
            public double dEndMileage;
            public double dRoadWidth;
            public int dRoadTotalNum;
        }

        static XRSetting _Setting = XRSetting.GetInstance();
        static RoadConfig _RoadConfig = RoadConfig.GetInstance();
        public static List<MilePart> _RoadPart = null;
        public static List<MilePart> _RoadPart10 = null;//整10米桩号分段
        private static double[] _SpeedVal = null;
        public static List<MilePart> _RoadPart1M = null;//1米桩号分段
        public static string[] _RoadGradeStr = { "高速公路", "一级公路", "二级公路", "三级公路", "四级公路" };
        public static string[]_ExcelProjectTypeStr = {"未知类型","等级公路2018","城镇道路"};
         public static string[]_ExcelRectTypeStr = {"未知类型","大框","小方格"};
        public static Dictionary<string, int> _RoadGradeDict;
        private static double[] _SRutDisVal = null;
        private static int[] _SRutDisMile = null;
        private static Disease[] _RoadDisList = null;
        private static Disease[] _RoadRepairList = null;
        private static double[] _rutThresh = new double[2];
        private static string[] _MarkVal = null;
        private static ProjectInfo _prjinfo = null;
        #region 从当前工程读取数据
        public static bool InitProData(DirectoryInfo prjdir, ProjectInfo prjinfo, int disval)
        {
            _RoadGradeDict = new Dictionary<string, int>();
            for (int i = 0; i < _RoadGradeStr.Length; ++i)
            {
                _RoadGradeDict.Add(_RoadGradeStr[i], i);
            }

            _SpeedVal = null;
            _prjinfo=prjinfo;

            if (_RoadPart != null)
            {
                _RoadPart.Clear();
                _RoadPart = null;
            }
            _RoadPart = new List<MilePart>();

            if (_RoadPart10 != null)
            {
                _RoadPart10.Clear();
                _RoadPart10 = null;
            }
            _RoadPart10 = new List<MilePart>();

            if (_RoadPart1M != null)
            {
                _RoadPart1M.Clear();
                _RoadPart1M = null;
            }
            _RoadPart1M = new List<MilePart>();
            MilePart spart = null;
            try
            {
                spart = new MilePart() { dmi = 0, roadtype = prjinfo._RoadType, mile = prjinfo._StartMile, roaddegree = _RoadGradeDict[prjinfo._RoadGrade], degreestr = prjinfo._RoadGrade };
            }
            catch
            {
                MessageBox.Show(string.Format("【等级公路】不包含【{0}】请检查工程数据！", prjinfo._RoadGrade));
                System.Environment.Exit(0);
            }
            _RoadPart.Add(spart);
            _RoadPart10.Add(spart);
            _RoadPart1M.Add(spart);
            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, disval, prjinfo._Direction, _RoadGradeStr, ref _RoadPart, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
            GlobalExcel.GetMarkInfo(prjinfo, prjdir, _RoadPart, ref _MarkVal);


            GlobalExcel.GetAllMilePart(prjdir.FullName, prjinfo, 1, prjinfo._Direction, _RoadGradeStr, ref _RoadPart1M, RoadDiseaseTypes.roadtypedict, _RoadGradeDict);
            if (_Setting.OutRut == 1 || (_Setting.OutRut == 2 && (_RoadGradeDict[prjinfo._RoadGrade] > 1)))
            {
                GlobalExcel.GetRutDisVal(prjinfo, prjdir, _RoadPart1M, ref _SRutDisVal, ref _SRutDisMile);
            }
            GlobalExcel.GetAllDis(prjdir.FullName, prjinfo, prjinfo._Direction, _RoadGradeDict, _SRutDisVal, _SRutDisMile, ref _RoadDisList, ref _RoadRepairList, _rutThresh, _RoadPart);

            return true;
        }

        public static void OutputDxf(string path, double beginMileage, double endMileage)
        {
            List<diseaseInfo> m_listDisease = new List<diseaseInfo>();
            int len = _RoadPart.Count - 1, dlen = _RoadDisList.Length;
            bool res = false;
            int typeidx = 0;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = _RoadPart[i].mile;
                int emile = _RoadPart[i + 1].mile;
                if (endMileage <= smile) continue;
                if (beginMileage > emile) continue;
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((_prjinfo._Direction > 0 && _RoadDisList[j].m_mile >= smile && _RoadDisList[j].m_mile < emile)
                    || (_prjinfo._Direction < 0 && _RoadDisList[j].m_mile <= smile && _RoadDisList[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[_RoadPart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                      _RoadDisList[j].RoadType, _RoadDisList[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = _RoadDisList[j].RoadDisType.Split('.');
                        diseaseInfo info;
                        info.mile = _RoadDisList[j].m_mile;
                        info.roadNum = _prjinfo._RoadNum;
                        info.diseaseType = s[0];
                        if (s.Length > 1)
                        {

                            info.diseaseDegree = s[1];
                        }
                        else
                        {

                            info.diseaseDegree = "无";
                        }
                        double leftx = _RoadDisList[j].rect.X * _RoadConfig.WidthScale;
                        double rightx = (_RoadDisList[j].rect.Width + _RoadDisList[j].rect.X) * _RoadConfig.WidthScale;

                        info.rectHeight = _RoadDisList[j].rect.Height * _RoadConfig.HeightScale;
                        info.rectWidth = _RoadDisList[j].rect.Width * _RoadConfig.WidthScale;
                        info.distToCenter = (rightx + leftx - _RoadConfig.DetectWidth) / 2;
                        info.diseaseArea = _RoadDisList[j].Area;
                        info.calcHeight = _RoadDisList[j].calcheight;
                        info.calcWidth = _RoadDisList[j].calcwidth;

                        if (leftx >= _RoadConfig.DetectWidth / 2 + 1.4
                            || rightx <= _RoadConfig.DetectWidth / 2 - 1.4
                            || leftx >= _RoadConfig.DetectWidth / 2 - 0.7 && rightx <= _RoadConfig.DetectWidth / 2 + 0.7)
                        {
                            info.bOnRoad = false;
                        }
                        else
                        {
                            info.bOnRoad = true;
                        }
                        m_listDisease.Add(info);

                    }
                    else
                    {
                        //string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        //File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
                //每隔100米输出一次dxf
                string strName = string.Format("{0:K0+000}-{1:K0+000}", smile, emile);
                string dxfPath = path + "//" + strName + ".dxf";
                int diseaseCount = m_listDisease.Count;
                int ptrSize = Marshal.SizeOf(typeof(diseaseInfo));
                IntPtr diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
                IntPtr ptr = diseaseInfoPtr;
                for (int m = 0; m < diseaseCount; m++)
                {
                    Marshal.StructureToPtr(m_listDisease[m], ptr, false);
                    ptr = (IntPtr)(long)ptr + ptrSize;
                }
                GridDiseaseInfo grid = new GridDiseaseInfo();
                grid.dBegMileage = smile;
                grid.dEndMileage = emile;
                grid.strName = strName;
                grid.strBegMile = beginMileage.ToString();
                grid.strEndMile = endMileage.ToString();
                grid.dRoadTotalNum = 1;//从当前工程一次只能读取一个车道的数据
                ptrSize = Marshal.SizeOf(typeof(GridDiseaseInfo));
                IntPtr GridInfoPtr = Marshal.AllocHGlobal(ptrSize);
                Marshal.StructureToPtr(grid, GridInfoPtr, false);

                OutputXRDxf(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, _prjinfo._Direction);
                Marshal.FreeHGlobal(diseaseInfoPtr);
                Marshal.FreeHGlobal(GridInfoPtr);
                m_listDisease.Clear();
            }
           
        }

        public static void OutputDxf2(string path, double beginMileage, double endMileage)
        {
            List<diseaseInfo> m_listDisease = new List<diseaseInfo>();
            int len = _RoadPart.Count - 1, dlen = _RoadDisList.Length;
            int j=0;
            while (j < dlen)
            {
                string[] s = _RoadDisList[j].RoadDisType.Split('.');
                diseaseInfo info;
                info.mile = _RoadDisList[j].m_mile;
                info.roadNum = _prjinfo._RoadNum;
                info.diseaseType = s[0];
                if (s.Length > 1)
                {

                    info.diseaseDegree = s[1];
                }
                else
                {

                    info.diseaseDegree = "无";
                }
                double leftx = _RoadDisList[j].rect.X * _RoadConfig.WidthScale;
                double rightx = (_RoadDisList[j].rect.Width + _RoadDisList[j].rect.X) * _RoadConfig.WidthScale;

                info.rectHeight = _RoadDisList[j].rect.Height * _RoadConfig.HeightScale;
                info.rectWidth = _RoadDisList[j].rect.Width * _RoadConfig.WidthScale;
                info.distToCenter = (rightx + leftx - _RoadConfig.DetectWidth) / 2;
                info.diseaseArea = _RoadDisList[j].Area;
                info.calcHeight = _RoadDisList[j].calcheight;
                info.calcWidth = _RoadDisList[j].calcwidth;

                if (leftx >= _RoadConfig.DetectWidth / 2 + 1.4
                    || rightx <= _RoadConfig.DetectWidth / 2 - 1.4
                    || leftx >= _RoadConfig.DetectWidth / 2 - 0.7 && rightx <= _RoadConfig.DetectWidth / 2 + 0.7)
                {
                    info.bOnRoad = false;
                }
                else
                {
                    info.bOnRoad = true;
                }
                m_listDisease.Add(info);
                ++j;
            }
                  
            //每隔100米输出一次dxf
            string strName = string.Format("{0:K0+000}-{1:K0+000}", beginMileage, endMileage);
            string dxfPath = path ;
            int diseaseCount = m_listDisease.Count;
            int ptrSize = Marshal.SizeOf(typeof(diseaseInfo));
            IntPtr diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
            IntPtr ptr = diseaseInfoPtr;
            for (int m = 0; m < diseaseCount; m++)
            {
                Marshal.StructureToPtr(m_listDisease[m], ptr, false);
                ptr = (IntPtr)(long)ptr + ptrSize;
            }
            GridDiseaseInfo grid = new GridDiseaseInfo();
            grid.dBegMileage = beginMileage;
            grid.dEndMileage = endMileage;
            grid.strName = strName;
            grid.strBegMile = beginMileage.ToString();
            grid.strEndMile = endMileage.ToString();
            grid.dRoadWidth = _RoadConfig.RealWidth;
            grid.dRoadTotalNum = 1;//从当前工程一次只能读取一个车道的数据
            ptrSize = Marshal.SizeOf(typeof(GridDiseaseInfo));
            IntPtr GridInfoPtr = Marshal.AllocHGlobal(ptrSize);
            Marshal.StructureToPtr(grid, GridInfoPtr, false);

            OutputXRDxf(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, _prjinfo._Direction);
            Marshal.FreeHGlobal(diseaseInfoPtr);
            Marshal.FreeHGlobal(GridInfoPtr);
            m_listDisease.Clear();
        }

        public static void OutputDxfProvinceRoad(string path, double beginMileage, double endMileage)
        {
            List<diseaseInfo> m_listDisease = new List<diseaseInfo>();
            int len = _RoadPart.Count - 1, dlen = _RoadDisList.Length;
            bool res = false;
            int typeidx = 0;
            for (int i = 0, j = 0; i < len; i++)//i区间索引，j病害索引
            {
                int smile = _RoadPart[i].mile;
                int emile = _RoadPart[i + 1].mile;
                if (endMileage <= smile) continue;
                if (beginMileage > emile) continue;
                RoadDiseaseTypes.Clear();
                while (j < dlen && ((_prjinfo._Direction > 0 && _RoadDisList[j].m_mile >= smile && _RoadDisList[j].m_mile < emile)
                    || (_prjinfo._Direction < 0 && _RoadDisList[j].m_mile <= smile && _RoadDisList[j].m_mile > emile)))
                {
                    res = RoadDiseaseTypes.DiseaseTypeDict[_RoadPart[i].roadtype].TryGetValue(string.Format("{0}.{1}",
                      _RoadDisList[j].RoadType, _RoadDisList[j].RoadDisType), out typeidx);
                    if (res)
                    {
                        string[] s = _RoadDisList[j].RoadDisType.Split('.');
                        diseaseInfo info;
                        info.mile = _RoadDisList[j].m_mile;
                        info.roadNum = _prjinfo._RoadNum;
                        info.diseaseType = s[0];
                        if (s.Length > 1)
                        {

                            info.diseaseDegree = s[1];
                        }
                        else
                        {

                            info.diseaseDegree = "无";
                        }
                        double leftx = _RoadDisList[j].rect.X * _RoadConfig.WidthScale;
                        double rightx = (_RoadDisList[j].rect.Width + _RoadDisList[j].rect.X) * _RoadConfig.WidthScale;

                        info.rectHeight = _RoadDisList[j].rect.Height * _RoadConfig.HeightScale;
                        info.rectWidth = _RoadDisList[j].rect.Width * _RoadConfig.WidthScale;
                        info.distToCenter = (rightx + leftx - _RoadConfig.DetectWidth) / 2;
                        info.diseaseArea = _RoadDisList[j].Area;
                        info.calcHeight = _RoadDisList[j].calcheight;
                        info.calcWidth = _RoadDisList[j].calcwidth;

                        if (leftx >= _RoadConfig.DetectWidth / 2 + 1.4
                            || rightx <= _RoadConfig.DetectWidth / 2 - 1.4
                            || leftx >= _RoadConfig.DetectWidth / 2 - 0.7 && rightx <= _RoadConfig.DetectWidth / 2 + 0.7)
                        {
                            info.bOnRoad = false;
                        }
                        else
                        {
                            info.bOnRoad = true;
                        }
                        m_listDisease.Add(info);

                    }
                    else
                    {
                        //string errval = string.Format("拉框病害信息：{0} 路面材质：{1}\r\n", arrdis[j].GetDisInfoStr(), GlobalExcel._RoadTypeStr[roadpart[i].roadtype]);
                        //File.AppendAllText(errlog, errval, Encoding.UTF8);
                    }
                    ++j;
                }
                //每隔100米输出一次dxf
                string strName = string.Format("{0:K0+000}-{1:K0+000}", smile, emile);
                string dxfPath = path + "//" + strName + ".dxf";
                int diseaseCount = m_listDisease.Count;
                int ptrSize = Marshal.SizeOf(typeof(diseaseInfo));
                IntPtr diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
                IntPtr ptr = diseaseInfoPtr;
                for (int m = 0; m < diseaseCount; m++)
                {
                    Marshal.StructureToPtr(m_listDisease[m], ptr, false);
                    ptr = (IntPtr)(long)ptr + ptrSize;
                }
                GridDiseaseInfo grid = new GridDiseaseInfo();
                grid.dBegMileage = smile;
                grid.dEndMileage = emile;
                grid.strName = strName;
                grid.strBegMile = beginMileage.ToString();
                grid.strEndMile = endMileage.ToString();
                grid.dRoadTotalNum = 1;//从当前工程一次只能读取一个车道的数据
                ptrSize = Marshal.SizeOf(typeof(GridDiseaseInfo));
                IntPtr GridInfoPtr = Marshal.AllocHGlobal(ptrSize);
                Marshal.StructureToPtr(grid, GridInfoPtr, false);

                OutputXRDxfProvinceRoad(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, _prjinfo._Direction);
                Marshal.FreeHGlobal(diseaseInfoPtr);
                Marshal.FreeHGlobal(GridInfoPtr);
                m_listDisease.Clear();
            }

        }
        #endregion

        #region 从Excel文件读取数据

        public static void OutputDxfByExcel(List<string> listExcelPath, int direction,string saveFolder, double beginMile, double endMile, StandardParmType roadType)
        {
            //打开Excel文件
            MSExcel.Application excelApp = new MSExcel.Application()
            {
                Visible = true,
                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                AlertBeforeOverwriting = false
            };
            List<diseaseInfo> m_listDisease = readDataFromExcel(excelApp, listExcelPath);
  

            string strName = string.Format("{0:K0+000}-{1:K0+000}", beginMile, endMile);
            string dxfPath = saveFolder;
            if (direction > 0)
                dxfPath += "\\上行";
            else if(direction < 0)
                dxfPath += "\\下行";
            int diseaseCount = m_listDisease.Count;
            int ptrSize = Marshal.SizeOf(typeof(diseaseInfo));
            IntPtr diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
            IntPtr ptr = diseaseInfoPtr;
            for (int m = 0; m < diseaseCount; m++)
            {
                Marshal.StructureToPtr(m_listDisease[m], ptr, false);
                ptr = (IntPtr)(long)ptr + ptrSize;
            }
            GridDiseaseInfo grid = new GridDiseaseInfo();
            grid.dBegMileage = beginMile;
            grid.dEndMileage = endMile;
            grid.strName = strName;
            grid.strBegMile = beginMile.ToString();
            grid.strEndMile = endMile.ToString();
            grid.dRoadWidth = _RoadConfig.RealWidth;
            grid.dRoadTotalNum = listExcelPath.Count;
            ptrSize = Marshal.SizeOf(typeof(GridDiseaseInfo));
            IntPtr GridInfoPtr = Marshal.AllocHGlobal(ptrSize);
            Marshal.StructureToPtr(grid, GridInfoPtr, false);

            Directory.CreateDirectory(dxfPath);
            switch (roadType)
            {
                case StandardParmType.DegreeRoad2007:
                    break;
                case StandardParmType.CityRoad:
                    OutputXRDxfCityRoad(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, direction);
                    break;
                case StandardParmType.RuralRoadBeijing:
                    break;
                case StandardParmType.DegreeRoad2018:
                    OutputXRDxf(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, direction);
                    break;
                case StandardParmType.DegreeRoad2001:
                    break;
                case StandardParmType.CityRoadShanghai:
                    break;
                case StandardParmType.RuralRoadLiaoning:
                    break;
                case StandardParmType.RuralRoadGuangxi:
                    break;
                case StandardParmType.RuralRoadChongqing:
                    break;
                case StandardParmType.RuralRoadHunan:
                    break;
                case StandardParmType.RuralRoadlowLevel:
                    break;
                default:
                    break;
            }
          
            Marshal.FreeHGlobal(diseaseInfoPtr);
            Marshal.FreeHGlobal(GridInfoPtr);
            m_listDisease.Clear();

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            excelApp.Quit();
        }
        public static void OutputDxfByExcel_ProvinceRoad(List<string> listExcelPath, int direction, string saveFolder, double beginMile, double endMile)
        {
            //打开Excel文件
            MSExcel.Application excelApp = new MSExcel.Application()
            {
                Visible = true,
                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                AlertBeforeOverwriting = false
            };
            List<diseaseInfo> m_listDisease = readDataFromExcel(excelApp, listExcelPath);


            string strName = string.Format("{0:K0+000}-{1:K0+000}", beginMile, endMile);
            string dxfPath = saveFolder;
            if (direction > 0)
                dxfPath += "\\上行";
            else if (direction < 0)
                dxfPath += "\\下行";
            int diseaseCount = m_listDisease.Count;
            int ptrSize = Marshal.SizeOf(typeof(diseaseInfo));
            IntPtr diseaseInfoPtr = Marshal.AllocHGlobal(diseaseCount * ptrSize);
            IntPtr ptr = diseaseInfoPtr;
            for (int m = 0; m < diseaseCount; m++)
            {
                Marshal.StructureToPtr(m_listDisease[m], ptr, false);
                ptr = (IntPtr)(long)ptr + ptrSize;
            }
            GridDiseaseInfo grid = new GridDiseaseInfo();
            grid.dBegMileage = beginMile;
            grid.dEndMileage = endMile;
            grid.strName = strName;
            grid.strBegMile = beginMile.ToString();
            grid.strEndMile = endMile.ToString();
            grid.dRoadWidth = _RoadConfig.RealWidth;
            grid.dRoadTotalNum = listExcelPath.Count;
            ptrSize = Marshal.SizeOf(typeof(GridDiseaseInfo));
            IntPtr GridInfoPtr = Marshal.AllocHGlobal(ptrSize);
            Marshal.StructureToPtr(grid, GridInfoPtr, false);

            Directory.CreateDirectory(dxfPath);
            OutputXRDxfProvinceRoad(dxfPath, diseaseCount, diseaseInfoPtr, GridInfoPtr, direction);
            Marshal.FreeHGlobal(diseaseInfoPtr);
            Marshal.FreeHGlobal(GridInfoPtr);
            m_listDisease.Clear();

            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            excelApp.Quit();
        }
        private static List<diseaseInfo> readDataFromExcel( MSExcel.Application excelApp,List<string> excelPath)
        {

            List<diseaseInfo> m_list = new List<diseaseInfo>();
            MyExcelData excelData = new MyExcelData();
            //MSExcel.Application excelApp = new MSExcel.Application()
            //{
            //    Visible = true,
            //    DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
            //    AlertBeforeOverwriting = false
            //};
            for (int i = 0; i < excelPath.Count(); i++)
            {
                ReadExcelData(excelApp, excelPath[i], 3, ref excelData);
                ConvertExcelData2Disease(excelData, ref m_list, i);
            }
        

            return m_list;
        }

        /// <summary>
        /// 读取原始报表的数据
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="fpath">excel文件的路径</param>
        /// <param name="startrow">开始有数据内容的行数</param>
        /// <param name="exceldata">输出数据</param>
        private static void ReadExcelData(MSExcel.Application excelApp, string fpath, int startrow, ref MyExcelData exceldata)
        {
            MSExcel.Workbook workbook = excelApp.Workbooks.Open(fpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet worksheet = workbook.Worksheets["病害列表"] as MSExcel.Worksheet;
            exceldata.datarow = GlobalExcel.judegeusedrow(worksheet, 1, startrow);
            exceldata.datacol = GlobalExcel.judegeusedcol(worksheet, startrow, 3);

            int headCols = GlobalExcel.judegeusedcol(worksheet, startrow - 1, 3);
            MSExcel.Range headrange = worksheet.get_Range(string.Format("A{0}:{1}{2}", 2, GlobalExcel.GetCol((char)(headCols - 1 + 'A')), 2));
            CheckRoadType((object[,])headrange.Value2, headCols, ref exceldata);
            if (exceldata.excelProjectType == 0 || exceldata.excelRectType == 0)
            {
                exceldata.datarow = 0;
                exceldata.datacol = 0;
                MessageBox.Show("未知类型");
                return;
            }
            MSExcel.Range workrange = worksheet.get_Range(string.Format("A{0}:{1}{2}", startrow, GlobalExcel.GetCol((char)(exceldata.datacol - 1 + 'A')), exceldata.datarow));
            exceldata.dataobj = (object[,])workrange.Value2;
            exceldata.datarow = exceldata.datarow - startrow + 1;

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void ConvertExcelData2Disease(MyExcelData excelData, ref List<diseaseInfo> lstInfo,int roadIndex)
        {
            int numval = 1;
            if (excelData.excelProjectType == 1 && excelData.excelRectType == 1) //等级路，大框
            {
                for (int j = 1; j <= excelData.datarow; ++j)
                {
                    diseaseInfo info = new diseaseInfo();
                    info.mile = Convert.ToInt32(excelData.dataobj[numval, 1]);//ConvertMile2Int((string)excelData.dataobj[numval, 1]);
                    info.roadNum = (roadIndex + 1).ToString();//车道数按传入excel的顺序来排
                    info.diseaseType = (string)excelData.dataobj[numval, 3];
                    info.diseaseDegree = (string)excelData.dataobj[numval, 4];

                    info.rectHeight = (double)excelData.dataobj[numval, 5];
                    info.rectWidth = (double)excelData.dataobj[numval, 6];
                    info.distToCenter = (double)excelData.dataobj[numval, 7];
                    if (!excelData.diseaseCenter2Edge)
                        info.distToCenter += _RoadConfig.RealWidth / 2;//病害距离中心改成距离边缘
                    info.diseaseArea = (double)excelData.dataobj[numval, 8];
                    info.calcHeight = (double)excelData.dataobj[numval, 9];
                    info.calcWidth = (double)excelData.dataobj[numval, 10];
                    //if ((string)excelData.dataobj[numval, 11] == "是")
                    //    info.bOnRoad = false;
                    //else if ((string)excelData.dataobj[numval, 11] == "否")
                    //    info.bOnRoad = true;

                    ++numval;
                    lstInfo.Add(info);
                }
            }
            else if (excelData.excelProjectType == 1 && excelData.excelRectType == 2) //等级路，小框
            { 
                MessageBox.Show("等级路，小框不支持导出dxf图");
               
            }
            else if (excelData.excelProjectType == 2 && excelData.excelRectType == 1) //城镇路，大框
            {
                for (int j = 1; j <= excelData.datarow; ++j)
                {
                    diseaseInfo info = new diseaseInfo();
                    info.mile = Convert.ToInt32(excelData.dataobj[numval, 1]);//ConvertMile2Int((string)excelData.dataobj[numval, 1]);
                    info.roadNum = (roadIndex + 1).ToString();//车道数按传入excel的顺序来排
                    info.diseaseType = (string)excelData.dataobj[numval, 3];
                    info.diseaseDegree = "无";

                    info.rectHeight = (double)excelData.dataobj[numval, 4];
                    info.rectWidth = (double)excelData.dataobj[numval, 5];
                    info.distToCenter = (double)excelData.dataobj[numval, 6];
                    info.diseaseArea = (double)excelData.dataobj[numval, 7];
                    info.calcHeight = (double)excelData.dataobj[numval, 8];
                    info.calcWidth = (double)excelData.dataobj[numval, 10];
                    info.bOnRoad = false;
               

                    ++numval;
                    lstInfo.Add(info);
                }
            }
            else if (excelData.excelProjectType == 2 && excelData.excelRectType == 2)//城镇路，小框
            {
                MessageBox.Show("城镇路，小框不支持导出dxf图");
             
            }
        }

        private static string[] _Degree2018_BigRect_ColStr = { "桩号", "车道", "病害类型", "病害程度", "病害框长度（m）", "病害框宽度（m）", "病害中心位置（距路面图像左边距离）（m）", "病害面积(m2)", "病害计算长度（m）", "病害计算宽度（m）", "路面图像名称", "路面图像相对路径", "路面材质", "备注" };
        private static string[] _Degree2018_BigRect_ROADIMG_ColStr = { "桩号", "车道", "病害类型", "病害程度", "病害框长度（m）", "病害框宽度（m）", "病害中心位置（距路面图像左边距离）（m）", "病害面积(m2)", "病害计算长度（m）", "病害计算宽度（m）", "路面材质", "备注" };
        private static string[] _Degree2018_BigRect_ROADIMG_QUFEN_ColStr = { "桩号", "车道", "病害类型", "病害框长度（m）", "病害框宽度（m）", "病害中心位置（距路面图像左边距离）（m）", "病害面积(m2)", "病害计算长度（m）", "病害计算宽度（m）", "路面材质", "备注" };
        private static string[] _Degree2018_BigRect_QUFEN_ColStr = { "桩号", "车道", "病害类型", "病害框长度（m）", "病害框宽度（m）", "病害中心位置（距路面图像左边距离）（m）", "病害面积(m2)", "病害计算长度（m）", "病害计算宽度（m）", "路面图像名称", "路面图像相对路径", "路面材质", "备注" };
        private static string[] _Degree2018_BigRect_ColStr_test = { "桩号", "车道", "病害类型", "病害程度", "病害框长度（m）", "病害框宽度（m）", "病害中心与中心线距离(m)", "病害面积(m2)", "病害计算长度（m）", "病害计算宽度（m）", "是否在轮迹带" };

        private static string[] _Degree2018_SmallRect_ColStr = { "桩号", "车道", "病害类型", "病害程度", "病害面积(m2)", "具体位置_距右侧标线位置(m)","路面图像名称", "路面图像相对路径", "路面材质" };
        private static string[] _City_ColStr = { "桩号", "车道", "病害类型", "病害框长度（m）", "病害框宽度（m）", "病害中心位置（距路面图像左边距离）（m）", "病害面积（m2）", "病害计算长度（m）", "病害深度（mm）", "病害计算宽度（m）", "路面图像名称", "路面图像相对路径", "路面材质"};                                                                                                                                                                                                                                                                                                    
        /// <summary>
        /// 根据报表的结构判断报表种类
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="fpath">excel文件的路径</param>
        /// <param name="startrow">开始有数据内容的行数</param>
        /// <param name="exceldata">输出数据</param>
        private static void CheckRoadType(object[,] dataobj, int datacol, ref MyExcelData exceldata)
        {
            //报表模板\等级公路 JTG H20-2018\路面病害面积统计表.xlsx
            if (datacol == _Degree2018_BigRect_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_BigRect_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 1;
                    return;
                }
            }
            if (datacol == _Degree2018_BigRect_ROADIMG_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_BigRect_ROADIMG_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 1;
                    return;
                }
            }
            if (datacol == _Degree2018_BigRect_ROADIMG_QUFEN_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_BigRect_ROADIMG_QUFEN_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 1;
                    return;
                }
            }
            if (datacol == _Degree2018_BigRect_QUFEN_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_BigRect_QUFEN_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 1;
                    return;
                }
            }
            //报表模板\等级公路 JTG H20-2018\路面病害面积统计表-Small.xlsx
            if (datacol == _Degree2018_BigRect_ColStr_test.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_BigRect_ColStr_test[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 1;
                    exceldata.diseaseCenter2Edge = false;//导出综合报表时会计算病害中心与中心线的距离，需要改成与左边缘的距离
                    return;
                }
            }
            //报表模板\等级公路 JTG H20-2018\路面病害面积统计表-Small.xlsx
            if (datacol == _Degree2018_SmallRect_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_Degree2018_SmallRect_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 1;
                    exceldata.excelRectType = 2;
                    return;
                }
            }
            //报表模板\城镇道路\路面病害面积统计表.xlsx
            if (datacol == _City_ColStr.Length)
            {
                bool bFind = true;
                for (int i = 0; i < datacol; i++)
                {
                    if (!((string)dataobj[1, i + 1]).Equals(_City_ColStr[i]))
                    {
                        bFind = false;
                        break;
                    }
                }
                if (bFind)
                {
                    exceldata.excelProjectType = 2;
                    exceldata.excelRectType = 1;
                    return;
                }
            }
            //未知Excel

            exceldata.excelProjectType = 0;
            exceldata.excelRectType = 0;
            return;

        }
        class MyExcelData
        {
            public int datarow = 0;
            public int datacol = 0;
            public int excelProjectType = 0;
            public int excelRectType = 0;
            public bool diseaseCenter2Edge = true;
            public object[,] dataobj = null;
        }
        private static int ConvertMile2Int(string strMile)
        {

           strMile = strMile.Substring(1, strMile.Length - 1);//先去掉首位的“K”
           strMile.Replace("+", "") ;//去掉"+"
            return int.Parse(strMile);     
        }

        #endregion
    }
}
