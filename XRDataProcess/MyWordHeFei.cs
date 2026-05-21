/*-----------------------------------------------------------------
//CopyRight (C) 2012 武汉汉宁轨道交通技术有限公司
//版权所有。
//MyWordSzechwanDQ
//安徽合肥报告
//
//
//创建标识:cwb 20230804
//修改标识:cwb 20230804
//修改描述: 读取用户提供的excel表格  获取出报告需要的数据 
 //------------------------------------------------------------------*/
#define RUN
#define PIE
#define bhhz
using DevExpress.DirectX.Common.Direct2D;
using DevExpress.Map.Native;
using DevExpress.Utils.Win.Hook;
using DevExpress.XtraBars.Docking2010.Views.Widget; 
using DevExpress.XtraCharts.Native;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraExport;
using DevExpress.XtraPrinting.Native;
using Framework.Office.Excel;
using Framework.Office.Work;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Ookii.Dialogs.WinForms;
using OpenTK.Platform.Windows;
using Org.BouncyCastle.Asn1.Crmf;
using Org.BouncyCastle.Bcpg;
using Org.BouncyCastle.Cms;
using SqlSugar;
using SqlSugar.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceModel.Description;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using MSExcel = Microsoft.Office.Interop.Excel;
using MSOffice = Microsoft.Office.Core;
using MSWord = Microsoft.Office.Interop.Word;
//using Framework.Office.Work;
namespace XRDataProcess
{

    public class HeFeiRaod
    {

        public HeFeiRaod(HeFeiRaod other)
        {
            this.eMile = other.eMile;
            this.RoadName = other.RoadName;
            this.RoadCode = other.RoadCode;
            this.PciGrad = other.PciGrad;
            this.PqiGrad = other.PqiGrad;
            this.RqiGrad = other.RqiGrad;
            this.PciValue = other.PciValue;
            this.PqiValue = other.PqiValue;
            this.RqiValue = other.RqiValue;
            this.RoadLen = other.RoadLen;
            this.sMile = other.sMile;
            this.RoadGrad = other.RoadGrad;
            this.RoadType = other.RoadType;
        }
        public HeFeiRaod()
        {

        }
        public List<HeFeiRaod> datas { get; set; } = new List<HeFeiRaod>();
        public string RoadName { get; set; }
        public string RoadCode { get; set; }
        /// <summary>
        /// 0 沥青
        /// 1 水泥
        /// 2 砂石
        /// </summary>
        public int RoadType { get; set; }
        //加权平均值
        public double PqiValue { get; set; }
        public double PciValue { get; set; }
        public double RqiValue { get; set; }

        public double DrValue { get; set; }

        public double IriValue { get; set; }

        public double sMile { get; set; }
        public double eMile { get; set; }
        public double RoadLen { get; set; }


        public string PqiGrad { get; set; }
        //算数平均值
        public string PciGrad { get; set; }
        public string RqiGrad { get; set; }
        /// <summary>
        ///  0 高速
        ///  1 一级
        ///  2 二级
        ///  3 三级
        ///  4 四级
        /// </summary>
        public int RoadGrad { get; set; }

    }

    public class DiseaseHeFei
    {
        public string RoadCode { get; set; }
        public string Name { get; set; }

        public string DamagedCondition { get; set; }

        public string RoadGrad { set; get; }

        public double Area { get; set; }
        public string RoadType { get; set; }

        public DiseaseHeFei()
        {

        }
        public DiseaseHeFei(DiseaseHeFei other)
        {
            this.Area = other.Area;
            this.Name = other.Name;
            this.RoadCode = other.RoadCode;
            this.DamagedCondition = other.DamagedCondition;
            this.RoadGrad = other.RoadGrad;
            this.RoadType = other.RoadType;

        }
        public override bool Equals(object obj)
        {
            DiseaseHeFei other = obj as DiseaseHeFei;

            if (other.Name == this.Name &&
                other.DamagedCondition == this.DamagedCondition &&
                this.RoadType == other.RoadType &&
                this.RoadGrad == other.RoadGrad)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
    public class MyWordHeFei
    {
        [DllImport("user32.dll")]
        public static extern bool IsClipboardFormatAvailable(uint format);
        private static string wordModlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"报告模板\安徽",
            "农村公路路面技术状况自动化检测_20251129.docx");
        private string outWordPath = "";
        private MSWord.Application wordApp = null;
        private MSWord.Selection currentSelection = null;
        private List<FileInfo> excelPathList = null;
        private List<HeFeiRaod> allDatas = new List<HeFeiRaod>();
        private string basePath = "";
        XRSetting _Setting = XRSetting.GetInstance();
        bool m_outPartEight = false;
        public MyWordHeFei(string sourceFile, string sourceTxtFile, bool needDis, bool outPartEight)
        {
            m_outPartEight = outPartEight;
            ReadDiseaseXml();
            NeedDis = needDis;
            //读取病害
            //FileInfo info = new FileInfo(sourceFile);
            if (!String.IsNullOrEmpty(sourceTxtFile))
            {
                readTxt(sourceTxtFile);
            }
            if (readExcel(sourceFile))
            {
                if (_Setting.heFeiContineIndex == 0)
                {
                    outWordPath = Path.Combine(
             Path.GetDirectoryName(sourceFile),
             Path.GetFileNameWithoutExtension(sourceFile) + ".docx");
                }
                else
                {
                    outWordPath = Path.Combine(
             Path.GetDirectoryName(sourceFile),
             Path.GetFileNameWithoutExtension(sourceFile) + "_Contine.docx");
                }

                if (File.Exists(outWordPath))
                {
                    File.Delete(outWordPath);
                }

            }

         
            string disExcelPath = Path.GetDirectoryName(sourceFile);
            basePath = disExcelPath;
            if (needDis)
            {
                DirectoryInfo dir = new DirectoryInfo(disExcelPath);
                excelPathList = dir.GetFiles("*病害统计*.xlsx", SearchOption.AllDirectories).ToList();
                string errorMsg = "";
                int valueTmep = 10;
                foreach (var item in allDatas)
                {
                    var result = GetTargetDiseaseExcel(item);
                    if (result.Count == 0)
                    {
                        valueTmep--;

                        errorMsg += item.RoadCode + "(" + item.RoadName + ")\t";
                        if (valueTmep == 0)
                        {
                            errorMsg += item.RoadCode + "(" + item.RoadName + ")\n";
                            valueTmep = 10;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(errorMsg))
                {
                    MessageBox.Show(errorMsg + "未找到对应病害表格");

                }
            }
          
        }

        double value;
        string grad = "";
        private static string 等级区间 = "90 80 70 60 0";
        string[] 区间值 = 等级区间.Split(' ');

        /// <summary>
        /// 项目名称
        /// </summary>
        private string m_proName { get; set; }
        /// <summary>
        /// 县区名称
        /// </summary>
        private string m_countyName { get; set; }
        /// <summary>
        /// 批准人
        /// </summary> 
        private string m_approverPeople { get; set; }

        /// <summary>
        /// 复核人
        /// </summary>
        private string m_reviewerPeople { get; set; }
        /// <summary>
        /// 检测人
        /// </summary>

        private string m_inspectorPeople { get; set; }

        /// <summary>
        /// 检测开始时间
        /// </summary>
        private DateTime s_time { get; set; }
        /// <summary>
        /// 检测结束时间
        /// </summary>
        private DateTime e_time { get; set; }


        /// <summary>
        /// 分组数量
        /// </summary>
        private string m_Count { get; set; }
        public bool NeedDis { get; }

        private bool readTxt(string sourceFile)
        {

            string[] txts = File.ReadAllLines(sourceFile);

            foreach (string txt in txts)
            {

                string[] txtSplit = txt.Split(':');
                if (txtSplit.Length > 1)
                {
                    string name = txtSplit[0];
                    string value = txtSplit[1];
                    switch (name)
                    {
                        case "项目名称":
                            m_proName = value;
                            break;
                        case "县区名称":
                            m_countyName = value;
                            break;
                        case "批准人":
                            m_approverPeople = value;
                            break;
                        case "复核人":
                            m_reviewerPeople = value;
                            break;
                        case "检测人":
                            m_inspectorPeople = value;
                            break;
                        case "检测开始时间":
                            {
                                DateTime dt;
                                if (DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                                {
                                    s_time = dt;
                                }
                                else
                                {
                                    s_time = new DateTime();
                                }
                            }
                            break;
                        case "检测结束时间":
                            {
                                DateTime dt;
                                if (DateTime.TryParseExact(value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                                {
                                    e_time = dt;
                                }
                                else
                                {
                                    e_time = new DateTime();
                                }
                            }
                            break;
                        case "分组数量":


                            m_Count = value;
                            break;

                        default:
                            break;
                    }

                }

            }


            return true;

        }
        private bool readExcel(string excelPath)
        {
            List<HeFeiRaod> TempDatas = new List<HeFeiRaod>();
            Dictionary<string, HeFeiRaod> TempDicDatas = new Dictionary<string, HeFeiRaod>();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                ReadExcelData(ref dt, excelPath, 0, 1, true);
                //  ReadExcelData(ref dt,excelPath, 1, 0); 
                //   ReadExcelData(ref dt,excelPath, 2, 0);   
            }
            catch (Exception ex)
            {
                throw ex;
            }
            #region 将数据导入内存
            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                DataRow data = dt.Rows[i];
                HeFeiRaod heFeiRaod = new HeFeiRaod();
                string grad = "";
                try
                {

                    heFeiRaod.RoadCode = data[2].ToString();
                    heFeiRaod.RoadName = data[3].ToString();
                    heFeiRaod.PqiValue = double.Parse(data[5].ToString());
                    heFeiRaod.PciValue = double.Parse(data[6].ToString());
                    heFeiRaod.RqiValue = double.Parse(data[7].ToString());
                    heFeiRaod.DrValue = double.Parse(data[20].ToString());
                    heFeiRaod.IriValue = double.Parse(data[21].ToString());
                    heFeiRaod.sMile = double.Parse(data[8].ToString());
                    heFeiRaod.eMile = double.Parse(data[9].ToString());
                    heFeiRaod.RoadLen = double.Parse(data[10].ToString());
                    heFeiRaod.PqiGrad = data[17].ToString();
                    heFeiRaod.PciGrad = data[18].ToString();
                    heFeiRaod.RqiGrad = data[19].ToString();
                    heFeiRaod.RoadType = data[16].ToString().Contains("沥青") ? 0 : data[16].ToString().Contains("水泥") ? 1 : 2;
                    grad = data[15].ToString();
                }
                catch (Exception ex)
                {
                    int errorRow = i + 2;
                    MessageBox.Show("读取的表格第[" + errorRow + "]行数据有问题，请检查是否使用了公式计算，或者有非法字符!");
                    return false;
                }

                if (grad.Contains("高速"))
                {
                    heFeiRaod.RoadGrad = 0;
                }
                if (grad.Contains("一级"))
                {
                    heFeiRaod.RoadGrad = 1;
                }
                if (grad.Contains("二级"))
                {
                    heFeiRaod.RoadGrad = 2;
                }
                if (grad.Contains("三级"))
                {
                    heFeiRaod.RoadGrad = 3;
                }
                if (grad.Contains("四级"))
                {
                    heFeiRaod.RoadGrad = 4;
                }

                TempDatas.Add(heFeiRaod);
            }

            //for (int i = 0; i < TempDatas.Count; i++)
            //{ 
            //    for (int t = 0; t < TempDatas.Count; t++)
            //    {

            //        if (TempDatas[i].RoadName != TempDatas[t].RoadName && TempDatas[i].RoadCode == TempDatas[t].RoadCode)
            //        {
            //            continue;
            //        }
            //    }
            //}


            string nowStr = "";
            for (int i = 0; i < TempDatas.Count; i++)
            {
                HeFeiRaod dataNow = TempDatas[i];
                nowStr = dataNow.RoadCode + dataNow.RoadName;
                if (TempDicDatas.ContainsKey(nowStr))
                {
                    TempDicDatas[nowStr].eMile = dataNow.eMile;
                    TempDicDatas[nowStr].RoadLen += dataNow.RoadLen;
                    TempDicDatas[nowStr].datas.Add(dataNow);
                }
                else
                {
                    HeFeiRaod data = new HeFeiRaod(dataNow);
                    data.datas.Add(dataNow);
                    TempDicDatas.Add(nowStr, data);
                }
            }
            foreach (var item in TempDicDatas)
            {
                double pci = 0;
                double pqi = 0;
                double rqi = 0;
                double dr = 0;
                double iri = 0;
                double len = 0;
                foreach (var data in item.Value.datas)
                {
                    len += data.RoadLen;
                    pci += data.PciValue;
                    pqi += data.PqiValue;
                    rqi += data.RqiValue;
                    dr += data.DrValue;
                    iri += data.IriValue;
                }
                item.Value.PciValue = pci / item.Value.datas.Count;
                item.Value.PqiValue = pqi / item.Value.datas.Count;
                item.Value.RqiValue = rqi / item.Value.datas.Count;
                item.Value.DrValue = dr / item.Value.datas.Count;
                item.Value.IriValue = iri / item.Value.datas.Count;
                item.Value.RoadLen = len;
                #region 修改评价值
                //90 80 70 60 0
                item.Value.PqiGrad = SetGrad(item.Value.PqiValue);
                item.Value.PciGrad = SetGrad(item.Value.PciValue);
                item.Value.RqiGrad = SetGrad(item.Value.RqiValue);
                #endregion
                allDatas.Add(item.Value);
            }


            foreach (var data in allDatas)
            {

                double len = 0;
                foreach (var item in data.datas)
                {
                    len += item.RoadLen;
                }

                data.RoadLen = len;
            }
            if (allDatas.Count == 0)
            {
                MessageBox.Show("Excel中未找到任何信息,请检查");
                return false;
            }
            #endregion

            calculateDatas();
            return true;
        }

        private string SetGrad(double value)
        {
            string grad = "";
            for (int i = 0; i < 区间值.Length - 1; i++)
            {
                double 上限, 下限;
                if (double.TryParse(区间值[i], out 上限) && double.TryParse(区间值[i + 1], out 下限))
                {
                    if (value >= 上限)
                    {
                        grad = "优";
                    }
                    else if (value >= 下限)
                    {
                        switch (i)
                        {
                            case 0:
                                grad = "良";
                                break;
                            case 1:
                                grad = "中";
                                break;
                            case 2:
                                grad = "次";
                                break;
                            case 3:
                                grad = "差";
                                break;
                        }
                        break;
                    }
                }
            }
            return grad;
        }


        #region Pqi
        //pqi加权平均值
        private double RoadPqiValue = 0;
        private string RqiRoadGrad = "";
        private string PqiRoadGrad = "";
        private string PciRoadGrad = "";
        //pqi 优良路率
        private double RoadPqiYLRate = 0;
        private double RoadPqiCCRate = 0;

        //个数比例
        private double RoadPqiYLCountRate = 0;
        private double RoadPciYLCountRate = 0;
        private double RoadRqiYLCountRate = 0;

        //PQI优良中次差路个数
        private int yPQIRoadCount = 0;
        private double yPQIRoadLength = 0;
        private double lPQIRoadLength = 0;
        private double zPQIRoadLength = 0;
        private double ciPQIRoadLength = 0;
        private double chaPQIRoadLength = 0;

        private double yPcIRoadLength = 0;
        private double lPcIRoadLength = 0;
        private double zPcIRoadLength = 0;
        private double ciPcIRoadLength = 0;
        private double chaPcIRoadLength = 0;

        private double yRQIRoadLength = 0;
        private double lRQIRoadLength = 0;
        private double zRQIRoadLength = 0;
        private double CiRQIRoadLength = 0;
        private double ChaRQIRoadLength = 0;


        private int lPQIRoadCount = 0;
        private int zPQIRoadCount = 0;
        private int ciPQIRoadCount = 0;


        private int chaPQIRoadCount = 0;


        //排名前五名称 及 后五
        private string pqiTopFiveRoadStr = "";
        private string pqiLastFiveRoadStr = "";

        private string rqiTopFiveRoadStr = "";
        private string rqiLastFiveRoadStr = "";

        private string pciTopFiveRoadStr = "";
        private string pciLastFiveRoadStr = "";
        //低于平均值道路数量
        private int pqiSubaverageCout = 0;
        private int pciSubaverageCout = 0;
        private int rqiSubaverageCout = 0;
        #endregion
        #region Rqi
        private double RoadRqiValue = 0;
        //rqi 优良路率
        private double RoadRqiYLRate = 0;
        private double RoadRqiCCRate = 0;
        #endregion
        #region Pci
        //pci加权平均值
        private double RoadPciValue = 0;
        //pci 优良路率  优良路长度/总长度
        private double RoadPciYLRate = 0;
        private double RoadPciCCRate = 0;

        private double yPCIRoadRate = 0;
        private double lPCIRoadRate = 0;
        private double zPCIRoadRate = 0;
        private double ciPCIRoadRate = 0;
        private double chaPCIRoadRate = 0;

        private int yPciCount = 0;
        private int lPciCount = 0;
        private int zPciCount = 0;
        private int ciPciCount = 0;
        private int chaPciCount = 0;



        #endregion
        private double yRQIRoadRate = 0;
        private double lRQIRoadRate = 0;
        private double zRQIRoadRate = 0;
        private double ciRQIRoadRate = 0;
        private double chaRQIRoadRate = 0;
        private int yRQICount = 0;
        private int lRQICount = 0;
        private int zRQICount = 0;
        private int ciRQICount = 0;
        private int chaRQICount = 0;


        private double RoadLength = 0;
        public void calculateDatas()
        {
            double tempPqis = 0;
            double tempRqis = 0;
            double tempPci = 0;
            //优良路长度
            double ylPciRoadLen = 0;
            //次差路长度
            double ccPciRoadLen = 0;

            double ylPqiRoadLen = 0;
            double ccPqiRoadLen = 0;

            double ylRqiRoadLen = 0;
            double ccRqiRoadLen = 0;
            //pqi平均值
            double pqiTemp = 0;
            double pciTemp = 0;
            double rqiTemp = 0;


            //计算得到路线总长度 
            for (int i = 0; i < allDatas.Count; i++)
            {

                pqiTemp += allDatas[i].PqiValue * allDatas[i].RoadLen;
                pciTemp += allDatas[i].PciValue * allDatas[i].RoadLen;
                rqiTemp += allDatas[i].RqiValue * allDatas[i].RoadLen;
                tempPqis += allDatas[i].PqiValue * allDatas[i].RoadLen;
                tempRqis += allDatas[i].RqiValue * allDatas[i].RoadLen;
                tempPci += allDatas[i].PciValue * allDatas[i].RoadLen;
                RoadLength += allDatas[i].RoadLen;
                var data = allDatas[i];
                #region PQI
                if (allDatas[i].PqiGrad == "优")
                {
                    yPQIRoadLength += allDatas[i].RoadLen;
                    yPQIRoadCount++;

                }
                if (allDatas[i].PqiGrad == "良")
                {
                    lPQIRoadLength += allDatas[i].RoadLen;
                    lPQIRoadCount++;
                }
                if (allDatas[i].PqiGrad == "中")
                {
                    zPQIRoadLength += allDatas[i].RoadLen;
                    zPQIRoadCount++;
                }
                if (allDatas[i].PqiGrad == "次")
                {
                    ciPQIRoadLength += allDatas[i].RoadLen;
                    ciPQIRoadCount++;
                }
                if (allDatas[i].PqiGrad == "差")

                {
                    chaPQIRoadLength += allDatas[i].RoadLen;
                    chaPQIRoadCount++;
                }
                #endregion
                #region PCI

                if (allDatas[i].PciGrad == "优")
                {

                    yPcIRoadLength += allDatas[i].RoadLen;
                    yPciCount++;
                }
                else if (allDatas[i].PciGrad == "良")
                {
                    lPcIRoadLength += allDatas[i].RoadLen;
                    lPciCount++;
                }
                else if (allDatas[i].PciGrad == "中")
                {
                    zPciCount++;
                    zPcIRoadLength += allDatas[i].RoadLen;
                }
                else if (allDatas[i].PciGrad == "次")
                {
                    ciPcIRoadLength += allDatas[i].RoadLen;
                    ciPciCount++;
                }
                else
                {
                    chaPcIRoadLength += allDatas[i].RoadLen;
                    chaPciCount++;
                }

                #endregion 
                #region RQI 
                if (allDatas[i].RqiGrad == "优")
                {
                    yRQIRoadLength += allDatas[i].RoadLen;
                    yRQICount++;
                }
                else if (allDatas[i].RqiGrad == "良")
                {
                    lRQIRoadLength += allDatas[i].RoadLen;

                    lRQICount++;
                }
                else if (allDatas[i].RqiGrad == "中")
                {
                    zRQIRoadLength += allDatas[i].RoadLen;
                    zRQICount++;
                }
                else if (allDatas[i].RqiGrad == "次")
                {
                    CiRQIRoadLength += allDatas[i].RoadLen;
                    ciRQICount++;
                }
                else
                {
                    ChaRQIRoadLength += allDatas[i].RoadLen;
                    chaRQICount++;
                }
                #endregion

                for (int t = 0; t < allDatas[i].datas.Count; t++)
                {

                    #region PQI 
                    if (allDatas[i].datas[t].PqiGrad == "优" || allDatas[i].datas[t].PqiGrad == "良")
                    {
                        ylPqiRoadLen += allDatas[i].datas[t].RoadLen;
                    }
                    if (allDatas[i].datas[t].PqiGrad == "次" || allDatas[i].datas[t].PqiGrad == "差")
                    {
                        ccPqiRoadLen += allDatas[i].datas[t].RoadLen;
                    }
                    #endregion
                    #region PCI
                    if (allDatas[i].datas[t].PciGrad == "优" || allDatas[i].datas[t].PciGrad == "良")
                    {
                        ylPciRoadLen += allDatas[i].datas[t].RoadLen;
                    }
                    if (allDatas[i].datas[t].PciGrad == "次" || allDatas[i].datas[t].PciGrad == "差")
                    {
                        ccPciRoadLen += allDatas[i].datas[t].RoadLen;
                    }
                    if (allDatas[i].datas[t].PciGrad == "优")
                    {
                        yPCIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    else if (allDatas[i].datas[t].PciGrad == "良")
                    {
                        lPCIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    else if (allDatas[i].datas[t].PciGrad == "中")
                    {
                        zPCIRoadRate += allDatas[i].datas[t].RoadLen;


                    }
                    else if (allDatas[i].datas[t].PciGrad == "次")
                    {
                        ciPCIRoadRate += allDatas[i].datas[t].RoadLen;

                    }
                    else
                    {
                        chaPCIRoadRate += allDatas[i].datas[t].RoadLen;

                    }

                    #endregion
                    #region RQI
                    if (allDatas[i].datas[t].RqiGrad == "优" || allDatas[i].datas[t].RqiGrad == "良")
                    {
                        ylRqiRoadLen += allDatas[i].datas[t].RoadLen;
                    }
                    if (allDatas[i].datas[t].RqiGrad == "次" || allDatas[i].datas[t].RqiGrad == "差")
                    {
                        ccRqiRoadLen += allDatas[i].datas[t].RoadLen;
                    }

                    if (allDatas[i].datas[t].RqiGrad == "优")
                    {
                        yRQIRoadRate += allDatas[i].datas[t].RoadLen;

                    }
                    else if (allDatas[i].datas[t].RqiGrad == "良")
                    {
                        lRQIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    else if (allDatas[i].datas[t].RqiGrad == "中")
                    {
                        zRQIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    else if (allDatas[i].datas[t].RqiGrad == "次")
                    {
                        ciRQIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    else
                    {
                        chaRQIRoadRate += allDatas[i].datas[t].RoadLen;
                    }
                    #endregion
                }
            }


            RoadPciYLRate = ylPciRoadLen / RoadLength;
            RoadPciCCRate = ccPciRoadLen / RoadLength;
            RoadRqiYLRate = ylRqiRoadLen / RoadLength;
            RoadRqiCCRate = ccRqiRoadLen / RoadLength;
            RoadPqiYLRate = ylPqiRoadLen / RoadLength;
            RoadPqiCCRate = ccPqiRoadLen / RoadLength;

            RoadPqiYLCountRate = ((double)yPQIRoadCount + (double)lPQIRoadCount) / (double)allDatas.Count;
            RoadPciYLCountRate = ((double)yPciCount + (double)lPciCount) / (double)allDatas.Count;
            RoadRqiYLCountRate = ((double)yRQICount + (double)lRQICount) / (double)allDatas.Count;
            RoadPqiValue = tempPqis / RoadLength;
            RoadPciValue = tempPci / RoadLength;
            RoadRqiValue = tempRqis / RoadLength;

            yPCIRoadRate /= RoadLength;
            lPCIRoadRate /= RoadLength;
            zPCIRoadRate /= RoadLength;
            ciPCIRoadRate /= RoadLength;
            chaPCIRoadRate /= RoadLength;


            yRQIRoadRate /= RoadLength;
            lRQIRoadRate /= RoadLength;
            zRQIRoadRate /= RoadLength;
            ciRQIRoadRate /= RoadLength;
            chaRQIRoadRate /= RoadLength;

            List<HeFeiRaod> topFiveProjects = allDatas.OrderByDescending(x => x.PqiValue).Take(5).ToList();
            foreach (var item in topFiveProjects)
            {
                pqiTopFiveRoadStr += item.RoadName + "、";
            }
            pqiTopFiveRoadStr = pqiTopFiveRoadStr.Substring(0, pqiTopFiveRoadStr.Count() - 1);

            topFiveProjects = allDatas.OrderByDescending(x => x.PciValue).Take(5).ToList();
            foreach (var item in topFiveProjects)
            {
                pciTopFiveRoadStr += item.RoadName + "、";
            }
            pciTopFiveRoadStr = pciTopFiveRoadStr.Substring(0, pciTopFiveRoadStr.Count() - 1);

            topFiveProjects = allDatas.OrderByDescending(x => x.RqiValue).Take(5).ToList();
            foreach (var item in topFiveProjects)
            {
                rqiTopFiveRoadStr += item.RoadName + "、";
            }
            rqiTopFiveRoadStr = rqiTopFiveRoadStr.Substring(0, rqiTopFiveRoadStr.Count() - 1);



            List<HeFeiRaod> bottomFiveProjects = allDatas.OrderBy(x => x.PqiValue).Take(5).ToList();
            foreach (var item in bottomFiveProjects)
            {
                pqiLastFiveRoadStr += item.RoadName + "、";
            }
            pqiLastFiveRoadStr = pqiLastFiveRoadStr.Substring(0, pqiLastFiveRoadStr.Count() - 1);

            bottomFiveProjects = allDatas.OrderBy(x => x.PciValue).Take(5).ToList();
            foreach (var item in bottomFiveProjects)
            {
                pciLastFiveRoadStr += item.RoadName + "、";
            }
            pciLastFiveRoadStr = pciLastFiveRoadStr.Substring(0, pciLastFiveRoadStr.Count() - 1);

            bottomFiveProjects = allDatas.OrderBy(x => x.RqiValue).Take(5).ToList();
            foreach (var item in bottomFiveProjects)
            {
                rqiLastFiveRoadStr += item.RoadName + "、";
            }
            rqiLastFiveRoadStr = rqiLastFiveRoadStr.Substring(0, rqiLastFiveRoadStr.Count() - 1);

            pqiSubaverageCout = allDatas.Where(t => t.PqiValue < RoadPqiValue).ToList().Count;
            pciSubaverageCout = allDatas.Where(t => t.PciValue < RoadPciValue).ToList().Count;
            rqiSubaverageCout = allDatas.Where(t => t.RqiValue < RoadRqiValue).ToList().Count;
            PqiRoadGrad = SetGrad(RoadPqiValue);
            RqiRoadGrad = SetGrad(RoadRqiValue);
            PciRoadGrad = SetGrad(RoadPciValue);

        }

        public bool outWord()
        {

            if (NeedDis)
            {
                if (excelPathList == null || excelPathList.Count < 1)
                {
                    MessageBox.Show("未在导入excel目录下查找到任何病害报表，请检查！");
                    return false;
                }
            }

            if (allDatas.Count == 0)
            {
                MessageBox.Show("未在导入excel中获得任何道路信息，请检查！");
                return false;
            }
            else
            {
                try
                {
                    //读取报表模板
                    wordApp = new MSWord.Application() { Visible = true };
                    MSWord.Document wordDoc = null;
                    if (_Setting.heFeiContineIndex != 0)
                    {

                        VistaFolderBrowserDialog fd0 = new VistaFolderBrowserDialog()
                        {
                            
                            Description = "检测到为增加数据模式，选择需要继续写入的报告",

                        };

                        if (fd0.ShowDialog() == DialogResult.OK)
                        {
                            if (fd0.SelectedPath != string.Empty)
                            {
                                wordModlePath = fd0.SelectedPath;
                            }
                        }
                    }
                    CWB_WordHelper.openWordApp(wordApp, wordModlePath, ref wordDoc);
                    CWB_WordHelper.saveWord(wordDoc, outWordPath);
                    //操作写入数据
                    WriteAllWordMarks(wordDoc);
                    wordDoc.Save();
                    wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);

                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n请人工检查报表是否生成完成");
                    return false;
                }
                finally
                {
                    CWB_WordHelper.disposeWord(wordApp);
                }

            }

        }

        private MSWord.Range wordrange = null;

        private async void WriteAllWordMarks(MSWord.Document wordDoc)
        {
            if (_Setting.heFeiContineIndex == 0)
            {

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
#if RUN

                    #region 公共部分


                    if (book.Name.Contains("p_name"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_countyName);
                        continue;
                    }
                    if (book.Name.Contains("approverPeople"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_approverPeople);
                        continue;
                    }
                    if (book.Name.Contains("reviewerPeople"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_reviewerPeople);
                        continue;
                    }
                    if (book.Name.Contains("inspectorPeople"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_inspectorPeople);
                        continue;
                    }

                    if (book.Name.Contains("StartTime"))
                    {
                        string startTime = s_time.ToString("yyyy年MM月dd日");
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }

                    if (book.Name.Contains("sYear"))
                    {
                        string startTime = s_time.Year.ToString();
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }

                    if (book.Name.Contains("sMonth"))
                    {
                        string startTime = s_time.Month.ToString();


                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }

                    if (book.Name.Contains("sDay"))
                    {
                        string startTime = s_time.Day.ToString();

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }

                    if (book.Name.Contains("eYear"))
                    {
                        string startTime = e_time.Year.ToString();


                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }
                    if (book.Name.Contains("eMonth"))
                    {
                        string startTime = e_time.Month.ToString();


                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }



                    if (book.Name.Contains("eDay"))
                    {
                        string startTime = e_time.Day.ToString();


                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(startTime);
                        continue;
                    }
                    if (book.Name.Contains("splitCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_Count);
                        continue;
                    }
                    if (book.Name.Contains("SumRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(allDatas.Count.ToString());
                        continue;
                    }
                    if (book.Name.Contains("SumRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(RoadLength.ToString());
                        continue;
                    }








                    if (book.Name.Contains("gldjJudgePqi"))
                    {

                        #region 获取各个公路等级评价 
                        double length0 = 0;
                        double length1 = 0;
                        double length2 = 0;
                        double length3 = 0;
                        double grad0PqiValue = 0;
                        double grad1PqiValue = 0;

                        double grad2PqiValue = 0;
                        double grad3PqiValue = 0;


                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 1)
                                {
                                    grad0PqiValue += item1.PqiValue * item1.RoadLen;

                                    length0 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 2)
                                {
                                    grad1PqiValue += item1.PqiValue * item1.RoadLen;

                                    length1 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }


                                if (item1.RoadGrad == 3)
                                {
                                    grad2PqiValue += item1.PqiValue * item1.RoadLen;

                                    length2 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 4)
                                {
                                    grad3PqiValue += item1.PqiValue * item1.RoadLen;

                                    length3 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);

                                }
                            }

                        }
                        grad0PqiValue = double.IsNaN(grad0PqiValue / length0) ? 0 : grad0PqiValue / length0;

                        grad1PqiValue = double.IsNaN(grad1PqiValue / length1) ? 0 : grad1PqiValue / length1;


                        grad2PqiValue = double.IsNaN(grad2PqiValue / length2) ? 0 : grad2PqiValue / length2;


                        grad3PqiValue = double.IsNaN(grad3PqiValue / length3) ? 0 : grad3PqiValue / length3;

                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";
                        string grad3 = "-";
                        if (length0 != 0)
                        {
                            grad0 = SetGrad(grad0PqiValue);
                        }
                        if (length1 != 0)
                        {
                            grad1 = SetGrad(grad1PqiValue);
                        }
                        if (length2 != 0)
                        {
                            grad2 = SetGrad(grad2PqiValue);
                        }
                        if (length3 != 0)
                        {
                            grad3 = SetGrad(grad3PqiValue);
                        }

                        #endregion

                        string text = string.Format("一级公路处于{0}等水平、二级公路处于{1}等水平、三级公路处于{2}等水平、四级公路处于{3}等水平。", grad0, grad1, grad2, grad3);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }



                    if (book.Name.Contains("gldjJudgePci"))
                    {

                        #region 获取各个公路等级评价 
                        double length0 = 0;
                        double length1 = 0;
                        double length2 = 0;
                        double length3 = 0;

                        double grad0PciValue = 0;
                        double grad1PciValue = 0;


                        double grad2PciValue = 0;
                        double grad3PciValue = 0;



                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 1)
                                {

                                    grad0PciValue += item1.PciValue * item1.RoadLen;

                                    length0 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 2)
                                {
                                    grad1PciValue += item1.PciValue * item1.RoadLen;

                                    length1 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }


                                if (item1.RoadGrad == 3)
                                {

                                    grad2PciValue += item1.PciValue * item1.RoadLen;

                                    length2 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 4)
                                {

                                    grad3PciValue += item1.PciValue * item1.RoadLen;

                                    length3 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);

                                }
                            }

                        }

                        grad0PciValue = double.IsNaN(grad0PciValue / length0) ? 0 : grad0PciValue / length0;


                        grad1PciValue = double.IsNaN(grad1PciValue / length1) ? 0 : grad1PciValue / length1;


                        grad2PciValue = double.IsNaN(grad2PciValue / length2) ? 0 : grad2PciValue / length2;



                        grad3PciValue = double.IsNaN(grad3PciValue / length3) ? 0 : grad3PciValue / length3;

                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";
                        string grad3 = "-";
                        if (length0 != 0)
                        {
                            grad0 = SetGrad(grad0PciValue);
                        }
                        if (length1 != 0)
                        {
                            grad1 = SetGrad(grad1PciValue);
                        }
                        if (length2 != 0)
                        {
                            grad2 = SetGrad(grad2PciValue);
                        }
                        if (length3 != 0)
                        {
                            grad3 = SetGrad(grad3PciValue);
                        }

                        #endregion

                        string text = string.Format("一级公路处于{0}等水平、二级公路处于{1}等水平、三级公路处于{2}等水平、四级公路处于{3}等水平。", grad0, grad1, grad2, grad3);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }


                    if (book.Name.Contains("gldjJudgeRqi"))
                    {

                        #region 获取各个公路等级评价 
                        double length0 = 0;
                        double length1 = 0;
                        double length2 = 0;
                        double length3 = 0;
                        double grad0RqiValue = 0;
                        double grad1RqiValue = 0;


                        double grad2RqiValue = 0;
                        double grad3RqiValue = 0;


                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 1)
                                {

                                    grad0RqiValue += item1.RqiValue * item1.RoadLen;
                                    length0 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 2)
                                {

                                    grad1RqiValue += item1.RqiValue * item1.RoadLen;
                                    length1 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }


                                if (item1.RoadGrad == 3)
                                {

                                    grad2RqiValue += item1.RqiValue * item1.RoadLen;
                                    length2 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);


                                }
                                if (item1.RoadGrad == 4)
                                {

                                    grad3RqiValue += item1.RqiValue * item1.RoadLen;
                                    length3 += item1.RoadLen;

                                    string temp = SetGrad(item1.PqiValue);

                                }
                            }

                        }

                        grad0RqiValue = double.IsNaN(grad0RqiValue / length0) ? 0 : grad0RqiValue / length0;


                        grad1RqiValue = double.IsNaN(grad1RqiValue / length1) ? 0 : grad1RqiValue / length1;


                        grad2RqiValue = double.IsNaN(grad2RqiValue / length2) ? 0 : grad2RqiValue / length2;


                        grad3RqiValue = double.IsNaN(grad3RqiValue / length3) ? 0 : grad3RqiValue / length3;
                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";
                        string grad3 = "-";
                        if (length0 != 0)
                        {
                            grad0 = SetGrad(grad0RqiValue);
                        }
                        if (length1 != 0)
                        {
                            grad1 = SetGrad(grad1RqiValue);
                        }
                        if (length2 != 0)
                        {
                            grad2 = SetGrad(grad2RqiValue);
                        }
                        if (length3 != 0)
                        {
                            grad3 = SetGrad(grad3RqiValue);
                        }

                        #endregion

                        string text = string.Format("一级公路处于{0}等水平、二级公路处于{1}等水平、三级公路处于{2}等水平、四级公路处于{3}等水平。", grad0, grad1, grad2, grad3);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }





                    if (book.Name.Contains("xzdjJudgePqi"))
                    {
                        double xiandaoLength = 0;
                        double xiangdaoLength = 0;
                        double cundaoLength = 0;
                        double xiandaoPqi = 0;
                        double xiandaoPci = 0;
                        double xiandaoRqi = 0;
                        double goodXianRoadLength = 0;
                        double goodXiangRoadLength = 0;
                        double goodCunRoadLength = 0;
                        //次差路长度
                        double ciChaXianRoadLength = 0;
                        double ciChaXiangRoadLength = 0;
                        double ciChaCunRoadLength = 0;
                        double xiangdaoPqi = 0;
                        double xiangdaoPci = 0;
                        double xiangdaoRqi = 0;

                        double cundaoPqi = 0;
                        double cundaoPci = 0;
                        double cundaoRqi = 0;
                        foreach (var item1 in allDatas)
                        {

                            foreach (var item in item1.datas)
                            {


                                if (item.RoadCode.StartsWith("X"))
                                {
                                    xiandaoLength += item.RoadLen;
                                    xiandaoPqi += item.PqiValue * item.RoadLen;
                                    xiandaoPci += item.PciValue * item.RoadLen;
                                    xiandaoRqi += item.RqiValue * item.RoadLen;
                                    string judge = SetGrad(item.PqiValue);
                                    if (judge == "优" || judge == "良")
                                    {
                                        goodXianRoadLength += item.RoadLen;
                                    }
                                    else if (judge == "次" || judge == "差")
                                    {
                                        ciChaXianRoadLength += item.RoadLen;
                                    }
                                }
                                if (item.RoadCode.StartsWith("Y"))
                                {
                                    xiangdaoLength += item.RoadLen;
                                    xiangdaoPci += item.PciValue * item.RoadLen;
                                    xiangdaoPqi += item.PqiValue * item.RoadLen;
                                    xiangdaoRqi += item.RqiValue * item.RoadLen;
                                    string judge = SetGrad(item.PqiValue);
                                    if (judge == "优" || judge == "良")
                                    {
                                        goodXiangRoadLength += item.RoadLen;
                                    }
                                    else if (judge == "次" || judge == "差")
                                    {
                                        ciChaXiangRoadLength += item.RoadLen;
                                    }
                                }
                                if (item.RoadCode.StartsWith("C"))
                                {
                                    cundaoLength += item.RoadLen;
                                    cundaoPci += item.PciValue * item.RoadLen;
                                    cundaoPqi += item.PqiValue * item.RoadLen;
                                    cundaoRqi += item.RqiValue * item.RoadLen;
                                    string judge = SetGrad(item.PqiValue);
                                    if (judge == "优" || judge == "良")
                                    {
                                        goodCunRoadLength += item.RoadLen;
                                    }
                                    else if (judge == "次" || judge == "差")
                                    {
                                        ciChaCunRoadLength += item.RoadLen;
                                    }
                                }

                            }
                        }
                        xiandaoPci /= xiandaoLength;
                        xiandaoPqi /= xiandaoLength;
                        xiandaoRqi /= xiandaoLength;
                        string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "-" : xiandaoPci.ToString("0.##");
                        string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "-" : xiandaoPqi.ToString("0.##");
                        string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "-" : xiandaoRqi.ToString("0.##");

                        xiangdaoPci /= xiangdaoLength;
                        xiangdaoPqi /= xiangdaoLength;
                        xiangdaoRqi /= xiangdaoLength;

                        string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "-" : xiangdaoPci.ToString("0.##");
                        string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "-" : xiangdaoPqi.ToString("0.##");
                        string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "-" : xiangdaoRqi.ToString("0.##");

                        cundaoPci /= cundaoLength;
                        cundaoPqi /= cundaoLength;
                        cundaoRqi /= cundaoLength;
                        string cundaoPciStr = double.IsNaN(cundaoPci) ? "-" : cundaoPci.ToString("0.##");
                        string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "-" : cundaoPqi.ToString("0.##");
                        string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "-" : cundaoRqi.ToString("0.##");
                        string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "-" : SetGrad(xiandaoPqi);
                        string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "-" : SetGrad(xiangdaoPqi);
                        string cundaojudge = double.IsNaN(cundaoPqi) ? "-" : SetGrad(cundaoPqi);

                        string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "-" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "-" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "-" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "-" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "-" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "-" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";

                        if (xiandaoLength != 0)
                        {
                            grad0 = SetGrad(xiandaoPqi);

                        }
                        if (xiangdaoLength != 0)
                        {
                            grad1 = SetGrad(xiangdaoPqi);
                        }
                        if (cundaoLength != 0)
                        {
                            grad2 = SetGrad(cundaoPqi);
                        }


                        string text = string.Format("县道处于{0}等水平、乡道处于{1}等水平、村道处于{2}等水平。", grad0, grad1, grad2);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }
                    if (book.Name.Contains("xzdjJudgePci"))
                    {
                        double xiandaoLength = 0;
                        double xiangdaoLength = 0;
                        double cundaoLength = 0;
                        double xiandaoPqi = 0;
                        double xiandaoPci = 0;
                        double xiandaoRqi = 0;
                        double goodXianRoadLength = 0;
                        double goodXiangRoadLength = 0;
                        double goodCunRoadLength = 0;
                        //次差路长度
                        double ciChaXianRoadLength = 0;
                        double ciChaXiangRoadLength = 0;
                        double ciChaCunRoadLength = 0;
                        double xiangdaoPqi = 0;
                        double xiangdaoPci = 0;
                        double xiangdaoRqi = 0;

                        double cundaoPqi = 0;
                        double cundaoPci = 0;
                        double cundaoRqi = 0;
                        foreach (var item in allDatas)
                        {

                            if (item.RoadCode.StartsWith("X"))
                            {
                                xiandaoLength += item.RoadLen;
                                xiandaoPqi += item.PqiValue * item.RoadLen;
                                xiandaoPci += item.PciValue * item.RoadLen;
                                xiandaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXianRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXianRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                xiangdaoLength += item.RoadLen;
                                xiangdaoPci += item.PciValue * item.RoadLen;
                                xiangdaoPqi += item.PqiValue * item.RoadLen;
                                xiangdaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXiangRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXiangRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("C"))
                            {
                                cundaoLength += item.RoadLen;
                                cundaoPci += item.PciValue * item.RoadLen;
                                cundaoPqi += item.PqiValue * item.RoadLen;
                                cundaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodCunRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaCunRoadLength += item.RoadLen;
                                }
                            }


                        }
                        xiandaoPci /= xiandaoLength;
                        xiandaoPqi /= xiandaoLength;
                        xiandaoRqi /= xiandaoLength;
                        string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "-" : xiandaoPci.ToString("0.##");
                        string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "-" : xiandaoPqi.ToString("0.##");
                        string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "-" : xiandaoRqi.ToString("0.##");

                        xiangdaoPci /= xiangdaoLength;
                        xiangdaoPqi /= xiangdaoLength;
                        xiangdaoRqi /= xiangdaoLength;

                        string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "-" : xiangdaoPci.ToString("0.##");
                        string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "-" : xiangdaoPqi.ToString("0.##");
                        string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "-" : xiangdaoRqi.ToString("0.##");

                        cundaoPci /= cundaoLength;
                        cundaoPqi /= cundaoLength;
                        cundaoRqi /= cundaoLength;
                        string cundaoPciStr = double.IsNaN(cundaoPci) ? "-" : cundaoPci.ToString("0.##");
                        string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "-" : cundaoPqi.ToString("0.##");
                        string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "-" : cundaoRqi.ToString("0.##");
                        string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "-" : SetGrad(xiandaoPqi);
                        string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "-" : SetGrad(xiangdaoPqi);
                        string cundaojudge = double.IsNaN(cundaoPqi) ? "-" : SetGrad(cundaoPqi);

                        string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "-" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "-" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "-" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "-" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "-" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "-" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";

                        if (xiandaoLength != 0)
                        {
                            grad0 = SetGrad(xiandaoPci);

                        }
                        if (xiangdaoLength != 0)
                        {
                            grad1 = SetGrad(xiangdaoPci);
                        }
                        if (cundaoLength != 0)
                        {
                            grad2 = SetGrad(cundaoPci);
                        }


                        string text = string.Format("县道处于{0}等水平、乡道处于{1}等水平、村道处于{2}等水平。", grad0, grad1, grad2);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }
                    if (book.Name.Contains("xzdjJudgeRqi"))
                    {
                        double xiandaoLength = 0;
                        double xiangdaoLength = 0;
                        double cundaoLength = 0;
                        double xiandaoPqi = 0;
                        double xiandaoPci = 0;
                        double xiandaoRqi = 0;
                        double goodXianRoadLength = 0;
                        double goodXiangRoadLength = 0;
                        double goodCunRoadLength = 0;
                        //次差路长度
                        double ciChaXianRoadLength = 0;
                        double ciChaXiangRoadLength = 0;
                        double ciChaCunRoadLength = 0;
                        double xiangdaoPqi = 0;
                        double xiangdaoPci = 0;
                        double xiangdaoRqi = 0;

                        double cundaoPqi = 0;
                        double cundaoPci = 0;
                        double cundaoRqi = 0;
                        foreach (var item in allDatas)
                        {

                            if (item.RoadCode.StartsWith("X"))
                            {
                                xiandaoLength += item.RoadLen;
                                xiandaoPqi += item.PqiValue * item.RoadLen;
                                xiandaoPci += item.PciValue * item.RoadLen;
                                xiandaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXianRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXianRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                xiangdaoLength += item.RoadLen;
                                xiangdaoPci += item.PciValue * item.RoadLen;
                                xiangdaoPqi += item.PqiValue * item.RoadLen;
                                xiangdaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXiangRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXiangRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("C"))
                            {
                                cundaoLength += item.RoadLen;
                                cundaoPci += item.PciValue * item.RoadLen;
                                cundaoPqi += item.PqiValue * item.RoadLen;
                                cundaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodCunRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaCunRoadLength += item.RoadLen;
                                }
                            }


                        }
                        xiandaoPci /= xiandaoLength;
                        xiandaoPqi /= xiandaoLength;
                        xiandaoRqi /= xiandaoLength;
                        string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "-" : xiandaoPci.ToString("0.##");
                        string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "-" : xiandaoPqi.ToString("0.##");
                        string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "-" : xiandaoRqi.ToString("0.##");

                        xiangdaoPci /= xiangdaoLength;
                        xiangdaoPqi /= xiangdaoLength;
                        xiangdaoRqi /= xiangdaoLength;

                        string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "-" : xiangdaoPci.ToString("0.##");
                        string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "-" : xiangdaoPqi.ToString("0.##");
                        string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "-" : xiangdaoRqi.ToString("0.##");

                        cundaoPci /= cundaoLength;
                        cundaoPqi /= cundaoLength;
                        cundaoRqi /= cundaoLength;
                        string cundaoPciStr = double.IsNaN(cundaoPci) ? "-" : cundaoPci.ToString("0.##");
                        string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "-" : cundaoPqi.ToString("0.##");
                        string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "-" : cundaoRqi.ToString("0.##");
                        string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "-" : SetGrad(xiandaoPqi);
                        string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "-" : SetGrad(xiangdaoPqi);
                        string cundaojudge = double.IsNaN(cundaoPqi) ? "-" : SetGrad(cundaoPqi);

                        string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "-" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "-" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "-" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "-" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "-" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "-" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";

                        if (xiandaoLength != 0)
                        {
                            grad0 = SetGrad(xiandaoRqi);

                        }
                        if (xiangdaoLength != 0)
                        {
                            grad1 = SetGrad(xiangdaoRqi);
                        }
                        if (cundaoLength != 0)
                        {
                            grad2 = SetGrad(cundaoRqi);
                        }


                        string text = string.Format("县道处于{0}等水平、乡道处于{1}等水平、村道处于{2}等水平。", grad0, grad1, grad2);
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }

                    if (book.Name.Contains("lmlxJudgeStr"))
                    {
                        double snLength = 0;
                        double lqLength = 0;
                        double allLength = 0;
                        double lqPqiValue = 0;
                        double snPqiValue = 0;
                        double lqPciValue = 0;
                        double snPciValue = 0;
                        double lqRqiValue = 0;
                        double snRqiValue = 0;
                        double goodLqLength = 0;
                        double goodSnLength = 0;
                        double cichaLqLength = 0;
                        double cichaSnLength = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                allLength += item1.RoadLen;
                                if (item1.RoadType == 0)
                                {
                                    lqLength += item1.RoadLen;
                                    lqPciValue += item1.PciValue * item1.RoadLen;
                                    lqPqiValue += item1.PqiValue * item1.RoadLen;
                                    lqRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodLqLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaLqLength += item1.RoadLen;
                                    }
                                }
                                if (item1.RoadType == 1)
                                {
                                    snLength += item1.RoadLen;
                                    snPciValue += item1.PciValue * item1.RoadLen;
                                    snPqiValue += item1.PqiValue * item1.RoadLen;
                                    snRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodSnLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaSnLength += item1.RoadLen;
                                    }
                                }
                            }
                        }
                        lqPqiValue /= lqLength;
                        snPqiValue /= snLength;
                        lqPciValue /= lqLength;
                        snPciValue /= snLength;
                        lqRqiValue /= lqLength;
                        snRqiValue /= snLength;
                        string lqPqiValueStr, snPqiValueStr, lqPciValueStr, snPciValueStr, lqRqiValueStr, snRqiValueStr;

                        if (lqLength == 0)
                        {
                            lqPqiValueStr = "-";
                            lqPciValueStr = "-";
                            lqRqiValueStr = "-";
                        }
                        else
                        {
                            lqPqiValueStr = lqPqiValue.ToString("0.##");
                            lqPciValueStr = lqPciValue.ToString("0.##");
                            lqRqiValueStr = lqRqiValue.ToString("0.##");
                        }
                        if (snLength == 0)
                        {
                            snPqiValueStr = "-";
                            snPciValueStr = "-";
                            snRqiValueStr = "-";
                        }
                        else
                        {
                            snPqiValueStr = snPqiValue.ToString("0.##");
                            snPciValueStr = snPciValue.ToString("0.##");
                            snRqiValueStr = snRqiValue.ToString("0.##");
                        }


                        string text = string.Format("沥青混凝土路面路况性能各指标（PQI：{0}、PCI：{1}、RQI：{2}），水泥混凝土路面各指标（PQI：{3}、PCI：{4}、RQI：{5}）。",
                           lqPqiValueStr, lqPciValueStr, lqRqiValueStr, snPqiValueStr, snPciValueStr, snRqiValueStr);

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }



                    if (book.Name.Contains("xzdhSumStr"))
                    {
                        double xiandaoLength = 0;
                        double xiangdaoLength = 0;
                        double cundaoLength = 0;
                        double xiandaoPqi = 0;
                        double xiandaoPci = 0;
                        double xiandaoRqi = 0;
                        double goodXianRoadLength = 0;
                        double goodXiangRoadLength = 0;
                        double goodCunRoadLength = 0;
                        //次差路长度
                        double ciChaXianRoadLength = 0;
                        double ciChaXiangRoadLength = 0;
                        double ciChaCunRoadLength = 0;
                        double xiangdaoPqi = 0;
                        double xiangdaoPci = 0;
                        double xiangdaoRqi = 0;

                        double cundaoPqi = 0;
                        double cundaoPci = 0;
                        double cundaoRqi = 0;
                        foreach (var item in allDatas)
                        {

                            if (item.RoadCode.StartsWith("X"))
                            {
                                xiandaoLength += item.RoadLen;
                                xiandaoPqi += item.PqiValue * item.RoadLen;
                                xiandaoPci += item.PciValue * item.RoadLen;
                                xiandaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXianRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXianRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                xiangdaoLength += item.RoadLen;
                                xiangdaoPci += item.PciValue * item.RoadLen;
                                xiangdaoPqi += item.PqiValue * item.RoadLen;
                                xiangdaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodXiangRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaXiangRoadLength += item.RoadLen;
                                }
                            }
                            if (item.RoadCode.StartsWith("C"))
                            {
                                cundaoLength += item.RoadLen;
                                cundaoPci += item.PciValue * item.RoadLen;
                                cundaoPqi += item.PqiValue * item.RoadLen;
                                cundaoRqi += item.RqiValue * item.RoadLen;
                                string judge = SetGrad(item.PqiValue);
                                if (judge == "优" || judge == "良")
                                {
                                    goodCunRoadLength += item.RoadLen;
                                }
                                else if (judge == "次" || judge == "差")
                                {
                                    ciChaCunRoadLength += item.RoadLen;
                                }
                            }


                        }
                        xiandaoPci /= xiandaoLength;
                        xiandaoPqi /= xiandaoLength;
                        xiandaoRqi /= xiandaoLength;
                        string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "-" : xiandaoPci.ToString("0.##");
                        string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "-" : xiandaoPqi.ToString("0.##");
                        string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "-" : xiandaoRqi.ToString("0.##");

                        xiangdaoPci /= xiangdaoLength;
                        xiangdaoPqi /= xiangdaoLength;
                        xiangdaoRqi /= xiangdaoLength;

                        string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "-" : xiangdaoPci.ToString("0.##");
                        string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "-" : xiangdaoPqi.ToString("0.##");
                        string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "-" : xiangdaoRqi.ToString("0.##");

                        cundaoPci /= cundaoLength;
                        cundaoPqi /= cundaoLength;
                        cundaoRqi /= cundaoLength;
                        string cundaoPciStr = double.IsNaN(cundaoPci) ? "-" : cundaoPci.ToString("0.##");
                        string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "-" : cundaoPqi.ToString("0.##");
                        string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "-" : cundaoRqi.ToString("0.##");
                        string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "-" : SetGrad(xiandaoPqi);
                        string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "-" : SetGrad(xiangdaoPqi);
                        string cundaojudge = double.IsNaN(cundaoPqi) ? "-" : SetGrad(cundaoPqi);

                        string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "-" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "-" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "-" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "-" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
                        string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "-" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
                        string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "-" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
                        string grad0 = "-";
                        string grad1 = "-";
                        string grad2 = "-";

                        if (xiandaoLength != 0)
                        {
                            grad0 = SetGrad(xiandaoPqi);

                        }
                        else
                        {
                            xiandaoPqi = 0;
                        }
                        if (xiangdaoLength != 0)
                        {
                            grad1 = SetGrad(xiangdaoPqi);
                        }
                        else
                        {
                            xiangdaoPqi = 0;
                        }
                        if (cundaoLength != 0)
                        {
                            grad2 = SetGrad(cundaoPqi);
                        }
                        else
                        {
                            cundaoPqi = 0;
                        }
                        List<string> values = new List<string>();
                        values.Add(xiangdaoPqi.ToString() + "_" + "乡道");
                        values.Add(xiandaoPqi.ToString() + "_" + "县道");
                        values.Add(cundaoPqi.ToString() + "_" + "村道");
                        values.Sort((t1, t2) =>
                        {
                            double temp1 = double.Parse(t1.Split('_').First());
                            double temp2 = double.Parse(t2.Split('_').First());

                            return temp2.CompareTo(temp1);


                        });

                        string text = string.Format("从路面技术状况指数PQI来看，{0}整体状况最好，其次为{1}，最低为{2}。", values.ElementAt(0).Split('_').Last(), values.ElementAt(1).Split('_').Last(), values.ElementAt(2).Split('_').Last());
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }

                    if (book.Name.Contains("lmlxJudgePciStr"))
                    {
                        double snLength = 0;
                        double lqLength = 0;
                        double allLength = 0;
                        double lqPqiValue = 0;
                        double snPqiValue = 0;
                        double lqPciValue = 0;
                        double snPciValue = 0;
                        double lqRqiValue = 0;
                        double snRqiValue = 0;
                        double goodLqLength = 0;
                        double goodSnLength = 0;
                        double cichaLqLength = 0;
                        double cichaSnLength = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                allLength += item1.RoadLen;
                                if (item1.RoadType == 0)
                                {
                                    lqLength += item1.RoadLen;
                                    lqPciValue += item1.PciValue * item1.RoadLen;
                                    lqPqiValue += item1.PqiValue * item1.RoadLen;
                                    lqRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodLqLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaLqLength += item1.RoadLen;
                                    }
                                }
                                if (item1.RoadType == 1)
                                {
                                    snLength += item1.RoadLen;
                                    snPciValue += item1.PciValue * item1.RoadLen;
                                    snPqiValue += item1.PqiValue * item1.RoadLen;
                                    snRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodSnLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaSnLength += item1.RoadLen;
                                    }
                                }
                            }
                        }
                        lqPqiValue /= lqLength;
                        snPqiValue /= snLength;
                        lqPciValue /= lqLength;
                        snPciValue /= snLength;
                        lqRqiValue /= lqLength;
                        snRqiValue /= snLength;

                        string lqStr, snStr;
                        if (lqLength == 0)
                        {
                            lqStr = "-";
                        }
                        else
                        {
                            lqStr = SetGrad(lqPciValue);

                        }
                        if (snLength == 0)
                        {
                            snStr = "-";
                        }
                        else
                        {
                            snStr = SetGrad(snPciValue);
                        }

                        string text = string.Format("沥青混凝土路面评价等级为{0}等水平，水泥混凝土路面评价等级为{1}等水平。",
                           lqStr, snStr);

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }



                    if (book.Name.Contains("lmlxJudgeRqiStr"))
                    {
                        double snLength = 0;
                        double lqLength = 0;
                        double allLength = 0;
                        double lqPqiValue = 0;
                        double snPqiValue = 0;
                        double lqPciValue = 0;
                        double snPciValue = 0;
                        double lqRqiValue = 0;
                        double snRqiValue = 0;
                        double goodLqLength = 0;
                        double goodSnLength = 0;
                        double cichaLqLength = 0;
                        double cichaSnLength = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                allLength += item1.RoadLen;
                                if (item1.RoadType == 0)
                                {
                                    lqLength += item1.RoadLen;
                                    lqPciValue += item1.PciValue * item1.RoadLen;
                                    lqPqiValue += item1.PqiValue * item1.RoadLen;
                                    lqRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodLqLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaLqLength += item1.RoadLen;
                                    }
                                }
                                if (item1.RoadType == 1)
                                {
                                    snLength += item1.RoadLen;
                                    snPciValue += item1.PciValue * item1.RoadLen;
                                    snPqiValue += item1.PqiValue * item1.RoadLen;
                                    snRqiValue += item1.RqiValue * item1.RoadLen;
                                    string temp = SetGrad(item1.PqiValue);
                                    if (temp == "优" || temp == "良")
                                    {
                                        goodSnLength += item1.RoadLen;
                                    }
                                    else
                                    {
                                        cichaSnLength += item1.RoadLen;
                                    }
                                }
                            }
                        }
                        lqPqiValue /= lqLength;
                        snPqiValue /= snLength;
                        lqPciValue /= lqLength;
                        snPciValue /= snLength;
                        lqRqiValue /= lqLength;
                        snRqiValue /= snLength;

                        string lqStr, snStr;
                        if (lqLength == 0)
                        {
                            lqStr = "-";
                        }
                        else
                        {
                            lqStr = SetGrad(lqRqiValue);

                        }
                        if (snLength == 0)
                        {
                            snStr = "-";
                        }
                        else
                        {
                            snStr = SetGrad(snRqiValue);
                        }

                        string text = string.Format("沥青混凝土路面平整度评价为{0}等水平，水泥混凝土路面平整度评价为{1}等水平。",
                           lqStr, snStr);

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(text);
                        continue;
                    }




                    #endregion
                    #region PQI
                    //pqi值
                    if (book.Name.Contains("PqiValue"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(RoadPqiValue.ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("PqiRoadGrad"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(PqiRoadGrad);
                        continue;
                    }
                    if (book.Name.Contains("PciRoadGrad"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(PciRoadGrad);
                        continue;
                    }
                    if (book.Name.Contains("RqiRoadGrad"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(RqiRoadGrad);
                        continue;
                    }
                    // 优良路率
                    if (book.Name.Contains("PqiGoodRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPqiYLRate * 100).ToString("0.##"));
                        continue;
                    }
                    //次差路率
                    if (book.Name.Contains("PqiCCRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPqiCCRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("yPQIRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(yPQIRoadCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("yPQIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(yPQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("lPQIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lPQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("zPQIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zPQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("ciPQIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ciPQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("chaPqiRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(chaPQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("yPcIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(yPcIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("lPcIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lPcIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("zPCIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zPcIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("ciPcIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ciPcIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("chaPcIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(chaPcIRoadLength.ToString());
                        continue;
                    }

                    if (book.Name.Contains("yRqIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(yRQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("lRqIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lRQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("zRqIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zRQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("ciRqIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(CiRQIRoadLength.ToString());
                        continue;
                    }
                    if (book.Name.Contains("chaRqIRoadLength"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ChaRQIRoadLength.ToString());
                        continue;
                    }



                    if (book.Name.Contains("lPQIRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lPQIRoadCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("zPQIRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zPQIRoadCount.ToString());

                        continue;
                    }
                    if (book.Name.Contains("ciPQIRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ciPQIRoadCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("chaPQIRoadCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(chaPQIRoadCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("pqiTopFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pqiTopFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("pqiLastFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pqiLastFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("pqiSubaverageCout"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pqiSubaverageCout.ToString());

                        continue;

                    }
                    if (book.Name.Contains("pqiSubaverageRate"))
                    {
                        double pqiSubaverageCoutRate = pqiSubaverageCout / double.Parse(allDatas.Count.ToString());
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((pqiSubaverageCoutRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RoadPqiYLCountRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPqiYLCountRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RoadPciYLCountRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPciYLCountRate * 100).ToString("0.##"));
                    }
                    if (book.Name.Contains("RoadRqiYLCountRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadRqiYLCountRate * 100).ToString("0.##"));
                    }


                    #region 图表
                    if (book.Name.Contains("pqiPic001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        System.Windows.Forms.Clipboard.Clear();
                        getPqiPicture001();
                        Thread.Sleep(GlobalWord.wd_sleep_us);

                        currentSelection.Range.Paste();

                        Thread.Sleep(GlobalWord.wd_sleep_us);

                        System.Windows.Forms.Clipboard.Clear();
                        currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        continue;
                    }
                    if (book.Name.Contains("PQITable001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPqiTable001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);

                    }
                    #endregion
                    #endregion
                    #region RQI
                    if (book.Name.Contains("RqiValue"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(RoadRqiValue.ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RqiYLRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadRqiYLRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RqiCCRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadRqiCCRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RqiEvaluate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(SetGrad(RoadRqiValue));
                        continue;
                    }

                    if (book.Name.Contains("yRQICount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;

                        //优

                        currentSelection.TypeText(yRQICount.ToString());
                        continue;

                    }
                    if (book.Name.Contains("lRQICount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lRQICount.ToString());
                        continue;
                    }
                    else if (book.Name.Contains("zRQICount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zRQICount.ToString());
                        continue;

                    }
                    if (book.Name.Contains("ciRQICount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ciRQICount.ToString());
                        continue;
                    }

                    if (book.Name.Contains("chaRQICount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(chaRQICount.ToString());
                        continue;
                    }
                    //优良rqi
                    if (book.Name.Contains("cichaRQIRate"))
                    {
                        double countRate = ((double)ciRQICount + (double)chaRQICount) / (double)allDatas.Count;
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((countRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("ylRQIRate"))
                    {
                        double countRate = ((double)yRQICount + (double)lRQICount) / (double)allDatas.Count;
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((countRate * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("RQITopFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(rqiTopFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("RQILastFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(rqiLastFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("RQISubaverageCout"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(rqiSubaverageCout.ToString());

                        continue;

                    }
                    if (book.Name.Contains("RQISubaverageRate"))
                    {
                        double RQISubaverageCoutRate = rqiSubaverageCout / double.Parse(allDatas.Count.ToString());
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RQISubaverageCoutRate * 100).ToString("0.##"));
                        continue;
                    }

                    #region 图表
                    if (book.Name.Contains("RQIPicture001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        while (true)
                        {
                            try
                            {
                                System.Windows.Forms.Clipboard.Clear();
                                getRQIPicture001();
                                currentSelection.Range.Paste();
                                Thread.Sleep(GlobalWord.wd_sleep_us);
                                currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                break;
                            }
                            catch (Exception ex)
                            {

                                Thread.Sleep(GlobalWord.wd_sleep_us);
                            }
                        }
                        continue;
                    }
                    if (book.Name.Contains("RQIPicture002"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        while (true)
                        {
                            try
                            {
                                System.Windows.Forms.Clipboard.Clear();
                                getRQIPicture002();
                                currentSelection.Range.Paste();
                                Thread.Sleep(GlobalWord.wd_sleep_us);
                                currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                break;
                            }
                            catch (Exception ex)
                            {

                                Thread.Sleep(GlobalWord.wd_sleep_us);
                            }
                        }
                        continue;
                    }

                    if (book.Name.Contains("RQITable001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getRqiTable001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);

                    }



                    #endregion
                    #endregion
                    #region PCI
                    if (book.Name.Contains("PciValue"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(RoadPciValue.ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("PciGoodRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPciYLRate * 100).ToString("0.##"));
                        continue;
                    }
                    //pci 次差路率
                    if (book.Name.Contains("PciCCRate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((RoadPciCCRate * 100).ToString("0.##"));
                        continue;
                    }
                    //pci评价
                    if (book.Name.Contains("PciEvaluate"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(SetGrad(RoadPciValue));
                        continue;
                    }
                    if (book.Name.Contains("RoadCount002"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(allDatas.Count.ToString());
                        continue;
                    }
                    if (book.Name.Contains("yPciCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;

                        //优

                        currentSelection.TypeText(yPciCount.ToString());
                        continue;

                    }
                    if (book.Name.Contains("lPciCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(lPciCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("zPciCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(zPciCount.ToString());
                        continue;

                    }
                    if (book.Name.Contains("ciPciCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ciPciCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("chaPciCount"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(chaPciCount.ToString());
                        continue;
                    }
                    if (book.Name.Contains("pciTopFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pciTopFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("pciLastFiveRoadStr"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pciLastFiveRoadStr);
                        currentSelection.Font.Bold = 1; // 加粗
                        continue;
                    }
                    if (book.Name.Contains("pciSubaverageCout"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(pciSubaverageCout.ToString());

                        continue;

                    }
                    if (book.Name.Contains("pciSubaverageRate"))
                    {
                        double pciSubaverageCoutRate = pciSubaverageCout / double.Parse(allDatas.Count.ToString());
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((pciSubaverageCoutRate * 100).ToString("0.##"));
                        continue;
                    }

                    #region 插入图表
                    if (book.Name.Contains("PciPicture001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPciPicture001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }

                    }
                    if (book.Name.Contains("PciPicture002"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPciPicture002))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }

                    }

                    if (book.Name.Contains("PCITable001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPciTable001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        continue;
                    }
                    #endregion
                    #endregion
                    #region 指标统计
                    if (book.Name.Contains("IndexStatisticsTable"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getIndexStatisticsTable(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);


                        continue;
                    }
                    #endregion

                    #region 第四部分
                    //县道长度
                    if (book.Name.Contains("RoadCountyLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("X"))
                            {
                                length += item.RoadLen;
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("RoadCountyCount"))
                    {
                        int length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("X"))
                            {
                                length++;
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString());
                        continue;
                    }

                    if (book.Name.Contains("RoadCountyRate"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("X"))
                            {
                                length += item.RoadLen;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    //乡道
                    if (book.Name.Contains("RoadVillageLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                length += item.RoadLen;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("RoadVillageCount"))
                    {
                        int length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                length++;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString());
                        continue;
                    }

                    if (book.Name.Contains("RoadVillageRate"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("Y"))
                            {
                                length += item.RoadLen;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    //村道
                    if (book.Name.Contains("RoadHamletLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("C"))
                            {
                                length += item.RoadLen;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("RoadHamletCount"))
                    {
                        int length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("C"))
                            {
                                length++;
                            }
                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString());
                        continue;
                    }
                    if (book.Name.Contains("RoadHamletRate"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            if (item.RoadCode.StartsWith("C"))
                            {
                                length += item.RoadLen;
                            }


                        }

                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    } 

                    //道路比例分布
                    if (book.Name.Contains("PavementDistributionRatioPic"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPart5Pic001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }

                    if (book.Name.Contains("PavementDistributionRatioTable"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPavementDistributionRatioTable001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        continue;
                    }


                    if (book.Name.Contains("Part5PqiPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart5Picture001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }

                    }
                    if (book.Name.Contains("Part5PciPci001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPart5Picture002))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }
                    if (book.Name.Contains("Part5RqiPci001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart5Picture003))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }

                    #endregion

                    #region 第五部分
                    if (book.Name.Contains("p_name"))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(m_countyName);
                        continue;
                    }
                    if (book.Name.Contains("FirstRoadLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 1)
                                {
                                    length += item1.RoadLen;
                                }
                            }



                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("TwoRoadLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 2)
                                {
                                    length += item1.RoadLen;
                                }
                            }



                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }

                    if (book.Name.Contains("ThreeRoadLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 3)
                                {
                                    length += item1.RoadLen;
                                }
                            }



                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("ThreeRoadRate"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 3)
                                {
                                    length += item1.RoadLen;
                                }
                            }

                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("OneRoadRateAdd"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 1)
                                {
                                    length += item1.RoadLen;
                                }
                            }

                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("TwoRoadRateAdd"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 2)
                                {
                                    length += item1.RoadLen;
                                }
                            }

                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("FourRoadLength"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 4)
                                {
                                    length += item1.RoadLen;
                                }
                            }

                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("FourRoadRate"))
                    {
                        double length = 0;
                        foreach (var item in allDatas)
                        {
                            foreach (var item1 in item.datas)
                            {
                                if (item1.RoadGrad == 4)
                                {
                                    length += item1.RoadLen;
                                }
                            }

                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }

                    if (book.Name.Contains("Part6Picture001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart6Pic001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }

                    if (book.Name.Contains("Part6Table001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPart6Table001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        continue;
                    }

                    if (book.Name.Contains("Part6PqiPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart6Picture001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }
                    if (book.Name.Contains("Part6PciPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart6Picture002))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }
                    if (book.Name.Contains("Part6RqiPci001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPart6Picture003))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }

                    }
                    #endregion
                    #region 第七部分
                    if (book.Name.Contains("lqRoadLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                if (item.RoadType == 0)
                                {
                                    length += item.RoadLen;
                                }
                            }



                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }

                    if (book.Name.Contains("lqRoadRate"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                if (item.RoadType == 0)
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }
                    if (book.Name.Contains("snRoadLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                if (item.RoadType == 1)
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("snRoadRate"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                if (item.RoadType == 1)
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText((length / RoadLength * 100).ToString("0.##"));
                        continue;
                    }

                    if (book.Name.Contains("Part7Picture001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPart7Pic001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }

                    if (book.Name.Contains("Part7Table001"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPart7Table001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        continue;
                    }

                    if (book.Name.Contains("Part7PqiPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;


                        if (PastePic(getPart7Picture001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }



                    }
                    if (book.Name.Contains("Part7PciPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        if (PastePic(getPart7Picture002))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }


                    }
                    if (book.Name.Contains("Part7RqiPci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;


                        if (PastePic(getPart7Picture003))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }

                    }

                    //附录
                    if (book.Name.Contains("Part10"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPart10Table002(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);

                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        continue;
                    }

                    if (book.Name.Contains("Part8CiChaPqiLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                string temp = SetGrad(item.PqiValue);
                                if (temp == "次" || temp == "差")
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }

                    if (book.Name.Contains("Part8CiChaPciLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                string temp = SetGrad(item.PciValue);
                                if (temp == "次" || temp == "差")
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }

                    if (book.Name.Contains("Part8CiChaRqiLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                string temp = SetGrad(item.RqiValue);
                                if (temp == "次" || temp == "差")
                                {
                                    length += item.RoadLen;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("Part8CiChaRciAndRqiLength"))
                    {
                        double length = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                string temp = SetGrad(item.PciValue);
                                string temp1 = SetGrad(item.RqiValue);
                                if (temp == "次" || temp == "差")
                                {
                                    if (temp1 == "次" || temp1 == "差")
                                    {
                                        length += item.RoadLen;
                                    }

                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(length.ToString("0.###"));
                        continue;
                    }
                    if (book.Name.Contains("Part8Pci"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;

                        if (PastePic(getPart8Picture001))
                        {
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            continue;
                        }



                    }
                    if (book.Name.Contains("Part8CiChaCount"))
                    {
                        int count = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {

                                string temp2 = SetGrad(item.PqiValue);

                                if (temp2 == "次" || temp2 == "差")
                                {
                                    count++;
                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(count.ToString());
                        continue;
                    }
                    if (book.Name.Contains("Part8CiChaPciPqiCount"))
                    {
                        int count = 0;
                        foreach (var item1 in allDatas)
                        {
                            foreach (var item in item1.datas)
                            {
                                string temp = SetGrad(item.PciValue);
                                string temp1 = SetGrad(item.RqiValue);
                                if (temp == "次" || temp == "差")
                                {
                                    if (temp1 == "次" || temp1 == "差")
                                    {
                                        count++;
                                    }

                                }
                            }
                        }
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(count.ToString());
                        continue;
                    }
                    if (book.Name.Contains("Part8CiChaTable002"))
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        int tableCnt = 1;
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;
                        getPart8Table001(ref dataRange, ref excelApp);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        continue;
                    }
                    #endregion
#endif
                }

                if (NeedDis)
                {
                    foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                    {
                        if (book.Name.Contains("bhhzTable001"))
                        {
#if bhhz
                            Dictionary<string, double> allDisFormXml = new Dictionary<string, double>();

                            List<FileInfo> disPaths = new List<FileInfo>();
                            foreach (HeFeiRaod road in allDatas)
                            {
                                string code = road.RoadCode;
                                foreach (FileInfo file in excelPathList)
                                {
                                    string name = file.Name;

                                    if (name.Contains(code))
                                    {
                                        if (disPaths.Contains(file))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            disPaths.Add(file);
                                        }
                                    }
                                }

                            }

                            List<DiseaseHeFei> allDis = null;
                            book.Select();


                            try
                            {
                                getDiseaseSumMessage(excelPathList, ref allDisFormXml, ref allDis);
                            }
                            catch (Exception ex)
                            {

                                Console.WriteLine("Error in getDiseaseSumMessage: " + ex.Message);
                            }




                            if (allDisFormXml.Count == 0)
                            {
                                continue;
                            }
                            book.Select();
                            currentSelection = wordApp.Selection;
                            int tableCnt = 1;
                            MSExcel.Range dataRange = null;
                            MSExcel.Application excelApp = null;
                            getDiseaseSumTable001(allDisFormXml, ref dataRange, ref excelApp);
                            CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                            CWB_ExcelHelper.disposeExcel(ref excelApp);
                            MSWord.Table tbl = currentSelection.Range.Tables[1];
                            CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);
                            AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                            FromatTable(wordApp, tbl, 0, 1, null, "BG_CCC1", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);

                            //根据路段的病害汇总表格

                            getDiseaseSumTable002(basePath + "/病害汇总excel.xlsx", allDisFormXml, allDis, ref excelApp);
                            CWB_ExcelHelper.disposeExcel(ref excelApp);

                            continue;
#endif
                        }
                    }
                }



            }

            try
            {
                if (m_outPartEight)
                {
                    var thread = new Thread(WorkWithComObjects);
                    thread.SetApartmentState(ApartmentState.STA);
                    thread.Start(wordDoc);
                    thread.Join();
                }
                // 更新页码
                wordDoc.Fields.Update();
                // 找到文档中的目录（Table of Contents）
                foreach (MSWord.TableOfContents toc in wordDoc.TablesOfContents)
                {
                    toc.Update();
                }
                CWB_WordHelper.ExportWordModel2(wordDoc);
            }
            catch (Exception ex)
            {


                throw ex;
            }

        }
        private MSExcel.Workbook CollectBookTemp = null;

        private Dictionary<string, List<(string, double)>> handelAllDiseaseToRoadPart(List<DiseaseHeFei> allDis)
        {
            Dictionary<string, List<DiseaseHeFei>> datas = new Dictionary<string, List<DiseaseHeFei>>();
            // 创建一个辅助字典，用于将 allDis 中的数据按照 RoadCode 分组
            Dictionary<string, List<DiseaseHeFei>> allDisGrouped = allDis.GroupBy(d => d.RoadCode)
                                                                        .ToDictionary(g => g.Key, g => g.ToList());
            foreach (var item in allDatas)
            {
                string roadCode = item.RoadCode;
                string name = item.RoadName;
                string key = name + "(" + roadCode + ")";

                // 如果 allDis 中存在当前路段的数据，则将其添加到 datas 中
                if (allDisGrouped.ContainsKey(roadCode))
                {
                    datas.Add(key, allDisGrouped[roadCode]);
                }
                else
                {
                    datas.Add(key, new List<DiseaseHeFei>());
                }
            }

            Dictionary<string, List<(string, double)>> result = new Dictionary<string, List<(string, double)>>();
            foreach (var item in datas)
            {
                if (!result.Keys.Contains(item.Key))
                {
                    result.Add(item.Key, new List<(string, double)>());
                }
                for (int i = 0; i < item.Value.Count; i++)
                {
                    var disItem = item.Value[i];
                    DiseaseFormXml disOnly = null;
                    List<DiseaseFormXml> findDis = AllDiseaseFormXml.Where(
                                    t => t.DiseaseName.Contains(disItem.Name)
                                    && t.RoadType.Contains(disItem.RoadType)
                                    && t.RoadGrad.Contains(disItem.RoadGrad)
                                    ).ToList();
                    if (findDis.Count == 0)
                    {
                        continue;
                    }
                    else if (findDis.Count == 1)
                    {
                        disOnly = findDis.First();
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(disItem.DamagedCondition) || disItem.DamagedCondition == "无")
                        {
                            disOnly = findDis.Where(t => t.DiseaseName.Split('.').Length == 1).FirstOrDefault();
                        }
                        else
                        {
                            disOnly = findDis.Where(t => t.DiseaseName.Contains(disItem.DamagedCondition)).FirstOrDefault();
                        }
                    }
                    //judge  type
                    int findIndex = result[item.Key].FindIndex(t => t.Item1 == disOnly.Number);
                    if (findIndex >= 0)
                    {
                        //find index
                        double nowValue = result[item.Key][findIndex].Item2;
                        double increaseAreaValue = nowValue + disItem.Area;
                        //remove old item
                        result[item.Key].RemoveAt(findIndex);

                        result[item.Key].Add((disOnly.Number, increaseAreaValue));
                    }
                    else
                    {
                        result[item.Key].Add((disOnly.Number, disItem.Area));
                    }
                }
            }
            return result;
        }


        private void getDiseaseSumMessage(List<FileInfo> disFileList, ref Dictionary<string, double> allDisFormXml, ref List<DiseaseHeFei> allDisFull)
        {
            allDisFull = new List<DiseaseHeFei>();
            var allDis = new List<DiseaseHeFei>();
            for (int i = 0; i < disFileList.Count; ++i)
            {
                var item = disFileList[i];


                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dtProject = new System.Data.DataTable();

                try
                {
                    ReadExcelData(ref dt, item.FullName, "病害列表", 2, true);
                    ReadExcelData(ref dtProject, item.FullName, "工程信息", 2, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(item.Name + "病害文件读取出错,请检查!");
                    throw ex;
                }


                int colCount = dt.Columns.Count;
                string roadGrad = dtProject.Rows[6][1].ToString();
                string roadCode = dtProject.Rows[2][1].ToString();

                foreach (DataRow row in dt.Rows)
                {
                    DiseaseHeFei dis = new DiseaseHeFei();
                    dis.RoadGrad = roadGrad;
                    dis.RoadCode = roadCode;
                    string disName = row[2].ToString();
                    string disGrad = row[3].ToString();
                    string roadType = "";
                    dis.DamagedCondition = disGrad;
                    dis.Name = disName;
                    if (string.IsNullOrWhiteSpace(disGrad) || disGrad == "无")
                    {
                        disGrad = "";
                    }
                    double area = 0;
                    try
                    {
                        if (colCount > 10) //大框
                        {
                            string temp = row[7].ToString();
                            string temp1 = row[12].ToString();
                            if (!string.IsNullOrWhiteSpace(temp1))
                            {
                                roadType = temp1;
                            }
                            else
                            {
                                continue;
                            }
                            if (!string.IsNullOrWhiteSpace(temp))
                            {
                                area = double.Parse(temp);
                            }
                            else
                            {
                                continue;
                            }

                        }
                        else
                        {
                            string temp = row[4].ToString();
                            string temp1 = row[8].ToString();
                            if (!string.IsNullOrWhiteSpace(temp1))
                            {
                                roadType = temp1;
                            }
                            else
                            {
                                continue;
                            }
                            if (!string.IsNullOrWhiteSpace(temp))
                            {
                                area = double.Parse(temp);
                            }
                            else
                            {
                                continue;
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                    if (!string.IsNullOrEmpty(disName))
                    {
                        dis.RoadType = roadType;
                        dis.Area = area;
                        DiseaseHeFei disCopy = new DiseaseHeFei(dis);
                        allDisFull.Add(disCopy);
                        bool hasDis = false;
                        for (int t = 0; t < allDis.Count; ++t)
                        {
                            var itemDis = allDis.ElementAt(t);

                            if (itemDis.RoadType == dis.RoadType
                                && itemDis.DamagedCondition == dis.DamagedCondition
                                && itemDis.Name == dis.Name
                                && itemDis.RoadGrad == dis.RoadGrad
                                )
                            {
                                allDis[t].Area += dis.Area;
                                hasDis = true;
                            }
                        }
                        if (!hasDis)//第一次出现该病害
                        {
                            allDis.Add(dis);
                        }
                    }

                }
            }

            foreach (var disItem in allDis)
            {
                DiseaseFormXml disOnly = null;
                List<DiseaseFormXml> findDis = AllDiseaseFormXml.Where(
                     t => t.DiseaseName.Contains(disItem.Name)
                     && t.RoadType.Contains(disItem.RoadType)
                     && t.RoadGrad.Contains(disItem.RoadGrad)
                     ).ToList();
                if (findDis.Count == 0)
                {
                    continue;
                }
                else if (findDis.Count == 1)
                {
                    disOnly = findDis.First();
                }
                else
                {
                    if (string.IsNullOrEmpty(disItem.DamagedCondition) || disItem.DamagedCondition == "无")
                    {
                        disOnly = findDis.Where(t => t.DiseaseName.Split('.').Length == 1).FirstOrDefault();
                    }
                    else
                    {
                        disOnly = findDis.Where(t => t.DiseaseName.Contains(disItem.DamagedCondition)).FirstOrDefault();
                    }
                }
                string disKey = "";

                if (disOnly == null)
                {
                    continue;
                }
                if (disOnly != null)
                {

                    disKey = disOnly.RoadType + "_" + disOnly.Number.ToString();
                }

                if (allDisFormXml.Keys.Contains(disKey))
                {
                    allDisFormXml[disKey] += disItem.Area;
                }
                else
                {
                    allDisFormXml.Add(disKey, disItem.Area);
                }
            }
        }


        void WorkWithComObjects(object obj)
        {
            MSWord.Document wordDoc = (MSWord.Document)obj;
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
#if PIE
                double picHeight = 150;//4.81
                double picWidth = 215;//7.55
                                      // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
                                      //各路线详细情况
                if (book.Name.Contains("Part9"))
                {

                    object oStyleName = "标题 3";

                    book.Select();
                    currentSelection = wordApp.Selection;
                    for (int t = _Setting.heFeiContineIndex; t < allDatas.Count; t++)
                    {

                        HeFeiRaod road = allDatas[t];



                        string title = road.RoadName + "(" + road.RoadCode + ")\n";
                        string evaluateStr = string.Format(
                            "        {0}路线长度为：{1}km，路面技术状况指数PQI为{2}，" +
                            "按现行规范评价为{3}等水平；路面破损状况PCI为{4}，" +
                            "按现行规范评价为{5}等水平；路面行驶质量指数RQI为{6}，" +
                            "按现行规范评价为{7}等水平。", road.RoadName, road.RoadLen, road.PqiValue.ToString("0.##"), road.PqiGrad,
                            road.PciValue.ToString("0.##"), road.PciGrad, road.RqiValue.ToString("0.##"), road.RqiGrad);


                        // 设置样式为标题2
                        oStyleName = "标题 3";
                        currentSelection.Range.set_Style(ref oStyleName);
                        //currentSelection.TypeParagraph(); 
                        currentSelection.TypeText(title);

                        oStyleName = "正文文本";
                        currentSelection.Range.set_Style(ref oStyleName);
                        currentSelection.TypeText(evaluateStr);
                        currentSelection.TypeParagraph();
                        MSExcel.Range dataRange = null;
                        MSExcel.Application excelApp = null;

                        //插入雷达图

                        while (true)
                        {
                            try
                            {
                                System.Windows.Forms.Clipboard.Clear();
                                string imagePath = getPart9Picture001(road, ref excelApp, picHeight, picWidth);
                                currentSelection.Range.InlineShapes.AddPicture(imagePath);
                                File.Delete(imagePath);

                                //currentSelection.Range.Paste();
                                // Thread.Sleep(GlobalWord.wd_sleep_us);

                                // currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                currentSelection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd); // 将光标折叠到行尾
                                currentSelection.EndKey(MSWord.WdUnits.wdLine); // 将光标移动到行尾  
                                currentSelection.TypeText("    ");

                                break;
                            }
                            catch (Exception ex)
                            {

                                Thread.Sleep(GlobalWord.wd_sleep_us);
                            }
                            finally
                            {
                                CWB_ExcelHelper.disposeExcel(ref excelApp);
                            }

                        }

                        bool hasDis = false;
                        //找到对应病害excel表格
                        List<FileInfo> targetFile = GetTargetDiseaseExcel(road);
                        if (NeedDis)
                        {
                            //插入饼图 
                            //  System.Windows.Forms.Clipboard.Clear(); 
                            while (true)
                            {
                                try
                                {
                                    System.Windows.Forms.Clipboard.Clear();
                                    if (targetFile.Count == 0)
                                    {
                                    }
                                    string imagePath = getDiseasePic(targetFile, ref excelApp, picHeight, picWidth);

                                    if (!string.IsNullOrEmpty(imagePath))
                                    {

                                        currentSelection.Range.InlineShapes.AddPicture(imagePath);
                                        File.Delete(imagePath);
                                        // currentSelection.Range.Paste();
                                        // Thread.Sleep(GlobalWord.wd_sleep_us);

                                        //  currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                        currentSelection.Collapse(MSWord.WdCollapseDirection.wdCollapseEnd); // 将光标折叠到行尾
                                        currentSelection.EndKey(MSWord.WdUnits.wdLine); // 将光标移动到行尾   
                                        currentSelection.TypeParagraph();
                                        hasDis = true;
                                    }
                                    else
                                    {
                                        hasDis = false;
                                    }
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    Thread.Sleep(GlobalWord.wd_sleep_us);
                                }
                                finally
                                {
                                    CWB_ExcelHelper.disposeExcel(ref excelApp);
                                }
                            }
                            if (hasDis)
                            {
                                currentSelection.TypeText("各性能指标统计图        ");
                                // currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oStyleName = "图号";
                                currentSelection.Range.set_Style(ref oStyleName);
                                // currentSelection.TypeParagraph();
                                currentSelection.TypeText("              病害数量统计分布图");
                                //   currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oStyleName = "图号";
                                currentSelection.Range.set_Style(ref oStyleName);
                                currentSelection.TypeParagraph();
                            }
                            else
                            {
                                currentSelection.TypeParagraph();
                                currentSelection.TypeText("各性能指标统计图");
                                currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                oStyleName = "图号";
                                currentSelection.Range.set_Style(ref oStyleName);
                                currentSelection.TypeParagraph();
                            }


                        }

                        currentSelection.TypeText("路面技术状况指数（PQI）分项指标统计表");

                        oStyleName = "表格标题";
                        // 设置样式为标题2
                        currentSelection.Range.set_Style(ref oStyleName);
                        currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        currentSelection.TypeParagraph();
                        int tableCnt = 1;
                        getPart9Table001(ref dataRange, ref excelApp, road);
                        CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                        CWB_ExcelHelper.disposeExcel(ref excelApp);
                        MSWord.Table tbl = currentSelection.Range.Tables[1];
                        //   CWB_WordHelper.DeleteCurrentSelectionLine(wordApp); 
                        AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                        tbl.Range.Font.Name = "宋体"; // 将字体设置为宋体
                        tbl.Range.Font.Size = 12; // 将字号设置为12磅（小四）
                        tbl.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        if (NeedDis)
                        {
                            Dictionary<string, double> allDisFormXml = new Dictionary<string, double>();
                            List<DiseaseHeFei> allDis = null;
                            getDiseaseSumMessage(targetFile, ref allDisFormXml, ref allDis);
                            if (allDisFormXml.Count == 0)
                            {
                                currentSelection.InsertBreak(MSWord.WdBreakType.wdPageBreak);
                                continue;
                            }
                            currentSelection.TypeText("各路面类型病害统计表");
                            oStyleName = "表格标题";
                            // 设置样式为标题2  
                            currentSelection.Range.set_Style(ref oStyleName);
                            currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            currentSelection.TypeParagraph();
                            tableCnt = 1;
                            getDiseaseSumTable001(allDisFormXml, ref dataRange, ref excelApp);
                            CWB_WordHelper.PastExcelTable2Word(wordDoc, dataRange, currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);

                            CWB_ExcelHelper.disposeExcel(ref excelApp);
                            tbl = currentSelection.Range.Tables[1];
                            //   CWB_WordHelper.DeleteCurrentSelectionLine(wordApp); 
                            AdjustTableFormatting(wordApp, "BG_CCC1", tbl, false);
                            tbl.Range.Font.Name = "宋体"; // 将字体设置为宋体
                            tbl.Range.Font.Size = 12; // 将字号设置为12磅（小四）
                                                      // 设置表格内文字居中对齐
                            tbl.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            currentSelection.TypeParagraph();
                        }
                        currentSelection.InsertBreak(MSWord.WdBreakType.wdPageBreak);

                        //object oStyleName = "标题 2";
                        //currentSelection.Range.set_Style(ref oStyleName);
                        //currentSelection.TypeParagraph();
                    }
                    //oStyleName = "标题 2";
                    //// 设置样式为标题2
                    //currentSelection.Range.set_Style(ref oStyleName);
                    //currentSelection.TypeText("景观图片");
                    //currentSelection.TypeParagraph();
                    //// 设置样式为标题2
                    //currentSelection.Range.set_Style(ref oStyleName);
                    //currentSelection.TypeText("路面损坏典型病害图");
                    //currentSelection.TypeParagraph();
                    continue;
                }
#endif
            }

        }


        private string SaveExcelChartAsImage(MSExcel.Chart chart)
        {
            string imagePath = basePath + "\\image.png";
            chart.Export(imagePath, "PNG", false);
            return imagePath;
        }
        public bool PastePic(System.Action method)
        {
            while (true)
            {
                try
                {
                    System.Windows.Forms.Clipboard.Clear();
                    method();
                    currentSelection.Range.Paste();
                    Thread.Sleep(GlobalWord.wd_sleep_us);
                    currentSelection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    return true;
                }
                catch (Exception ex)
                {

                    Thread.Sleep(GlobalWord.wd_sleep_us);
                }
            }
            return true;
        }


        #region 准备插入的图表
        private void getRqiTable001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "路线编码";
            worksheet.Cells[1, 2].Value = "路线名称";
            worksheet.Cells[1, 3].Value = "RQI算术平均值";
            worksheet.Cells[1, 4].Value = "优良路率（%）";
            worksheet.Cells[1, 5].Value = "次差路率（%）";
            worksheet.Cells[1, 6].Value = "评价等级";
            int rowCount = 1;
            foreach (var item in allDatas)
            {
                rowCount++;
                worksheet.Cells[rowCount, 1].Value = item.RoadCode;
                worksheet.Cells[rowCount, 2].Value = item.RoadName;
                worksheet.Cells[rowCount, 3].Value = item.RqiValue.ToString("0.##");
                double subLen = 0;
                //优良路
                double ylRoadLen = 0;
                double ciRoadLen = 0;

                for (int i = 0; i < item.datas.Count; i++)
                {
                    var data = item.datas[i];
                    subLen += data.RoadLen;
                    if (data.RqiGrad == "优" | data.RqiGrad == "良")
                    {
                        ylRoadLen += data.RoadLen;
                    }
                    else if (data.RqiGrad == "次" | data.RqiGrad == "差")
                    {
                        ciRoadLen += data.RoadLen;
                    }
                }
                worksheet.Cells[rowCount, 4].Value = (ylRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 5].Value = (ciRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 6].Value = item.RqiGrad;
            }


            dataRange = worksheet.Range["A1:F" + rowCount.ToString()];



        }
        private void getPqiTable001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "路线编码";
            worksheet.Cells[1, 2].Value = "路线名称";
            worksheet.Cells[1, 3].Value = "PQI算术平均值";
            worksheet.Cells[1, 4].Value = "优良路率（%）";
            worksheet.Cells[1, 5].Value = "次差路率（%）";
            worksheet.Cells[1, 6].Value = "评价等级";
            int rowCount = 1;
            foreach (var item in allDatas)
            {
                rowCount++;
                worksheet.Cells[rowCount, 1].Value = item.RoadCode;
                worksheet.Cells[rowCount, 2].Value = item.RoadName;
                worksheet.Cells[rowCount, 3].Value = item.PqiValue.ToString("0.##");
                double subLen = 0;
                //优良路
                double ylRoadLen = 0;
                double ciRoadLen = 0;

                for (int i = 0; i < item.datas.Count; i++)
                {
                    var data = item.datas[i];
                    subLen += data.RoadLen;
                    if (data.PqiGrad == "优" | data.PqiGrad == "良")
                    {
                        ylRoadLen += data.RoadLen;
                    }

                    else if (data.PqiGrad == "次" | data.PqiGrad == "差")
                    {
                        ciRoadLen += data.RoadLen;
                    }
                }
                worksheet.Cells[rowCount, 4].Value = (ylRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 5].Value = (ciRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 6].Value = item.PqiGrad;
            }


            dataRange = worksheet.Range["A1:F" + rowCount.ToString()];



        }

        private void getPciTable001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "路线编码";
            worksheet.Cells[1, 2].Value = "路线名称";
            worksheet.Cells[1, 3].Value = "PCI算术平均值";
            worksheet.Cells[1, 4].Value = "优良路率（%）";
            worksheet.Cells[1, 5].Value = "次差路率（%）";
            worksheet.Cells[1, 6].Value = "评价等级";
            int rowCount = 1;
            foreach (var item in allDatas)
            {
                rowCount++;
                worksheet.Cells[rowCount, 1].Value = item.RoadCode;
                worksheet.Cells[rowCount, 2].Value = item.RoadName;
                worksheet.Cells[rowCount, 3].Value = item.PciValue.ToString("0.##");
                double subLen = 0;
                //优良路
                double ylRoadLen = 0;
                double ciRoadLen = 0;

                for (int i = 0; i < item.datas.Count; i++)
                {
                    var data = item.datas[i];
                    subLen += data.RoadLen;
                    if (data.PciGrad == "优" | data.PciGrad == "良")
                    {
                        ylRoadLen += data.RoadLen;
                    }
                    else if (data.PciGrad == "次" | data.PciGrad == "差")
                    {
                        ciRoadLen += data.RoadLen;
                    }
                }
                worksheet.Cells[rowCount, 4].Value = (ylRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 5].Value = (ciRoadLen / subLen * 100).ToString("0.##");
                worksheet.Cells[rowCount, 6].Value = item.PciGrad;
            }


            dataRange = worksheet.Range["A1:F" + rowCount.ToString()];



        }
        public static bool IsClipboardContainsImaged()
        {
            const uint CF_BITMAP = 2;
            const uint CF_DIB = 8;
            const uint CF_ENHMETAFILE = 14;

            return IsClipboardFormatAvailable(CF_BITMAP) ||
                   IsClipboardFormatAvailable(CF_DIB) ||
                   IsClipboardFormatAvailable(CF_ENHMETAFILE);
        }
        private void getPart9Table001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp, HeFeiRaod road)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;
            double length = 0;
            double goodPqiLength = 0;
            double liangPqiLength = 0;
            double zhongPqiLength = 0;
            double ciPqiLength = 0;
            double chaPqiLength = 0;


            double goodPciLength = 0;
            double liangPciLength = 0;
            double zhongPciLength = 0;
            double ciPciLength = 0;
            double chaPciLength = 0;

            double goodRqiLength = 0;
            double liangRqiLength = 0;
            double zhongRqiLength = 0;
            double ciRqiLength = 0;
            double chaRqiLength = 0;

            foreach (var item in road.datas)
            {
                length += item.RoadLen;
                string temp = SetGrad(item.PqiValue);
                string temp1 = SetGrad(item.PciValue);
                string temp2 = SetGrad(item.RqiValue);

                switch (temp)
                {
                    case "优":
                        goodPqiLength += item.RoadLen;
                        break;
                    case "良":
                        liangPqiLength += item.RoadLen;
                        break;
                    case "中":
                        zhongPqiLength += item.RoadLen;
                        break;
                    case "次":
                        ciPqiLength += item.RoadLen;
                        break;
                    case "差":
                        chaPqiLength += item.RoadLen;
                        break;
                    default:
                        break;
                }
                switch (temp1)
                {
                    case "优":
                        goodPciLength += item.RoadLen;
                        break;
                    case "良":
                        liangPciLength += item.RoadLen;
                        break;
                    case "中":
                        zhongPciLength += item.RoadLen;
                        break;
                    case "次":
                        ciPciLength += item.RoadLen;
                        break;
                    case "差":
                        chaPciLength += item.RoadLen;
                        break;
                    default:
                        break;
                }
                switch (temp2)
                {
                    case "优":
                        goodRqiLength += item.RoadLen;
                        break;
                    case "良":
                        liangRqiLength += item.RoadLen;
                        break;
                    case "中":
                        zhongRqiLength += item.RoadLen;
                        break;
                    case "次":
                        ciRqiLength += item.RoadLen;
                        break;
                    case "差":
                        chaRqiLength += item.RoadLen;
                        break;
                    default:
                        break;
                }

            }


            worksheet.Cells[1, 1].Value = "指标";
            worksheet.Cells[1, 2].Value = "算数平均值";
            worksheet.Cells[1, 3].Value = "优良路率";
            worksheet.Cells[1, 4].Value = "次差路率";

            worksheet.Cells[2, 1].Value = "PQI";
            worksheet.Cells[2, 2].Value = road.PqiValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = ((goodPqiLength + liangPqiLength) / length * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 4].Value = ((ciPqiLength + chaPqiLength) / length * 100).ToString("0.##") + "%";

            worksheet.Cells[3, 1].Value = "PCI";
            worksheet.Cells[3, 2].Value = road.PciValue.ToString("0.##");
            worksheet.Cells[3, 3].Value = ((goodPciLength + liangPciLength) / length * 100).ToString("0.##") + "%";
            worksheet.Cells[3, 4].Value = ((ciPciLength + chaPciLength) / length * 100).ToString("0.##") + "%";

            worksheet.Cells[4, 1].Value = "RQI";
            worksheet.Cells[4, 2].Value = road.RqiValue.ToString("0.##");
            worksheet.Cells[4, 3].Value = ((goodRqiLength + liangRqiLength) / length * 100).ToString("0.##") + "%";
            worksheet.Cells[4, 4].Value = ((ciRqiLength + chaRqiLength) / length * 100).ToString("0.##") + "%";
            dataRange = worksheet.Range["A1:D4"];

        }

        private void getPart7Table001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {

            double snLength = 0;
            double lqLength = 0;
            double lqPqiValue = 0;
            double snPqiValue = 0;
            double lqPciValue = 0;
            double snPciValue = 0;
            double lqRqiValue = 0;
            double snRqiValue = 0;
            double goodLqLength = 0;
            double goodSnLength = 0;
            double cichaLqLength = 0;
            double cichaSnLength = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {

                    if (item1.RoadType == 0)
                    {
                        lqLength += item1.RoadLen;
                        lqPciValue += item1.PciValue * item1.RoadLen;
                        lqPqiValue += item1.PqiValue * item1.RoadLen;
                        lqRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodLqLength += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaLqLength += item1.RoadLen;
                        }
                    }
                    if (item1.RoadType == 1)
                    {
                        snLength += item1.RoadLen;
                        snPciValue += item1.PciValue * item1.RoadLen;
                        snPqiValue += item1.PqiValue * item1.RoadLen;
                        snRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodSnLength += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaSnLength += item1.RoadLen;
                        }
                    }
                }
            }
            lqPqiValue /= lqLength;
            snPqiValue /= snLength;
            lqPciValue /= lqLength;
            snPciValue /= snLength;
            lqRqiValue /= lqLength;
            snRqiValue /= snLength;
            string goodLqRate = (goodLqLength / lqLength * 100).ToString("0.##");
            string goodSnRate = (goodSnLength / snLength * 100).ToString("0.##");
            string cichaLqRate = (cichaLqLength / lqLength * 100).ToString("0.##");
            string cichaSnRate = (cichaSnLength / snLength * 100).ToString("0.##");
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路等级";
            worksheet.Cells[1, 2].Value = "PQI";
            worksheet.Cells[1, 3].Value = "PCI";
            worksheet.Cells[1, 4].Value = "RQI";
            worksheet.Cells[1, 5].Value = "优良路率（%）";
            worksheet.Cells[1, 6].Value = "次差路率（%）";
            worksheet.Cells[1, 7].Value = "评价等级";

            worksheet.Cells[2, 1].Value = "水泥混凝土路面";
            if (snLength == 0)
            {
                worksheet.Cells[2, 2].Value = "-";
                worksheet.Cells[2, 3].Value = "-";
                worksheet.Cells[2, 4].Value = "-";
                worksheet.Cells[2, 5].Value = "-";
                worksheet.Cells[2, 6].Value = "-";
                worksheet.Cells[2, 7].Value = "-";


            }
            else
            {
                worksheet.Cells[2, 2].Value = snPqiValue.ToString("0.##");
                worksheet.Cells[2, 3].Value = snPciValue.ToString("0.##");
                worksheet.Cells[2, 4].Value = snRqiValue.ToString("0.##");
                worksheet.Cells[2, 5].Value = goodSnRate;
                worksheet.Cells[2, 6].Value = cichaSnRate;
                worksheet.Cells[2, 7].Value = SetGrad(snPqiValue);

            }

            worksheet.Cells[3, 1].Value = "沥青混凝土路面";

            if (lqLength == 0)
            {
                worksheet.Cells[3, 2].Value = "-";
                worksheet.Cells[3, 3].Value = "-";
                worksheet.Cells[3, 4].Value = "-";
                worksheet.Cells[3, 5].Value = "-";
                worksheet.Cells[3, 6].Value = "-";
                worksheet.Cells[3, 7].Value = "-";
            }
            else
            {
                worksheet.Cells[3, 2].Value = lqPqiValue.ToString("0.##");
                worksheet.Cells[3, 3].Value = lqPciValue.ToString("0.##");
                worksheet.Cells[3, 4].Value = lqRqiValue.ToString("0.##");
                worksheet.Cells[3, 5].Value = goodLqRate;
                worksheet.Cells[3, 6].Value = cichaLqRate;
                worksheet.Cells[3, 7].Value = SetGrad(lqPqiValue);
            }

            dataRange = worksheet.Range["A1:G3"];

        }



        private void getDiseaseSumTable001(Dictionary<string, double> allDisFormXml, ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;
            List<string> snDis = new List<string>();
            List<string> lqDis = new List<string>();
            foreach (var item in allDisFormXml)
            {
                string number = item.Key.Split('_').Last();
                string area = item.Value.ToString();
                string sumDis = number + "_" + area;
                if (item.Key.Contains("水泥"))
                {
                    snDis.Add(sumDis);
                }
                else
                {
                    lqDis.Add(sumDis);
                }
            }

            //根据number排个序 
            snDis.Sort((t1, t2) =>
            {
                double number = double.Parse(t1.Split('_').First().PadRight(3, '0'));
                double number2 = double.Parse(t2.Split('_').First().PadRight(3, '0'));
                return number > number2 ? 1 : (number < number2 ? -1 : 0);
            }
            );
            lqDis.Sort((t1, t2) =>
            {
                double number = double.Parse(t1.Split('_').First().PadRight(3, '0'));
                double number2 = double.Parse(t2.Split('_').First().PadRight(3, '0'));
                return number > number2 ? 1 : (number < number2 ? -1 : 0);
            }
           );
            int colCount = Math.Max(snDis.Count, lqDis.Count);

            worksheet.Cells[1, 1].Value = "路面病害类型";
            worksheet.Cells[2, 1].Value = "水泥混凝土路面病害总面积";
            worksheet.Cells[3, 1].Value = "路面病害类型";
            worksheet.Cells[4, 1].Value = "沥青混凝土路面病害总面积";

            for (int i = 0; i < snDis.Count; i++)
            {
                string[] str = snDis[i].Split('_');
                worksheet.Cells[1, i + 2].Value = str.First();
                worksheet.Cells[2, i + 2].Value = double.Parse(str.Last()).ToString("0.##");
            }

            for (int i = 0; i < lqDis.Count; i++)
            {
                string[] str = lqDis[i].Split('_');
                worksheet.Cells[3, i + 2].Value = str.First();
                worksheet.Cells[4, i + 2].Value = double.Parse(str.Last()).ToString("0.##");
            }
            string colStr = GlobalExcel.GetCol((char)('A' + colCount));
            if (lqDis.Count == 0 && snDis.Count != 0)
            {
                dataRange = worksheet.Range[string.Format("A1:{0}2", colStr)];

            }
            if (snDis.Count == 0 && lqDis.Count != 0)
            {
                dataRange = worksheet.Range[string.Format("A3:{0}4", colStr)];
            }

            if (lqDis.Count != 0 && snDis.Count != 0)
            {
                dataRange = worksheet.Range[string.Format("A1:{0}4", colStr)];

            }

        }


        /// <summary>
        /// 分路段进行病害统计
        /// </summary>
        /// <param name="allDisFormXml"></param>
        /// <param name="dataRange"></param>
        /// <param name="excelApp"></param>
        private void getDiseaseSumTable002(string savePath, Dictionary<string, double> allDisFormXml, List<DiseaseHeFei> allDis, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;
            Dictionary<string, List<(string, double)>> result = handelAllDiseaseToRoadPart(allDis);
            //获得所有的病害编号
            List<string> numbers = new List<string>();
            foreach (var item in result)
            {
                foreach (var item1 in item.Value)
                {
                    if (!numbers.Contains(item1.Item1))
                    {
                        numbers.Add(item1.Item1);
                    }
                }
            }
            numbers.Sort();
            worksheet.Cells[1, 1].Value = "项目名称及编号";
            for (int i = 0; i < numbers.Count; i++)
            {
                worksheet.Cells[1, 2 + i].Value = numbers[i];
            }

            for (int i = 0; i < result.Count; i++)
            {
                var ele = result.ElementAt(i);
                worksheet.Cells[2 + i, 1] = ele.Key;

                for (int j = 0; j < numbers.Count; j++)
                {
                    worksheet.Cells[2 + i, 2 + j].Value = 0; //先把所有单元填充为0
                }
                for (int t = 0; t < ele.Value.Count; t++)
                {
                    //找到指定位置填写入病害面积
                    int findIndex = numbers.FindIndex(d => d == ele.Value[t].Item1);
                    worksheet.Cells[2 + i, 2 + findIndex].Value = ele.Value[t].Item2; //先把所有单元填充为0
                }

            }
            worksheet.SaveAs(savePath);
        }

        private void getPart8Table001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {



            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "路线";
            worksheet.Cells[1, 2].Value = "检评总里程(km)";
            worksheet.Cells[1, 3].Value = "PQI次差里程(km)";
            worksheet.Cells[1, 4].Value = "PCI次差里程(km)";
            worksheet.Cells[1, 5].Value = "RQI次差里程(km)";
            worksheet.Cells[1, 6].Value = "PCI&RQI双指标次差里程(km)";
            int index = 1;
            foreach (var item in allDatas)
            {
                index++;
                double length = 0;
                double pqiCiChaLen = 0;
                double pciCiChaLen = 0;
                double rqiCiChaLen = 0;
                double rqiAndPciCiChaLen = 0;
                foreach (var item1 in item.datas)
                {
                    length += item1.RoadLen;
                    string temp = SetGrad(item1.PciValue);
                    if (temp == "次" || temp == "差")
                    {
                        pciCiChaLen += item1.RoadLen;
                    }
                    string temp2 = SetGrad(item1.PqiValue);
                    if (temp2 == "次" || temp2 == "差")
                    {
                        pqiCiChaLen += item1.RoadLen;
                    }
                    string temp1 = SetGrad(item1.RqiValue);
                    if (temp1 == "次" || temp1 == "差")
                    {
                        rqiCiChaLen += item1.RoadLen;
                    }

                    if (temp == "次" || temp == "差")
                    {
                        if (temp1 == "次" || temp1 == "差")
                            rqiAndPciCiChaLen += item1.RoadLen;
                    }
                }
                worksheet.Cells[index, 1].Value = item.RoadCode;
                worksheet.Cells[index, 2].Value = length.ToString();
                worksheet.Cells[index, 3].Value = pqiCiChaLen.ToString();
                worksheet.Cells[index, 4].Value = pciCiChaLen.ToString();
                worksheet.Cells[index, 5].Value = rqiCiChaLen.ToString();
                worksheet.Cells[index, 6].Value = rqiAndPciCiChaLen.ToString();
            }
            dataRange = worksheet.Range["A1:F" + index.ToString()];
        }

        private void getPart6Table001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;

            double grad0PqiValue = 0;
            double grad1PqiValue = 0;
            double grad0PciValue = 0;
            double grad1PciValue = 0;
            double grad0RqiValue = 0;
            double grad1RqiValue = 0;


            double grad2PqiValue = 0;
            double grad3PqiValue = 0;
            double grad2PciValue = 0;
            double grad3PciValue = 0;
            double grad2RqiValue = 0;
            double grad3RqiValue = 0;

            double goodGrad0Length = 0;
            double goodGrad1Length = 0;
            double cichaGrad0Length = 0;
            double cichaGrad1Length = 0;


            double goodGrad2Length = 0;
            double goodGrad3Length = 0;
            double cichaGrad2Length = 0;
            double cichaGrad3Length = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    if (item1.RoadGrad == 1)
                    {
                        grad0PqiValue += item1.PqiValue * item1.RoadLen;
                        grad0PciValue += item1.PciValue * item1.RoadLen;
                        grad0RqiValue += item1.RqiValue * item1.RoadLen;
                        length0 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad0Length += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaGrad0Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 2)
                    {
                        grad1PqiValue += item1.PqiValue * item1.RoadLen;
                        grad1PciValue += item1.PciValue * item1.RoadLen;
                        grad1RqiValue += item1.RqiValue * item1.RoadLen;
                        length1 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad1Length += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaGrad1Length += item1.RoadLen;
                        }

                    }


                    if (item1.RoadGrad == 3)
                    {
                        grad2PqiValue += item1.PqiValue * item1.RoadLen;
                        grad2PciValue += item1.PciValue * item1.RoadLen;
                        grad2RqiValue += item1.RqiValue * item1.RoadLen;
                        length2 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad2Length += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaGrad2Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 4)
                    {
                        grad3PqiValue += item1.PqiValue * item1.RoadLen;
                        grad3PciValue += item1.PciValue * item1.RoadLen;
                        grad3RqiValue += item1.RqiValue * item1.RoadLen;
                        length3 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad3Length += item1.RoadLen;
                        }
                        else if (temp == "次" || temp == "差")
                        {
                            cichaGrad3Length += item1.RoadLen;
                        }
                    }
                }

            }
            grad0PqiValue = double.IsNaN(grad0PqiValue / length0) ? 0 : grad0PqiValue / length0;
            grad0PciValue = double.IsNaN(grad0PciValue / length0) ? 0 : grad0PciValue / length0;
            grad0RqiValue = double.IsNaN(grad0RqiValue / length0) ? 0 : grad0RqiValue / length0;

            grad1PqiValue = double.IsNaN(grad1PqiValue / length1) ? 0 : grad1PqiValue / length1;
            grad1PciValue = double.IsNaN(grad1PciValue / length1) ? 0 : grad1PciValue / length1;
            grad1RqiValue = double.IsNaN(grad1RqiValue / length1) ? 0 : grad1RqiValue / length1;


            grad2PqiValue = double.IsNaN(grad2PqiValue / length2) ? 0 : grad2PqiValue / length2;
            grad2PciValue = double.IsNaN(grad2PciValue / length2) ? 0 : grad2PciValue / length2;
            grad2RqiValue = double.IsNaN(grad2RqiValue / length2) ? 0 : grad2RqiValue / length2;

            grad3PqiValue = double.IsNaN(grad3PqiValue / length3) ? 0 : grad3PqiValue / length3;
            grad3PciValue = double.IsNaN(grad3PciValue / length3) ? 0 : grad3PciValue / length3;
            grad3RqiValue = double.IsNaN(grad3RqiValue / length3) ? 0 : grad3RqiValue / length3;




            string goodGrad0Rate = double.IsNaN(goodGrad0Length / length0) ? "0" : (goodGrad0Length / length0 * 100).ToString("0.##");
            string goodGrad1Rate = double.IsNaN(goodGrad1Length / length1) ? "0" : (goodGrad1Length / length1 * 100).ToString("0.##");

            string goodGrad3Rate = double.IsNaN(goodGrad3Length / length3) ? "0" : (goodGrad3Length / length3 * 100).ToString("0.##");
            string goodGrad2Rate = double.IsNaN(goodGrad2Length / length2) ? "0" : (goodGrad2Length / length2 * 100).ToString("0.##");

            string cichaGrad0Rate = double.IsNaN(cichaGrad0Length / length0) ? "0" : (cichaGrad0Length / length0 * 100).ToString("0.##");
            string cichaGrad1Rate = double.IsNaN(cichaGrad1Length / length1) ? "0" : (cichaGrad1Length / length1 * 100).ToString("0.##");

            string cichaGrad2Rate = double.IsNaN(cichaGrad2Length / length2) ? "0" : (cichaGrad2Length / length2 * 100).ToString("0.##");
            string cichaGrad3Rate = double.IsNaN(cichaGrad3Length / length3) ? "0" : (cichaGrad3Length / length3 * 100).ToString("0.##");

            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路等级";
            worksheet.Cells[1, 2].Value = "PQI";
            worksheet.Cells[1, 3].Value = "PCI";
            worksheet.Cells[1, 4].Value = "RQI";
            worksheet.Cells[1, 5].Value = "优良路率（%）";
            worksheet.Cells[1, 6].Value = "次差路率（%）";
            worksheet.Cells[1, 7].Value = "评价等级";

            if (length0 == 0)
            {
                worksheet.Cells[2, 1].Value = "一级公路";
                worksheet.Cells[2, 2].Value = "-";
                worksheet.Cells[2, 3].Value = "-";
                worksheet.Cells[2, 4].Value = "-";
                worksheet.Cells[2, 5].Value = "-";
                worksheet.Cells[2, 6].Value = "-";
                worksheet.Cells[2, 7].Value = "-";
            }
            else
            {
                worksheet.Cells[2, 1].Value = "一级公路";
                worksheet.Cells[2, 2].Value = grad0PqiValue.ToString("0.##");
                worksheet.Cells[2, 3].Value = grad0PciValue.ToString("0.##");
                worksheet.Cells[2, 4].Value = grad0RqiValue.ToString("0.##");
                worksheet.Cells[2, 5].Value = goodGrad0Rate;
                worksheet.Cells[2, 6].Value = cichaGrad0Rate;
                worksheet.Cells[2, 7].Value = SetGrad(grad0PqiValue);
            }
            if (length1 == 0)
            {
                worksheet.Cells[3, 1].Value = "二级公路";
                worksheet.Cells[3, 2].Value = "-";
                worksheet.Cells[3, 3].Value = "-";
                worksheet.Cells[3, 4].Value = "-";
                worksheet.Cells[3, 5].Value = "-";
                worksheet.Cells[3, 6].Value = "-";
                worksheet.Cells[3, 7].Value = "-";
            }
            else
            {
                worksheet.Cells[3, 1].Value = "二级公路";
                worksheet.Cells[3, 2].Value = grad1PqiValue.ToString("0.##");
                worksheet.Cells[3, 3].Value = grad1PciValue.ToString("0.##");
                worksheet.Cells[3, 4].Value = grad1RqiValue.ToString("0.##");
                worksheet.Cells[3, 5].Value = goodGrad1Rate;
                worksheet.Cells[3, 6].Value = cichaGrad1Rate;
                worksheet.Cells[3, 7].Value = SetGrad(grad1PqiValue);
            }


            if (length2 == 0)
            {
                worksheet.Cells[4, 1].Value = "三级公路";
                worksheet.Cells[4, 2].Value = "-";
                worksheet.Cells[4, 3].Value = "-";
                worksheet.Cells[4, 4].Value = "-";
                worksheet.Cells[4, 5].Value = "-";
                worksheet.Cells[4, 6].Value = "-";
                worksheet.Cells[4, 7].Value = "-";
            }
            else
            {
                worksheet.Cells[4, 1].Value = "三级公路";
                worksheet.Cells[4, 2].Value = grad2PqiValue.ToString("0.##");
                worksheet.Cells[4, 3].Value = grad2PciValue.ToString("0.##");
                worksheet.Cells[4, 4].Value = grad2RqiValue.ToString("0.##");
                worksheet.Cells[4, 5].Value = goodGrad2Rate;
                worksheet.Cells[4, 6].Value = cichaGrad2Rate;
                worksheet.Cells[4, 7].Value = SetGrad(grad2PqiValue);
            }
            if (length3 == 0)
            {
                worksheet.Cells[5, 1].Value = "四级公路";
                worksheet.Cells[5, 2].Value = "-";
                worksheet.Cells[5, 3].Value = "-";
                worksheet.Cells[5, 4].Value = "-";
                worksheet.Cells[5, 5].Value = "-";
                worksheet.Cells[5, 6].Value = "-";
                worksheet.Cells[5, 7].Value = "-";
            }
            else
            {
                worksheet.Cells[5, 1].Value = "四级公路";
                worksheet.Cells[5, 2].Value = grad3PqiValue.ToString("0.##");
                worksheet.Cells[5, 3].Value = grad3PciValue.ToString("0.##");
                worksheet.Cells[5, 4].Value = grad3RqiValue.ToString("0.##");
                worksheet.Cells[5, 5].Value = goodGrad3Rate;
                worksheet.Cells[5, 6].Value = cichaGrad3Rate;
                worksheet.Cells[5, 7].Value = SetGrad(grad3PqiValue);
            }






            dataRange = worksheet.Range["A1:G5"];



        }
        private List<DiseaseFormXml> AllDiseaseFormXml = new List<DiseaseFormXml>();
        private void ReadDiseaseXml()
        {
            AllDiseaseFormXml.Clear();
            string xmlPath = "\\ConfigFile\\DiseaseToNum.xml";
            XmlDocument document = new XmlDocument();
            XmlElement elem;
            XmlNodeList xmlNodeList;
            document.Load(AppDomain.CurrentDomain.BaseDirectory + xmlPath);
            elem = document.DocumentElement;
            xmlNodeList = elem.ChildNodes;
            foreach (XmlNode xmlRoadGradNode in xmlNodeList)
            {
                var roadTypes = xmlRoadGradNode.ChildNodes;
                string raodGradName = xmlRoadGradNode.Name;
                foreach (XmlNode roadType in roadTypes)
                {
                    //各个病害
                    var roadDisease = roadType.ChildNodes;
                    string roadTypeName = roadType.Name;
                    foreach (XmlNode disease in roadDisease)
                    {

                        string fullName = raodGradName + "." + roadTypeName + "." + disease.Name;
                        //各个病害
                        string weight = disease.Attributes["权重"].Value;
                        string number = disease.Attributes["编号"].Value;
                        DiseaseFormXml temp = new DiseaseFormXml(raodGradName, roadTypeName, disease.Name
                          , weight, number);
                        AllDiseaseFormXml.Add(temp);
                    }
                }
            }

        }
        private void getPart10Table002(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {




            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "路线编码";
            worksheet.Cells[1, 2].Value = "检评长度（km）";
            worksheet.Cells[1, 3].Value = "技术等级";
            worksheet.Cells[1, 4].Value = "";
            worksheet.Cells[1, 5].Value = "";
            worksheet.Cells[1, 6].Value = "";
            worksheet.Cells[1, 7].Value = "路面类型";
            worksheet.Cells[1, 8].Value = "";

            worksheet.Cells[2, 1].Value = "";
            worksheet.Cells[2, 2].Value = "";
            worksheet.Cells[2, 3].Value = "一级公路(km)";
            worksheet.Cells[2, 4].Value = "二级公路(km)";
            worksheet.Cells[2, 5].Value = "三级公路(km)";
            worksheet.Cells[2, 6].Value = "四级公路(km)";
            worksheet.Cells[2, 7].Value = "沥青混凝土路面（km）";
            worksheet.Cells[2, 8].Value = "水泥混凝土路面（km）";


            MSExcel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]];
            range.Merge(); // 合并单元格

            range = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]];
            range.Merge(); // 合并单元格
            range = worksheet.Range[worksheet.Cells[1, 3], worksheet.Cells[1, 6]];
            range.Merge(); // 合并单元格
            range = worksheet.Range[worksheet.Cells[1, 7], worksheet.Cells[1, 8]];
            range.Merge(); // 合并单元格

            int index = 2;


            foreach (var item in allDatas)
            {
                index++;
                double grad1Length = 0;
                double grad2Length = 0;
                double grad3Length = 0;
                double grad4Length = 0;
                double lqLength = 0;
                double snLength = 0;
                foreach (var road in item.datas)
                {
                    switch (road.RoadGrad)
                    {
                        case 0:
                            break;

                        case 1:
                            grad1Length += road.RoadLen;
                            break;
                        case 2:
                            grad2Length += road.RoadLen;
                            break;
                        case 3:
                            grad3Length += road.RoadLen;
                            break;
                        case 4:
                            grad4Length += road.RoadLen;
                            break;
                        default:
                            break;
                    }
                    switch (road.RoadType)
                    {
                        case 0:
                            lqLength += road.RoadLen;
                            break;
                        case 1:
                            snLength += road.RoadLen;
                            break;
                        default:
                            break;
                    }
                }
                worksheet.Cells[index, 1].Value = item.RoadCode;
                worksheet.Cells[index, 2].Value = item.RoadLen;
                worksheet.Cells[index, 3].Value = grad1Length == 0 ? "-" : grad1Length.ToString("0.###");
                worksheet.Cells[index, 4].Value = grad2Length == 0 ? "-" : grad2Length.ToString("0.###");
                worksheet.Cells[index, 5].Value = grad3Length == 0 ? "-" : grad3Length.ToString("0.###");
                worksheet.Cells[index, 6].Value = grad4Length == 0 ? "-" : grad4Length.ToString("0.###");
                worksheet.Cells[index, 7].Value = lqLength == 0 ? "-" : lqLength.ToString("0.###");
                worksheet.Cells[index, 8].Value = snLength == 0 ? "-" : snLength.ToString("0.###");
            }
            dataRange = worksheet.Range["A1:H" + index.ToString()];



        }
        private void getPavementDistributionRatioTable001(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            double xiandaoLength = 0;
            double xiangdaoLength = 0;
            double cundaoLength = 0;
            double xiandaoPqi = 0;
            double xiandaoPci = 0;
            double xiandaoRqi = 0;
            double goodXianRoadLength = 0;
            double goodXiangRoadLength = 0;
            double goodCunRoadLength = 0;
            //次差路长度
            double ciChaXianRoadLength = 0;
            double ciChaXiangRoadLength = 0;
            double ciChaCunRoadLength = 0;
            double xiangdaoPqi = 0;
            double xiangdaoPci = 0;
            double xiangdaoRqi = 0;

            double cundaoPqi = 0;
            double cundaoPci = 0;
            double cundaoRqi = 0;
          //  List<string> testValues = new List<string>();
            foreach (var item in allDatas)
            {
                //string line = $"{item.RoadCode},{item.RoadLen},{item.PqiValue},{item.PciValue},{item.RqiValue}\n";
                //for (int i = 0; i < item.datas.Count; i++)
                //{
                //    var project = item.datas[i];
                //    string line1 = $"\t{project.RoadCode},{project.RoadLen},{project.PqiValue},{project.PciValue},{project.RqiValue}\n";
                //    line += line1;
                //}
               
                if (item.RoadCode.StartsWith("X"))
                {
                  //  testValues.Add(line);
                    xiandaoLength += item.RoadLen;
                    xiandaoPqi += item.PqiValue * item.RoadLen;
                    xiandaoPci += item.PciValue * item.RoadLen;
                    xiandaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXianRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXianRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("Y"))
                {
                    xiangdaoLength += item.RoadLen;
                    xiangdaoPci += item.PciValue * item.RoadLen;
                    xiangdaoPqi += item.PqiValue * item.RoadLen;
                    xiangdaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXiangRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXiangRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("C"))
                {
                    cundaoLength += item.RoadLen;
                    cundaoPci += item.PciValue * item.RoadLen;
                    cundaoPqi += item.PqiValue * item.RoadLen;
                    cundaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodCunRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaCunRoadLength += item.RoadLen;
                    }
                }


            }

            //File.WriteAllLines("C:\\Users\\cwb\\Desktop\\安徽报告\\出报告素材\\细则.txt", testValues);
            xiandaoPci /= xiandaoLength;
            xiandaoPqi /= xiandaoLength;
            xiandaoRqi /= xiandaoLength;
            string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "-" : xiandaoPci.ToString("0.##");
            string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "-" : xiandaoPqi.ToString("0.##");
            string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "-" : xiandaoRqi.ToString("0.##");

            xiangdaoPci /= xiangdaoLength;
            xiangdaoPqi /= xiangdaoLength;
            xiangdaoRqi /= xiangdaoLength;

            string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "-" : xiangdaoPci.ToString("0.##");
            string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "-" : xiangdaoPqi.ToString("0.##");
            string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "-" : xiangdaoRqi.ToString("0.##");

            cundaoPci /= cundaoLength;
            cundaoPqi /= cundaoLength;
            cundaoRqi /= cundaoLength;
            string cundaoPciStr = double.IsNaN(cundaoPci) ? "-" : cundaoPci.ToString("0.##");
            string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "-" : cundaoPqi.ToString("0.##");
            string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "-" : cundaoRqi.ToString("0.##");
            string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "-" : SetGrad(xiandaoPqi);
            string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "-" : SetGrad(xiangdaoPqi);
            string cundaojudge = double.IsNaN(cundaoPqi) ? "-" : SetGrad(cundaoPqi);

            string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "-" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "-" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "-" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
            string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "-" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "-" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "-" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";

            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路等级";
            worksheet.Cells[1, 2].Value = "PQI";
            worksheet.Cells[1, 3].Value = "PCI";
            worksheet.Cells[1, 4].Value = "RQI";
            worksheet.Cells[1, 5].Value = "优良路率（%）";
            worksheet.Cells[1, 6].Value = "次差路率（%）";
            worksheet.Cells[1, 7].Value = "评价等级";

            worksheet.Cells[2, 1].Value = "县道";
            worksheet.Cells[2, 2].Value = xiandaoPqiStr;
            worksheet.Cells[2, 3].Value = xiandaoPciStr;
            worksheet.Cells[2, 4].Value = xiandaoRqiStr;
            worksheet.Cells[2, 5].Value = goodXianRoadRate;
            worksheet.Cells[2, 6].Value = cichaXianRoadRate;
            worksheet.Cells[2, 7].Value = xiandaoJudge;

            worksheet.Cells[3, 1].Value = "乡道";
            worksheet.Cells[3, 2].Value = xiangdaoPqiStr;
            worksheet.Cells[3, 3].Value = xiangdaoPciStr;
            worksheet.Cells[3, 4].Value = xiangdaoRqiStr;
            worksheet.Cells[3, 5].Value = goodXiangRoadRate;
            worksheet.Cells[3, 6].Value = cichaXiangRoadRate;
            worksheet.Cells[3, 7].Value = xiangdaojudge;

            worksheet.Cells[4, 1].Value = "村道";
            worksheet.Cells[4, 2].Value = cundaoPqiStr;
            worksheet.Cells[4, 3].Value = cundaoPciStr;
            worksheet.Cells[4, 4].Value = cundaoRqiStr;
            worksheet.Cells[4, 5].Value = goodCunRoadRate;
            worksheet.Cells[4, 6].Value = cichaCunRoadRate;
            worksheet.Cells[4, 7].Value = cundaojudge;

            dataRange = worksheet.Range["A1:G4"];



        }
        private void getRQIPicture001()
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "优";
            worksheet.Cells[1, 2].Value = "良";
            worksheet.Cells[1, 3].Value = "中";
            worksheet.Cells[1, 4].Value = "次";
            worksheet.Cells[1, 5].Value = "差";

            worksheet.Cells[2, 1].Value = (yRQIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 2].Value = (lRQIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 3].Value = (zRQIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 4].Value = (ciRQIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 5].Value = (chaRQIRoadRate * 100).ToString("0.##") + "%";

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:E2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = false;

            // 设置数据标签可见
            series1.HasDataLabels = true;

            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getRQIPicture002()
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路名称";
            worksheet.Cells[1, 2].Value = "RQI";
            int rowCount = 1;
            var datas = allDatas.OrderByDescending(t => t.RqiValue).ToList();

            if (datas.Count <= 20)
            {
                for (int i = 0; i < datas.Count; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].RqiValue;

                }
            }
            else
            {
                for (int i = 0; i < 20; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].RqiValue;


                }
            }



            // 根据数据创建折线图
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:B" + rowCount.ToString()];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = false;
            chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPart6Pic001()
        {

            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    if (item1.RoadGrad == 1)
                    {
                        length0 += item1.RoadLen;
                    }
                    if (item1.RoadGrad == 2)
                    {
                        length1 += item1.RoadLen;
                    }
                    if (item1.RoadGrad == 3)
                    {
                        length2 += item1.RoadLen;
                    }
                    if (item1.RoadGrad == 4)
                    {
                        length3 += item1.RoadLen;
                    }
                }



            }
            double allLenth = length0 + length1 + length2 + length3;
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "一级公路";
            worksheet.Cells[1, 2].Value = "二级公路";
            worksheet.Cells[1, 3].Value = "三级公路";
            worksheet.Cells[1, 4].Value = "四级公路";

            worksheet.Cells[2, 1].Value = (length0 / allLenth * 100) + "%";
            worksheet.Cells[2, 2].Value = (length1 / allLenth * 100) + "%";
            worksheet.Cells[2, 3].Value = (length2 / allLenth * 100) + "%";
            worksheet.Cells[2, 4].Value = (length3 / allLenth * 100) + "%";

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;

            MSExcel.Range dataRange = worksheet.Range["A1:D2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xl3DPie;

            chart.ChartColor = 13;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelBestFit);
            //series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = true;

            // 设置数据标签可见
            //series1.HasDataLabels = false;

            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPart7Pic001()
        {

            double snLength = 0;
            double lqLength = 0;
            double allLength = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    allLength += item1.RoadLen;
                    if (item1.RoadType == 0)
                    {
                        lqLength += item1.RoadLen;
                    }
                    if (item1.RoadType == 1)
                    {
                        snLength += item1.RoadLen;
                    }
                }
            }
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "水泥混凝土路面";
            worksheet.Cells[1, 2].Value = "沥青混凝土路面";

            worksheet.Cells[2, 1].Value = (snLength / allLength * 100) + "%";
            worksheet.Cells[2, 2].Value = (lqLength / allLength * 100) + "%";

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;

            MSExcel.Range dataRange = worksheet.Range["A1:B2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xl3DPie;

            chart.ChartColor = 13;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelBestFit);
            //series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // 移除右边的标注
            chart.HasLegend = true;

            // 设置数据标签可见
            //series1.HasDataLabels = false;

            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPart5Pic001()
        {

            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            foreach (var item in allDatas)
            {
                if (item.RoadCode.StartsWith("X"))
                {
                    length0 += item.RoadLen;
                }
                if (item.RoadCode.StartsWith("Y"))
                {
                    length1 += item.RoadLen;
                }
                if (item.RoadCode.StartsWith("C"))
                {
                    length2 += item.RoadLen;
                }
            }
            double rate0 = length0 / RoadLength;
            double rate1 = length1 / RoadLength;
            double rate2 = length2 / RoadLength;


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "县道";
            worksheet.Cells[1, 2].Value = "乡道";
            worksheet.Cells[1, 3].Value = "村道";

            worksheet.Cells[2, 1].Value = (rate0 * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 2].Value = (rate1 * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 3].Value = (rate2 * 100).ToString("0.##") + "%";

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;

            MSExcel.Range dataRange = worksheet.Range["A1:C2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xl3DPie;

            chart.ChartColor = 13;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelBestFit);
            //series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = true;

            // 设置数据标签可见
            //series1.HasDataLabels = false;

            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private string getDiseasePic(List<FileInfo> fileInfo, ref MSExcel.Application excelApp, double height, double width)
        {
            //读取excel信息
            if (fileInfo == null)
            {
                return "";
            }
            Dictionary<string, double> allDisFormXml = new Dictionary<string, double>();
            List<DiseaseHeFei> allDis = new List<DiseaseHeFei>();
            foreach (var item in fileInfo)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dtProject = new System.Data.DataTable();
                ReadExcelData(ref dt, item.FullName, "病害列表", 2, true);
                ReadExcelData(ref dtProject, item.FullName, "工程信息", 2, true);
                int colCount = dt.Columns.Count;
                string roadGrad = dtProject.Rows[6][1].ToString();


                foreach (DataRow row in dt.Rows)
                {
                    DiseaseHeFei dis = new DiseaseHeFei();
                    dis.RoadGrad = roadGrad;
                    string disName = row[2].ToString();
                    string disGrad = row[3].ToString();
                    string roadType = "";
                    dis.DamagedCondition = disGrad;
                    dis.Name = disName;
                    if (string.IsNullOrWhiteSpace(disGrad) || disGrad == "无")
                    {
                        disGrad = "";
                    }
                    double area = 0;
                    try
                    {
                        if (colCount > 10) //大框
                        {
                            string temp = row[7].ToString();
                            string temp1 = row[12].ToString();
                            if (!string.IsNullOrWhiteSpace(temp1))
                            {
                                roadType = temp1;
                            }
                            else
                            {
                                continue;
                            }
                            if (!string.IsNullOrWhiteSpace(temp))
                            {
                                area = double.Parse(temp);
                            }
                            else
                            {
                                continue;
                            }

                        }
                        else
                        {
                            string temp = row[4].ToString();
                            string temp1 = row[8].ToString();
                            if (!string.IsNullOrWhiteSpace(temp1))
                            {
                                roadType = temp1;
                            }
                            else
                            {
                                continue;
                            }
                            if (!string.IsNullOrWhiteSpace(temp))
                            {
                                area = double.Parse(temp);
                            }
                            else
                            {
                                continue;
                            }
                        }

                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                    if (!string.IsNullOrEmpty(disName))
                    {
                        dis.RoadType = roadType;
                        dis.Area = area;
                        bool hasDis = false;
                        for (int t = 0; t < allDis.Count; ++t)
                        {
                            var itemDis = allDis.ElementAt(t);

                            if (itemDis.RoadType == dis.RoadType
                                && itemDis.DamagedCondition == dis.DamagedCondition
                                && itemDis.Name == dis.Name
                                && itemDis.RoadGrad == dis.RoadGrad
                                )
                            {
                                allDis[t].Area += dis.Area;
                                hasDis = true;
                            }
                        }
                        if (!hasDis)//第一次出现该病害
                        {
                            allDis.Add(dis);
                        }
                    }

                }
            }

            foreach (var disItem in allDis)
            {
                DiseaseFormXml disOnly = null;
                List<DiseaseFormXml> findDis = AllDiseaseFormXml.Where(
                     t => t.DiseaseName.Contains(disItem.Name)
                     && t.RoadType.Contains(disItem.RoadType)
                     && t.RoadGrad.Contains(disItem.RoadGrad)
                     ).ToList();
                if (findDis.Count == 0)
                {
                    continue;
                }
                else if (findDis.Count == 1)
                {
                    disOnly = findDis.First();
                }
                else
                {
                    if (string.IsNullOrEmpty(disItem.DamagedCondition) || disItem.DamagedCondition == "无")
                    {

                        disOnly = findDis.Where(t => t.DiseaseName.Split('.').Length == 1).FirstOrDefault();
                    }
                    else
                    {
                        disOnly = findDis.Where(t => t.DiseaseName.Contains(disItem.DamagedCondition)).FirstOrDefault();
                    }
                }
                string disKey = "";

                if (disOnly == null)
                {
                    continue;
                }
                if (disOnly != null)
                {

                    disKey = disOnly.DiseaseName;
                }

                if (allDisFormXml.Keys.Contains(disKey))
                {
                    allDisFormXml[disKey] += disItem.Area;
                }
                else
                {
                    allDisFormXml.Add(disKey, disItem.Area);
                }
            }
            if (allDisFormXml.Count <= 0)
            {
                return "";
            }
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            int rowCount = 1;
            if (allDisFormXml.Count >= 5)
            {
                height = 250;//4.81
            }
            double sumArea = 0;
            foreach (var item in allDisFormXml)
            {
                sumArea += item.Value;
            }
            foreach (var item in allDisFormXml)
            {
                worksheet.Cells[rowCount, 1].Value = item.Key;
                worksheet.Cells[rowCount, 2].Value = (item.Value / sumArea * 100).ToString("0.##") + "%";
                rowCount++;
            }
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, width, height);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //4.8,7.5
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;

            MSExcel.Range dataRange = worksheet.Range["A1:B" + (rowCount - 1).ToString()];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            //series1.ChartType = MSExcel.XlChartType.xlPie; MSExcel.XlChartType.xl3DPie;
            series1.ChartType = MSExcel.XlChartType.xl3DPie;
            // 设置图例文本

            series1.XValues = worksheet.Range["A1:A" + (rowCount - 1)];
            chart.ChartColor = 13;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelBestFit);
            //series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // 移除右边的标注
            chart.HasLegend = true;

            // 设置数据标签可见
            series1.HasDataLabels = true;
            // chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelBestFit);
            // 设置数据标签位置为最佳位置
            series1.HasLeaderLines = true;
            // 调整标签布局，防止重叠
            //  chart.ApplyLayout(3); // 设置标签布局为自动调整
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            //  chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            //chart.ChartTitle.Font.Name = "宋体";
            //chart.ChartTitle.Font.Size = 12;

            chart.Legend.Font.Name = "宋体";
            chart.Legend.Font.Size = 12;
            // 设置轴标签字体
            // chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Name = "宋体";
            // chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Size = 12;
            MSExcel.DataLabels dataLabels = series1.DataLabels();
            dataLabels.Font.Name = "宋体";
            dataLabels.Font.Size = 12;

            return SaveExcelChartAsImage(chart);
            // 设置标签布局为自动调整

            //CWB_ExcelHelper.disposeExcel(ref excelApp);

        }

        private void getPciPicture001()
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "优";
            worksheet.Cells[1, 2].Value = "良";
            worksheet.Cells[1, 3].Value = "中";
            worksheet.Cells[1, 4].Value = "次";
            worksheet.Cells[1, 5].Value = "差";

            worksheet.Cells[2, 1].Value = (yPCIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 2].Value = (lPCIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 3].Value = (zPCIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 4].Value = (ciPCIRoadRate * 100).ToString("0.##") + "%";
            worksheet.Cells[2, 5].Value = (chaPCIRoadRate * 100).ToString("0.##") + "%";

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:E2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = false;

            // 设置数据标签可见
            series1.HasDataLabels = true;

            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);

        }

        private void getPart5Picture001()
        {
            double xiandaoLength = 0;
            double xiangdaoLength = 0;
            double cundaoLength = 0;


            double xiandaoPqi = 0;
            double xiandaoPci = 0;
            double xiandaoRqi = 0;
            double goodXianRoadLength = 0;
            double goodXiangRoadLength = 0;
            double goodCunRoadLength = 0;
            //次差路长度
            double ciChaXianRoadLength = 0;
            double ciChaXiangRoadLength = 0;
            double ciChaCunRoadLength = 0;
            double xiangdaoPqi = 0;
            double xiangdaoPci = 0;
            double xiangdaoRqi = 0;

            double cundaoPqi = 0;
            double cundaoPci = 0;
            double cundaoRqi = 0;
            foreach (var item in allDatas)
            {

                if (item.RoadCode.StartsWith("X"))
                {
                    xiandaoLength += item.RoadLen;
                    xiandaoPqi += item.PqiValue * item.RoadLen;
                    xiandaoPci += item.PciValue * item.RoadLen;
                    xiandaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXianRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")

                    {
                        ciChaXianRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("Y"))
                {
                    xiangdaoLength += item.RoadLen;
                    xiangdaoPci += item.PciValue * item.RoadLen;
                    xiangdaoPqi += item.PqiValue * item.RoadLen;
                    xiangdaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXiangRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXiangRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("C"))
                {
                    cundaoLength += item.RoadLen;
                    cundaoPci += item.PciValue * item.RoadLen;
                    cundaoPqi += item.PqiValue * item.RoadLen;
                    cundaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodCunRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaCunRoadLength += item.RoadLen;
                    }
                }


            }
            xiandaoPci /= xiandaoLength;
            xiandaoPqi /= xiandaoLength;
            xiandaoRqi /= xiandaoLength;
            string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "0" : xiandaoPci.ToString("0.##");
            string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "0" : xiandaoPqi.ToString("0.##");
            string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "0" : xiandaoRqi.ToString("0.##");

            xiangdaoPci /= xiangdaoLength;
            xiangdaoPqi /= xiangdaoLength;
            xiangdaoRqi /= xiangdaoLength;

            string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "0" : xiangdaoPci.ToString("0.##");
            string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "0" : xiangdaoPqi.ToString("0.##");
            string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "0" : xiangdaoRqi.ToString("0.##");

            cundaoPci /= cundaoLength;
            cundaoPqi /= cundaoLength;
            cundaoRqi /= cundaoLength;
            string cundaoPciStr = double.IsNaN(cundaoPci) ? "0" : cundaoPci.ToString("0.##");
            string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "0" : cundaoPqi.ToString("0.##");
            string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "0" : cundaoRqi.ToString("0.##");
            string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "0" : SetGrad(xiandaoPqi);
            string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "0" : SetGrad(xiangdaoPqi);
            string cundaojudge = double.IsNaN(cundaoPqi) ? "0" : SetGrad(cundaoPqi);

            string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "0%" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "0%" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "0%" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
            string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "0%" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "0%" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "0%" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "县道";
            worksheet.Cells[1, 3].Value = "乡道";
            worksheet.Cells[1, 4].Value = "村道";


            worksheet.Cells[2, 1].Value = "PQI(优良路率)";
            worksheet.Cells[2, 2].Value = xiandaoPqiStr;
            worksheet.Cells[2, 3].Value = xiangdaoPqiStr;
            worksheet.Cells[2, 4].Value = cundaoPqiStr;


            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:D2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);


            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;


            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodXianRoadRate, goodXiangRoadRate, goodCunRoadRate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPart6Picture001()
        {
            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;

            double grad0PqiValue = 0;
            double grad1PqiValue = 0;
            double grad0PciValue = 0;
            double grad1PciValue = 0;
            double grad0RqiValue = 0;
            double grad1RqiValue = 0;

            double grad2PqiValue = 0;
            double grad3PqiValue = 0;
            double grad2PciValue = 0;
            double grad3PciValue = 0;
            double grad2RqiValue = 0;
            double grad3RqiValue = 0;


            double goodGrad0Length = 0;
            double goodGrad1Length = 0;
            double goodGrad2Length = 0;
            double goodGrad3Length = 0;
            double cichaGrad0Length = 0;
            double cichaGrad1Length = 0;
            double cichaGrad2Length = 0;
            double cichaGrad3Length = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    if (item1.RoadGrad == 1)
                    {
                        grad0PqiValue += item1.PqiValue * item1.RoadLen;
                        grad0PciValue += item1.PciValue * item1.RoadLen;
                        grad0RqiValue += item1.RqiValue * item1.RoadLen;
                        length0 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad0Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad0Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 2)
                    {
                        grad1PqiValue += item1.PqiValue * item1.RoadLen;
                        grad1PciValue += item1.PciValue * item1.RoadLen;
                        grad1RqiValue += item1.RqiValue * item1.RoadLen;
                        length1 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad1Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad1Length += item1.RoadLen;
                        }

                    }


                    if (item1.RoadGrad == 3)
                    {
                        grad2PqiValue += item1.PqiValue * item1.RoadLen;
                        grad2PciValue += item1.PciValue * item1.RoadLen;
                        grad2RqiValue += item1.RqiValue * item1.RoadLen;
                        length2 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad2Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad2Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 4)
                    {
                        grad3PqiValue += item1.PqiValue * item1.RoadLen;
                        grad3PciValue += item1.PciValue * item1.RoadLen;
                        grad3RqiValue += item1.RqiValue * item1.RoadLen;
                        length3 += item1.RoadLen;

                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad3Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad3Length += item1.RoadLen;
                        }
                    }
                }

            }

            grad0PqiValue = double.IsNaN(grad0PqiValue / length0) ? 0 : grad0PqiValue / length0;
            grad0PciValue = double.IsNaN(grad0PciValue / length0) ? 0 : grad0PciValue / length0;
            grad0RqiValue = double.IsNaN(grad0RqiValue / length0) ? 0 : grad0RqiValue / length0;

            grad1PqiValue = double.IsNaN(grad1PqiValue / length1) ? 0 : grad1PqiValue / length1;
            grad1PciValue = double.IsNaN(grad1PciValue / length1) ? 0 : grad1PciValue / length1;
            grad1RqiValue = double.IsNaN(grad1RqiValue / length1) ? 0 : grad1RqiValue / length1;

            grad3PqiValue = double.IsNaN(grad3PqiValue / length3) ? 0 : grad3PqiValue / length3;
            grad3PciValue = double.IsNaN(grad3PciValue / length3) ? 0 : grad3PciValue / length3;
            grad3RqiValue = double.IsNaN(grad3RqiValue / length3) ? 0 : grad3RqiValue / length3;


            grad2PqiValue = double.IsNaN(grad2PqiValue / length2) ? 0 : grad2PqiValue / length2;
            grad2PciValue = double.IsNaN(grad2PciValue / length2) ? 0 : grad2PciValue / length2;
            grad2RqiValue = double.IsNaN(grad2RqiValue / length2) ? 0 : grad2RqiValue / length2;

            string goodGrad0Rate = double.IsNaN(goodGrad0Length / length0) ? "0%" : (goodGrad0Length / length0 * 100).ToString("0.##") + "%";
            string goodGrad1Rate = double.IsNaN(goodGrad1Length / length1) ? "0%" : (goodGrad1Length / length1 * 100).ToString("0.##") + "%";
            string goodGrad2Rate = double.IsNaN(goodGrad2Length / length2) ? "0%" : (goodGrad2Length / length2 * 100).ToString("0.##") + "%";
            string goodGrad3Rate = double.IsNaN(goodGrad3Length / length3) ? "0%" : (goodGrad3Length / length3 * 100).ToString("0.##") + "%";

            string cichaGrad2Rate = double.IsNaN(cichaGrad2Length / length2) ? "0%" : (cichaGrad2Length / length2 * 100).ToString("0.##") + "%";
            string cichaGrad3Rate = double.IsNaN(cichaGrad3Length / length2) ? "0%" : (cichaGrad3Length / length3 * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "一级公路";
            worksheet.Cells[1, 3].Value = "二级公路";
            worksheet.Cells[1, 4].Value = "三级公路";
            worksheet.Cells[1, 5].Value = "四级公路";


            worksheet.Cells[2, 1].Value = "PQI(优良路率)";
            worksheet.Cells[2, 2].Value = grad0PqiValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = grad1PqiValue.ToString("0.##");
            worksheet.Cells[2, 4].Value = grad2PqiValue.ToString("0.##");
            worksheet.Cells[2, 5].Value = grad3PqiValue.ToString("0.##");
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:E2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;


            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodGrad0Rate, goodGrad1Rate, goodGrad2Rate, goodGrad3Rate);

            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPart7Picture001()
        {
            double snLength = 0;
            double lqLength = 0;

            double lqPqiValue = 0;
            double snPqiValue = 0;
            double lqPciValue = 0;
            double snPciValue = 0;
            double lqRqiValue = 0;
            double snRqiValue = 0;
            double goodLqLength = 0;
            double goodSnLength = 0;
            double cichaLqLength = 0;
            double cichaSnLength = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {

                    if (item1.RoadType == 0)
                    {
                        lqLength += item1.RoadLen;
                        lqPciValue += item1.PciValue * item1.RoadLen;
                        lqPqiValue += item1.PqiValue * item1.RoadLen;
                        lqRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodLqLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaLqLength += item1.RoadLen;
                        }
                    }
                    if (item1.RoadType == 1)
                    {
                        snLength += item1.RoadLen;
                        snPciValue += item1.PciValue * item1.RoadLen;
                        snPqiValue += item1.PqiValue * item1.RoadLen;
                        snRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodSnLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaSnLength += item1.RoadLen;
                        }
                    }
                }
            }
            lqPqiValue /= lqLength;
            snPqiValue /= snLength;
            lqPciValue /= lqLength;
            snPciValue /= snLength;
            lqRqiValue /= lqLength;
            snRqiValue /= snLength;
            string goodLqRate = (goodLqLength / lqLength * 100).ToString("0.##");
            string goodSnRate = (goodSnLength / snLength * 100).ToString("0.##");
            string cichaLqRate = (cichaLqLength / lqLength * 100).ToString("0.##");
            string cichaSnRate = (cichaSnLength / snLength * 100).ToString("0.##");


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 设置字体

            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;
            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "沥青路面";
            worksheet.Cells[1, 3].Value = "水泥路面";

            worksheet.Cells[2, 1].Value = "PQI(优良路率)";
            worksheet.Cells[2, 2].Value = lqPqiValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = snPqiValue.ToString("0.##");
            // worksheet.Cells[2, 2].Value = lqPqiValue.ToString("0.##") + "(" + goodLqRate + ")";
            //  worksheet.Cells[2, 3].Value = snPqiValue.ToString("0.##") + "(" + goodSnRate + ")";
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            //设置标题文字
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:C2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);
            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;


            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);

            setChartFont(series1, chart, goodLqRate, goodSnRate);

            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }


        private void setChartFont(MSExcel.Series series1, MSExcel.Chart chart, string x1Str, string x2Str
            )
        {
            MSExcel.DataLabels dataLabels = series1.DataLabels();
            dataLabels.Item(1).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x1Str));
            dataLabels.Item(2).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x2Str));

            //chart.ChartTitle.Font.Name = "宋体";
            //chart.ChartTitle.Font.Size = 12;

            chart.Legend.Font.Name = "宋体";
            chart.Legend.Font.Size = 12;
            // 设置轴标签字体
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Name = "宋体";
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Size = 12;
            dataLabels.Font.Name = "宋体";
            dataLabels.Font.Size = 12;
        }

        private void setChartFont(MSExcel.Series series1, MSExcel.Chart chart, string x1Str, string x2Str, string x3Str, string x4Str)
        {
            MSExcel.DataLabels dataLabels = series1.DataLabels();
            dataLabels.Item(1).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x1Str));
            dataLabels.Item(2).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x2Str));
            dataLabels.Item(3).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x3Str));
            dataLabels.Item(4).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x4Str));

            //chart.ChartTitle.Font.Name = "宋体";
            //chart.ChartTitle.Font.Size = 12;

            chart.Legend.Font.Name = "宋体";
            chart.Legend.Font.Size = 12;
            // 设置轴标签字体
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Name = "宋体";
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Size = 12;
            dataLabels.Font.Name = "宋体";
            dataLabels.Font.Size = 12;
        }
        private void setChartFont(MSExcel.Series series1, MSExcel.Chart chart, string x1Str, string x2Str, string x3Str)
        {
            MSExcel.DataLabels dataLabels = series1.DataLabels();
            dataLabels.Item(1).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x1Str));
            dataLabels.Item(2).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x2Str));
            dataLabels.Item(3).Format.TextFrame2.TextRange.InsertAfter(string.Format("({0})", x3Str));

            //chart.ChartTitle.Font.Name = "宋体";
            //chart.ChartTitle.Font.Size = 12;

            chart.Legend.Font.Name = "宋体";
            chart.Legend.Font.Size = 12;
            // 设置轴标签字体
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Name = "宋体";
            chart.Axes(MSExcel.XlAxisType.xlCategory).TickLabels.Font.Size = 12;
            dataLabels.Font.Name = "宋体";
            dataLabels.Font.Size = 12;
        }
        private void getPart5Picture002()
        {
            double xiandaoLength = 0;
            double xiangdaoLength = 0;
            double cundaoLength = 0;


            double xiandaoPqi = 0;
            double xiandaoPci = 0;
            double xiandaoRqi = 0;
            double goodXianRoadLength = 0;
            double goodXiangRoadLength = 0;
            double goodCunRoadLength = 0;
            //次差路长度
            double ciChaXianRoadLength = 0;
            double ciChaXiangRoadLength = 0;
            double ciChaCunRoadLength = 0;
            double xiangdaoPqi = 0;
            double xiangdaoPci = 0;
            double xiangdaoRqi = 0;

            double cundaoPqi = 0;
            double cundaoPci = 0;
            double cundaoRqi = 0;
            foreach (var item in allDatas)
            {

                if (item.RoadCode.StartsWith("X"))
                {
                    xiandaoLength += item.RoadLen;
                    xiandaoPqi += item.PqiValue * item.RoadLen;
                    xiandaoPci += item.PciValue * item.RoadLen;
                    xiandaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PciValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXianRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXianRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("Y"))
                {
                    xiangdaoLength += item.RoadLen;
                    xiangdaoPci += item.PciValue * item.RoadLen;
                    xiangdaoPqi += item.PqiValue * item.RoadLen;
                    xiangdaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PciValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXiangRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXiangRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("C"))
                {
                    cundaoLength += item.RoadLen;
                    cundaoPci += item.PciValue * item.RoadLen;
                    cundaoPqi += item.PqiValue * item.RoadLen;
                    cundaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.PciValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodCunRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaCunRoadLength += item.RoadLen;
                    }
                }


            }
            xiandaoPci /= xiandaoLength;
            xiandaoPqi /= xiandaoLength;
            xiandaoRqi /= xiandaoLength;
            string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "0" : xiandaoPci.ToString("0.##");
            string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "0" : xiandaoPqi.ToString("0.##");
            string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "0" : xiandaoRqi.ToString("0.##");

            xiangdaoPci /= xiangdaoLength;
            xiangdaoPqi /= xiangdaoLength;
            xiangdaoRqi /= xiangdaoLength;

            string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "0" : xiangdaoPci.ToString("0.##");
            string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "0" : xiangdaoPqi.ToString("0.##");
            string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "0" : xiangdaoRqi.ToString("0.##");

            cundaoPci /= cundaoLength;
            cundaoPqi /= cundaoLength;
            cundaoRqi /= cundaoLength;
            string cundaoPciStr = double.IsNaN(cundaoPci) ? "0" : cundaoPci.ToString("0.##");
            string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "0" : cundaoPqi.ToString("0.##");
            string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "0" : cundaoRqi.ToString("0.##");
            string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "0" : SetGrad(xiandaoPqi);
            string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "0" : SetGrad(xiangdaoPqi);
            string cundaojudge = double.IsNaN(cundaoPqi) ? "0" : SetGrad(cundaoPqi);

            string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "0%" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "0%" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "0%" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
            string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "0%" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "0%" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "0%" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "县道";
            worksheet.Cells[1, 3].Value = "乡道";
            worksheet.Cells[1, 4].Value = "村道";

            worksheet.Cells[2, 1].Value = "PCI(优良路率)";
            worksheet.Cells[2, 2].Value = xiandaoPciStr;
            worksheet.Cells[2, 3].Value = xiangdaoPciStr;
            worksheet.Cells[2, 4].Value = cundaoPciStr;

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:D2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;


            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            setChartFont(series1, chart, goodXianRoadRate, goodXiangRoadRate, goodCunRoadRate);
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPart6Picture002()
        {
            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;


            double grad2PqiValue = 0;
            double grad3PqiValue = 0;
            double grad0PciValue = 0;
            double grad1PciValue = 0;
            double grad2PciValue = 0;
            double grad3PciValue = 0;
            double grad2RqiValue = 0;
            double grad3RqiValue = 0;


            double goodGrad0Length = 0;
            double goodGrad1Length = 0;
            double goodGrad2Length = 0;
            double goodGrad3Length = 0;
            double cichaGrad2Length = 0;
            double cichaGrad3Length = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    if (item1.RoadGrad == 1)
                    {

                        grad0PciValue += item1.PciValue * item1.RoadLen;

                        length0 += item1.RoadLen;

                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad0Length += item1.RoadLen;
                        }
                        else
                        {

                        }

                    }
                    if (item1.RoadGrad == 2)
                    {
                        grad1PciValue += item1.PciValue * item1.RoadLen;
                        length1 += item1.RoadLen;

                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad1Length += item1.RoadLen;
                        }
                        else
                        {
                        }

                    }
                    if (item1.RoadGrad == 3)
                    {
                        grad2PqiValue += item1.PqiValue * item1.RoadLen;
                        grad2PciValue += item1.PciValue * item1.RoadLen;
                        grad2RqiValue += item1.RqiValue * item1.RoadLen;
                        length2 += item1.RoadLen;

                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad2Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad2Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 4)
                    {
                        grad3PqiValue += item1.PqiValue * item1.RoadLen;
                        grad3PciValue += item1.PciValue * item1.RoadLen;
                        grad3RqiValue += item1.RqiValue * item1.RoadLen;
                        length3 += item1.RoadLen;

                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad3Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad3Length += item1.RoadLen;
                        }
                    }
                }

            }
            grad3PqiValue = double.IsNaN(grad3PqiValue / length3) ? 0 : grad3PqiValue / length3;
            grad3PciValue = double.IsNaN(grad3PciValue / length3) ? 0 : grad3PciValue / length3;
            grad3RqiValue = double.IsNaN(grad3RqiValue / length3) ? 0 : grad3RqiValue / length3;


            grad2PqiValue = double.IsNaN(grad2PqiValue / length2) ? 0 : grad2PqiValue / length2;

            grad0PciValue = double.IsNaN(grad0PciValue / length0) ? 0 : grad0PciValue / length0;
            grad1PciValue = double.IsNaN(grad1PciValue / length1) ? 0 : grad1PciValue / length1;

            grad2PciValue = double.IsNaN(grad2PciValue / length2) ? 0 : grad2PciValue / length2;
            grad2RqiValue = double.IsNaN(grad2RqiValue / length2) ? 0 : grad2RqiValue / length2;

            string goodGrad3Rate = double.IsNaN(goodGrad3Length / length3) ? "0%" : (goodGrad3Length / length3 * 100).ToString("0.##") + "%";
            string goodGrad2Rate = double.IsNaN(goodGrad2Length / length2) ? "0%" : (goodGrad2Length / length2 * 100).ToString("0.##") + "%";
            string goodGrad0Rate = double.IsNaN(goodGrad0Length / length0) ? "0%" : (goodGrad0Length / length0 * 100).ToString("0.##") + "%";
            string goodGrad1Rate = double.IsNaN(goodGrad1Length / length1) ? "0%" : (goodGrad1Length / length1 * 100).ToString("0.##") + "%";

            string cichaGrad2Rate = double.IsNaN(cichaGrad2Length / length2) ? "0%" : (cichaGrad2Length / length2 * 100).ToString("0.##") + "%";
            string cichaGrad3Rate = double.IsNaN(cichaGrad3Length / length2) ? "0%" : (cichaGrad3Length / length3 * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "一级公路";
            worksheet.Cells[1, 3].Value = "二级公路";
            worksheet.Cells[1, 4].Value = "三级公路";
            worksheet.Cells[1, 5].Value = "四级公路";

            worksheet.Cells[2, 1].Value = "PCI(优良路率)";
            worksheet.Cells[2, 2].Value = grad0PciValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = grad1PciValue.ToString("0.##");
            worksheet.Cells[2, 4].Value = grad2PciValue.ToString("0.##");
            worksheet.Cells[2, 5].Value = grad3PciValue.ToString("0.##");


            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:E2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodGrad0Rate, goodGrad1Rate, goodGrad2Rate, goodGrad3Rate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private string getPart9Picture001(HeFeiRaod road, ref MSExcel.Application excelApp, double height, double width)
        {

            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "PQI";
            worksheet.Cells[1, 2].Value = "PCI";
            worksheet.Cells[1, 3].Value = "RQI";

            worksheet.Cells[2, 1].Value = road.PqiValue;
            worksheet.Cells[2, 2].Value = road.PciValue;
            worksheet.Cells[2, 3].Value = road.RqiValue;

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, width, height);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:C2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlRadar;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementPrimaryValueAxisNone);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendNone);
            //  chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            return SaveExcelChartAsImage(chart);
            // CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPart8Picture001()
        {
            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;
            foreach (var item1 in allDatas)
            {
                foreach (var item in item1.datas)
                {
                    string temp = SetGrad(item.PqiValue);

                    if (temp == "次" || temp == "差")
                    {
                        length0 += item.RoadLen;


                    }
                }
            }
            foreach (var item1 in allDatas)
            {
                foreach (var item in item1.datas)
                {
                    string temp = SetGrad(item.PciValue);

                    if (temp == "次" || temp == "差")
                    {
                        length1 += item.RoadLen;


                    }
                }
            }
            foreach (var item1 in allDatas)
            {
                foreach (var item in item1.datas)
                {
                    string temp = SetGrad(item.RqiValue);

                    if (temp == "次" || temp == "差")
                    {
                        length2 += item.RoadLen;


                    }
                }
            }

            foreach (var item1 in allDatas)
            {
                foreach (var item in item1.datas)
                {
                    string temp = SetGrad(item.RqiValue);
                    string temp1 = SetGrad(item.PciValue);

                    if (temp == "次" || temp == "差")
                    {
                        if (temp1 == "次" || temp1 == "差")
                            length3 += item.RoadLen;


                    }
                }
            }
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "PQI";
            worksheet.Cells[1, 2].Value = "PCI";
            worksheet.Cells[1, 3].Value = "RQI";
            worksheet.Cells[1, 4].Value = "PCI&RQI";

            worksheet.Cells[2, 1].Value = length0;
            worksheet.Cells[2, 2].Value = length1;
            worksheet.Cells[2, 3].Value = length2;
            worksheet.Cells[2, 4].Value = length3;

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:D2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板

            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPart7Picture002()
        {
            double snLength = 0;
            double lqLength = 0;

            double lqPqiValue = 0;
            double snPqiValue = 0;
            double lqPciValue = 0;
            double snPciValue = 0;
            double lqRqiValue = 0;
            double snRqiValue = 0;
            double goodLqLength = 0;
            double goodSnLength = 0;
            double cichaLqLength = 0;
            double cichaSnLength = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {

                    if (item1.RoadType == 0)
                    {
                        lqLength += item1.RoadLen;
                        lqPciValue += item1.PciValue * item1.RoadLen;
                        lqPqiValue += item1.PqiValue * item1.RoadLen;
                        lqRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodLqLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaLqLength += item1.RoadLen;
                        }
                    }
                    if (item1.RoadType == 1)
                    {
                        snLength += item1.RoadLen;
                        snPciValue += item1.PciValue * item1.RoadLen;
                        snPqiValue += item1.PqiValue * item1.RoadLen;
                        snRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.PciValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodSnLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaSnLength += item1.RoadLen;
                        }
                    }
                }
            }
            lqPqiValue /= lqLength;
            snPqiValue /= snLength;
            lqPciValue /= lqLength;
            snPciValue /= snLength;
            lqRqiValue /= lqLength;
            snRqiValue /= snLength;
            string goodLqRate = (goodLqLength / lqLength * 100).ToString("0.##");
            string goodSnRate = (goodSnLength / snLength * 100).ToString("0.##");
            string cichaLqRate = (cichaLqLength / lqLength * 100).ToString("0.##");
            string cichaSnRate = (cichaSnLength / snLength * 100).ToString("0.##");



            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "沥青路面";
            worksheet.Cells[1, 3].Value = "水泥路面";

            worksheet.Cells[2, 1].Value = "PCI(优良路率)";
            worksheet.Cells[2, 2].Value = lqPciValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = snPciValue.ToString("0.##");
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;
            MSExcel.Range dataRange = worksheet.Range["A1:C2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodLqRate, goodSnRate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPart5Picture003()
        {
            double xiandaoLength = 0;
            double xiangdaoLength = 0;
            double cundaoLength = 0;


            double xiandaoPqi = 0;
            double xiandaoPci = 0;
            double xiandaoRqi = 0;
            double goodXianRoadLength = 0;
            double goodXiangRoadLength = 0;
            double goodCunRoadLength = 0;
            //次差路长度
            double ciChaXianRoadLength = 0;
            double ciChaXiangRoadLength = 0;
            double ciChaCunRoadLength = 0;
            double xiangdaoPqi = 0;
            double xiangdaoPci = 0;
            double xiangdaoRqi = 0;

            double cundaoPqi = 0;
            double cundaoPci = 0;
            double cundaoRqi = 0;
            foreach (var item in allDatas)
            {

                if (item.RoadCode.StartsWith("X"))
                {
                    xiandaoLength += item.RoadLen;
                    xiandaoPqi += item.PqiValue * item.RoadLen;
                    xiandaoPci += item.PciValue * item.RoadLen;
                    xiandaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.RqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXianRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXianRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("Y"))
                {
                    xiangdaoLength += item.RoadLen;
                    xiangdaoPci += item.PciValue * item.RoadLen;
                    xiangdaoPqi += item.PqiValue * item.RoadLen;
                    xiangdaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.RqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodXiangRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaXiangRoadLength += item.RoadLen;
                    }
                }
                if (item.RoadCode.StartsWith("C"))
                {
                    cundaoLength += item.RoadLen;
                    cundaoPci += item.PciValue * item.RoadLen;
                    cundaoPqi += item.PqiValue * item.RoadLen;
                    cundaoRqi += item.RqiValue * item.RoadLen;
                    string judge = SetGrad(item.RqiValue);
                    if (judge == "优" || judge == "良")
                    {
                        goodCunRoadLength += item.RoadLen;
                    }
                    else if (judge == "次" || judge == "差")
                    {
                        ciChaCunRoadLength += item.RoadLen;
                    }
                }



            }
            xiandaoPci /= xiandaoLength;
            xiandaoPqi /= xiandaoLength;
            xiandaoRqi /= xiandaoLength;
            string xiandaoPciStr = double.IsNaN(xiandaoPci) ? "0" : xiandaoPci.ToString("0.##");
            string xiandaoPqiStr = double.IsNaN(xiandaoPqi) ? "0" : xiandaoPqi.ToString("0.##");
            string xiandaoRqiStr = double.IsNaN(xiandaoRqi) ? "0" : xiandaoRqi.ToString("0.##");

            xiangdaoPci /= xiangdaoLength;
            xiangdaoPqi /= xiangdaoLength;
            xiangdaoRqi /= xiangdaoLength;

            string xiangdaoPciStr = double.IsNaN(xiangdaoPci) ? "0" : xiangdaoPci.ToString("0.##");
            string xiangdaoPqiStr = double.IsNaN(xiangdaoPqi) ? "0" : xiangdaoPqi.ToString("0.##");
            string xiangdaoRqiStr = double.IsNaN(xiangdaoRqi) ? "0" : xiangdaoRqi.ToString("0.##");

            cundaoPci /= cundaoLength;
            cundaoPqi /= cundaoLength;
            cundaoRqi /= cundaoLength;
            string cundaoPciStr = double.IsNaN(cundaoPci) ? "0" : cundaoPci.ToString("0.##");
            string cundaoPqiStr = double.IsNaN(cundaoPqi) ? "0" : cundaoPqi.ToString("0.##");
            string cundaoRqiStr = double.IsNaN(cundaoRqi) ? "0" : cundaoRqi.ToString("0.##");
            string xiandaoJudge = double.IsNaN(xiandaoPqi) ? "0" : SetGrad(xiandaoPqi);
            string xiangdaojudge = double.IsNaN(xiangdaoPqi) ? "0" : SetGrad(xiangdaoPqi);
            string cundaojudge = double.IsNaN(cundaoPqi) ? "0" : SetGrad(cundaoPqi);

            string goodXianRoadRate = double.IsNaN(goodXianRoadLength / xiandaoLength) ? "0%" : (goodXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string goodXiangRoadRate = double.IsNaN(goodXiangRoadLength / xiangdaoLength) ? "0%" : (goodXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string goodCunRoadRate = double.IsNaN(goodCunRoadLength / cundaoLength) ? "0%" : (goodCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";
            string cichaXianRoadRate = double.IsNaN(ciChaXianRoadLength / xiandaoLength) ? "0%" : (ciChaXianRoadLength / xiandaoLength * 100).ToString("0.##") + "%";
            string cichaXiangRoadRate = double.IsNaN(ciChaXiangRoadLength / xiangdaoLength) ? "0%" : (ciChaXiangRoadLength / xiangdaoLength * 100).ToString("0.##") + "%";
            string cichaCunRoadRate = double.IsNaN(ciChaCunRoadLength / cundaoLength) ? "0%" : (ciChaCunRoadLength / cundaoLength * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "县道";
            worksheet.Cells[1, 3].Value = "乡道";
            worksheet.Cells[1, 4].Value = "村道";

            worksheet.Cells[2, 1].Value = "RQI(优良路率)";
            worksheet.Cells[2, 2].Value = xiandaoRqiStr;
            worksheet.Cells[2, 3].Value = xiangdaoRqiStr;
            worksheet.Cells[2, 4].Value = cundaoRqiStr;
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:D2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            //series2.ChartType = MSExcel.XlChartType.xlXYScatter;
            //series2.AxisGroup = MSExcel.XlAxisGroup.xlSecondary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodXianRoadRate, goodXiangRoadRate, goodCunRoadRate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPart6Picture003()
        {
            double length0 = 0;
            double length1 = 0;
            double length2 = 0;
            double length3 = 0;




            double grad2PqiValue = 0;
            double grad3PqiValue = 0;
            double grad2PciValue = 0;
            double grad3PciValue = 0;
            double grad0RqiValue = 0;
            double grad1RqiValue = 0;
            double grad2RqiValue = 0;
            double grad3RqiValue = 0;


            double goodGrad0Length = 0;
            double goodGrad1Length = 0;
            double goodGrad2Length = 0;
            double goodGrad3Length = 0;
            double cichaGrad2Length = 0;
            double cichaGrad3Length = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {
                    if (item1.RoadGrad == 1)
                    {
                        grad0RqiValue += item1.RqiValue * item1.RoadLen;
                        length0 += item1.RoadLen;

                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad0Length += item1.RoadLen;
                        }
                        else
                        {
                        }

                    }
                    if (item1.RoadGrad == 2)
                    {

                        grad1RqiValue += item1.RqiValue * item1.RoadLen;
                        length1 += item1.RoadLen;

                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad1Length += item1.RoadLen;
                        }
                        else
                        {

                        }

                    }
                    if (item1.RoadGrad == 3)
                    {
                        grad2PqiValue += item1.PqiValue * item1.RoadLen;
                        grad2PciValue += item1.PciValue * item1.RoadLen;
                        grad2RqiValue += item1.RqiValue * item1.RoadLen;
                        length2 += item1.RoadLen;

                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad2Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad2Length += item1.RoadLen;
                        }

                    }
                    if (item1.RoadGrad == 4)
                    {
                        grad3PqiValue += item1.PqiValue * item1.RoadLen;
                        grad3PciValue += item1.PciValue * item1.RoadLen;
                        grad3RqiValue += item1.RqiValue * item1.RoadLen;
                        length3 += item1.RoadLen;

                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodGrad3Length += item1.RoadLen;
                        }
                        else
                        {
                            cichaGrad3Length += item1.RoadLen;
                        }
                    }
                }

            }
            grad3PqiValue = double.IsNaN(grad3PqiValue / length3) ? 0 : grad3PqiValue / length3;
            grad3PciValue = double.IsNaN(grad3PciValue / length3) ? 0 : grad3PciValue / length3;
            grad3RqiValue = double.IsNaN(grad3RqiValue / length3) ? 0 : grad3RqiValue / length3;


            grad2PqiValue = double.IsNaN(grad2PqiValue / length2) ? 0 : grad2PqiValue / length2;
            grad2PciValue = double.IsNaN(grad2PciValue / length2) ? 0 : grad2PciValue / length2;
            grad2RqiValue = double.IsNaN(grad2RqiValue / length2) ? 0 : grad2RqiValue / length2;
            grad0RqiValue = double.IsNaN(grad0RqiValue / length0) ? 0 : grad0RqiValue / length0;
            grad1RqiValue = double.IsNaN(grad1RqiValue / length1) ? 0 : grad1RqiValue / length1;

            string goodGrad0Rate = double.IsNaN(goodGrad0Length / length0) ? "0%" : (goodGrad0Length / length0 * 100).ToString("0.##") + "%";
            string goodGrad1Rate = double.IsNaN(goodGrad1Length / length1) ? "0%" : (goodGrad1Length / length1 * 100).ToString("0.##") + "%";
            string goodGrad3Rate = double.IsNaN(goodGrad3Length / length3) ? "0%" : (goodGrad3Length / length3 * 100).ToString("0.##") + "%";
            string goodGrad2Rate = double.IsNaN(goodGrad2Length / length2) ? "0%" : (goodGrad2Length / length2 * 100).ToString("0.##") + "%";

            string cichaGrad2Rate = double.IsNaN(cichaGrad2Length / length2) ? "0%" : (cichaGrad2Length / length2 * 100).ToString("0.##") + "%";
            string cichaGrad3Rate = double.IsNaN(cichaGrad3Length / length2) ? "0%" : (cichaGrad3Length / length3 * 100).ToString("0.##") + "%";


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "三级公路";
            worksheet.Cells[1, 3].Value = "四级公路";
            worksheet.Cells[1, 4].Value = "四级公路";
            worksheet.Cells[1, 5].Value = "四级公路";

            worksheet.Cells[2, 1].Value = "RQI(优良路率)";
            worksheet.Cells[2, 2].Value = grad0RqiValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = grad1RqiValue.ToString("0.##");
            worksheet.Cells[2, 4].Value = grad2RqiValue.ToString("0.##");
            worksheet.Cells[2, 5].Value = grad3RqiValue.ToString("0.##");


            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:E2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            //series2.ChartType = MSExcel.XlChartType.xlXYScatter;
            // series2.AxisGroup = MSExcel.XlAxisGroup.xlSecondary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodGrad0Rate, goodGrad1Rate, goodGrad2Rate, goodGrad3Rate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }


        private void getPart7Picture003()
        {
            double snLength = 0;
            double lqLength = 0;

            double lqPqiValue = 0;
            double snPqiValue = 0;
            double lqPciValue = 0;
            double snPciValue = 0;
            double lqRqiValue = 0;
            double snRqiValue = 0;
            double goodLqLength = 0;
            double goodSnLength = 0;
            double cichaLqLength = 0;
            double cichaSnLength = 0;
            foreach (var item in allDatas)
            {
                foreach (var item1 in item.datas)
                {

                    if (item1.RoadType == 0)
                    {
                        lqLength += item1.RoadLen;
                        lqPciValue += item1.PciValue * item1.RoadLen;
                        lqPqiValue += item1.PqiValue * item1.RoadLen;
                        lqRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodLqLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaLqLength += item1.RoadLen;
                        }
                    }
                    if (item1.RoadType == 1)
                    {
                        snLength += item1.RoadLen;
                        snPciValue += item1.PciValue * item1.RoadLen;
                        snPqiValue += item1.PqiValue * item1.RoadLen;
                        snRqiValue += item1.RqiValue * item1.RoadLen;
                        string temp = SetGrad(item1.RqiValue);
                        if (temp == "优" || temp == "良")
                        {
                            goodSnLength += item1.RoadLen;
                        }
                        else
                        {
                            cichaSnLength += item1.RoadLen;
                        }
                    }
                }
            }
            lqPqiValue /= lqLength;
            snPqiValue /= snLength;
            lqPciValue /= lqLength;
            snPciValue /= snLength;
            lqRqiValue /= lqLength;
            snRqiValue /= snLength;
            string goodLqRate = (goodLqLength / lqLength * 100).ToString("0.##");
            string goodSnRate = (goodSnLength / snLength * 100).ToString("0.##");
            string cichaLqRate = (cichaLqLength / lqLength * 100).ToString("0.##");
            string cichaSnRate = (cichaSnLength / snLength * 100).ToString("0.##");


            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;

            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "";
            worksheet.Cells[1, 2].Value = "沥青路面";
            worksheet.Cells[1, 3].Value = "水泥路面";

            worksheet.Cells[2, 1].Value = "RQI(优良路率)";
            worksheet.Cells[2, 2].Value = lqRqiValue.ToString("0.##");
            worksheet.Cells[2, 3].Value = snRqiValue.ToString("0.##");

            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:C2"];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);


            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;

            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementDataLabelShow);
            chart.SetElement(MSOffice.MsoChartElementType.msoElementChartTitleNone);
            // 移除标题
            chart.HasTitle = false;
            // 移除右边的标注
            chart.HasLegend = false;
            // 设置数据标签可见
            series1.HasDataLabels = true;
            chart.SetElement(MSOffice.MsoChartElementType.msoElementLegendRight);
            // chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            setChartFont(series1, chart, goodLqRate, goodSnRate);
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }

        private void getPciPicture002()
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路名称";
            worksheet.Cells[1, 2].Value = "PCI";
            int rowCount = 1;
            var datas = allDatas.OrderByDescending(t => t.PciValue).ToList();


            if (datas.Count <= 20)
            {
                for (int i = 0; i < datas.Count; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].PciValue;

                }
            }
            else
            {
                for (int i = 0; i < 20; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].PciValue;

                }
            }



            // 根据数据创建折线图
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:B" + rowCount.ToString()];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = false;
            chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getPqiPicture001()
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            MSExcel.Application excelApp = new MSExcel.Application();
            excelApp.Visible = false;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1].Value = "道路名称";
            worksheet.Cells[1, 2].Value = "PQI";
            int rowCount = 1;

            var datas = allDatas.OrderByDescending(t => t.PqiValue).ToList();
            if (datas.Count <= 20)
            {
                for (int i = 0; i < datas.Count; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].PqiValue;

                }
            }
            else
            {
                for (int i = 0; i < 20; i++)
                {
                    rowCount++;
                    worksheet.Cells[rowCount, 1].Value = datas[i].RoadName;
                    worksheet.Cells[rowCount, 2].Value = datas[i].PqiValue;

                }
            }

            // 根据数据创建折线图
            MSExcel.ChartObjects chartObjects = worksheet.ChartObjects();
            MSExcel.ChartObject chartObject = chartObjects.Add(100, 100, 650, 293);
            // 200/72 = 2.78英寸。  2.78 * 2.54 = 7.0612厘米。
            //22.97   
            MSExcel.Chart chart = chartObject.Chart;


            MSExcel.Range dataRange = worksheet.Range["A1:B" + rowCount.ToString()];
            chart.SetSourceData(dataRange);

            // 设置第一系列为柱状图
            MSExcel.Series series1 = chart.FullSeriesCollection(1);

            series1.ChartType = MSExcel.XlChartType.xlColumnClustered;
            series1.AxisGroup = MSExcel.XlAxisGroup.xlPrimary;
            // 移除标题
            chart.HasTitle = false;

            // 移除右边的标注
            chart.HasLegend = false;
            chart.Axes(MSExcel.XlAxisType.xlValue).MaximumScale = 100;
            // 将图表复制到剪贴板
            chartObject.CopyPicture(MSExcel.XlPictureAppearance.xlScreen, MSExcel.XlCopyPictureFormat.xlBitmap);
            CWB_ExcelHelper.disposeExcel(ref excelApp);
        }
        private void getIndexStatisticsTable(ref MSExcel.Range dataRange, ref MSExcel.Application excelApp)
        {
            // 填充一些示例数据
            // 创建一个新的Excel应用程序实例
            excelApp = new MSExcel.Application();
            excelApp.Visible = true;
            // 新建一个工作簿
            MSExcel.Workbook workbook = excelApp.Workbooks.Add();
            MSExcel.Worksheet worksheet = workbook.ActiveSheet;
            worksheet.Cells[1, 1].Value = "路线编码";
            worksheet.Cells[1, 2].Value = "路线名称";
            worksheet.Cells[1, 3].Value = "起点桩号";
            worksheet.Cells[1, 4].Value = "终点桩号"; 
            worksheet.Cells[1, 5].Value = "路段长度";
            worksheet.Cells[1, 6].Value = "RQI";
            worksheet.Cells[1, 7].Value = "PQI";
            worksheet.Cells[1, 8].Value = "PCI";
            worksheet.Cells[1, 9].Value = "DR";
            worksheet.Cells[1, 10].Value = "IRI";
            int rowCount = 1;
            foreach (var item in allDatas)
            {
                rowCount++;
                worksheet.Cells[rowCount, 1].Value = item.RoadCode;
                worksheet.Cells[rowCount, 2].Value = item.RoadName;
                worksheet.Cells[rowCount, 3].Value = item.sMile;
                worksheet.Cells[rowCount, 4].Value = item.eMile; 
                worksheet.Cells[rowCount, 5].Value = item.RoadLen;
                worksheet.Cells[rowCount, 6].Value = item.RqiValue.ToString("0.##");
                worksheet.Cells[rowCount, 7].Value = item.PqiValue.ToString("0.##");
                worksheet.Cells[rowCount, 8].Value = item.PciValue.ToString("0.##");
                worksheet.Cells[rowCount, 9].Value = item.DrValue.ToString("0.##");
                worksheet.Cells[rowCount, 10].Value = item.IriValue.ToString("0.##");
            }
            dataRange = worksheet.Range["A1:J" + rowCount.ToString()];
        }
        #endregion
        public void Disposed()
        {

            //CollectBookTemp.Close(Type.Missing, Type.Missing, Type.Missing);

            //CWB_ExcelHelper.disposeExcel(ref tempApp);
            //CWB_ExcelHelper.disposeExcel(ref excelApp);
            //if (File.Exists(tempModuleExcelPath))
            //{
            //    File.Delete(tempModuleExcelPath);
            //}
        }

        private List<FileInfo> GetTargetDiseaseExcel(HeFeiRaod road)
        {
            string code = road.RoadCode;
            if (code.Length == 13)
            {
                code = code.Substring(0, 10);
            }
            var disExcels = excelPathList.Where(t => t.Name.Contains(code)).ToList();
            if (disExcels.Count > 0)
            {
                return disExcels;
            }
            else
            {
                return new List<FileInfo>();
            }


        }
        public void ReadExcelData(ref System.Data.DataTable dataTable, string filePath, int sheetNum, int startRow, bool hasHead = false)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    NPOI.SS.UserModel.IWorkbook workbook = new XSSFWorkbook(fs);
                    ISheet sheet = workbook.GetSheetAt(sheetNum);

                    // Get the header row
                    IRow headerRow = sheet.GetRow(0);
                    int cellCount = headerRow.LastCellNum;
                    // Create columns in DataTable
                    if (hasHead)
                    {
                        for (int i = 0; i < cellCount; i++)
                        {
                            DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                            dataTable.Columns.Add(column);
                        }
                    }
                    // Read data rows starting from the specified row
                   
                    for (int i = startRow; i <= sheet.LastRowNum; i++)
                    {
                        IRow dataRow = sheet.GetRow(i);
                        DataRow row = dataTable.NewRow();

                        
                        for (int j = 0; j < cellCount; j++)
                        {
                            try
                            {
                                ICell cell = dataRow.GetCell(j);
                                if (cell != null)
                                {
                                    if (cell.CellType == CellType.Formula)
                                    {
                                        //IFormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);
                                        IFormulaEvaluator evaluator = new XSSFFormulaEvaluator(workbook);
                                        evaluator.EvaluateInCell(dataRow.GetCell(j));
                                    }
                                    row[j] = cell.ToString();
                                }
                            }
                            catch (Exception ex)
                            {

                                throw ex;
                            }
                           
                        }
                        dataTable.Rows.Add(row);
                    }
                }
                catch (Exception ex )
                {

                    throw ex;
                }
               
            }
        }

        public void ReadExcelData(ref System.Data.DataTable dataTable, string filePath, string sheetName, int startRow, bool hasHead = false)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                NPOI.SS.UserModel.IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheet(sheetName);

                // Get the header row
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                // Create columns in DataTable
                if (hasHead)
                {
                    for (int i = 0; i < cellCount; i++)
                    {
                        DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                        dataTable.Columns.Add(column);
                    }
                }
                // Read data rows starting from the specified row
                for (int i = startRow; i <= sheet.LastRowNum; i++)
                {
                    IRow dataRow = sheet.GetRow(i);
                    DataRow row = dataTable.NewRow();

                    for (int j = 0; j < cellCount; j++)
                    {
                        if (dataRow.GetCell(j) != null)
                        {
                            row[j] = dataRow.GetCell(j).ToString();

                        }
                    }
                    dataTable.Rows.Add(row);
                }
            }
        }

        // private static float[] width_t1 = { 16.5f, 8.7f, 16.5f, 13.0f, 14.0f, 15.0f, 16.0f };
        private static float[] width_t1 = { 16.5f, 8.7f, 16.5f, 13.0f, 14.0f, 15.0f, 16.0f };

        private void AdjustTableFormatting(MSWord.Application wordApp, string StyleName, MSWord.Table table, bool isSetEveryCell)
        {
            // 设置表格样式
            table.set_Style(StyleName);

            // 设置表格标题
            //table.Title = tableTitle;

            // 设置列宽
            table.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100.0f;

            if (isSetEveryCell)
            {
                for (int i = 1; i <= table.Columns.Count; i++)
                {
                    if (i < width_t1.Length)
                    {
                        try
                        {
                            table.Columns[i].SetWidth(width_t1[i], Microsoft.Office.Interop.Word.WdRulerStyle.wdAdjustNone);
                        }
                        catch (Exception)
                        {

                        }
                    }

                    for (int j = 1; j <= table.Rows.Count; j++)
                    {
                        try
                        {
                            table.Cell(j, i).PreferredWidth = width_t1[width_t1.Length - 1];
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
            }
        }


        public static void FromatTable(MSWord.Application wordApp, MSWord.Table temptable, int headRowCnt, int realheadRowcnt, object headStayle, object oStyleName, float height, int wd_sleep_us, int colnum = 0, bool IsSetEveryCell = false, int roadnum = 0)
        {
            //wordApp.ScreenUpdating = false;
            temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitContent);
            temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            // 空格替换            
            object replaceAll = MSWord.WdReplace.wdReplaceAll;
            object[] Parameters = new object[15];
            Parameters[1] = Type.Missing;
            Parameters[2] = Type.Missing;
            Parameters[3] = Type.Missing;
            Parameters[4] = Type.Missing;
            Parameters[5] = Type.Missing;
            Parameters[6] = Type.Missing;
            Parameters[7] = Type.Missing;
            Parameters[8] = Type.Missing;
            Parameters[11] = Type.Missing;
            Parameters[12] = Type.Missing;
            Parameters[13] = Type.Missing;
            Parameters[14] = Type.Missing;
            Parameters[10] = replaceAll;
            object myFind = temptable.Range.Find;
            object findText = " ";
            object replaceText = "";
            Parameters[0] = findText;
            Parameters[9] = replaceText;
            myFind.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, myFind, Parameters);
            Thread.Sleep(wd_sleep_us);


            MSWord.Selection currentSelection = null;
            //  oStyleName = "报告表格内容（通用居中 小五）";
            if (!IsSetEveryCell)
            {
                temptable.Range.Select();
                currentSelection = wordApp.Selection;
                CWB_WordHelper.SetStyle(currentSelection, oStyleName, false, wd_sleep_us);
            }
            else
            {
                temptable.Range.Select();
                currentSelection = wordApp.Selection;
                CWB_WordHelper.SetStyle(currentSelection, oStyleName, false, wd_sleep_us);

                for (int i = 1; i < 25; i++)
                {
                    for (int t = 1; t <= headRowCnt; t++)
                    {
                        try
                        {
                            temptable.Cell(t, i).Range.set_Style(ref headStayle);
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }
            temptable.Range.Font.Name = "宋体"; // 将字体设置为宋体
            temptable.Range.Font.Size = 12; // 将字号设置为12磅（小四）
            #region 20240108注释


            //temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            //temptable.PreferredWidth = 100.0f;
            //temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            //temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            //temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            //temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            //temptable.AllowAutoFit = false;
            //temptable.LeftPadding = 0.0f;
            //temptable.RightPadding = 0.0f;
            //temptable.TopPadding = 0.0f;
            //temptable.BottomPadding = 0.0f;
            //temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            //temptable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            //temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            //temptable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            //temptable.Rows.AllowBreakAcrossPages = 0;
            //temptable.ApplyStyleHeadingRows = true;

            //height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            //temptable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);

            //wordApp.ScreenUpdating = true;
            #endregion
        }
    }
}
