/*-----------------------------------------------------------------
//CopyRight (C) 2012 武汉汉宁轨道交通技术有限公司
//版权所有。
//MyWordSzechwanDQ
//四川公路院报告
//
//
//创建标识:cwb 20220711
//修改标识:cwb 20220715 
//修改描述: 读取用户提供的excel表格  获取出报告需要的数据
//修改标识:cwb 20220720
//修改描述：根据客户需求调整表格列宽等细节
//修改标识:cwb 20220730
//修改描述：处理重复标题行bug
//修改标识：cwb20220810
//修改描述：处理图片插入问题（改变了插入顺序就行了）
 //------------------------------------------------------------------*/
#define 概况
#define 图片
#define 文字结果 
#define 表格
#define 附录表格
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using Framework.Office.Work;
using Framework.Other;
using Framework.Log;
using DevExpress.XtraExport;
using Framework.Office.Excel;
using System.Reflection;
using System.Windows.Forms;
using System.Data;

namespace XRDataProcess
{
    /// <summary>
    /// 四川道桥所
    /// word报告输出
    /// </summary>
    class MyWordSzechwanDQ
    {
        private static MyLogger log = new MyLogger(typeof(MyWordSzechwanDQ));
        private static MSExcel.Application excelApp = null;
        #region 概况
        /// <summary>
        /// 输出文件夹
        /// </summary>
        private string outPath = "";
        /// <summary>
        /// 报告名称
        /// </summary>
        private string reportName = ""; //S32西绵高速

        /// <summary>
        /// 工程概况
        /// </summary>
        private string projectProfileStr = ""; //S32西绵高速起于绵阳市三台县永明镇，与绵阳绕城高速公路相接，经绵阳市游仙区、三台县、盐亭县，西充县，止于南充市顺庆区同仁乡，与广南高速公路相接。路线全长124.358km，于2018年12月29日全线建成通车。


        /// <summary>
        /// 路线编码
        /// </summary>
        private string roadNumStr = "";
        /// <summary>
        /// 起点桩号
        /// </summary>
        private string startMileStr = "";
        private string endMileStr = "";

        /// <summary>
        ///  检测里程
        /// </summary>
        private string mileageStr = "";

        /// <summary>
        /// 方向
        /// </summary>
        private string directionStr = "";

        #endregion
        /// <summary>
        /// 保存所有需要导入的图片
        /// </summary>
        private Dictionary<string, MSExcel.Shape> pcisDic = null;

        /// <summary>
        ///  保存总的工程概况
        /// </summary>
        private Dictionary<string, string> itemInformationDic = null;
        /// <summary>
        ///   保存表格
        /// </summary>
        private Dictionary<string, MSExcel.Range> tablesFromExcel = null;
        /// <summary>
        /// 附录
        /// </summary>
        private Dictionary<string, MSExcel.Range> flTablesFromExcel = null;
        /// <summary>
        /// 保存需要导入的结果分析
        /// </summary>
        private Dictionary<string, string> resultStrDic = null;

        private static string baseModlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"报告模板\四川公路院");
        private static string wordModlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"报告模板\四川公路院", "2022报告模板.docx");
        private static string txtModlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"报告模板\四川公路院", "2022各路段项目概况汇总.txt");
        private static string exlceModlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"报告模板\四川公路院", "模板.xlsx");
        private MSWord.Selection currentSelection = null;
        private MSWord.Range wordrange = null;
        
        public MyWordSzechwanDQ(string sourcePath)
        { 
                pcisDic = new Dictionary<string, Microsoft.Office.Interop.Excel.Shape>();

                itemInformationDic = new Dictionary<string, string>();

                tablesFromExcel = new Dictionary<string, MSExcel.Range>();
                flTablesFromExcel = new Dictionary<string, MSExcel.Range>();
                resultStrDic = new Dictionary<string, string>();
                picNum = 1; intTableNum = 1; flTableCnt = 1;

                outPath = Path.GetDirectoryName(sourcePath);
                // DirectoryInfo info = new DirectoryInfo(sourcePath);
            

        }
        public void Disposed()
        {
           
            CollectBookTemp.Close(Type.Missing, Type.Missing, Type.Missing);
            
            CWB_ExcelHelper.disposeExcel(ref tempApp);
            CWB_ExcelHelper.disposeExcel( ref excelApp);
            if (File.Exists(tempModuleExcelPath))
            {
                File.Delete(tempModuleExcelPath);
            }
        }
     
        private MSWord.Application wordApp = null;
       public bool readModuleTxt()
        {
            if (!File.Exists(txtModlePath))
            {
                log.Error("模板文件不存在");
                return false;
            }
            if (itemInformationDic.Keys.Count == 0)
            {
                string[] str = File.ReadAllLines(txtModlePath, Encoding.UTF8);
                for (int i = 0; i < str.Length - 1; i += 2)
                {
                    string place = str[i];
                    string info = str[i + 1];
                    itemInformationDic.Add(place, info);
                }
                return true;
            }
            return true;
        }
        public void getTextInfoData()
        {
            if (itemInformationDic.Keys.Contains(reportName))
            {
                projectProfileStr = itemInformationDic[reportName];
            }

        }
        private string ceshu = null;
        public bool WriteWord(ProcessOperator p)
        {
            wordApp = new MSWord.Application() { Visible = true };
            MSWord.Document wordDoc = null;

            string outWordPath = Path.Combine(outPath, reportName + ".docx");
            if (File.Exists(outWordPath))
            {
                File.Delete(outWordPath);
            }
            string headerTxt = $"2022年度{this.reportName}路面技术状况检测报告";
            CWB_WordHelper.openWordApp(wordApp, wordModlePath, ref wordDoc);
            CWB_WordHelper.saveWord(wordDoc, outWordPath);
            //设置页头
            CWB_WordHelper.setPageHeader(wordDoc, wordApp, headerTxt, "2");

            WriteAllWordMarks(wordDoc);
            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            CWB_WordHelper.disposeWord(wordApp);
         
            return true;

        }
        /// <summary>
        /// 记录加入字典的图片 个数  从1开始  用来拼接字符串  找到对应的书签
        /// </summary>
        private int picNum = 1;
        /// <summary>
        /// 写入所有书签
        /// </summary>

        private void WriteAllWordMarks(MSWord.Document wordDoc)
        {

            int tableCnt = 1;
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {

                if (book.Name == "pci3_14_左幅")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;
                    while (true)
                    {
                        try
                        {
                            System.Windows.Forms.Clipboard.Clear();
                            pqiSheet.Shapes.Item(2).Copy();
                            Thread.Sleep(GlobalWord.wd_sleep_us*2);
                            currentSelection.PasteSpecial(Link: false, DataType: MSWord.WdPasteDataType.wdPasteEnhancedMetafile, Placement: MSWord.WdInlineShapeType.wdInlineShapeChart, DisplayAsIcon: false);
                            var shape = currentSelection.Range.ShapeRange;
                            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            shape.Width = (float)(14.70 * 0.3937008 * 72);
                            shape.Height = (float)(7.40 * 0.3937008 * 72);
                            shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapInline;
                            break;

                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us*2);
                        }
                    }

                    continue;

                }

            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
#if 概况

#region 概况

                if (book.Name.Contains("工程名称"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(reportName);
                    continue;
                }
                if (book.Name.Contains("工程概况"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(projectProfileStr);
                    continue;
                }
                if (book.Name.Contains("行车方向"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(directionStr);
                    continue;
                }
                if (book.Name.Contains("行车方向"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(directionStr);
                    continue;
                }
                if (book.Name.Contains("路线编码"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(roadNumStr);
                    continue;
                }
                if (book.Name.Contains("起点桩号"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(startMileStr);
                    continue;
                }
                if (book.Name.Contains("止点桩号"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(endMileStr);
                    continue;
                }
                if (book.Name.Contains("检测里程"))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(mileageStr);
                    continue;
                }
                if (book.Name == "main_table")
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    MSWord.Table table = currentSelection.Range.Tables[1];
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        if (table.Cell(i, 2).Range.Text.Contains(reportName))
                        {
                            table.Cell(i, 3).Range.Select();
                            currentSelection = wordApp.Selection;
                            currentSelection.Font.Color = MSWord.WdColor.wdColorRed;
                            currentSelection.TypeText("★");
                            ceshu = table.Cell(i, 1).Range.Text;
                        }
                    }
                    continue;
                }
#endregion
#endif

#if 图片
#region pic
                if (pcisDic.Keys.Contains(book.Name))
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    bool notfinished = true;
                    int pox = 0;
                    do
                    {
                        try
                        {
                          
                            System.Windows.Forms.Clipboard.Clear();
                            pcisDic[book.Name].Copy();
                            Thread.Sleep(GlobalWord.wd_sleep_us*2);
                            currentSelection.PasteSpecial(Link: false, DataType: MSWord.WdPasteDataType.wdPasteEnhancedMetafile, Placement: MSWord.WdInlineShapeType.wdInlineShapeChart, DisplayAsIcon: false);
                            var shape = currentSelection.Range.ShapeRange;
                            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            shape.Width = (float)(14.70 * 0.3937008 * 72);
                            shape.Height = (float)(7.40 * 0.3937008 * 72);
                            shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapInline;
                            notfinished = false;

                           Thread.Sleep(GlobalWord.wd_sleep_us  );
                            break;
                        }
                        catch (System.Exception ex)
                        {
                            notfinished = true;
                            Thread.Sleep(GlobalWord.wd_sleep_us * 2);
                            pox++;
                        }
                    } while (notfinished);
                    continue;
                }
#endregion
#endif
#if 文字结果

                if (resultStrDic.Keys.Contains(book.Name))
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    // currentSelection.Font.Color = MSWord.WdColor.wdColorYellow;
                    object oStyleName = "结果分析格式";
                    currentSelection.Font.Color = MSWord.WdColor.wdColorWhite;
                    currentSelection.set_Style(oStyleName);
                    currentSelection.TypeText(resultStrDic[book.Name]);
                    continue;

                }

#endif
#if 表格
                //其他表格
                if (tablesFromExcel.Keys.Contains(book.Name))
                {
                    book.Select();
                    currentSelection = wordApp.Selection;
                    CWB_WordHelper.PastExcelTable2Word(wordDoc, tablesFromExcel[book.Name], currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                    MSWord.Table tbl = currentSelection.Range.Tables[1];
                    CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);
                    //tbl.Delete();    
                    switch (book.Name)
                    {
                        case "tb_1":
                            FromatTable(wordApp, tbl, 0, 1, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 1);
                            break;
                        case "tb_2":
                        case "tb_4":
                        case "tb_7":
                        case "tb_5":

                        case "tb_8":
                            FromatTable(wordApp, tbl, 0, 2, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 2);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 2);
                            break;
                        case "tb_15":
                        case "tb_13":
                        case "tb_14":
                        case "tb_10":
                            FromatTable(wordApp, tbl, 0, 2, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 15);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 2);
                            break;
                            /*
                            FromatTable(wordApp, tbl, 0, 2, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 5);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 2);
                            break;
                            */
                        case "tb_6":
                            FromatTable(wordApp, tbl, 0, 1, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 6);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 1);
                            break;
                        case "tb_12":
                        case "tb_3":
                        case "tb_9":
                            tbl.Cell(1, 7).Range.Text = "等级\r\n评定";
                            FromatTable(wordApp, tbl, 0, 1, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 12);
                            tbl.set_Style("报告表格");
                            handelHead(wordApp, tbl, 1);
                            break;
                        case "tb_11":
                            FromatTable(wordApp, tbl, 0, 2, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true);
                            tbl.set_Style("报告表格"); handelHead(wordApp, tbl, 2);
                            break;
                        default:
                            FromatTable(wordApp, tbl, 0, 1, null, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true);
                            tbl.set_Style("报告表格"); handelHead(wordApp, tbl, 1);
                            break;
                    }
                    continue;
                }
#endif
                
#if 附录表格
                 //附录表格
                if (flTablesFromExcel.Keys.Contains(book.Name))
                {
                    book.Select();
                    currentSelection = wordApp.Selection;
                    CWB_WordHelper.PastExcelTable2Word(wordDoc, flTablesFromExcel[book.Name], currentSelection, GlobalWord.wd_sleep_us, ref tableCnt);
                    CWB_WordHelper.DeleteCurrentSelectionLine(wordApp);
                    MSWord.Table tbl = currentSelection.Range.Tables[1];
                    if (book.Name.Contains("fl_3"))
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 3, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 3); tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 3);
                    }
                    else if (book.Name.Contains("fl_1"))
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 2, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 1);
                        tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 2);
                    }
                    else if (book.Name.Contains("fl_2"))
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 2, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 2); tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 2);
                    }
                    else if (book.Name.Contains("fl_4"))
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 2, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 4); tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 2);
                    }
                    else if (book.Name.Contains("fl4_6"))
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 2, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true, colnum: 46); tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 2);
                    }

                    else
                    {
                        FromatTableFL(wordApp, tbl, "表格表头文字", 2, "表格文字2", 0.6f, GlobalWord.wd_sleep_us, IsSetEveryCell: true); tbl.set_Style("报告表格");
                        handelHead(wordApp, tbl, 2);
                    }
                    continue;
                }            

            foreach (var data in DataToExcel)
                {

                    if (book.Name.Contains(data.Key))
                    {
                        string[] strs = book.Name.Split('_');
                        int len = strs.Length;
                        int colIndex = int.Parse(strs[len - 2]);
                        int rowIndex = int.Parse(strs[len - 1]);

                        object[,] datas = data.Value;
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        string Stayle = "表格文字2";
                        currentSelection.Range.set_Style(Stayle);
                        currentSelection.TypeText(datas[colIndex, rowIndex].ToString());
                        continue;
                    }
                }
#endif
            }
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "pci3_14_全幅")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;
                    while (true)
                    {
                        try
                        {
                            System.Windows.Forms.Clipboard.Clear();
                            pqiSheet.Shapes.Item(3).Copy();
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                            currentSelection.PasteSpecial(Link: false, DataType: MSWord.WdPasteDataType.wdPasteEnhancedMetafile, Placement: MSWord.WdInlineShapeType.wdInlineShapeChart, DisplayAsIcon: false);
                            var shape = currentSelection.Range.ShapeRange;
                            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            shape.Width = (float)(14.70 * 0.3937008 * 72);
                            shape.Height = (float)(7.40 * 0.3937008 * 72);
                            shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapInline;

                            break;
                        }
                        catch
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                        }
                    }
                    continue;

                }

            }
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name.Contains("ceshu"))
                {
                    if ( !string.IsNullOrEmpty(ceshu))
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(ceshu.Substring(0, ceshu.Length - 2) + " " + reportName);
                        continue;
                    }
                    else
                    {
                        book.Range.Select();
                        currentSelection = wordApp.Selection;
                        currentSelection.TypeText(reportName);
                        continue;
                    }
                  
                }

            }
            System.Windows.Forms.Clipboard.Clear();
            CWB_WordHelper.ExportWordModel2(wordDoc);
        }
        private MSExcel.Worksheet pqiSheet = null;
        private string excelPathName = null;
        private MSExcel.Workbook CollectBookMain = null;

        private MSExcel.Worksheet pciSheet= null;
        private MSExcel.Worksheet rqiSheet= null;
        private MSExcel.Worksheet rdiSheet= null;
        private MSExcel.Worksheet pbiSheet = null;
        private MSExcel.Worksheet otherSheet = null;
        /// <summary>
        /// 读取excel 给所需数据赋值
        /// 
        /// </summary>
        public bool ReadExcel(string info,int index)
        {
            try
            {
                excelApp = new MSExcel.Application() { Visible = true, DisplayAlerts = false, AlertBeforeOverwriting = false };

                FileInfo excelSource = new FileInfo(info);
                reportName = excelSource.Name.Split('.').First();
                roadNumStr = OtherHelper.removeChineseLetter(reportName);
                excelPathName = excelSource.FullName;

                CollectBookMain = excelApp.Workbooks.Open(excelSource.FullName, Type.Missing,
                          true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                         Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Thread.Sleep(1000); //CollectBookMain.Save();
                pciSheet = CollectBookMain.Sheets["PCI"] as MSExcel.Worksheet;
                rqiSheet = CollectBookMain.Sheets["RQI"] as MSExcel.Worksheet;
                rdiSheet = CollectBookMain.Sheets["RDI"] as MSExcel.Worksheet;
                pbiSheet = CollectBookMain.Sheets["PBI"] as MSExcel.Worksheet;
                pqiSheet = CollectBookMain.Sheets["PQI"] as MSExcel.Worksheet;
                otherSheet = CollectBookMain.Sheets["历年指标衰减"] as MSExcel.Worksheet;

                getData(
                    pciSheet,
                    rqiSheet,
                    rdiSheet,
                    pbiSheet,
                    pqiSheet,
                    otherSheet,index
                    );
                //       p._backgroundWorker.ReportProgress(20);
                log.Info("获取excel内的数据成功");
                return true;
            }
            catch (Exception ex)
            {
                
                throw ;
            }
           
               
            //
        }
        /// <summary>
        /// 从excel 表格中获得数据
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="pciSheet"></param>
        /// <param name="rqiSheet"></param>
        /// <param name="rdiSheet"></param>
        /// <param name="pbiSheet"></param>
        /// <param name="pqiSheet"></param>
        /// <param name="otherSheet"></param>
        private void getData(

            MSExcel.Worksheet pciSheet,
            MSExcel.Worksheet rqiSheet,
            MSExcel.Worksheet rdiSheet,
            MSExcel.Worksheet pbiSheet,
            MSExcel.Worksheet pqiSheet,
            MSExcel.Worksheet otherSheet,int index)
        {
            initCommonData(pqiSheet);
#region 处理附录 使用模板合成
            //   flAddTableDic(pqiSheet, 4, 1, "A2:J");
            //  flAddTableDic(pciSheet, 4, 1, "A2:Q");

            handelFl1Table(pqiSheet,index);
            handelFl2Table(pciSheet);

            //处理附录3表格，需要合并几个表
            handelFl3Table(rqiSheet, rdiSheet, pbiSheet);

            handelFl4Table(pciSheet, rdiSheet, pqiSheet);
            handelFl5Table(pqiSheet);
            //附录6_1
            handelFl6_1Table(pciSheet, pqiSheet);
            //表4-5

            //表4-6
            handelFl7_1Table(pbiSheet);
            //附录6_2
            handel4_6(pqiSheet);
#endregion

            handlePCI(pciSheet, 3);
            handleRQI(rqiSheet, 4);
            handleRDI(rdiSheet, 5);
            handlePBI(pbiSheet, 3);
            handlePQI(pqiSheet, 1);
            handleOther(otherSheet, 5);

        }
        private string tempModuleExcelPath = null;
        private MSExcel.Application tempApp = null;
        private MSExcel.Workbook CollectBookTemp = null;
        /// <summary>
        /// 路线名称     桩号   检测方向
        /// </summary>
        private object[,] rqiLeftHeadDatas = null;
        private object[,] rqiRightHeadDatas = null;
        /// <summary>
        /// 左右幅的行数
        /// </summary>
        private int rightRowCountLength = 0;
        private int LeftRowCountLength = 0;
        /// <summary>
        /// 附录3需要将 pqi rdi pbi 三个表的  一部分拼起来
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <param name="t3"></param>
        private void handelFl1Table(MSExcel.Worksheet t1,int index)
        {
            object[,] rqiHeadDatas = null;
            int rightRowCount = 0;
            int leftRowCount = 0;
            object[,] leftTemp = null;  
            CWB_WordHelper.getExcelRangeData(t1, 4, 4, "A4:D", ref rqiHeadDatas, ref rightRowCount);
            CWB_WordHelper.getExcelRangeData(t1, 21, 4, "U4:X", ref leftTemp, ref leftRowCount);

            rightRowCountLength = rightRowCount;
            LeftRowCountLength = leftRowCount;
            rightRowCount -= 3;
            leftRowCount -= 3;
            object[,] rightDataHeade = new object[rightRowCount, 4];
            object[,] leftDataHeade = new object[leftRowCount, 4];

            for (int t = 1; t < 5; t++)
            {
                for (int i = 1; i <= rightRowCount; i++)
                {


                    rightDataHeade[i - 1, t - 1] = rqiHeadDatas[i, t];
                    
                    rightDataHeade[i - 1, 3] = "右幅";
                }
                for (int i = 1; i <= leftRowCount; i++)
                {


                  
                    leftDataHeade[i - 1, t - 1] = leftTemp[i, t];
                    leftDataHeade[i - 1, 3] = "左幅";
                }
            }
            rqiLeftHeadDatas = leftDataHeade;
            rqiRightHeadDatas = rightDataHeade;
            MSExcel.Range srcrange = null;
            object[,] rightsDatas = null;
            srcrange = t1.get_Range("E4:J" + (rightRowCount + 3).ToString());
            rightsDatas = (object[,])srcrange.Value2;
            object[,] leftsDatas = null;
            srcrange = t1.get_Range("Y4:AD" + (leftRowCount + 3).ToString());
            leftsDatas = (object[,])srcrange.Value2;


            object[,] allData = new object[leftRowCount+rightRowCount, 10];

            for (int i = 0; i < 10; i++)
            {
                for (int t = 0; t < leftRowCount + rightRowCount; t++)
                {
                    if (i < 4)
                    {


                        if (t < rightRowCount)
                        {
                            if (i == 1 || i == 2)
                            {
                                allData[t, i] = handleMileToStr(rightDataHeade[t, i]);
                            }
                            else
                                allData[t, i] = rightDataHeade[t, i];

                        }
                        else
                        {
                            if (i == 1 || i == 2)
                            {
                                allData[t, i] = handleMileToStr(leftDataHeade[t - rightRowCount, i]);
                            }
                            else
                                allData[t, i] = leftDataHeade[t - rightRowCount, i];

                        }
                    }
                    else
                    {

                        if (t < rightRowCount)
                        {
                            allData[t, i] = rightsDatas[t + 1, i + 1 - 4];

                        }
                        else
                        {
                            allData[t, i] = leftsDatas[t + 1 - rightRowCount, i + 1 - 4];

                        }

                    }

                }
            }
            try
            {
                if (File.Exists(Path.Combine(baseModlePath, $"temp{index}.xlsx")))
                {
                    File.Delete(Path.Combine(baseModlePath, $"temp{index}.xlsx"));
                }
            }
            catch (Exception)
            {

                
            }
         
            tempApp = new MSExcel.Application()

            {
                Visible = false,
                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                AlertBeforeOverwriting = false
            };

           
            CollectBookTemp = tempApp.Workbooks.Open(exlceModlePath, Type.Missing,
                                 true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                 Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            tempModuleExcelPath = Path.Combine(baseModlePath, $"temp{index}.xlsx");
            CollectBookTemp.SaveAs(tempModuleExcelPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CollectBookTemp.Save();
            MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录1"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:J{0}", rightRowCount+leftRowCount + 2));
            rangeTemp.Value2 = allData;
            rangeTemp = sheet.get_Range(string.Format("A1:J{0}", rightRowCount + leftRowCount + 2));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }
        /// <summary>
        /// 附录3需要将 pqi rdi pbi 三个表的  一部分拼起来
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <param name="t3"></param>
        private void handelFl2Table(MSExcel.Worksheet t2)
        {


            int rightRowNum = 0;
            int leftRowNum = 0;
            object[,] rightDataS = null;
            object[,] temp = null;
            CWB_WordHelper.getExcelRangeData(t2, 17, 4, "A4:Q", ref rightDataS, ref rightRowNum);
            CWB_WordHelper.getExcelRangeData(t2, 26, 4, "Z4:AP", ref temp, ref leftRowNum);
            MSExcel.Range srcrange = null;
            object[,] leftDatas = null;
            srcrange = t2.get_Range("Z4:AP" + leftRowNum.ToString());
            rightRowNum -= 3;
            leftRowNum -= 3;
            leftDatas = (object[,])srcrange.Value2;
            object[,] allDatas = new object[rightRowNum+leftRowNum, 17];

            for (int i = 0; i < 17; i++)
            {
                for (int t = 0; t < rightRowNum +leftRowNum; t++)
                {


                    if (t < rightRowNum)
                    {
                        if (i == 0 || i == 1)
                        {
                            allDatas[t, i] = handleMileToStr(rightDataS[t + 1, i + 1]);
                        }
                        else
                        {
                            allDatas[t, i] = rightDataS[t + 1, i + 1];
                        }


                    }
                    else
                    {
                        if (i == 0 || i == 1)
                        {
                            allDatas[t, i] = handleMileToStr(leftDatas[t + 1 - rightRowNum, i + 1]);
                        }
                        else
                        {
                            allDatas[t, i] = leftDatas[t + 1 - rightRowNum, i + 1];
                        }

                    }
                }
            }
            MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录2"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:Q{0}", rightRowNum + leftRowNum + 2));
            rangeTemp.Value2 = allDatas;
            rangeTemp = sheet.get_Range(string.Format("A1:Q{0}", rightRowNum + leftRowNum + 2));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }
        /// <summary>
        /// 附录3需要将 pqi rdi pbi 三个表的  一部分拼起来
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <param name="t3"></param>
        private void handelFl3Table(MSExcel.Worksheet t1, MSExcel.Worksheet t2, MSExcel.Worksheet t3)
        {

            object[,] srcobjPqi = null;
            int rightRowNum = 0;
            CWB_WordHelper.getExcelRangeData(t1, 15, 4, "O4:T", ref srcobjPqi, ref rightRowNum);

            MSExcel.Range srcrange = null;
            object[,] srcobjRdi = null;
            srcrange = t2.get_Range("V4:Z" + rightRowNum.ToString());
            srcobjRdi = (object[,])srcrange.Value2;
            object[,] srcobjPbi = null;
            srcrange = t3.get_Range("T4:X" + rightRowNum.ToString());
            srcobjPbi = (object[,])srcrange.Value2;

            rightRowNum -= 3;
            object[,] allData = new object[rightRowNum, 16];
            for (int i = 0; i < rightRowNum; i++)
            {
                for (int t = 0; t < 6; t++)
                {
                    if (t == 0 || t == 1)
                    {
                        allData[i, t] = handleMileToStr(srcobjPqi[i + 1, t + 1]);
                    }
                    else
                    {
                        allData[i, t] = srcobjPqi[i + 1, t + 1];

                    }
                }
                for (int d = 6; d < 11; d++)
                {
                    allData[i, d] = srcobjRdi[i + 1, d - 5];
                }
                for (int j = 11; j < 16; j++)
                {
                    allData[i, j] = srcobjPbi[i + 1, j - 10];
                }

            }
            int leftRowNum = 0;
            object[,] allData2 = handelFl3_1Table(t1, t2, t3,ref leftRowNum);

            object[,] allDatass = new object[rightRowNum + leftRowNum, 16];


            for (int i = 0; i < 16; i++)
            {
                for (int t = 0; t < rightRowNum + leftRowNum; t++)
                {
                    if (t < rightRowNum)
                    {
                        allDatass[t, i] = allData[t, i];
                    }
                    else
                    {
                        allDatass[t, i] = allData2[t - rightRowNum, i];
                    }
                }
            }


            MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录3"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A4:P{0}", (rightRowNum + leftRowNum) + 3));
            rangeTemp.Value2 = allDatass;
            rangeTemp = sheet.get_Range(string.Format("A1:P{0}", (rightRowNum + leftRowNum) + 3));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }
        /// <summary>
        /// rqi rdi pbi
        /// </summary>
        /// <param name="t1"></param>
        /// <param name="t2"></param>
        /// <param name="t3"></param>
        /// <returns></returns>
        private object[,] handelFl3_1Table(MSExcel.Worksheet t1, MSExcel.Worksheet t2, MSExcel.Worksheet t3,ref int num)
        {


            object[,] srcobjPqi = null;
            int leftRowNum = 0;
            CWB_WordHelper.getExcelRangeData(t1, 28, 4, "AB4:AG", ref srcobjPqi, ref leftRowNum);

            MSExcel.Range srcrange = null;
            object[,] srcobjRdi = null;
            srcrange = t2.get_Range("AL4:AP" + leftRowNum.ToString());
            srcobjRdi = (object[,])srcrange.Value2;
            object[,] srcobjPbi = null;
            srcrange = t3.get_Range("AI4:AM" + leftRowNum.ToString());
            srcobjPbi = (object[,])srcrange.Value2;

            leftRowNum -= 3;
            num = leftRowNum;
            object[,] allData = new object[leftRowNum, 16];
            for (int i = 0; i < leftRowNum; i++)
            {
                for (int t = 0; t < 6; t++)
                {
                    if (t == 0 || t == 1)
                    {
                        allData[i, t] = handleMileToStr(srcobjPqi[i + 1, t + 1]);
                    }
                    else
                    {
                        allData[i, t] = srcobjPqi[i + 1, t + 1];

                    }
                }
                for (int d = 6; d < 11; d++)
                {
                    allData[i, d] = srcobjRdi[i + 1, d - 5];
                }
                for (int j = 11; j < 16; j++)
                {
                    allData[i, j] = srcobjPbi[i + 1, j - 10];
                }

            }
            return allData;
        }
        /// <summary>
        /// 调整桩号格式
        /// </summary>
        /// <returns></returns>
        private string handleMileToStr(object o)
        {
           
              string s =   string.Format("{0:K0+000}", int.Parse(o.ToString()));
              return s;
           
           
        }

        /// <summary>
        /// pci rdi pqi
        /// </summary>
        /// <param name="t1">pci</param>
        /// <param name="t2"></param>
        /// <param name="t3">pqi</param>
        private void handelFl4Table(MSExcel.Worksheet t1, MSExcel.Worksheet t2, MSExcel.Worksheet t3)
        {


            //   private object[,] rqiHeadDatas = null;
            // private int lengthData = 0;
            int rightRowNum = rightRowCountLength;
            int leftRowNum = LeftRowCountLength;
            MSExcel.Range srcrange = null;
            object[,] objPci = null;
            srcrange = t1.get_Range("X4:X" + rightRowNum.ToString());
            objPci = (object[,])srcrange.Value2;
            object[,] objPci1 = null;
            srcrange = t1.get_Range("AW4:AW" +leftRowNum.ToString());
            objPci1 = (object[,])srcrange.Value2;


            object[,] objRdi = null;
            srcrange = t2.get_Range("AG4:AG" + rightRowNum.ToString());
            objRdi = (object[,])srcrange.Value2;
            object[,] objRdi1 = null;
            srcrange = t2.get_Range("AW4:AW" + leftRowNum.ToString());
            objRdi1 = (object[,])srcrange.Value2;

            object[,] allData = new object[rightRowNum + leftRowNum, 6];
            rightRowNum -= 3;
            leftRowNum -= 3;
            for (int i = 0; i < rightRowNum + leftRowNum; i++)
            {
                for (int t = 0; t < 4; t++)
                {
                    //桩号
                    if (t == 1 || t == 2)
                    {
                        if (i < rightRowNum)
                        {
                            allData[i, t] = handleMileToStr(rqiRightHeadDatas[i, t]);
                        }
                        else
                        {
                            allData[i, t] = handleMileToStr(rqiLeftHeadDatas[i - rightRowNum, t]);
                        }


                    }
                    else
                    {
                        if (i < rightRowNum)
                        {
                            allData[i, t] = rqiRightHeadDatas[i, t];
                        }
                        else
                        {
                            allData[i, t] = rqiLeftHeadDatas[i - rightRowNum, t];
                        }
                    }
                }
                for (int d = 4; d < 5; d++)
                {
                    if (i < rightRowNum)
                    {
                        allData[i, d] = objPci[i + 1, 1];
                    }
                    else
                    {
                        allData[i, d] = objPci1[i + 1 - rightRowNum, 1];
                    }

                }
                for (int j = 5; j < 6; j++)
                {
                    if (i < rightRowNum)
                    {
                        allData[i, j] = objRdi[i + 1, 1];
                    }
                    else
                    {
                        allData[i, j] = objRdi1[i + 1 - rightRowNum, 1];
                    }

                }
            }
            /*CollectBookTemp.SaveAs(Path.Combine(baseModlePath, "temp.xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
         */   MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录4"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:F{0}", (leftRowNum+rightRowNum) + 2));
            rangeTemp.Value2 = allData;
            rangeTemp = sheet.get_Range(string.Format("A1:F{0}", (leftRowNum + rightRowNum) + 2));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }

        private object[,] allData_fl5 = null;
        private object[,] allData_fl5_1 = null;
        private void handelFl5Table(MSExcel.Worksheet pqiSheet)
        {

            // rqiRightHeadDatas
            // rqiLeftHeadDatas
            int rightRowNum = 0;
            rightRowNum = rightRowCountLength;
            int leftRowNum = LeftRowCountLength; 
            MSExcel.Range srcrange = null;
            object[,] objPciRight = null;
            srcrange = pqiSheet.get_Range("E4:J" + rightRowNum.ToString());
            objPciRight = (object[,])srcrange.Value2;


            object[,] objPciLeft = null;
            srcrange = pqiSheet.get_Range("Y4:AD" + leftRowNum.ToString());
            objPciLeft = (object[,])srcrange.Value2;

            rightRowNum -= 3;
            leftRowNum -= 3;
            allData_fl5 = new object[rightRowNum, 10];
            for (int i = 0; i < rightRowNum; i++)
            {
                for (int t = 0; t < 10; t++)
                {
                    if (t < 4)
                    {
                        allData_fl5[i, t] = rqiRightHeadDatas[i, t];
                    }
                    else
                    {
                        allData_fl5[i, t] = objPciRight[i + 1, t + 1 - 4];
                    }

                }
            }

            allData_fl5_1 = new object[leftRowNum, 10];
            for (int i = 0; i <leftRowNum; i++)
            {
                for (int t = 0; t < 10; t++)
                {
                    if (t < 4)
                    {
                        allData_fl5_1[i, t] = rqiLeftHeadDatas[i, t];
                    }
                    else
                    {
                        allData_fl5_1[i, t] = objPciLeft[i + 1, t + 1 - 4];
                    }

                }
            }

            System.Collections.Generic.List<int> indexs = new List<int>();
            for (int i = 0; i < rightRowNum; i++)
            {

                if (Double.Parse(allData_fl5[i, 5].ToString()) < 80 || Double.Parse(allData_fl5[i, 6].ToString()) < 80 ||
                    Double.Parse(allData_fl5[i, 7].ToString()) < 80 || Double.Parse(allData_fl5[i, 8].ToString()) < 80)
                {
                    indexs.Add(i);
                }
            }
            System.Collections.Generic.List<int> indexs1 = new List<int>();
            for (int i = 0; i < leftRowNum; i++)
            {
                try
                {
                    if (Double.Parse(allData_fl5_1[i, 5].ToString()) < 80 || Double.Parse(allData_fl5_1[i, 6].ToString()) < 80 ||
                 Double.Parse(allData_fl5_1[i, 7].ToString()) < 80 || Double.Parse(allData_fl5_1[i, 8].ToString()) < 80)
                    {
                        indexs1.Add(i);
                    }
                }
               
                catch (Exception)
                {

                    continue;
                }
            }



            object[,] allData = new object[indexs.Count + indexs1.Count, 10];
            for (int i = 0; i < indexs.Count; i++)
            {
                for (int t = 0; t < 10; t++)
                {
                    //桩号
                    if (t == 1 || t == 2)
                    {
                        allData[i, t] = handleMileToStr(allData_fl5[indexs[i], t]);
                    }
                    else
                    {
                        allData[i, t] = allData_fl5[indexs[i], t];
                    }
                }
            }
            for (int i = indexs.Count; i < indexs1.Count + indexs.Count; i++)
            {
                for (int t = 0; t < 10; t++)
                {
                    //桩号
                    if (t == 1 || t == 2)
                    {
                        allData[i, t] = handleMileToStr(allData_fl5_1[indexs1[i - indexs.Count], t]);
                    }
                    else
                    {
                        allData[i, t] = allData_fl5_1[indexs1[i - indexs.Count], t];
                    }
                }
            }
            int rowCount =  indexs.Count + indexs1.Count;
            /*CollectBookTemp.SaveAs(Path.Combine(baseModlePath, "temp.xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
               */
            MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录5"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:J{0}", rowCount + 2));
            rangeTemp.Value2 = allData;
            rangeTemp = sheet.get_Range(string.Format("A1:J{0}", rowCount + 2));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }
        private void handelFl6_1Table(MSExcel.Worksheet pciSheet, MSExcel.Worksheet pqiSheet)
        {
            int rightRowCount = rightRowCountLength;
           
            //  private object[,] allData_fl5 = null;
            //  private object[,] allData_fl5_1 = null;
            MSExcel.Range srcrange = null;
            object[,] objPciRight = null;
            srcrange = pqiSheet.get_Range("P4:S" + rightRowCount.ToString());
            objPciRight = (object[,])srcrange.Value2;
            System.Collections.Generic.List<object[]> indexRights = new List<object[]>();
            for (int i = 1; i <= rightRowCount - 3; i++)
            {
                //任意一个为1 则需要记录 同时还要记录其等级
                if (double.Parse(objPciRight[i, 1].ToString()) != 0)
                {
                    indexRights.Add(new object[] { i - 1, "A级" });
                }
                else if (double.Parse(objPciRight[i, 2].ToString()) != 0)
                {
                    indexRights.Add(new object[] { i - 1, "B级" });
                }
                else if (double.Parse(objPciRight[i, 3].ToString()) != 0)
                {
                    indexRights.Add(new object[] { i - 1, "C级" });
                }
                else if (double.Parse(objPciRight[i, 4].ToString()) != 0)
                {
                    indexRights.Add(new object[] { i - 1, "D级" });
                }
            }

            object[,] dataRight = new object[indexRights.Count, 9];
            int index1 = 0;
            foreach (var item in indexRights)
            {
                for (int t = 0; t < rightRowCount; t++)
                {
                    if (t == int.Parse(item[0].ToString()))
                    {
                        for (int i = 0; i < 9; i++)
                        {
                            if (i < 8)
                            {
                                if (i == 0 || i == 1)
                                {
                                    dataRight[index1, i] = handleMileToStr(allData_fl5[t, i + 1]);

                                }
                                else
                                {
                                    dataRight[index1, i] = allData_fl5[t, i + 1];

                                }



                            }
                            if (i == 8)
                            {
                                dataRight[index1, i] = item[1].ToString();
                            }

                        }
                        index1++;
                    }

                }
            }
            index1 = 0;

            object[,] objPciLeft = null;
            int leftRowCount = LeftRowCountLength;
            srcrange = pqiSheet.get_Range("AJ4:AM" + leftRowCount.ToString());
            objPciLeft = (object[,])srcrange.Value2;
            System.Collections.Generic.List<object[]> indexLefts = new List<object[]>();
            for (int i = 1; i <= leftRowCount - 3; i++)
            {
                //任意一个为1 则需要记录 同时还要记录其等级
                if (double.Parse(objPciLeft[i, 1].ToString()) != 0)
                {
                    indexLefts.Add(new object[] { i - 1, "A级" });
                }
                else if (double.Parse(objPciLeft[i, 2].ToString()) != 0)
                {
                    indexLefts.Add(new object[] { i - 1, "B级" });
                }
                else if (double.Parse(objPciLeft[i, 3].ToString()) != 0)
                {
                    indexLefts.Add(new object[] { i - 1, "C级" });
                }
                else if (double.Parse(objPciLeft[i, 4].ToString()) != 0)
                {
                    indexLefts.Add(new object[] { i - 1, "D级" });
                }
            }

            object[,] dataLeft = new object[indexLefts.Count, 9];

            foreach (var item in indexLefts)
            {

                for (int t = 0; t < rightRowCount; t++)
                {
                    if (t == int.Parse(item[0].ToString()))
                    {
                        for (int i = 0; i < 9; i++)
                        {


                            if (i < 8)
                            {
                                if (i == 0 || i == 1)
                                {
                                    dataLeft[index1, i] = handleMileToStr(allData_fl5_1[t, i + 1]);

                                }
                                else
                                {
                                    dataLeft[index1, i] = allData_fl5_1[t, i + 1];

                                }
                            }
                            if (i == 8)
                            {
                                dataLeft[index1, i] = item[1].ToString();
                            }

                        }
                        index1++;

                    }

                }

            }

            rightRowCount = indexLefts.Count + indexRights.Count;
            object[,] allData = new object[rightRowCount, 9];
            for (int i = 0; i < rightRowCount; i++)
            {
                for (int t = 0; t < 9; t++)
                {
                    //if (i < indexLefts.Count)
                    //{
                    //    allData[i, t] = dataRight[i, t];
                    //}
                    //else
                    //{
                    //    allData[i, t] = dataLeft[i - indexLefts.Count, t];
                    //}
                    if (i < indexRights.Count)
                    {
                        allData[i, t] = dataRight[i, t];
                    }
                    else
                    {
                        allData[i, t] = dataLeft[i -indexRights.Count, t];
                    }

                }
            }




           /* CollectBookTemp.SaveAs(Path.Combine(baseModlePath, "temp.xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
           */ MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录6_1"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:I{0}", rightRowCount + 2));
            rangeTemp.Value2 = allData;
            rangeTemp = sheet.get_Range(string.Format("A1:I{0}", rightRowCount + 2));
            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
            flTableCnt++;
        }
        private int zhdu = 0;
        private int zhongdu = 0;
        private object[,] hadelFL7(MSExcel.Worksheet Sheet, string range, out int length)
        {
            object[,] srcobjPqi = null;
            int userownum = 0;
            CWB_WordHelper.getExcelRangeData(Sheet, 1, 4, range, ref srcobjPqi, ref userownum);

            System.Collections.Generic.List<int> indexs = new List<int>();
            userownum -= 3;
            for (int i = 1; i <= userownum; i++)
            {
                if (srcobjPqi[i,7]!=null)
                {
                    if (!srcobjPqi[i, 7].ToString().Equals("无跳车") && !srcobjPqi[i, 7].ToString().Equals("轻度跳车"))
                    {
                        indexs.Add(i);
                    }
                    if (srcobjPqi[i, 7].ToString().Equals("中度跳车"))
                    {
                        zhdu++;
                    }
                    if (srcobjPqi[i, 7].ToString().Equals("重度跳车"))
                    {
                        zhongdu++;
                    }

                }
            }
            object[,] allData = new object[indexs.Count, 8];
            for (int i = 0; i < indexs.Count; i++)
            {
                for (int t = 0; t < 8; t++)
                {

                    //桩号
                    if (t == 0 || t == 1)
                    {

                        //5+1 最多6列数据
                        allData[i, t] = handleMileToStr(srcobjPqi[int.Parse(indexs[i].ToString()), t + 1]);


                    }
                    else
                    {

                        if (t <= 6)
                        {
                            allData[i, t] = srcobjPqi[int.Parse(indexs[i].ToString()), t + 1];
                        }
                    }
                    allData[i, 7] = "调平处治";


                }
            }
            length = indexs.Count;
            return allData;
        }
        private string handelSpcData(object data, int weishu)
        {
            try
            {
                double eps = 1e-10;
                string s1 = data.ToString();
                double data1 = double.Parse(data.ToString());
                string result = data.ToString();
                if (data1 - Math.Floor(data1) < eps)
                {
                    //没有小数
                    int dataInt = int.Parse(data.ToString());
                    result = dataInt.ToString();
                }
                else
                {
                    result = data1.ToString("f" + weishu);
                }

                return result;



            }
            catch (Exception)
            {

                return data.ToString();
            }




        }
        /// <summary>
        /// 根据客户需要 保留小数
        /// </summary>
        /// <param name="data"></param>
        /// <param name="weishu"></param>
        /// <returns></returns>
        private string handelSpcData1(object data, int weishu)
        {
            try
            {

                double data1 = double.Parse(data.ToString());
                string result = "";


                result = data1.ToString("f" + weishu);


                return result;



            }
            catch (Exception)
            {

                return data.ToString();
            }




        }
        Dictionary<string, object[,]> DataToExcel = new Dictionary<string, object[,]>();
        private object[,] dataToTable = null;
        private void handel4_6(MSExcel.Worksheet Sheet)
        {
            object[,] srcobjPqi = null;
            int userownum = 0;
            CWB_WordHelper.getExcelRangeData(Sheet, 5, 6, "AZ6:BC", ref srcobjPqi, ref userownum);

            object[,] data = new object[2, 5];
            data[0, 0] = handelSpcData(srcobjPqi[1, 3], 3);
            data[0, 1] = handelSpcData(srcobjPqi[2, 3], 3);
            data[0, 2] = handelSpcData(srcobjPqi[3, 3], 3);
            data[0, 3] = handelSpcData(srcobjPqi[4, 3], 3);
            data[1, 0] = handelSpcData1(srcobjPqi[1, 4], 2);
            data[1, 1] = handelSpcData1(srcobjPqi[2, 4], 2);
            data[1, 2] = handelSpcData1(srcobjPqi[3, 4], 2);
            data[1, 3] = handelSpcData1(srcobjPqi[4, 4], 2);
            double sum1 = 0;
            for (int i = 0; i < 4; i++)
            {
                sum1 += double.Parse(data[0, i].ToString());
            }
            double sum2 = 0;
            for (int i = 0; i < 4; i++)
            {
                sum2 += double.Parse(data[1, i].ToString());
            }
            data[0, 4] = sum1;
            data[1, 4] = sum2;
            DataToExcel.Add("定dt4_5", data);
            string str = $"{reportName}“中次差路段”长度为{ data[0, 0]}km、“督办路段”长度为{data[0, 1]}km，“修复性养护路段”长度为{data[0, 2]}km，“预防性养护路段”长度为{data[0, 3]}km";
            resultStrDic.Add("str4_3_2", str);
            resultStrDic.Add("重度跳车个数", zhongdu.ToString());
            resultStrDic.Add("中度跳车个数", zhdu.ToString());


            MSExcel.Range srcrange = null;
            object[,] datas = null;
            srcrange = Sheet.get_Range("BB11:BC13" + userownum.ToString());
            datas = (object[,])srcrange.Value2;

            object[,] data1 = new object[2, 3];

            handelObjectS(data1, datas, 2, 3);
            data1[1, 0] = handelSpcData1(data1[1, 0], 2);
            data1[1, 1] = handelSpcData1(data1[1, 1], 2);
            data1[1, 2] = handelSpcData1(data1[1, 2], 2);
            DataToExcel.Add("定dt4_7", data1);

            string yanghuStr = $"其中，若仅对单指标中次差路段和督办路段（A+B级）进行养护处治，则需要养护的里程为{data1[0, 0]}km，养护比例为{data1[1, 0]}%；若对单指标中次差路段、督办路段和修复性养护路段（A+B+C级）进行养护处治，则需要养护的里程为{data1[0, 1]}km，养护比例为{data1[1, 1]}%；若对单指标中次差路段、督办路段、修复性养护路段及预防性养护路段（A+B+C+D级）进行养护处治，则需要养护的里程为{data1[0, 2]}km，养护比例为{data1[1, 2]}%。";
            resultStrDic.Add("养护策略文字", yanghuStr);
        }
        /// <summary>
        /// data2为源
        /// </summary>
        /// <param name="data1"></param>
        /// <param name="data2"></param>
        /// <param name="colLen"></param>
        /// <param name="rowLen"></param>
        private void handelObjectS(object[,] data1, object[,] data2, int colLen, int rowLen)
        {
            for (int i = 0; i < colLen; i++)
            {
                for (int row = 0; row < rowLen; row++)
                {
                    if (i == 0)
                    {
                        data1[i, row] = handelSpcData(data2[row + 1, i + 1], 3);
                    }
                    else
                    {
                        data1[i, row] = handelSpcData(data2[row + 1, i + 1], 2);
                    }

                }
            }
        }
        private void handelFl7_1Table(MSExcel.Worksheet Sheet)
        {  //  private object[,] allData_fl5 = null;
           //  private object[,] allData_fl5_1 = null;

            object[,] allData1 = null;
            object[,] allData2 = null;
            int userownum = 0;
            int length1;
            int length2;
            zhdu = 0;
            zhongdu = 0;
            allData1 = hadelFL7(Sheet, "A4:G", out length1);
            allData2 = hadelFL7(Sheet, "I4:O", out length2);
            userownum = length1 + length2;
            object[,] allData = new object[userownum, 8];
            for (int i = 0; i < userownum; i++)
            {
                for (int t = 0; t < 8; t++)
                {
                    if (i < length1)
                    {
                        allData[i, t] = allData1[i, t];
                    }
                    else
                    {
                        allData[i, t] = allData2[i - length1, t];
                    }
                    allData[i, 7] = "调平处治";
                }
            }



           /* CollectBookTemp.SaveAs(Path.Combine(baseModlePath, "temp.xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
               MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);*/
            MSExcel.Worksheet sheet = CollectBookTemp.Sheets["附录6_2"] as MSExcel.Worksheet;
            var rangeTemp = sheet.get_Range(string.Format("A3:H{0}", userownum + 2));
            rangeTemp.Value2 = allData;
            rangeTemp = sheet.get_Range(string.Format("A1:H{0}", userownum + 2));

           // CollectBook.Save();

            flTablesFromExcel.Add($"fl_{flTableCnt}", rangeTemp);
           flTablesFromExcel.Add($"fl4_6", rangeTemp);

            specialTable4_7 = rangeTemp;
            flTableCnt++;
           // CollectBook.Save();
        }
        private MSExcel.Range specialTable4_7 = null;
        private int flTableCnt = 1;
        private void flAddTableDic(MSExcel.Worksheet sheet, int row, int col, string rangStr)
        {

            MSExcel.Range r1 = CWB_WordHelper.getExcelRange(sheet, col, row, rangStr);
            flTablesFromExcel.Add($"fl_{flTableCnt}", r1);
            flTableCnt++;
        }

        /// <summary>
        /// 先初始化一下必要的数据  利用读取pqi表格
        /// </summary>
        /// <param name="sheet"></param>
        private void initCommonData(MSExcel.Worksheet sheet)
        {
            int startInt = 0;
            object[,] srcobj = null;
            int userownum = 0;
            CWB_WordHelper.getExcelRangeData(sheet, 1, 4, "A2:O", ref srcobj, ref userownum);


            if (srcobj[3, 2] != null)
            {
                //有效数据的第一行
                this.directionStr = srcobj[3, 4].ToString().Trim();
                startInt = int.Parse(srcobj[3, 2].ToString().Trim());

                this.startMileStr = startInt.ToString("K000+000");

            }
            //最后一行
            if (srcobj[userownum - 1, 3] != null)
            {
                int endMile = int.Parse(srcobj[userownum - 1, 3].ToString().Trim());
                double endmileDouble = endMile;
                this.endMileStr = endMile.ToString("K000+000");
                mileageStr = (Math.Abs(endMile - startInt) *0.001).ToString();

            }


        }

        private void handlePCI(MSExcel.Worksheet sheet, int picNum)
        {
            addPic2Dic(sheet, picNum);

            object[,] pciData = null;
            MSExcel.Range srcrange = sheet.get_Range("AZ5:BF8");
            // srcrange.Cells[1, 1] = "wocaocaosd";
            pciData = (object[,])srcrange.Value2;

            addTable2Dic(srcrange);
            MSExcel.Range srcrange1 = sheet.get_Range("AZ11:BI18");
            object[,] pciData1 = null;
            pciData1 = (object[,])srcrange1.Value2;
            string temp = $"根据检测结果，{reportName}右幅路面损坏状况指数（PCI）为{handelObject2Str(pciData[2, 6])}，评价等级为“{handelObject2Str(pciData[2, 7])}”，优良路率为{handelObject2Str(pciData1[4, 4])}%，次差路段率为{handelObject2Str(pciData1[4, 5])}%；左幅路面损坏状况指数（PCI）为{handelObject2Str(pciData[3, 6])}，评价等级为“{handelObject2Str(pciData[3, 7])}”，优良路率为{handelObject2Str(pciData1[6, 4])}%，次差路段率为{handelObject2Str(pciData1[6, 5])}%；全幅路面损坏状况指数（PCI）为{handelObject2Str(pciData[4, 6])}，评价等级为“{handelObject2Str(pciData[4, 7])}”，优良路率为{handelObject2Str(pciData1[8, 4])}%，次差路段率为{handelObject2Str(pciData1[8, 5])}%。";


            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("pciResult", temp);


            for (int i = 4; i < 11; i++)
            {
                handelPolish(srcrange1, pciData1, 3, i, "100.00");
            }

            addTable2Dic(srcrange1);

            // string temp = $"根据检测结果，{reportName}{directionStr}路面损坏状况指数（PCI）为{handelObject2Str(pciData[2, 6])}%，评价等级为“{handelObject2Str(pciData[2, 7])}。";

        }

        private void handelPolish(MSExcel.Range range, object[,] data, int startRow, int colIndex, string maxStr)
        {
            range.Cells[startRow, colIndex] = polishingNum(data[startRow, colIndex], maxStr);
        }

        private string polishingNum(object o, string maxStr)
        {

            string s = o.ToString();
            int num = s.Length;
            int lenNum = maxStr.Length;
            if (num < lenNum)
            {
                bool left = true;

                for (int i = 0; i < lenNum - num; i++)
                {
                    if (left)
                    {
                        s = "" + s;
                        left = false;
                    }
                    else
                    {
                        s += "";
                        left = true;
                    }

                }
            }
            return s;
        }
        private string handelObject2Str(object str)
        {
            double result = 0;
            bool isOk = double.TryParse(str.ToString(), out result);
            if (isOk)
            {
                return result.ToString("f2");
            }
            else
            {
                return str.ToString();
            }

        }
        private void handleRQI(MSExcel.Worksheet sheet, int picNum)
        {
            needChange = false;
            addPic2Dic(sheet, picNum);

            object[,] data1 = null;
            MSExcel.Range srcrange1 = sheet.get_Range("AO5:AW11");
            data1 = (object[,])srcrange1.Value2;
            addTable2Dic(srcrange1);

            MSExcel.Range srcrange2 = sheet.get_Range("AO14:AX21");

            object[,] data2 = (object[,])srcrange2.Value2;
            addTable2Dic(srcrange2);

            MSExcel.Range srcrange3 = sheet.get_Range("AO24:AX31");
            object[,] data3 = (object[,])srcrange3.Value2;
            addTable2Dic(srcrange3);


            //  string temp = $"根据检测结果，{reportName}{directionStr}路面行驶质量指数（RQI）为{handelObject2Str(data1[3, 6])}%，评价等级为“{handelObject2Str(data1[3, 9])}”。以每公里为评定单元，该段优良路率为{handelObject2Str(data2[4, 8])}%，次差路段比例{handelObject2Str(data2[4, 9])}%；以每10m为评定单元，该段优良路率为{handelObject2Str(data3[4, 8])}%，次差路率为{handelObject2Str(data3[4, 9])}%。";
            string temp = $"根据检测结果，{reportName}右幅路面行驶质量指数（RQI）为{handelObject2Str(data1[3, 6])}，评价等级为“{handelObject2Str(data1[3, 9])}”；左幅路面行驶质量指数（RQI）为{handelObject2Str(data1[5, 6])}，评价等级为“{handelObject2Str(data1[5, 9])}”；全幅路面行驶质量指数（RQI）为{handelObject2Str(data1[7, 6])}，评价等级为“{handelObject2Str(data1[7, 9])}”。\r\n以每公里为评定单元，该段右幅优良路率为{handelObject2Str(data2[4, 4])}%，次差路段率为{handelObject2Str(data2[4, 5])}%；左幅优良路率为{handelObject2Str(data2[6, 4])}%，次差路段率为{handelObject2Str(data2[6, 5])}%；全幅优良路率为{handelObject2Str(data2[8, 4])}%，次差路段率为{handelObject2Str(data2[8, 5])}%。以每10m为评定单元，该段右幅优良路率为{handelObject2Str(data3[4, 4])}%，次差路率为{handelObject2Str(data3[4, 5])}%；左幅优良路率为{handelObject2Str(data3[6, 4])}%，次差路率为{handelObject2Str(data3[6, 5])}%；全幅优良路率为{handelObject2Str(data3[8, 4])}%，次差路率为{handelObject2Str(data3[8, 5])}%。";
            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("rqiResult", temp);
        }
        private void handleRDI(MSExcel.Worksheet sheet, int picNum)
        {
            addPic2Dic(sheet, picNum);

            object[,] data1 = null;
            MSExcel.Range srcrange1 = sheet.get_Range("AY5:BG8");
            data1 = (object[,])srcrange1.Value2;
            addTable2Dic(srcrange1);

            MSExcel.Range srcrange2 = sheet.get_Range("AY11:BH18");

            object[,] data2 = (object[,])srcrange2.Value2;
            addTable2Dic(srcrange2);

            MSExcel.Range srcrange3 = sheet.get_Range("AY21:BH28");
            object[,] data3 = (object[,])srcrange3.Value2;
            addTable2Dic(srcrange3);


            //  string temp = $"根据检测结果，{reportName}{directionStr}路面车辙指数（RDI）为{handelObject2Str(data1[2, 8])}%，评价等级为“{handelObject2Str(data1[2, 9])}”。以每公里为评定单元，该段优良路率为{handelObject2Str(data2[4, 8])}%，次差路段率为{handelObject2Str(data2[4, 9])}%；以每10m为评定单元，该段优良路率为{handelObject2Str(data3[4, 8])}%，次差路率为{handelObject2Str(data3[4, 9])}%。";
            string temp = $"根据检测结果，{reportName}{directionStr}路面车辙指数（RDI）为{handelObject2Str(data1[2, 8])}，评价等级为“{handelObject2Str(data1[2, 9])}”；左幅路面车辙指数（RDI）为{handelObject2Str(data1[3, 8])}，评价等级为“{handelObject2Str(data1[3, 9])}”；全幅路面车辙指数（RDI）为{handelObject2Str(data1[4, 8])}，评价等级为“{handelObject2Str(data1[4, 9])}”。\r\n以每公里为评定单元，该段右幅优良路率为{handelObject2Str(data2[4, 4])}%，次差路率为{handelObject2Str(data2[4, 5])}%；左幅优良路率为{handelObject2Str(data2[6, 4])}%，次差路率为{handelObject2Str(data2[6, 5])}%；全幅优良路率为{handelObject2Str(data2[8, 4])}%，次差路率为{handelObject2Str(data2[8, 5])}%。以每10m为评定单元，该段右幅优良路率为{handelObject2Str(data3[4, 4])}%，次差路率为{handelObject2Str(data3[4, 5])}%；左幅优良路率为{handelObject2Str(data3[6, 4])}%，次差路率为{handelObject2Str(data3[6, 5])}%；全幅优良路率为{handelObject2Str(data3[8, 4])}%，次差路率为{handelObject2Str(data3[8, 5])}%。";

            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("rdiResult", temp);

        }
        private void handlePBI(MSExcel.Worksheet sheet, int picNum)
        {
            addPic2Dic(sheet, picNum);

            object[,] data1 = null;
            MSExcel.Range srcrange1 = sheet.get_Range("AU5:BC11");
            data1 = (object[,])srcrange1.Value2;
            addTable2Dic(srcrange1);

            MSExcel.Range srcrange2 = sheet.get_Range("AU14:BD21");

            object[,] data2 = (object[,])srcrange2.Value2;
            addTable2Dic(srcrange2);

            MSExcel.Range srcrange3 = sheet.get_Range("AU24:AZ29");
            object[,] data3 = (object[,])srcrange3.Value2;
            addTable2Dic(srcrange3);


            // string temp = $"根据检测结果，{reportName}{directionStr}路面跳车指数（PBI）为{handelObject2Str(data1[3, 6])}%，评价等级为“{handelObject2Str(data1[3, 9])}”，优良路率为{handelObject2Str(data2[4, 8])}%，次差路率为{handelObject2Str(data2[4, 9])}%。";
            string temp = $"根据检测结果，{reportName}右幅路面跳车指数（PBI）为{handelObject2Str(data1[3, 6])}，评价等级为“{handelObject2Str(data1[3, 9])}”，优良路率为{handelObject2Str(data2[4, 4])}%，次差路段率为{handelObject2Str(data2[4, 5])}%；左幅路面跳车指数（PBI）为{handelObject2Str(data1[5, 6])}，评价等级为“{handelObject2Str(data1[5, 9])}”，优良路率为{handelObject2Str(data2[6, 4])}%，次差路率为{handelObject2Str(data2[6, 5])}%；全幅路面跳车指数（PBI）为{handelObject2Str(data1[7, 6])}，评价等级为“{handelObject2Str(data1[7, 9])}”，优良路率为{handelObject2Str(data2[8, 4])}%，次差路率为{handelObject2Str(data2[8, 5])}%。";

            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("pbiResult", temp);

        }

        private void handlePQI(MSExcel.Worksheet sheet, int picNum)
        {
            addPic2Dic(sheet, picNum);
            object[,] data1 = null;
            MSExcel.Range srcrange1 = sheet.get_Range("AO5:AW20");
            data1 = (object[,])srcrange1.Value2;
            addTable2Dic(srcrange1);
            MSExcel.Range srcrange2 = sheet.get_Range("AO23:AX34");
            addTable2Dic(srcrange2);
            object[,] data2 = null;
            data2 = (object[,])srcrange2.Value2;
            MSExcel.Range srcrange3 = sheet.get_Range("AO38:AX49");
            addTable2Dic(srcrange3);
            object[,] data3 = null;
            data3 = (object[,])srcrange3.Value2;
            object[,] allData = new object[10, 7];
            // 处理上面两个表格的数据 sum/2

            for (int i = 0; i < 10; i++)
            {

                for (int t = 0; t < 7; t++)
                {
                    allData[i, t] = ((double.Parse(data2[i + 3, t + 4].ToString()) + double.Parse(data3[i + 3, t + 4].ToString())) / 2).ToString("f2");
                }
            }

            MSExcel.Worksheet sheetTemp = CollectBookTemp.Sheets["table3_15"] as MSExcel.Worksheet;
            var rangeTemp = sheetTemp.get_Range("C3:G12");
            rangeTemp.Value2 = allData;
            rangeTemp = sheetTemp.get_Range("A1:G12");
            //addTable2Dic(rangeTemp);
            MSExcel.Range srcrange4 = sheet.get_Range("AO53:AX64");
            addTable2Dic(srcrange4);
            string temp = $"根据检测结果，{reportName}右幅路面技术状况指数（PQI）为{handelObject2Str(data1[2, 6])}，评价等级为“{handelObject2Str(data1[2, 9])}”，优良路率为{handelObject2Str(data2[4, 4])}%，次差路段率为{handelObject2Str(data2[4, 5])}%；左幅路面技术状况指数（PQI）为{handelObject2Str(data1[7, 6])}，评价等级为“{handelObject2Str(data1[7, 9])}”，优良路率为{handelObject2Str(data3[4, 4])}%，次差路段率为{handelObject2Str(data3[4, 5])}%；全幅路面技术状况指数（PQI）为{handelObject2Str(data1[12, 6])}，评价等级为“{handelObject2Str(data1[12, 9])}”，优良路率为{allData[1, 0]}%，次差路段率为{allData[1, 1]}%。";
            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("pqiResult", temp);

           
        }
        private void handleOther(MSExcel.Worksheet sheet, int picNum)
        {

            addPic2Dic(sheet, picNum);
            object[,] data1 = null;
            MSExcel.Range srcrange1 = sheet.get_Range("B2:G7");
            addTable2Dic(srcrange1);
            data1 = (object[,])srcrange1.Value2;
            string temp = $"{reportName}连续5年路面技术状况检测结果显示：";

            //System.Windows.Forms.Clipboard.Clear();
            resultStrDic.Add("otherResult", temp);
        }
        //客户后面修改了模块插入了多张图片 
        bool needChange = true;
        /// <summary>
        /// 添加图片到列表
        /// </summary>
        /// <param name="page">当前小结要放入的图片张数(excel中需要存在)</param>
        private void addPic2Dic(MSExcel.Worksheet sheet, int page)
        {
            for (int i = 1; i < page + 1; i++)
            {
                //pci3_4 不需要插入
                if (picNum == 4)
                {
                    picNum++;
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);
                    picNum++;
                }
                //ccc   左幅
                else if (picNum == 7 && needChange)
                {
                    //客户后面修改了模块插入了多张图片 
                    string name = $"pci3_7_左幅";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);
                    needChange = false; picNum++;
                }
                else if (picNum == 7 && !needChange)
                {
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);

                    needChange = true;
                }
                else if (picNum == 10 && needChange)
                {
                    string name = $"pci3_10_左幅";
                    pcisDic.Add(name, sheet.Shapes.Item(i + 1) as Microsoft.Office.Interop.Excel.Shape);
                    needChange = false; picNum++;
                }
                else if (picNum == 10 && !needChange)
                {
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);

                    needChange = true;
                }
                else if (picNum == 11)
                {
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i - 1) as Microsoft.Office.Interop.Excel.Shape);
                    picNum++;
                }
                else if (picNum == 13 && needChange)
                {
                    string name = $"pci3_13_左幅";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);
                    needChange = false; picNum++;
                }
                else if (picNum == 13 && !needChange)
                {
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);

                    needChange = true;
                }
                else
                {
                    string name = $"pci3_{picNum}";
                    pcisDic.Add(name, sheet.Shapes.Item(i) as Microsoft.Office.Interop.Excel.Shape);
                    picNum++;
                }

            }

        }
        private int intTableNum = 1;
        private void addTable2Dic(MSExcel.Range srcrange)
        {
            tablesFromExcel.Add($"tb_{intTableNum}", srcrange);
            intTableNum++;
        }

        private static float[] width_FL1 = { 24.0f, 10.0f, 10.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f };
        private static float[] width_FL2 = { 8.0f, 8.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 6.0f, 5.6f, 5.0f };
        private static float[] width_FL4 = { 32.6f, 16.7f, 15.7f, 11.3f, 11.7f, 11.7f };
        private static float[] width_FL46 = {14.6f,14.6f,8.2f,12.1f,12.1f,12.1f,13.2f,13.2f };
        private static float[] width_FL3 = { 8.2f, 8.2f, 5.5f, 6.7f, 6.3f, 5.5f, 6.1f, 6.1f, 8.0f, 5.0f, 5.8f, 6.0f, 6.0f, 6.0f, 6.0f, 4.6f };
        //附录
        public void FromatTableFL(MSWord.Application wordApp, MSWord.Table temptable, object headStayle, int headRowCnt, object oStyleName, float height, int wd_sleep_us, int colnum = 0, bool IsSetEveryCell = false, int roadnum = 0)
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
            switch (colnum)
            {
                case 1:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_FL1[width_FL1.Length - 1];
                        for (int i = 0; i < width_FL1.Length; ++i)
                        {
                            // if (width_FL1[i] != width_FL1[width_FL1.Length - 1])
                            // {
                            // for (int j = headRowCnt+1; j <= rownum; ++j)
                            // {
                            try
                            {
                                temptable.Cell(headRowCnt + 1, i + 1).PreferredWidth = width_FL1[i];
                            }
                            catch (Exception)
                            {


                            }

                            // }
                            //}
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;

                case 2:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_FL2[width_FL2.Length - 1];
                        for (int i = 0; i < width_FL2.Length; ++i)
                        {
                            //if (width_FL2[i] != width_FL2[width_FL2.Length - 1])
                            // {
                            // for (int j = headRowCnt + 1; j <= rownum; ++j)
                            //   {
                            try
                            {
                                temptable.Cell(headRowCnt + 1, i + 1).PreferredWidth = width_FL2[i];
                            }
                            catch (Exception)
                            {


                            }

                            // }
                            //  }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 3:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_FL3[width_FL3.Length - 1];
                        for (int i = 0; i < width_FL3.Length; ++i)
                        {

                            try
                            {
                                temptable.Cell(headRowCnt + 1, i + 1).PreferredWidth = width_FL3[i];
                            }
                            catch (Exception)
                            {


                            }

                            // }
                            //}
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 4:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_FL4[width_FL4.Length - 1];
                        for (int i = 0; i < width_FL4.Length; ++i)
                        {

                            try
                            {
                                temptable.Cell(3, i + 1).PreferredWidth = width_FL4[i];
                            }
                            catch (Exception)
                            {


                            }

                            // }
                            //}
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 46:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_FL46[width_FL46.Length - 1];
                        for (int i = 1; i < width_FL46.Length; ++i)
                        {
                            if (width_FL46[i] != width_FL46[width_FL46.Length - 1])
                            {
                                for (int j = 1; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_FL46[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 100.0f;
                    }
                    break;
                default: break;
            }
            temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;


            temptable.AllowPageBreaks = true;
            temptable.LeftPadding = 0.0f;
            temptable.RightPadding = 0.0f;
            temptable.TopPadding = 0.0f;
            temptable.BottomPadding = 0.0f;

            temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;
            
            /*      try
                  {
                      for (int i = 1; i <= headRowCnt; i++)
                      {

                              temptable.Rows[i].HeadingFormat = -1;
                      }


                  }
                  catch
                  {
                      temptable.Cell(1, 1).Select();
                      GlobalWord.wordAppHeadingFormat(wordApp);


                  }*/
            temptable.Rows.AllowBreakAcrossPages = 0;
            temptable.ApplyStyleHeadingRows = true;

            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            temptable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            temptable.AllowAutoFit = false;  //自动调整以适应内容
                                             //根据窗口自动调整表格
            //wordApp.ScreenUpdating = true;
        }
        private static void handelHead(MSWord.Application wordApp, MSWord.Table temptable, int headRowCnt)
        {
            int rowsCount = temptable.Rows.Count; temptable.ApplyStyleHeadingRows = true;
            for (int i = 1; i <= rowsCount; i++)
            {
                for (int t = 1; t < 3; t++)
                {

                    try
                    {
                        temptable.Cell(i, t).Select();
                        if (i > headRowCnt)
                        {
                            wordApp.ActiveWindow.Selection.Rows.HeadingFormat = 0;


                        }
                        else
                        {
                            wordApp.ActiveWindow.Selection.Rows.HeadingFormat = -1;
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }
        private static float[] width_t1 = { 16.5f, 8.7f, 16.5f, 13.0f, 14.0f, 15.0f, 16.0f };
        private static float[] width_t2 = { 7.4f, 21.8f, 11.7f, 11.9f, 11.0f, 9.0f, 9.0f, 9.0f, 9.0f };
        private static float[] width_t15 = { 7.5f, 20.0f, 13.5f, 12.0f, 11.0f, 9.0f, 9.0f, 9.0f, 9.0f };
        private static float[] width_t3 = { 14.0f, 5.9f, 14.0f, 8.0f, 14.2f, 10.0f, 10.0f, 15f, 8.9f };
        private static float[] width_t12 = { 15.80f,4.20f, 15.90f,6.70f, 13.20f, 10.00f, 10.20f, 14.80f, 8.90f };
        private static float[] width_t12_1 = { 33.20f, 8.00f, 14.70f, 10.00f, 10.20f, 14.80f, 8.90f };
        private static float[] width_t5 = { 8.0f, 17.0f, 15.0f, 15.0f, 9.0f, 9.0f, 9.0f, 9.0f, 9.0f };
        private static float[] width_t6 = { 15.2f, 4.6f, 15.2f, 8.0f, 12.0f, 12.0f, 12.0f, 11.0f, 10.0f };

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
            switch (colnum)
            {
                case 1:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t1[width_t1.Length - 1];
                        for (int i = 1; i < width_t1.Length; ++i)
                        {
                            if (width_t1[i] != width_t1[width_t1.Length - 1])
                            {
                                for (int j = 1; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t1[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 2:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t2[width_t2.Length - 1];
                        for (int i = 0; i < width_t2.Length; ++i)
                        {
                            if (width_t2[i] != width_t2[width_t2.Length - 1])
                            {
                                for (int j = 3; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t2[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 3:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t3[width_t3.Length - 1];



                        for (int i = 1; i < width_t3.Length; ++i)
                        {
                            if (width_t3[i] != width_t3[width_t3.Length - 1])
                            {
                                for (int j = 2; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t3[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 12:
                    {


                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t12[width_t12.Length - 1];

                        for (int i = 0; i < width_t12_1.Length; ++i)
                        {
                            if (width_t12_1[i] != width_t12_1[width_t12_1.Length - 1])
                            {

                                try
                                {
                                    temptable.Cell(1, i + 1).PreferredWidth = width_t12_1[i];

                                }
                                catch (Exception)
                                {

                                }

                            }
                        }


                        for (int i = 0; i < width_t12.Length; ++i)
                        {
                            if (width_t12[i] != width_t12[width_t12.Length - 1])
                            {
                                for (int j = 2; j <= rownum; ++j)
                                {

                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t12[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }


                    }
                    break;
                case 15:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t15[width_t15.Length - 1];
                        for (int i = 0; i < width_t15.Length; ++i)
                        {
                            if (width_t15[i] != width_t15[width_t15.Length - 1])
                            {
                                for (int j = 3; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t15[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 5:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t5[width_t5.Length - 1];
                        for (int i = 0; i < width_t5.Length; ++i)
                        {
                            if (width_t5[i] != width_t5[width_t5.Length - 1])
                            {
                                for (int j = 3; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t5[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                case 6:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_t6[width_t6.Length - 1];
                        for (int i = 0; i < width_t6.Length; ++i)
                        {
                            if (width_t6[i] != width_t6[width_t6.Length - 1])
                            {
                                for (int j = 2; j <= rownum; ++j)
                                {
                                    try
                                    {
                                        temptable.Cell(j, i + 1).PreferredWidth = width_t6[i];

                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                     temptable.PreferredWidth = 100.0f;
                    }
                    break;
                default: break;
            }
            temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;


            temptable.AllowAutoFit = false;
            temptable.LeftPadding = 0.0f;
            temptable.RightPadding = 0.0f;
            temptable.TopPadding = 0.0f;
            temptable.BottomPadding = 0.0f;

            temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            /* try
             {
                 temptable.Rows[1].HeadingFormat = -1;

             }
             catch
             {
                 temptable.Cell(1, 1).Select(); temptable.ApplyStyleHeadingRows = true;
                 GlobalWord.wordAppHeadingFormat(wordApp);
             }*/
            temptable.Rows.AllowBreakAcrossPages = 0;
            temptable.ApplyStyleHeadingRows = true;

            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            temptable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);

            wordApp.ScreenUpdating = true;
        }

        public void wordAppHeadingFormat(MSWord.Application wordApp)
        {
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Rows.HeadingFormat = (int)MSWord.WdConstants.wdToggle;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }

    }
}
