using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.Threading;
using Framework.Log;
using System.Reflection;

namespace Framework.Office.Work
{
    /// <summary>
    /// CWB word帮助类
    /// </summary>
    /// 
    public class CWB_WordHelper
    {
        private static MyLogger log = new MyLogger(typeof(CWB_WordHelper));
        public CWB_WordHelper()
        {

        }
     
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wordApp"></param>
        /// <param name="temptable"></param>
        /// <param name="oStyleName"></param>
        /// <param name="height"></param>
        /// <param name="wd_sleep_us"></param>
        /// <param name="colnum"></param>
        /// <param name="IsSetEveryCell"></param>
        /// <param name="roadnum"></param>
        public static void FromatTable(MSWord.Application wordApp, MSWord.Table temptable, object oStyleName, float height, int wd_sleep_us, int colnum = 0, bool IsSetEveryCell = false, int roadnum = 0)
        {
            //wordApp.ScreenUpdating = false;

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
                SetStyle(currentSelection, oStyleName, false, wd_sleep_us);
            }
            else
            {
                for (int i = 1; i < 5; ++i)
                {
                    try
                    {
                        temptable.Cell(1, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception) { }
                }
                for (int i = 5; i < 9; ++i)
                {
                    try
                    {
                        temptable.Cell(1, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception  ) { }
                }
                for (int i = 5; i < 13; ++i)
                {
                    try
                    {
                        temptable.Cell(2, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception ) { }
                }
                for (int i = 1; i < 13; ++i)
                {
                    for (int j = 0; j < roadnum; ++j)
                    {
                        try
                        {
                            temptable.Cell(j + 3, i).Range.set_Style(ref oStyleName);
                        }
                        catch (Exception) { }
                    }
                }
                for (int i = 1; i < 10; ++i)
                {
                    try
                    {
                        temptable.Cell(roadnum + 3, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception) { }
                }
            }
            //switch (colnum)
            {
                //case 1:
                //    {
                //        int rownum = temptable.Rows.Count;
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_lq[width_lq.Length - 1];
                //        for (int i = 0; i < width_lq.Length; ++i)
                //        {
                //            if (width_lq[i] != width_lq[width_lq.Length - 1])
                //            {
                //                for (int j = 2; j <= rownum; ++j)
                //                {
                //                    temptable.Cell(j, i + 1).PreferredWidth = width_lq[i];
                //                }
                //            }
                //        }
                //        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.PreferredWidth = 105.0f;
                //    }
                //    break;
                //case 2:
                //    {
                //        int rownum = temptable.Rows.Count;
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_sn[width_sn.Length - 1];
                //        for (int i = 0; i < width_sn.Length; ++i)
                //        {
                //            if (width_sn[i] != width_sn[width_sn.Length - 1])
                //            {
                //                for (int j = 2; j <= rownum; ++j)
                //                {
                //                    temptable.Cell(j, i + 1).PreferredWidth = width_sn[i];
                //                }
                //            }
                //        }
                //        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.PreferredWidth = 105.0f;
                //    }
                //    break;
                //case 3:
                //    {
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_pci[width_pci.Length - 1];
                //        for (int i = 0; i < width_pci.Length; ++i)
                //        {
                //            if (width_pci[i] != width_pci[width_pci.Length - 1])
                //            {
                //                temptable.Columns[i + 1].PreferredWidth = width_pci[i];
                //            }
                //        }
                //        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.PreferredWidth = 105.0f;
                //    }
                //    break;
                //case 4:
                //    {
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_rqi[1];
                //        for (int i = 0; i < width_rqi.Length; ++i)
                //        {
                //            if (width_rqi[i] != width_rqi[1])
                //            {
                //                temptable.Columns[i + 1].PreferredWidth = width_rqi[i];
                //            }
                //        }
                //        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.PreferredWidth = 105.0f;
                //    }
                //    break;
                //case 5:
                //    {
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_mt[1];
                //        for (int i = 0; i < width_mt.Length; ++i)
                //        {
                //            if (width_mt[i] != width_mt[1])
                //            {
                //                temptable.Columns[i + 1].PreferredWidth = width_mt[i];
                //            }
                //        }
                //        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.PreferredWidth = 105.0f;
                //    }
                //    break;
                //case 6:
                //    {
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = 10;
                //    }
                //    break;
                //case 7:
                //    {
                //        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                //        temptable.Range.Cells.PreferredWidth = width_hz[1];
                //        for (int i = 0; i < 4; ++i)
                //        {
                //            try
                //            {
                //                temptable.Cell(1, i + 1).PreferredWidth = width_hz[i];
                //            }
                //            catch (Exception ex) { }
                //        }
                //        for (int i = 4; i < 12; ++i)
                //        {
                //            try
                //            {
                //                temptable.Cell(2, i + 1).PreferredWidth = width_hz[i];
                //            }
                //            catch (Exception ex) { }
                //        }
                //        for (int i = 0; i < roadnum; ++i)
                //        {
                //            for (int j = 0; j < 12; ++j)
                //            {
                //                try
                //                {
                //                    temptable.Cell(i + 2, j + 1).PreferredWidth = width_hz[j];
                //                }
                //                catch (Exception ex) { }
                //            }
                //        }
                //        for (int i = 0; i < 8; ++i)
                //        {
                //            try
                //            {
                //                temptable.Cell(roadnum + 3, i + 2).PreferredWidth = width_hz[i + 4];
                //            }
                //            catch (Exception ex) { }
                //        }
                //    }
                //    break;
                //default: break;
            }
            temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            temptable.AllowAutoFit = false;
            temptable.LeftPadding = 0.0f;
            temptable.RightPadding = 0.0f;
            temptable.TopPadding = 0.0f;
            temptable.BottomPadding = 0.0f;

            temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            temptable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                temptable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                temptable.Cell(1, 1).Select();
                wordAppHeadingFormat(wordApp);
            }
            temptable.Rows.AllowBreakAcrossPages = 0;
            temptable.ApplyStyleHeadingRows = true;

            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            temptable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);

            //wordApp.ScreenUpdating = true;
        }
 
        public static void ExportWordModel1( MSWord.Document wordDoc)
        {
         
            //更新目录
            int count = wordDoc.TablesOfContents.Count;
            for (int i = 0; i < count; i++)
            {
                wordDoc.TablesOfContents[i + 1].Update();
            }
        }
        public static void ExportWordModel2(MSWord.Document wordDoc)
        {

            //更新目录
            int count = wordDoc.TablesOfContents.Count;
            for (int i = 0; i < count; i++)
            {
                wordDoc.TablesOfContents[i + 1].UpdatePageNumbers();
            }
            
        }
        public static void wordAppHeadingFormat(MSWord.Application wordApp)
        {
            while (true)
            {
                try
                {
                    wordApp.ActiveWindow.Selection.Rows.HeadingFormat = -1;
                    break;
                }
                catch
                {
                    Thread.Sleep(200);
                }
            }
        }

        private static object missing = System.Reflection.Missing.Value;
      public static void DeleteCurrentSelectionLine(MSWord._Application application)
        {
           
            object wdLine = MSWord.WdUnits.wdLine;
            object wdCharacter = MSWord.WdUnits.wdCharacter;
            object wdExtend = MSWord.WdMovementType.wdExtend;
            object count = 0;
            MSWord.Selection selection = application.Selection;
           selection.HomeKey(ref wdLine, ref missing);
            selection.MoveDown(ref wdLine, ref count, ref wdExtend);
            selection.Delete(ref wdCharacter, ref missing);
        }
        public static void saveWord(MSWord.Document wordDoc, string outWordPath)
        {
            wordDoc.SaveAs(outWordPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        public static bool openWordApp(MSWord.Application wordApp, string wordModlePath, ref MSWord.Document wordDoc)
        {
            MSWord.Document doc = wordDoc = wordApp.Documents.Open(wordModlePath,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;
            wordDoc.ShowGrammaticalErrors = false;
            wordDoc.ShowRevisions = false;
            if (doc == null)
            {
                return false;

            }
            return true;
        }
        /// <summary>
        /// 模板word 设置页眉
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="wordApp"></param>
        /// <param name="headStr">需要插入的文字</param>
        /// <param name="PageNum">需要设置页眉的页数，从1开始，默认1</param>
        public static void setPageHeader(MSWord.Document wordDoc, MSWord.Application wordApp,string headStr, string PageNum = "1")
        {
            object What = MSWord.WdGoToItem.wdGoToPage;
            object Which = MSWord.WdGoToDirection.wdGoToNext;
            object Name = PageNum;
            wordDoc.ActiveWindow.Selection.GoTo(ref What, ref Which, ref Name);
            wordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
            var headerTxt = wordApp.Selection.HeaderFooter.Range.Text;
            headerTxt =headStr;

            wordApp.Selection.InsertAfter(headerTxt);
            wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//退出页眉设置
        }

        public  static void WriteTablePicStr(MSWord.Selection currentSelection, object oStyleName, string str, int sleep_us)
        {
            currentSelection.TypeText(str);
            SetStyle(currentSelection, oStyleName, false,sleep_us);
        }
       public static void SetStyle(MSWord.Selection currentSelection, object oStyleName, bool IsTypeParagraph,int sleep_us)
        {
            while (true)
            {
                Thread.Sleep(sleep_us);
                try
                {
                    currentSelection.set_Style(ref oStyleName);
                    break;
                }
                catch (Exception)
                {
                    Thread.Sleep(sleep_us);
                }
            }
            if (IsTypeParagraph)
            {
                Thread.Sleep(sleep_us);
                while (true)
                {
                    try
                    {
                        currentSelection.TypeParagraph();
                        break;
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(sleep_us);
                    }
                }
            }
        }
        public static void WriteText2Word(MSWord.Selection currentSelection, string str, int sleep_us)
        {
            Thread.Sleep(sleep_us);
            while (true)
            {
                try
                {
                    currentSelection.TypeText(str);
                    break;
                }
                catch (Exception)
                {
                    Thread.Sleep(sleep_us);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="srcrange">excel表格区域</param>
        /// <param name="currentSelection">插入位置</param>
        /// <param name="wordtablecnt"></param>
        /// <param name="tableheader">表头</param>
        /// <param name="oStyleName">文字格式</param>
        /// <param name="sleep_us">插入延时</param>
        /// <param name="IsGetTable"></param>
        /// <returns></returns>
        public static MSWord.Table PastExcelTable2Word(MSWord.Document wordDoc, MSExcel.Range srcrange, MSWord.Selection currentSelection,
       int sleep_us, ref int wordtablecnt, string tableheader = "", object oStyleName = null,  bool IsGetTable = true)
        {
        
            if (!string.IsNullOrEmpty( tableheader))
            {
                WriteText2Word(currentSelection, tableheader, sleep_us);

            }
            if (oStyleName!=null)
            {
                SetStyle(currentSelection, oStyleName, true, sleep_us);
            }
         
            

            while (true)
            {
             
                try
                {
                    System.Windows.Forms.Clipboard.Clear();
                   
                    srcrange.Copy();
                    Thread.Sleep(sleep_us);
                    currentSelection.PasteExcelTable(false, false, false);
                    ++wordtablecnt;
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(sleep_us);
                    log.Warn(ex.Message);
                }
            }

            //oStyleName = "报告表下空行";


           if (oStyleName!=null)
            SetStyle(currentSelection, oStyleName, true,sleep_us);

            MSWord.Table curtable = null;
            if (IsGetTable)
            {
                while (true)
                {
                    try
                    {
                        curtable = wordDoc.Tables[wordtablecnt];
                        break;
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(sleep_us);
                    }
                }
            }
            return curtable;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="col">计算列数</param>
        /// <param name="row">计算行数  </param>
        /// <param name="rangStr">想要获得的区域 如  "A2:O" </param>
        /// <param name="srcobj"></param>
        /// <param name="userownum">计算得到的数据域行数</param>
       public static void getExcelRangeData(MSExcel.Worksheet sheet, int col, int row, string rangStr, ref object[,] srcobj, ref int userownum)
        {
            userownum = judegeusedrow(sheet, col, row);
            MSExcel.Range srcrange = sheet.get_Range(rangStr + userownum.ToString());
            srcobj = (object[,])srcrange.Value2; //3.1之后为正式数据

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="col">计算行数 开始的行数</param>
        /// <param name="row">计算行数  开始的列数</param>
        /// <param name="rangStr">想要获得的区域 如  "A2:O"</param>
        /// <returns></returns>
        public static MSExcel.Range getExcelRange(MSExcel.Worksheet sheet, int col, int row, string rangStr)
        {
            int userownum = judegeusedrow(sheet, col, row);
            MSExcel.Range srcrange = sheet.get_Range(rangStr + userownum.ToString());
            return srcrange;
        }

        public static void  disposeWord<T>(T app)  where T: MSWord._Application 
        {
            
            if (app==null)
            {
                return;
            }
            try
            {
                app.Quit();
                int generation = System.GC.GetGeneration(app);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                app = default(T);

            }
            catch (Exception)
            {

                
            }
          
        }




        //根据指定列倒序
        public static void ReflectionColnum(MSExcel.Worksheet _Worksheet, MSExcel.Range srcrange, MSExcel.Range sortrange)
        {
            srcrange.Sort(sortrange, MSExcel.XlSortOrder.xlAscending,
                Type.Missing, Type.Missing,
                MSExcel.XlSortOrder.xlAscending,
                Type.Missing,
                MSExcel.XlSortOrder.xlAscending,
                MSExcel.XlYesNoGuess.xlNo,
                Type.Missing,
                Type.Missing,
                MSExcel.XlSortOrientation.xlSortColumns,
                MSExcel.XlSortMethod.xlPinYin,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal);
        }
        public static void ReflectionColnumDescending(MSExcel.Worksheet _Worksheet, MSExcel.Range srcrange, MSExcel.Range sortrange)
        {
            srcrange.Sort(sortrange, MSExcel.XlSortOrder.xlDescending,
                Type.Missing, Type.Missing,
                MSExcel.XlSortOrder.xlDescending,
                Type.Missing,
                MSExcel.XlSortOrder.xlDescending,
                MSExcel.XlYesNoGuess.xlNo,
                Type.Missing,
                Type.Missing,
                MSExcel.XlSortOrientation.xlSortColumns,
                MSExcel.XlSortMethod.xlPinYin,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal,
                MSExcel.XlSortDataOption.xlSortNormal);
        }
        //判断某列有数值的行数, Column从1开始
        public static int judegeusedrow(MSExcel.Worksheet worksheet, int Column, int startrow = 2)
        {
            int usedrowcnt = startrow - 1;
            MSExcel.Range trange = worksheet.get_Range(String.Format("{0}:{0}", GetCol((char)('A' + Column - 1))));
            object[,] tobj = (object[,])trange.Value2;
            for (int i = startrow; ; i++)
            {
                if (tobj[i, 1] != null)
                {
                    ++usedrowcnt;
                }
                else
                {
                    break;
                }
            }
            return usedrowcnt;
        }

        //判断某行有数值的列数
        public static int judegeusedcol(MSExcel.Worksheet worksheet, int Rownum, int startcol = 2)
        {
            int usedcolcnt = startcol - 1;
            MSExcel.Range trange = worksheet.get_Range(String.Format("{0}:{0}", Rownum));
            object[,] tobj = (object[,])trange.Value2;
            for (int i = startcol; ; i++)
            {
                if (tobj[1, i] != null)
                {
                    ++usedcolcnt;
                }
                else
                {
                    break;
                }
            }
            return usedcolcnt;
        }
        public static string GetCol(char instr)
        {
            string outstr;
            if (instr > 'Z')
            {
                instr = (char)(instr - 'A');
                outstr = ((char)(instr / 26 + 'A' - 1)).ToString() + ((char)(instr % 26 + 'A')).ToString();
            }
            else
            {
                outstr = instr.ToString();
            }
            return outstr;
        }
        public static void PastExcelPic2Word(MSWord.Selection currentSelection, MSExcel.Worksheet srcsheet, int PicIdx, string PicName, int wd_sleep_us=1000,bool IsTypeParagraph = true)
        {
            if (IsTypeParagraph)
            {
                currentSelection.TypeParagraph();
            }

            object oStyleName = "报告图与下段同页";
            currentSelection.set_Style(ref oStyleName);
            Thread.Sleep(wd_sleep_us);

            currentSelection.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            bool notfinished = true;
            do
            {
                try
                {
                    MSExcel.Shape excelshape = srcsheet.Shapes.Item(PicIdx) as Microsoft.Office.Interop.Excel.Shape;
                    System.Windows.Forms.Clipboard.Clear();
                    excelshape.Copy();
                    currentSelection.PasteAndFormat(MSWord.WdRecoveryType.wdChartPicture);
                    Thread.Sleep(wd_sleep_us);
                    notfinished = false;
                }
                catch (System.Exception )
                {
                    notfinished = true;
                }
            } while (notfinished);

            currentSelection.MoveRight();
            currentSelection.TypeParagraph();
            currentSelection.TypeText(PicName);
            oStyleName = "报告图标题3";
            currentSelection.set_Style(ref oStyleName);
            Thread.Sleep(wd_sleep_us);
        }
    }
}
