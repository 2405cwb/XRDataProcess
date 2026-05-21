using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OperateIniFile;
using MSWord = Microsoft.Office.Interop.Word;
using MSExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;

namespace XRDataProcess
{
    class MyWord
    {
        static XRSetting _Setting = XRSetting.GetInstance();

        private static bool _HasYHTable = false;

        private class DiseaseArea
        {
            public string _name;
            public double _area;
        }

        // 07标准，桂兴达-百米公里报告
        public static void OutputDoc(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            string srcdoc = string.Format(@"{0}\报告模板\路线报告模板_百米公里.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = excelpath.Replace(".xlsx",".docx");
            
            MSWord.Document wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbook1000 = excelApp.Workbooks.Open(excelpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbook100 = excelApp.Workbooks.Open(excelpath.Replace("_1000m.xlsx", "_100m.xlsx"),
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWordString(wordDoc, workbook1000);
            ExportWord(wordApp, excelApp, wordDoc, workbook1000, workbook100);
            if (_HasYHTable)
            {
                ExportWordString(wordDoc, workbook1000);
            }

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            workbook1000.Close(Type.Missing, Type.Missing, Type.Missing);
            workbook100.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void ExportWord(MSWord.Application wordApp, MSExcel.Application excelApp,
            MSWord.Document wordDoc, MSExcel.Workbook workbook1000, MSExcel.Workbook workbook100)
        {
            MSExcel.Worksheet excelsheet = null;
            MSExcel.Range excelrange = null;
            MSWord.Range wordrange = null;
            MSExcel.Shape excelshape = null;
            int excelrow = 0;
            string datastr;
            bool ishasrut = false;
            #region 表
            excelsheet = workbook1000.Sheets["分项指标统计表"] as MSExcel.Worksheet;
            datastr = ((MSExcel.Range)excelsheet.Cells[6, 2]).Value.ToString();
            try {
                if (Convert.ToDouble(datastr) > 0)
                { ishasrut = true; datastr = "A2:I6"; }
                else { datastr = "A2:I5"; }
            }
            catch { datastr = "A2:I5"; }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表1路面使用性能分项统计表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图1路面使用性能分项统计图");
            
            excelsheet = workbook1000.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            excelrow = GlobalExcel.judegeusedrow(excelsheet, 5, 2);
            if(ishasrut) { datastr = String.Format("A2:K{0}", excelrow); }
            else { datastr = String.Format("A2:J{0}", excelrow); }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表2技术状况明细表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图2技术状况明细图");

            if (_HasYHTable)
            {
                excelsheet = workbook1000.Sheets["养护需求建议表"] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 4, 2);
                excelrange = excelsheet.get_Range(String.Format("A2:G{0}", excelrow));
                excelrange.Copy();

                wordrange = GlobalWord.GetMarkRange(wordDoc, "表3养护建议表");
                GlobalWord.wordAppGoTo(wordApp, wordrange);
                GlobalWord.wordAppSize(wordApp, 12f);
                GlobalWord.wordAppTypeText(wordApp, "{年}年度{路线代码}{路线名}的路面大中修养护计划建议如下表。");
                GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                GlobalWord.wordAppFontBold(wordApp);
                GlobalWord.wordAppFontName(wordApp, "黑体");
                GlobalWord.wordAppTypeText(wordApp, "表4-1 {路线名}({路线代码})路面大中修养护建议表");
                GlobalWord.wordAppSelectionPaste(wordApp);
            }

            excelsheet = workbook100.Sheets["技术状况明细表"] as MSExcel.Worksheet;            
            excelshape = excelsheet.Shapes.Item(2) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图3破损率明细图");            
            excelshape = excelsheet.Shapes.Item(3) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图4平整度明细图");

            excelrow = GlobalExcel.judegeusedrow(excelsheet, 6, 2);
            excelrange = excelsheet.get_Range(String.Format("E2:N{0}", excelrow));
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表4检测数据明细表");

            #endregion
            
            #region 表格格式化
            int tidx = 0;
            int lasttab = wordDoc.Tables.Count;
            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                GlobalWord.wordAppGoTo(wordApp, temptable.Range);
                if (tidx++ == 0)
                    continue;
                if (tidx == lasttab)
                {
                    for (int i = 0; i < 6; ++i)
                    {
                        while (true)
                        {
                            try
                            {
                                temptable.Columns[2].Delete();
                                break;
                            }
                            catch
                            {
                                Thread.Sleep(200);
                            }
                        }
                    }
                }
                temptable.Range.Font.Name = "Times New Roman";
                temptable.Range.Font.Size = 10.5f;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;

                temptable.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAuto;
                try
                {
                    temptable.Rows[1].HeadingFormat = -1;
                }
                catch 
                {
                    temptable.Cell(1, 1).Select();
                    GlobalWord.wordAppHeadingFormat(wordApp);
                }
                temptable.Rows.AllowBreakAcrossPages = 0;
                temptable.ApplyStyleHeadingRows = true;

            }
            #endregion
        }

        #region 桂兴达 低等级农村路
        public static void _OutputDoc_low(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            string srcdoc = string.Format(@"{0}\报告模板\低等级农村路\路线报告模板_百米公里.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = excelpath.Replace(".xlsx", ".docx");

            MSWord.Document wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbook1000 = excelApp.Workbooks.Open(excelpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbook100 = excelApp.Workbooks.Open(excelpath.Replace("_1000m.xlsx", "_100m.xlsx"),
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWordString_low(wordDoc, workbook1000);
            _18ExportWord(wordApp, excelApp, wordDoc, workbook1000, workbook100);
            if (_HasYHTable)
            {
                ExportWordString_low(wordDoc, workbook1000);
            }

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            workbook1000.Close(Type.Missing, Type.Missing, Type.Missing);
            workbook100.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        // 18标准，桂兴达-十米公里报告
        public static void _18OutputDoc_low_10_1000(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            string srcdoc = string.Format(@"{0}\报告模板\低等级农村路\路线报告模板_十米公里.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = excelpath.Replace(".xlsx", ".docx");

            MSWord.Document wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbook1000 = excelApp.Workbooks.Open(excelpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbook10 = excelApp.Workbooks.Open(excelpath.Replace("_1000m.xlsx", "_10m.xlsx"),
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWordString(wordDoc, workbook1000);
            _18ExportWord_10_1000(wordApp, excelApp, wordDoc, workbook1000, workbook10);
            if (_HasYHTable)
            {
                ExportWordString(wordDoc, workbook1000);
            }

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            workbook1000.Close(Type.Missing, Type.Missing, Type.Missing);
            workbook10.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        #endregion
        // 18标准，桂兴达-百米公里报告
        public static void _18OutputDoc(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            string srcdoc = string.Format(@"{0}\报告模板\路线报告模板_百米公里.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = excelpath.Replace(".xlsx", ".docx");

            MSWord.Document wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbook1000 = excelApp.Workbooks.Open(excelpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbook100 = excelApp.Workbooks.Open(excelpath.Replace("_1000m.xlsx", "_100m.xlsx"),
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            ExportWordString(wordDoc, workbook1000);
            _18ExportWord(wordApp, excelApp, wordDoc, workbook1000, workbook100);
            if (_HasYHTable)
            {
                ExportWordString(wordDoc, workbook1000);
            }

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            workbook1000.Close(Type.Missing, Type.Missing, Type.Missing);
            workbook100.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void _18ExportWord(MSWord.Application wordApp, MSExcel.Application excelApp,
           MSWord.Document wordDoc, MSExcel.Workbook workbook1000, MSExcel.Workbook workbook100)
        {
            MSExcel.Worksheet excelsheet = null;
            MSExcel.Range excelrange = null;
            MSWord.Range wordrange = null;
            MSExcel.Shape excelshape = null;
            int excelrow = 0;
            string datastr;
            bool ishasrut = false;
            #region 表
            excelsheet = workbook1000.Sheets["分项指标统计表"] as MSExcel.Worksheet;
            datastr = ((MSExcel.Range)excelsheet.Cells[6, 2]).Value.ToString();
            try
            {
                if (Convert.ToDouble(datastr) > 0)
                { ishasrut = true; datastr = "A2:I6"; }
                else { datastr = "A2:I5"; }
            }
            catch { datastr = "A2:I6"; }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表1路面使用性能分项统计表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图1路面使用性能分项统计图");

            excelsheet = workbook1000.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            excelrow = GlobalExcel.judegeusedrow(excelsheet, 5, 4);
            if (ishasrut) { datastr = String.Format("A2:K{0}", excelrow); }
            else { datastr = String.Format("A2:J{0}", excelrow); }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表2技术状况明细表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图2技术状况明细图");

            if (_HasYHTable)
            {
                excelsheet = workbook1000.Sheets["养护需求建议表"] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 4, 3);
                excelrange = excelsheet.get_Range(String.Format("A2:G{0}", excelrow));
                excelrange.Copy();

                wordrange = GlobalWord.GetMarkRange(wordDoc, "表3养护建议表");
                GlobalWord.wordAppGoTo(wordApp, wordrange);
                GlobalWord.wordAppSize(wordApp, 12f);
                GlobalWord.wordAppTypeText(wordApp, "{年}年度{路线代码}{路线名}的路面大中修养护计划建议如下表。");
                GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                GlobalWord.wordAppFontBold(wordApp);
                GlobalWord.wordAppFontName(wordApp, "黑体");
                GlobalWord.wordAppTypeText(wordApp, "表4-1 {路线名}({路线代码})路面大中修养护建议表");
                GlobalWord.wordAppSelectionPaste(wordApp);
            }

            excelsheet = workbook100.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            excelshape = excelsheet.Shapes.Item(2) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图3破损率明细图");
            excelshape = excelsheet.Shapes.Item(3) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图4平整度明细图");

            excelrow = GlobalExcel.judegeusedrow(excelsheet, 6, 4);
            excelrange = excelsheet.get_Range(String.Format("E2:P{0}", excelrow));
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表4检测数据明细表");

            #endregion

            #region 表格格式化
            int tidx = 0;
            int lasttab = wordDoc.Tables.Count;
            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                GlobalWord.wordAppGoTo(wordApp, temptable.Range);
                if (tidx++ == 0)
                    continue;
                if (tidx == lasttab)
                {
                    for (int i = 0; i < 8; ++i)
                    {
                        int kk = 0;
                        while (true)
                        {
                            try
                            {
                                temptable.Columns[2].Delete();
                                break;
                            }
                            catch
                            {
                                Thread.Sleep(200);
                                kk++;
                                if (kk > 10) break;
                            }
                        }
                    }
                }
                temptable.Range.Font.Name = "Times New Roman";
                temptable.Range.Font.Size = 10.5f;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;

                temptable.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAuto;
                try
                {
                    temptable.Rows[1].HeadingFormat = -1;
                }
                catch
                {
                    temptable.Cell(1, 1).Select();
                    GlobalWord.wordAppHeadingFormat(wordApp);
                }
                temptable.Rows.AllowBreakAcrossPages = 0;
                temptable.ApplyStyleHeadingRows = true;

            }
            #endregion
        }
 
        // 18标准，桂兴达-十米公里报告
        public static void _18OutputDoc_10_1000(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            string srcdoc = string.Format(@"{0}\报告模板\路线报告模板_十米公里.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = excelpath.Replace(".xlsx", ".docx");

            MSWord.Document wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbook1000 = excelApp.Workbooks.Open(excelpath, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbook10 = excelApp.Workbooks.Open(excelpath.Replace("_1000m.xlsx", "_10m.xlsx"),
                Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWordString(wordDoc, workbook1000);
            _18ExportWord_10_1000(wordApp, excelApp, wordDoc, workbook1000, workbook10);
            if (_HasYHTable)
            {
                ExportWordString(wordDoc, workbook1000);
            }

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            workbook1000.Close(Type.Missing, Type.Missing, Type.Missing);
            workbook10.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
        private static void _18ExportWord_10_1000(MSWord.Application wordApp, MSExcel.Application excelApp,
           MSWord.Document wordDoc, MSExcel.Workbook workbook1000, MSExcel.Workbook workbook10)
        {
            MSExcel.Worksheet excelsheet = null;
            MSExcel.Range excelrange = null;
            MSWord.Range wordrange = null;
            MSExcel.Shape excelshape = null;
            int excelrow = 0;
            string datastr;
            bool ishasrut = false;
            #region 表
            excelsheet = workbook1000.Sheets["分项指标统计表"] as MSExcel.Worksheet;
            datastr = ((MSExcel.Range)excelsheet.Cells[6, 2]).Value.ToString();
            try
            {
                if (Convert.ToDouble(datastr) > 0)
                { ishasrut = true; datastr = "A2:I6"; }
                else { datastr = "A2:I5"; }
            }
            catch { datastr = "A2:I6"; }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表1路面使用性能分项统计表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图1路面使用性能分项统计图");

            excelsheet = workbook1000.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            excelrow = GlobalExcel.judegeusedrow(excelsheet, 5, 4);
            if (ishasrut) { datastr = String.Format("A2:K{0}", excelrow); }
            else { datastr = String.Format("A2:J{0}", excelrow); }
            excelrange = excelsheet.get_Range(datastr);
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表2技术状况明细表");
            excelshape = excelsheet.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图2技术状况明细图");

            if (_HasYHTable)
            {
                excelsheet = workbook1000.Sheets["养护需求建议表"] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 4, 3);
                excelrange = excelsheet.get_Range(String.Format("A2:G{0}", excelrow));
                excelrange.Copy();

                wordrange = GlobalWord.GetMarkRange(wordDoc, "表3养护建议表");
                GlobalWord.wordAppGoTo(wordApp, wordrange);
                GlobalWord.wordAppSize(wordApp, 12f);
                GlobalWord.wordAppTypeText(wordApp, "{年}年度{路线代码}{路线名}的路面大中修养护计划建议如下表。");
                GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                GlobalWord.wordAppFontBold(wordApp);
                GlobalWord.wordAppFontName(wordApp, "黑体");
                GlobalWord.wordAppTypeText(wordApp, "表4-1 {路线名}({路线代码})路面大中修养护建议表");
                GlobalWord.wordAppSelectionPaste(wordApp);
            }

            excelsheet = workbook10.Sheets["技术状况明细表"] as MSExcel.Worksheet;
            excelshape = excelsheet.Shapes.Item(2) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图3破损率明细图");
            excelshape = excelsheet.Shapes.Item(3) as Microsoft.Office.Interop.Excel.Shape;
            GlobalWord.PastExcel2Word(wordApp, excelshape, wordDoc, "图4平整度明细图");

            excelrow = GlobalExcel.judegeusedrow(excelsheet, 6, 4);
            excelrange = excelsheet.get_Range(String.Format("E2:P{0}", excelrow));
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表4检测数据明细表");

            excelsheet = workbook10.Sheets["病害列表"] as MSExcel.Worksheet;
            excelrow = GlobalExcel.judegeusedrow(excelsheet, 1, 3);
            if (_Setting.SelectDrawDis == 0)
                excelrange = excelsheet.get_Range(String.Format("A2:K{0}", excelrow));
            else
                excelrange = excelsheet.get_Range(String.Format("A2:F{0}", excelrow));
            GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表5路面病害明细表");

            bool IsLQExit = false;
            try
            {
                excelsheet = workbook10.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                IsLQExit = true;
            }
            catch 
            {
                IsLQExit = false;
            }
            if (IsLQExit)
            {
                if (_Setting.SelectDrawDis == 0)
                    excelrange = excelsheet.get_Range("A2:F25");
                else
                    excelrange = excelsheet.get_Range("A2:F16");
                GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表6沥青路面病害统计表");
            }

            bool IsSNExit = false;
            try
            {
                excelsheet = workbook10.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                IsSNExit = true;
            }
            catch
            {
                IsSNExit = false;
            }
            if (IsSNExit)
            {
                if (_Setting.SelectDrawDis == 0) 
                    excelrange = excelsheet.get_Range("A2:F24");
                else
                    excelrange = excelsheet.get_Range("A2:F15");
                GlobalWord.PastExcel2Word(wordApp, excelrange, wordDoc, "表6水泥路面病害统计表");
            }
            #endregion

            #region 表格格式化
            int tidx = 0;
            int lasttab = wordDoc.Tables.Count;
            if (IsSNExit && IsSNExit)
                lasttab = lasttab - 3;
            else
                lasttab = lasttab - 2;

            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                GlobalWord.wordAppGoTo(wordApp, temptable.Range);
                if (tidx++ == 0)
                    continue;
                if (tidx == lasttab)
                {
                    for (int i = 0; i < 8; ++i)
                    {
                        int kk = 0;
                        while (true)
                        {
                            try
                            {
                                temptable.Columns[2].Delete();
                                break;
                            }
                            catch
                            {
                                Thread.Sleep(200);
                                kk++;
                                if (kk > 10) break;
                            }
                        }
                    }
                }
                temptable.Range.Font.Name = "Times New Roman";
                temptable.Range.Font.Size = 10.5f;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                temptable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                temptable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;

                temptable.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAuto;
                try
                {
                    temptable.Rows[1].HeadingFormat = -1;
                }
                catch
                {
                    temptable.Cell(1, 1).Select();
                    GlobalWord.wordAppHeadingFormat(wordApp);
                }
                temptable.Rows.AllowBreakAcrossPages = 0;
                temptable.ApplyStyleHeadingRows = true;

            }
            #endregion
        }
        
        //简单的字段替换
        public static void ExportWordString(MSWord.Document wordDoc, MSExcel.Workbook workbook1000)
        {
            MSExcel.Worksheet worksheet = workbook1000.Sheets["路线信息表"] as MSExcel.Worksheet;
            MSExcel.Range workrange = worksheet.get_Range("B2:B20");
            object[,] prj = (object[,])workrange.Value2;
            _HasYHTable = false;

            Dictionary<string, string> datas = new Dictionary<string, string>();

            // 获取占比前三的病害类型
            bool IsExitSheet = false;
            object[,] objlq = null;
            object[,] objsn = null;
            try
            {
                worksheet = workbook1000.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                IsExitSheet = true;
            }
            catch
            {
                IsExitSheet = false;
            }
            if (IsExitSheet)
            {
                workrange = worksheet.get_Range("D29:E39");
                objlq = (object[,])workrange.Value2;
                if (objlq[11, 1] == null)
                {
                    workrange = worksheet.get_Range("D20:E30");
                    objlq = (object[,])workrange.Value2;
                }
            }

            IsExitSheet = false;
            try
            {
                worksheet = workbook1000.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                IsExitSheet = true;
            }
            catch
            {
                IsExitSheet = false;
            }
            if (IsExitSheet)
            {
                workrange = worksheet.get_Range("D28:E38");
                objsn = (object[,])workrange.Value2;
                if (objsn[11, 1] == null)
                {
                    workrange = worksheet.get_Range("D19:E29");
                    objsn = (object[,])workrange.Value2;
                    if (objsn[2, 2] == null)
                    {
                        objsn[2, 2] = ((MSExcel.Range)worksheet.Cells[5, 5]).Value;
                    }
                }
            }

            List<DiseaseArea> disarea = new List<DiseaseArea>();
            if (objlq != null)
            {
                for (int i = 0; i < 11; ++i)
                {
                    DiseaseArea tda = new DiseaseArea();
                    tda._name = objlq[i + 1, 1].ToString();
                    tda._area = Convert.ToDouble(objlq[i + 1, 2]);
                    disarea.Add(tda);
                }
            }
            if (objsn != null)
            {
                for (int i = 0; i < 11; ++i)
                {
                    DiseaseArea tda = new DiseaseArea();
                    tda._name = objsn[i + 1, 1].ToString();
                    tda._area = Convert.ToDouble(objsn[i + 1, 2]);
                    disarea.Add(tda);
                }
            }
            DiseaseArea[] disareaarr = disarea.ToArray();
            Array.Sort(disareaarr, delegate(DiseaseArea x, DiseaseArea y) { return y._area.CompareTo(x._area); });
            string maindis = disareaarr[0]._name;
            if (disareaarr[0]._area != 0)
            {
                if (disareaarr[1]._area != 0)
                {
                    maindis = maindis + "、" + disareaarr[1]._name;
                    if (disareaarr[2]._area != 0)
                    {
                        maindis = maindis + "、" + disareaarr[2]._name;
                    }
                }
                datas.Add("{主要病害}", maindis);
            }
            else
            {
                datas.Add("{路线代码}{路线名}的主要路面病害有{主要病害}。", "");
            } 

            if (prj[5, 1] == null) datas.Add("{路线代码}", "未知代码");
            else  datas.Add("{路线代码}", prj[5, 1].ToString());
            if (prj[6, 1] == null) datas.Add("{路线名}", "未知路线名");
            else datas.Add("{路线名}", prj[6, 1].ToString());
            datas.Add("{采集日期}", string.Format("{0:0000年00月00日}", prj[1, 1]));
            if (prj[3, 1] == null) datas.Add("{检测员}", "未知检测员");
            else datas.Add("{检测员}", prj[3, 1].ToString());
            datas.Add("{天气}", prj[4, 1].ToString());
            if (prj[7, 1] == null) datas.Add("{市}", "未知市");
            else datas.Add("{市}", prj[7, 1].ToString());
            if (prj[8, 1] == null) datas.Add("{管养单位}", "未知管养单位");
            else datas.Add("{管养单位}", prj[8, 1].ToString());
            string yanghu = "";
            string[] yanghutypes = { "大修", "中修", "预防性养护", "日常养护" };

            double tyhval = 0;
            int tmpi = 0;
            if (_Setting.YHType == 0 || _Setting.YHType == 2) tmpi = 2;
            else if (_Setting.YHType == 1) tmpi = 3;
            for (int i = 0; i < tmpi; i++)
            {
                tyhval = Math.Round(Convert.ToDouble(prj[i + 9, 1]) * 1000) * 0.001;
                if (tyhval != 0.0)
                {
                    yanghu += string.Format("{0}{1:0.000}公里，", yanghutypes[i], prj[i + 9, 1]);
                    _HasYHTable = true;
                }
            }
            tyhval = Math.Round(Convert.ToDouble(prj[12, 1]) * 1000) * 0.001;
            if (tyhval != 0.0)
            {
                yanghu += string.Format("日常养护{0}公里", tyhval);
            }
            else 
            {
                yanghu = yanghu.Remove(yanghu.Length - 1);
            }           

            datas.Add("{养护评价}", yanghu);
            datas.Add("{年}", prj[1, 1].ToString().Substring(0, 4));
            datas.Add("{实际检测里程}", prj[13, 1].ToString());
            datas.Add("{路面宽度}", prj[14, 1].ToString());
            datas.Add("{行车方向}", prj[15, 1].ToString());
            datas.Add("{整体路况}", prj[16, 1].ToString());
            datas.Add("{技术等级}", prj[17, 1].ToString());
            datas.Add("{路面类型}", prj[18, 1].ToString());
            datas.Add("{路线长度}", prj[18, 1].ToString());

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
            int kk = 0;
            while (true)
            {
                try
                {
                    foreach (var item in datas)
                    {
                        object myFind = wordDoc.Content.Find;
                        object findText = item.Key;
                        object replaceText = item.Value;

                        Parameters[0] = findText;
                        Parameters[9] = replaceText;
                        myFind.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, myFind, Parameters);
                    }
                    break;
                }
                catch
                {
                    Thread.Sleep(500);
                    ++kk;
                }
                if (kk > 10) break;          
            }
            Thread.Sleep(50);         
        }



        //简单的字段替换 
        /// <summary>
        /// 低等级农村路
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="workbook1000"></param>
        public static void ExportWordString_low(MSWord.Document wordDoc, MSExcel.Workbook workbook1000)
        {
            MSExcel.Worksheet worksheet = workbook1000.Sheets["路线信息表"] as MSExcel.Worksheet;
            MSExcel.Range workrange = worksheet.get_Range("B2:B20");
            object[,] prj = (object[,])workrange.Value2;
            _HasYHTable = false;

            Dictionary<string, string> datas = new Dictionary<string, string>();

            // 获取占比前三的病害类型
            bool IsExitSheet = false;
            object[,] objlq = null;
            object[,] objsn = null;
            try
            {
                worksheet = workbook1000.Sheets["沥青病害统计表"] as MSExcel.Worksheet;
                IsExitSheet = true;
            }
            catch
            {
                IsExitSheet = false;
            }
            if (IsExitSheet)
            {
                workrange = worksheet.get_Range("C23:D27");
                objlq = (object[,])workrange.Value2;
            }

            IsExitSheet = false;
            try
            {
                worksheet = workbook1000.Sheets["水泥病害统计表"] as MSExcel.Worksheet;
                IsExitSheet = true;
            }
            catch
            {
                IsExitSheet = false;
            }
            if (IsExitSheet)
            {
                workrange = worksheet.get_Range("C22:D27");
                objsn = (object[,])workrange.Value2;
            }

            List<DiseaseArea> disarea = new List<DiseaseArea>();
            if (objlq != null)
            {
                for (int i = 0; i < 5; ++i)
                {
                    DiseaseArea tda = new DiseaseArea();
                    tda._name = objlq[i + 1, 1].ToString();
                    tda._area = Convert.ToDouble(objlq[i + 1, 2]);
                    disarea.Add(tda);
                }
            }
            if (objsn != null)
            {
                for (int i = 0; i < 6; ++i)
                {
                    DiseaseArea tda = new DiseaseArea();
                    tda._name = objsn[i + 1, 1].ToString();
                    tda._area = Convert.ToDouble(objsn[i + 1, 2]);
                    disarea.Add(tda);
                }
            }
            DiseaseArea[] disareaarr = disarea.ToArray();
            Array.Sort(disareaarr, delegate (DiseaseArea x, DiseaseArea y) { return y._area.CompareTo(x._area); });
            string maindis = disareaarr[0]._name;
            if (disareaarr[0]._area != 0)
            {
                if (disareaarr[1]._area != 0)
                {
                    maindis = maindis + "、" + disareaarr[1]._name;
                    if (disareaarr[2]._area != 0)
                    {
                        maindis = maindis + "、" + disareaarr[2]._name;
                    }
                }
                datas.Add("{主要病害}", maindis);
            }
            else
            {
                datas.Add("{路线代码}{路线名}的主要路面病害有{主要病害}。", "");
            }

            if (prj[5, 1] == null) datas.Add("{路线代码}", "未知代码");
            else datas.Add("{路线代码}", prj[5, 1].ToString());
            if (prj[6, 1] == null) datas.Add("{路线名}", "未知路线名");
            else datas.Add("{路线名}", prj[6, 1].ToString());
            datas.Add("{采集日期}", string.Format("{0:0000年00月00日}", prj[1, 1]));
            if (prj[3, 1] == null) datas.Add("{检测员}", "未知检测员");
            else datas.Add("{检测员}", prj[3, 1].ToString());
            datas.Add("{天气}", prj[4, 1].ToString());
            if (prj[7, 1] == null) datas.Add("{市}", "未知市");
            else datas.Add("{市}", prj[7, 1].ToString());
            if (prj[8, 1] == null) datas.Add("{管养单位}", "未知管养单位");
            else datas.Add("{管养单位}", prj[8, 1].ToString());
            string yanghu = "";
            string[] yanghutypes = { "大修", "中修", "预防性养护", "日常养护" };

            double tyhval = 0;
            int tmpi = 0;
            if (_Setting.YHType == 0 || _Setting.YHType == 2) tmpi = 2;
            else if (_Setting.YHType == 1) tmpi = 3;
            for (int i = 0; i < tmpi; i++)
            {
                tyhval = Math.Round(Convert.ToDouble(prj[i + 9, 1]) * 1000) * 0.001;
                if (tyhval != 0.0)
                {
                    yanghu += string.Format("{0}{1:0.000}公里，", yanghutypes[i], prj[i + 9, 1]);
                    _HasYHTable = true;
                }
            }
            tyhval = Math.Round(Convert.ToDouble(prj[12, 1]) * 1000) * 0.001;
            if (tyhval != 0.0)
            {
                yanghu += string.Format("日常养护{0}公里", tyhval);
            }
            else
            {
                yanghu = yanghu.Remove(yanghu.Length - 1);
            }

            datas.Add("{养护评价}", yanghu);
            datas.Add("{年}", prj[1, 1].ToString().Substring(0, 4));
            datas.Add("{实际检测里程}", prj[13, 1].ToString());
            datas.Add("{路面宽度}", prj[14, 1].ToString());
            datas.Add("{行车方向}", prj[15, 1].ToString());
            datas.Add("{整体路况}", prj[16, 1].ToString());
            datas.Add("{技术等级}", prj[17, 1].ToString());
            datas.Add("{路面类型}", prj[18, 1].ToString());
            //datas.Add("{路线长度}", prj[18, 1].ToString());
            datas.Add("{路线长度}", prj[13, 1].ToString());

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
            int kk = 0;
            while (true)
            {
                try
                {
                    foreach (var item in datas)
                    {
                        object myFind = wordDoc.Content.Find;
                        object findText = item.Key;
                        object replaceText = item.Value;

                        Parameters[0] = findText;
                        Parameters[9] = replaceText;
                        myFind.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, myFind, Parameters);
                    }
                    break;
                }
                catch
                {
                    Thread.Sleep(500);
                    ++kk;
                }
                if (kk > 10) break;
            }
            Thread.Sleep(50);
        }

        //中南安环双车道报表合并
        public static void OutputZNExcel(MSExcel._Application excelApp, string leftpath, string rightpath, string destpath)
        {
            string srcxls = string.Format(@"{0}\报表模板\等级公路 JTG H20-2007\中南安环\两车道检测记录表.xlsx",
               System.Windows.Forms.Application.StartupPath);
            FileInfo tfile = new FileInfo(leftpath);
            string prjname = " ";
            string[] tstr = tfile.Name.Split(new char[] { '_', '.' });
            string destxls;
            try
            {
                prjname = string.Format("{0}_{1}", tstr[0], tstr[1]);
                destxls = string.Format("{0}_检测记录表{1}.xlsx", prjname, tstr[tstr.Length - 2]);
            }
            catch 
            {
                destxls = "检测记录表.xlsx";
            }
            destxls = string.Format("{0}\\{1}", destpath, destxls);

            MSExcel.Workbook workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Workbook workbookleft = excelApp.Workbooks.Open(leftpath, Type.Missing,
                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Workbook workbookright = excelApp.Workbooks.Open(rightpath,
                Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            MergeExcel(workbookleft, workbookright, workbook, prjname);

            workbook.Save();
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            workbookleft.Close(Type.Missing, Type.Missing, Type.Missing);
            workbookright.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        public static void MergeExcel(MSExcel.Workbook workbookleft, MSExcel.Workbook workbookright,
            MSExcel.Workbook workbook, string prjname)
        {
            MSExcel.Worksheet sheetleft = workbookleft.Sheets["汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetright = workbookright.Sheets["汇总表"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetIRI = workbook.Sheets["平整度"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetRUT = workbook.Sheets["车辙"] as MSExcel.Worksheet;
            MSExcel.Worksheet sheetDR = workbook.Sheets["破损率"] as MSExcel.Worksheet;

            int leftuserow, rightuserow, userow;
            leftuserow = GlobalExcel.judegeusedrow(sheetleft, 1, 4) - 1;
            rightuserow = GlobalExcel.judegeusedrow(sheetright, 1, 4) - 1;
            userow = leftuserow < rightuserow?leftuserow:rightuserow;

            MSExcel.Range srcrange, destrange;
            srcrange = sheetleft.get_Range(string.Format("A4:B{0}", userow));

            //桩号
            destrange = sheetIRI.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range("A10");
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range("A10");
            srcrange.Copy(destrange);

            //平整度
            sheetIRI.Cells[7, 3] = prjname;
            srcrange = sheetleft.get_Range(string.Format("E4:F{0}", userow));
            destrange = sheetIRI.get_Range("C10");
            srcrange.Copy(destrange);
            srcrange = sheetright.get_Range(string.Format("E4:F{0}", userow));
            destrange = sheetIRI.get_Range("E10");
            srcrange.Copy(destrange);
            destrange = sheetIRI.get_Range(string.Format("G10:G{0}",userow+6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //车辙
            sheetRUT.Cells[7, 3] = prjname;
            srcrange = sheetleft.get_Range(string.Format("I4:J{0}", userow));
            destrange = sheetRUT.get_Range("C10");
            srcrange.Copy(destrange);
            srcrange = sheetright.get_Range(string.Format("I4:J{0}", userow));
            destrange = sheetRUT.get_Range("E10");
            srcrange.Copy(destrange);
            destrange = sheetRUT.get_Range(string.Format("G10:G{0}", userow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);

            //破损率
            sheetDR.Cells[7, 3] = prjname;
            srcrange = sheetleft.get_Range(string.Format("G4:H{0}", userow));
            destrange = sheetDR.get_Range("C10");
            srcrange.Copy(destrange);
            srcrange = sheetright.get_Range(string.Format("G4:H{0}", userow));
            destrange = sheetDR.get_Range("E10");
            srcrange.Copy(destrange);
            destrange = sheetDR.get_Range(string.Format("G10:G{0}", userow + 6));
            GlobalExcel.SetBorderLine(destrange, 63);
        }

        //////////////////////////////////////////////////////
        //中交国通报表合并
        public static void OutputZJGTExcel(MSExcel._Application excelApp, string leftpath, string destpath)
        {
            int sheetnum = 7;
            List<FileInfo>[] xlsfilenames = new List<FileInfo>[sheetnum];//平整度、车辙、构造深度、GPS高程、原始车辙、路面材质

            int[] startrow = { 11, 10, 10, 10, 2, 3, 2 };
            int[] colnum = { 9, 9, 7, 8, 25, 8, 3 };

            List<string>[] sheetname = new List<string>[sheetnum];
            for (int i = 0; i < sheetnum; ++i)
            {
                xlsfilenames[i] = new List<FileInfo>();
                sheetname[i] = new List<string>();
                if (i == 0) { sheetname[i].Add("1000米"); sheetname[i].Add("100米"); sheetname[i].Add("20米"); }
                if (i == 1) { sheetname[i].Add("1000米"); sheetname[i].Add("100米"); sheetname[i].Add("10米"); }
                if (i == 2) { sheetname[i].Add("1000米"); sheetname[i].Add("100米"); sheetname[i].Add("10米"); }
                if (i == 3) { sheetname[i].Add("1000米"); sheetname[i].Add("5米"); }
                if (i == 4) { sheetname[i].Add("PQI (2)"); }
                if (i == 5) { sheetname[i].Add("10米"); }
                if (i == 6) { sheetname[i].Add("5米"); sheetname[i].Add("10米"); sheetname[i].Add("20米"); sheetname[i].Add("100米"); sheetname[i].Add("1000米"); }
            }

            DirectoryInfo xlsdir = new DirectoryInfo(leftpath);
            FileInfo[] xlsfiles = xlsdir.GetFiles("*.xlsx");
            foreach (FileInfo tfile in xlsfiles)
            {
                if (tfile.Name.Contains("路面平整度"))
                {
                    xlsfilenames[0].Add(tfile);
                }
                else if (tfile.Name.Contains("路面车辙"))
                {
                    xlsfilenames[1].Add(tfile);
                }
                else if (tfile.Name.Contains("路面抗滑性能"))
                {
                    xlsfilenames[2].Add(tfile);
                }
                else if (tfile.Name.Contains("路面高程GPS"))
                {
                    xlsfilenames[3].Add(tfile);
                }
                else if (tfile.Name.Contains("大表"))
                {
                    xlsfilenames[4].Add(tfile);
                }
                else if (tfile.Name.Contains("原始车辙"))
                {
                    xlsfilenames[5].Add(tfile);
                }
                else if (tfile.Name.Contains("路面材质"))
                {
                    xlsfilenames[6].Add(tfile);
                }
            }

            for (int i = 0; i < sheetnum; ++i)
            {
                if (xlsfilenames[i].Count > 0)
                {
                    MergeZJGTExcel(excelApp, xlsfilenames[i], sheetname[i], string.Format("{0}\\{1}", destpath, xlsfilenames[i][0].Name.Replace("_","")),
                        startrow[i], colnum[i]);
                }
            }
        }
        //待合并报表名列表, 待合并表单名, 合并表名, 合并模板路径, 合并起始行, 合并列数
        public static void MergeZJGTExcel(MSExcel._Application excelApp, List<FileInfo> filenames, List<string> sheetnames,
            string destfilename, int startrow, int colnum)
        {
            if (!destfilename.Contains("大表"))
            {
                if (destfilename.Contains("上行"))
                {
                    filenames.Sort((x, y) => x.Name.CompareTo(y.Name));
                }
                else if (destfilename.Contains("下行"))
                {
                    filenames.Sort((x, y) => -x.Name.CompareTo(y.Name));
                }
            }
            MSExcel.Workbook _destWorkbook = null;
            MSExcel.Workbook _srcWorkbook = null;

            if (File.Exists(destfilename))
            {
                File.Delete(destfilename);
            }
            File.Copy(filenames[0].FullName, destfilename);
            _destWorkbook = excelApp.Workbooks.Open(destfilename, Type.Missing,
                    false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet _destWorksheet = null;
            MSExcel.Worksheet _srcWorksheet = null;
            MSExcel.Range srcrange = null, destrange = null;

            foreach (string sheetname in sheetnames)
            {
                _destWorksheet = _destWorkbook.Sheets[sheetname] as MSExcel.Worksheet;
                int userow = 0, destrownum = startrow;
                for (int i = 0; i < filenames.Count; ++i)
                {
                    _srcWorkbook = excelApp.Workbooks.Open(filenames[i].FullName, Type.Missing,
                        true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    _srcWorksheet = _srcWorkbook.Sheets[sheetname] as MSExcel.Worksheet;
                    userow = GlobalExcel.judegeusedrow(_srcWorksheet, 2, startrow);
                    if (i > 0)
                    {
                        srcrange = _srcWorksheet.get_Range(string.Format("A{0}:{1}{2}", startrow, (char)('A' + colnum), userow));
                        destrange = _destWorksheet.get_Range(string.Format("A{0}", destrownum));
                        srcrange.Copy(destrange);
                    }
                    destrownum += userow - startrow + 1;
                    _srcWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                }
            }
            _destWorkbook.Save();
            _destWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
        }

        public static void OutputDocTest(MSWord.Application wordApp, MSExcel.Application excelApp, string excelpath)
        {
            MSWord.Application app = null;
            MSWord.Document doc = null;
            string srcDoc = "D://Doc.docx";
            string dstDoc = srcDoc.Replace("Doc",".DocNew");
            string fileName = "D:\\0.jpg";
            object linkToFile = false;
            object saveWithDocument = true;
            app = new MSWord.Application();
            doc = app.Documents.Add();
            doc.Activate();
            doc.Select();

            //string Lable = "书签名";
            //foreach (MSWord.Bookmark book in doc.Bookmarks)
            //{
            //    if (book.Name == Lable)
            //    {                        
            //        book.Select();                
            //        break;
            //    }
            //}


            MSWord.Selection currentSelection = app.Selection;
            MSWord.InlineShape tmppic = currentSelection.InlineShapes.AddPicture(fileName);
            tmppic.Width = 1191.26f;
            tmppic.Height = 767f;
            tmppic.Select();
 
            MSWord.Shape shape = doc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            //shape.Width = 483.5f;
            //shape.Height = 412.625f;
            //object range = app.Selection.Range;
            //MSWord.Shape shape = app.ActiveDocument.Shapes.AddPicture(fileName, ref linkToFile, ref saveWithDocument, ref range);
            //shape.Width = 100f;
            //shape.Height = 120f;
            shape.IncrementLeft(-90.11f);
            shape.IncrementTop(-74.37f);
            shape.WrapFormat.Type = MSWord.WdWrapType.wdWrapFront;
            doc.SaveAs(dstDoc);
            doc.Close();
      
        }

    }
}
