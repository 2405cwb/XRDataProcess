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
    class MyWordCitySH2013
    {
        #region 上海浦公模板
        public static void OutputMode7DocXls(MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclist)
        {
            // 每一段路的统计
            WriteExcelTJ(excelApp, srcpath, srclist);
        }

        public static void OutputMode7DocAppendix(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclist)
        {
            // 附录
            WriteAppendix2Docx(wordApp, excelApp, srcpath, srclist);
        }

        public static void OutputMode7DocSummary(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath)
        {
            // 第六章、第七章的总结
            WriteSummary2Docx(wordApp, excelApp, srcpath);
        }

        public static void OutputMode7DocHeader(MSWord.Application wordApp, string srcpath,
            ProjectProjectClass tproject, ReportProjectClass treport)
        {
            // 报告的头部
            WriteHeader2Docx(wordApp, srcpath, tproject, treport);
        }

        public static void OutputMode7DocMerge(MSWord.Application wordApp, string srcpath, List<RoadPartProjectClass> srclist)
        {
            // 合并到一起
            WriteAll2Docx(wordApp, srcpath, srclist);
        }

        public static void OutputMode7Doc(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclist)
        {
            // 每一段的报告主体
            WriteMainRoad2Docx(wordApp, excelApp, srcpath, srclist);
        }

        private static float[] width_lq = { 5.0f, 7.0f, 7.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f, 5.0f };
        private static float[] width_sn = { 3.0f, 6.0f, 6.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f, 4.0f };
        private static float[] width_pci = { 10.0f, 15.0f, 15.0f, 18.0f, 15.0f, 15.0f, 15.0f };
        private static float[] width_rqi = { 8.0f, 10.0f, 10.0f, 10.0f, 10.0f, 15.0f, 8.0f, 8.0f };
        private static float[] width_mt = { 8.0f, 10.0f, 10.0f, 10.0f, 10.0f, 15.0f, 10.0f };
        private static float[] width_hz = { 12.0f, 7.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f };//各车道路面技术状况评价结果汇总表

        /// <summary>
        /// 格式化表格
        /// </summary>
        /// <param name="wordApp"></param>
        /// <param name="temptable"></param>
        /// <param name="height"></param>
        /// <param name="colnum">0-不设置列宽，1-沥青病害表列宽，2-水泥病害表列宽，3-PCI表列宽，4-RQI表列宽，5-PQI表列宽, 6-车道技术状况表，7-汇总表</param>
        /// <param name="IsLast"></param>
        /// <param name="roadnum"></param>
        private static void FromatTable(MSWord.Application wordApp, MSWord.Table temptable, float height, int colnum = 0, bool IsSetEveryCell = false, int roadnum = 0)
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
            Thread.Sleep(GlobalWord.wd_sleep_us);

            object oStyleName = null;
            MSWord.Selection currentSelection = null;
            oStyleName = "报告表格内容（通用居中 小五）";
            if (!IsSetEveryCell)
            {
                temptable.Range.Select();
                currentSelection = wordApp.Selection;
                SetStyle(currentSelection, oStyleName, false);
            }
            else
            {
                for (int i = 1; i < 5; ++i)
                {
                    try
                    {
                        temptable.Cell(1, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception ex) { }
                }
                for (int i = 5; i < 9; ++i)
                {
                    try
                    {
                        temptable.Cell(1, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception ex) { }
                }
                for (int i = 5; i < 13; ++i)
                {
                    try
                    {
                        temptable.Cell(2, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception ex) { }
                }
                for (int i = 1; i < 13; ++i)
                {
                    for (int j = 0; j < roadnum; ++j)
                    {
                        try
                        {
                            temptable.Cell(j + 3, i).Range.set_Style(ref oStyleName);
                        }
                        catch (Exception ex) { }
                    }
                }
                for (int i = 1; i < 10; ++i)
                {
                    try
                    {
                        temptable.Cell(roadnum + 3, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception ex) { }
                }
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

            switch (colnum)
            {
                case 1:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_lq[width_lq.Length - 1];
                        for (int i = 0; i < width_lq.Length; ++i)
                        {
                            if (width_lq[i] != width_lq[width_lq.Length - 1])
                            {
                                for (int j = 2; j <= rownum; ++j)
                                {
                                    temptable.Cell(j, i + 1).PreferredWidth = width_lq[i];
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    } break;
                case 2:
                    {
                        int rownum = temptable.Rows.Count;
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_sn[width_sn.Length - 1];
                        for (int i = 0; i < width_sn.Length; ++i)
                        {
                            if (width_sn[i] != width_sn[width_sn.Length - 1])
                            {
                                for (int j = 2; j <= rownum; ++j)
                                {
                                    temptable.Cell(j, i + 1).PreferredWidth = width_sn[i];
                                }
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    } break;
                case 3:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_pci[width_pci.Length - 1];
                        for (int i = 0; i < width_pci.Length; ++i)
                        {
                            if (width_pci[i] != width_pci[width_pci.Length - 1])
                            {
                                temptable.Columns[i + 1].PreferredWidth = width_pci[i];
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    } break;
                case 4:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_rqi[1];
                        for (int i = 0; i < width_rqi.Length; ++i)
                        {
                            if (width_rqi[i] != width_rqi[1])
                            {
                                temptable.Columns[i + 1].PreferredWidth = width_rqi[i];
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    } break;
                case 5:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_mt[1];
                        for (int i = 0; i < width_mt.Length; ++i)
                        {
                            if (width_mt[i] != width_mt[1])
                            {
                                temptable.Columns[i + 1].PreferredWidth = width_mt[i];
                            }
                        }
                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    } break;
                case 6:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = 10;
                    } break;
                case 7:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_hz[1];
                        for (int i = 0; i < 4; ++i)
                        {
                            try
                            {
                                temptable.Cell(1, i + 1).PreferredWidth = width_hz[i];
                            }
                            catch (Exception ex) { }
                        }
                        for (int i = 4; i < 12; ++i)
                        {
                            try
                            {
                                temptable.Cell(2, i + 1).PreferredWidth = width_hz[i];
                            }
                            catch (Exception ex) { }
                        }
                        for (int i = 0; i < roadnum; ++i)
                        {
                            for (int j = 0; j < 12; ++j)
                            {
                                try
                                {
                                    temptable.Cell(i + 2, j + 1).PreferredWidth = width_hz[j];
                                }
                                catch (Exception ex) { }
                            }
                        }
                        for (int i = 0; i < 8; ++i)
                        {
                            try
                            {
                                temptable.Cell(roadnum + 3, i + 2).PreferredWidth = width_hz[i + 4];
                            }
                            catch (Exception ex) { }
                        }
                    } break;
                default: break;
            }

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
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            temptable.Rows.AllowBreakAcrossPages = 0;
            temptable.ApplyStyleHeadingRows = true;

            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            temptable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);

            //wordApp.ScreenUpdating = true;
        }
        
        private static void SetStyle(MSWord.Selection currentSelection, object oStyleName, bool IsTypeParagraph)
        {
            while (true)
            {
                Thread.Sleep(GlobalWord.wd_sleep_us2);
                try
                {
                    currentSelection.set_Style(ref oStyleName);
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(GlobalWord.wd_sleep_us2);
                }
            }
            if (IsTypeParagraph)
            {
                Thread.Sleep(GlobalWord.wd_sleep_us2);
                while (true)
                {
                    try
                    {
                        currentSelection.TypeParagraph();
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(GlobalWord.wd_sleep_us2);
                    }
                }
            }
        }
        private static void WriteText2Word(MSWord.Selection currentSelection, string str)
        {
            Thread.Sleep(GlobalWord.wd_sleep_us2);
            while (true)
            {
                try
                {
                    currentSelection.TypeText(str);
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(GlobalWord.wd_sleep_us2);
                }
            }
        }
        private static MSWord.Table PastExcelTable2Word(MSWord.Document wordDoc, MSExcel.Range srcrange, MSWord.Selection currentSelection,
            ref int wordtablecnt, string tableheader, object oStyleName, bool IsGetTable = true)
        {
            WriteText2Word(currentSelection, tableheader);
            SetStyle(currentSelection, oStyleName, true);

            while (true)
            {
                try
                {
                    System.Windows.Forms.Clipboard.Clear();
                    srcrange.Copy();
                    currentSelection.PasteExcelTable(false, false, false);
                    ++wordtablecnt;
                    Thread.Sleep(GlobalWord.wd_sleep_us);
                    break;
                }
                catch (Exception ex)
                {
                    Thread.Sleep(GlobalWord.wd_sleep_us);
                }
            }

            oStyleName = "报告表下空行";
            SetStyle(currentSelection, oStyleName, true);

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
                    catch (Exception ex)
                    {
                        Thread.Sleep(GlobalWord.wd_sleep_us);
                    }
                }
            }
            return curtable;
        }

        // 生成附录文档
        private static void WriteAppendix2Docx(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclists)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\路段附录模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            MSWord.Document wordDoc = null;
            MSExcel.Workbook[] srcbooks = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            MSWord.Table curtable = null;

            string[] typeheaderstrs = { "路面破损检测与评定结果表",
                                          "路面平整度检测与评定结果表", 
                                          "路面构造深度检测与评定结果表" };

            string destdoc = srcinfo.DirectoryName + "//附录.docx";
            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;
            wordDoc.ShowGrammaticalErrors = false;
            wordDoc.ShowRevisions = false;

            MSWord.Selection currentSelection = null;

            wordDoc.Paragraphs.Last.Range.Select();
            currentSelection = wordApp.Selection;
            //wordApp.ScreenUpdating = false;

            int[] delcol = { 8, 7, 5 };
            int wordtablecnt = 0;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                bool IsRoadPartHeader = true;
                srcbooks = new MSExcel.Workbook[srclist.m_lanelist.Count];
                for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                {
                    while (true)
                    {
                        try
                        {
                            srcbooks[j] = excelApp.Workbooks.Open(srclist.m_lanelist[j].m_xlsxpath, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            break;
                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                        }
                    }
                }
                for (int i = 0; i < 3; ++i)
                {
                    bool IsSheetTypeHeader = true;
                    for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                    {
                        srcsheet = srcbooks[j].Sheets["控制信息"] as MSExcel.Worksheet;
                        srcrange = srcsheet.get_Range("B2:B34");
                        object[,] srcobj = (object[,])srcrange.Value2;
                        srcrange = srcsheet.get_Range("H2:H11");
                        object[,] srcobjroad = (object[,])srcrange.Value2;
                        object oStyleName = null;
                        string tableheader = "";
                        string tabelheaderapp = "";
                        int userownum = 0;

                        if (IsRoadPartHeader)
                        {
                            WriteText2Word(currentSelection, srcobj[7, 1].ToString() + "路面技术状况检测结果");
                            oStyleName = "报告附录1";
                            SetStyle(currentSelection, oStyleName, true);
                            IsRoadPartHeader = false;

                            oStyleName = "报告附表1（隐藏）";
                            SetStyle(currentSelection, oStyleName, true);
                        }

                        if (IsSheetTypeHeader)
                        {
                            if (i == 0)
                            {
                                WriteText2Word(currentSelection, "单元路面病害面积统计及" + typeheaderstrs[i]);
                            }
                            else
                            {
                                WriteText2Word(currentSelection, "单元" + typeheaderstrs[i]);
                            }
                            oStyleName = "报告附录2";
                            SetStyle(currentSelection, oStyleName, true);
                            IsSheetTypeHeader = false;
                        }

                        tabelheaderapp = srcobj[7, 1].ToString() + srcobjroad[3, 1].ToString() + srcobjroad[4, 1].ToString().Replace("车道", "") + "车道";

                        // 病害面积统计表
                        if (i == 0)
                        {
                            bool islq = false;
                            // 沥青
                            try
                            {
                                srcsheet = srcbooks[j].Sheets["病害面积计算（沥青）"] as MSExcel.Worksheet;
                                islq = true;
                            }
                            catch (Exception ex)
                            {
                                islq = false;
                            }
                            if (islq)
                            {
                                ((MSExcel.Range)srcsheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                if (userownum > 2)
                                {
                                    srcrange = srcsheet.get_Range("A1:Q" + userownum.ToString());
                                    srcrange.ColumnWidth = 12.0f;

                                    tableheader = tabelheaderapp + "沥青路面病害面积统计表";
                                    curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告附表2");
                                    curtable.Cell(2, 2).Range.Text = "起点\r\n桩号";
                                    curtable.Cell(2, 3).Range.Text = "终点\r\n桩号";
                                    FromatTable(wordApp, curtable, 0.5f, 1);
                                    wordDoc.Paragraphs.Last.Range.Select();
                                    currentSelection = wordApp.Selection;
                                }
                            }

                            bool issn = false;
                            // 水泥
                            try
                            {
                                srcsheet = srcbooks[j].Sheets["病害面积计算（水泥）"] as MSExcel.Worksheet;
                                issn = true;
                            }
                            catch (Exception ex2)
                            {
                                issn = false;
                            }
                            if (issn)
                            {
                                ((MSExcel.Range)srcsheet.Cells[System.Reflection.Missing.Value, 5]).EntireColumn.Delete();
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                if (userownum > 2)
                                {
                                    srcrange = srcsheet.get_Range("A1:R" + userownum.ToString());
                                    srcrange.ColumnWidth = 12.0f;

                                    tableheader = tabelheaderapp + "水泥路面病害面积统计表";
                                    curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告附表2");
                                    curtable.Cell(2, 2).Range.Text = "起点\r\n桩号";
                                    curtable.Cell(2, 3).Range.Text = "终点\r\n桩号";
                                    FromatTable(wordApp, curtable, 0.5f, 2);
                                    wordDoc.Paragraphs.Last.Range.Select();
                                    currentSelection = wordApp.Selection;
                                }
                            }

                            // PCI的附录表
                            srcsheet = srcbooks[j].Sheets["PCI"] as MSExcel.Worksheet;
                            userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                            srcrange = srcsheet.get_Range("A1:I" + userownum.ToString());

                            tableheader = tabelheaderapp + "路面破损检测与评定结果表";
                            curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告附表2");
                            for (int k = 1; k < delcol.Length; ++k)
                            {
                                while (true)
                                {
                                    try
                                    {
                                        curtable.Columns[delcol[k]].Delete();
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Thread.Sleep(GlobalWord.wd_sleep_us);
                                    }
                                }
                            }
                            FromatTable(wordApp, curtable, 0.5f, 3);
                            wordDoc.Paragraphs.Last.Range.Select();
                            currentSelection = wordApp.Selection;
                        }
                        else
                        {
                            if (i == 1)
                            {
                                srcsheet = srcbooks[j].Sheets["RQI"] as MSExcel.Worksheet;
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                srcrange = srcsheet.get_Range("A1:K" + userownum.ToString());
                            }
                            else
                            {
                                srcsheet = srcbooks[j].Sheets["TD"] as MSExcel.Worksheet;
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                srcrange = srcsheet.get_Range("A1:J" + userownum.ToString());
                            }

                            tableheader = tabelheaderapp + typeheaderstrs[i];
                            curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告附表2");
                            for (int k = 0; k < delcol.Length; ++k)
                            {
                                while (true)
                                {
                                    try
                                    {
                                        curtable.Columns[delcol[k]].Delete();
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Thread.Sleep(GlobalWord.wd_sleep_us);
                                    }
                                }
                            }

                            if (i == 1)
                            {
                                FromatTable(wordApp, curtable, 0.5f, 4);
                            }
                            else
                            {
                                FromatTable(wordApp, curtable, 0.5f, 5);
                            }
                            wordDoc.Paragraphs.Last.Range.Select();
                            currentSelection = wordApp.Selection;
                        }
                    }
                }
                for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                {
                    srcbooks[j].Close(false, Type.Missing, Type.Missing);
                }
                wordDoc.Save();
                generation = System.GC.GetGeneration(wordApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                generation = System.GC.GetGeneration(excelApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
            }

            //wordApp.ScreenUpdating = true;
            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        //生成路段的统计信息
        private static void WriteExcelTJ(MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclists)
        {
            int generation;
            FileInfo srcinfo = new FileInfo(srcpath);

            MSExcel.Workbook srcbook = null;
            MSExcel.Workbook destbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Worksheet destsheet = null;
            MSExcel.Worksheet destsheet2 = null;
            MSExcel.Range srcrange = null;
            MSExcel.Range destrange = null;

            MSExcel.ChartObject chartobj = null;
            MSExcel.Chart chart = null;

            if (!Directory.Exists(srcinfo.DirectoryName + "\\路段统计"))
            {
                Directory.CreateDirectory(srcinfo.DirectoryName + "\\路段统计");
            }

            string CollectSrcXls = string.Format(@"{0}\报告模板\城镇道路\模板5\报告汇总.xlsx", System.Windows.Forms.Application.StartupPath);
            string CollectDestXls = srcinfo.DirectoryName + "\\路段统计\\报告汇总.xlsx";
            MSExcel.Workbook CollectBook = excelApp.Workbooks.Open(CollectSrcXls, false, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CollectBook.SaveAs(CollectDestXls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet CollectDestSheet = CollectBook.Sheets["Sheet1"] as MSExcel.Worksheet;
            MSExcel.Worksheet CollectYHSheet = CollectBook.Sheets["Sheet3"] as MSExcel.Worksheet;
            int CollectRow = 3;

            int Collect_yhstartrowidx = 2;
            int Collect_yhendrowidx = 2;
            int roadpartidx = 0;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                ++roadpartidx;
                string srcxls = string.Format(@"{0}\报告模板\城镇道路\模板5\路段汇总.xlsx", System.Windows.Forms.Application.StartupPath);
                string destxls = srcinfo.DirectoryName
                    + "\\路段统计\\路段"
                    + srclist.m_roadpart.m_id + "#"
                    + srclist.m_roadpart.m_roadinfo.m_code + "_"
                    + srclist.m_roadpart.m_roadinfo.m_name + "（"
                    + srclist.m_roadpart.m_startlocation + "-"
                    + srclist.m_roadpart.m_endlocation + "）"
                    + ".xlsx";

                destbook = excelApp.Workbooks.Open(srcxls, Type.Missing, false,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                destbook.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                double[] lqdisarea = new double[13];
                double[] sndisarea = new double[14];

                int lqunitnum = 0;
                int snunitnum = 0;
                int maxlen = 0;
                int curlen = 0;

                object[,] pciobj = new object[srclist.m_lanelist.Count, 18];
                object[,] rqiobj = new object[srclist.m_lanelist.Count, 13];
                object[,] tdobj = new object[srclist.m_lanelist.Count, 13];
                object[,] pqiobj = new object[srclist.m_lanelist.Count, 13];

                excelApp.ScreenUpdating = false;
                int yhstartrowidx = 2;
                int yhendrowidx = 2;
                for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                {
                    srcbook = excelApp.Workbooks.Open(srclist.m_lanelist[j].m_xlsxpath, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    destsheet = destbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                    srcsheet = srcbook.Sheets["统计"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("H23:P26");
                    object[,] obj1 = (object[,])srcrange.Value2;

                    srcrange = srcsheet.get_Range("H14:S14");
                    object[,] obj2 = (object[,])srcrange.Value2;
                    destrange = destsheet.get_Range(string.Format("AI{0}:AT{0}", j + 3));
                    System.Windows.Forms.Clipboard.Clear();
                    srcrange.Copy();
                    destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValuesAndNumberFormats);

                    // PCI、RQI、PQI、TD
                    int[] srcidx = { 5, 7, 3, 9 };
                    int[] destidx = { 11, 17, 23, 29 };
                    for (int k = 0; k < destidx.Length; ++k)
                    {
                        destsheet.Cells[3 + j, destidx[k]] = obj2[1, 1];
                        for (int i = 0; i < 4; ++i)
                        {
                            destsheet.Cells[3 + j, destidx[k] + i + 1] = obj1[1 + i, srcidx[k]];
                        }
                    }

                    //病害面积统计
                    srcrange = srcsheet.get_Range("C4:C16");
                    obj1 = (object[,])srcrange.Value2;
                    for (int t = 0; t < 13; ++t)
                    {
                        lqdisarea[t] = lqdisarea[t] + Convert.ToDouble(obj1[t + 1, 1]);
                    }

                    srcrange = srcsheet.get_Range("C24:C37");
                    obj1 = (object[,])srcrange.Value2;
                    for (int t = 0; t < 14; ++t)
                    {
                        sndisarea[t] = sndisarea[t] + Convert.ToDouble(obj1[t + 1, 1]);
                    }

                    //单元划分的沥青水泥数量
                    srcsheet = srcbook.Sheets["单元划分"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("F:F");
                    obj1 = (object[,])srcrange.Value2;
                    int tt = 1;
                    while (true)
                    {
                        ++tt;
                        if (obj1[tt, 1] != null)
                        {
                            if (obj1[tt, 1].ToString() == "沥青")
                            {
                                ++lqunitnum;
                            }
                            else if (obj1[tt, 1].ToString() == "水泥")
                            {
                                ++snunitnum;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    //路段信息
                    srcsheet = srcbook.Sheets["控制信息"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("A1:Z35");
                    obj2 = (object[,])srcrange.Value2;

                    curlen = Convert.ToInt32(obj2[8, 5]);
                    if (curlen > maxlen)
                    {
                        maxlen = curlen;
                        destsheet2 = destbook.Sheets["Sheet2"] as MSExcel.Worksheet;
                        destrange = destsheet2.get_Range("A1:Z35");
                        destrange.Value2 = obj2;
                    }

                    //各个指标的数据重新写入路段的sheet3
                    pciobj[j, 0] = roadpartidx;
                    pciobj[j, 1] = obj2[9, 2];
                    pciobj[j, 2] = obj2[8, 2];
                    pciobj[j, 3] = obj2[3, 5];
                    pciobj[j, 4] = obj2[4, 5];
                    pciobj[j, 5] = obj2[11, 2];

                    rqiobj[j, 0] = obj2[11, 2];
                    tdobj[j, 0] = obj2[11, 2];
                    pqiobj[j, 0] = obj2[11, 2];

                    // 养护对策
                    srcsheet = srcbook.Sheets["单元对策"] as MSExcel.Worksheet;
                    int yhrow = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                    Object[,] yhobj = new Object[yhrow - 1, 5];
                    for (int yhi = 0; yhi < yhrow - 1; ++yhi)
                    {
                        yhobj[yhi, 0] = roadpartidx;
                        yhobj[yhi, 1] = obj2[9, 2];
                        yhobj[yhi, 2] = obj2[8, 2];
                        yhobj[yhi, 3] = obj2[3, 5];
                        yhobj[yhi, 4] = obj2[4, 5];
                    }
                    yhendrowidx = yhstartrowidx + yhrow - 2;
                    destsheet2 = destbook.Sheets["Sheet4"] as MSExcel.Worksheet;
                    destrange = destsheet2.get_Range(string.Format("A{0}:E{1}", yhstartrowidx, yhendrowidx));
                    destrange.Value2 = yhobj;
                    Collect_yhendrowidx = Collect_yhstartrowidx + yhrow - 2;
                    destrange = CollectYHSheet.get_Range(string.Format("A{0}:E{1}", Collect_yhstartrowidx, Collect_yhendrowidx));
                    destrange.Value2 = yhobj;

                    srcrange = srcsheet.get_Range(string.Format("B2:N{0}", yhrow));
                    srcrange.Copy();
                    destrange = destsheet2.get_Range(string.Format("F{0}:R{1}", yhstartrowidx, yhendrowidx));
                    destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
                    destrange = CollectYHSheet.get_Range(string.Format("F{0}:R{1}", Collect_yhstartrowidx, Collect_yhendrowidx));
                    destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);

                    yhstartrowidx = yhendrowidx + 1;
                    Collect_yhstartrowidx = Collect_yhendrowidx + 1;

                    // 车道统计信息
                    srcsheet = srcbook.Sheets["统计"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("H2:S26");
                    obj2 = (object[,])srcrange.Value2;

                    pciobj[j, 6] = obj2[13, 1];
                    rqiobj[j, 1] = obj2[13, 1];
                    tdobj[j, 1] = obj2[13, 1];
                    pqiobj[j, 1] = obj2[13, 1];

                    pciobj[j, 8] = obj2[22, 4];
                    pciobj[j, 9] = obj2[22, 5];
                    pciobj[j, 10] = obj2[23, 4];
                    pciobj[j, 11] = obj2[23, 5];
                    pciobj[j, 12] = obj2[24, 4];
                    pciobj[j, 13] = obj2[24, 5];
                    pciobj[j, 14] = obj2[25, 4];
                    pciobj[j, 15] = obj2[25, 5];
                    pciobj[j, 16] = obj2[13, 5];
                    pciobj[j, 17] = obj2[13, 6];

                    rqiobj[j, 3] = obj2[22, 6];
                    rqiobj[j, 4] = obj2[22, 7];
                    rqiobj[j, 5] = obj2[23, 6];
                    rqiobj[j, 6] = obj2[23, 7];
                    rqiobj[j, 7] = obj2[24, 6];
                    rqiobj[j, 8] = obj2[24, 7];
                    rqiobj[j, 9] = obj2[25, 6];
                    rqiobj[j, 10] = obj2[25, 7];
                    rqiobj[j, 11] = obj2[13, 7];
                    rqiobj[j, 12] = obj2[13, 8];

                    tdobj[j, 3] = obj2[22, 8].ToString() == "/" ? 0 : obj2[22, 8];
                    tdobj[j, 4] = obj2[22, 9].ToString() == "/" ? 0 : obj2[22, 9];
                    tdobj[j, 5] = obj2[23, 8].ToString() == "/" ? 0 : obj2[23, 8];
                    tdobj[j, 6] = obj2[23, 9].ToString() == "/" ? 0 : obj2[23, 9];
                    tdobj[j, 7] = obj2[24, 8].ToString() == "/" ? 0 : obj2[24, 8];
                    tdobj[j, 8] = obj2[24, 9].ToString() == "/" ? 0 : obj2[24, 9];
                    tdobj[j, 9] = obj2[25, 8].ToString() == "/" ? 0 : obj2[25, 8];
                    tdobj[j, 10] = obj2[25, 9].ToString() == "/" ? 0 : obj2[25, 9];
                    tdobj[j, 11] = obj2[13, 11].ToString() == "/" ? 0 : obj2[13, 11];
                    tdobj[j, 12] = obj2[13, 12].ToString() == "/" ? 0 : obj2[13, 12];

                    pqiobj[j, 3] = obj2[22, 2];
                    pqiobj[j, 4] = obj2[22, 3];
                    pqiobj[j, 5] = obj2[23, 2];
                    pqiobj[j, 6] = obj2[23, 3];
                    pqiobj[j, 7] = obj2[24, 2];
                    pqiobj[j, 8] = obj2[24, 3];
                    pqiobj[j, 9] = obj2[25, 2];
                    pqiobj[j, 10] = obj2[25, 3];
                    pqiobj[j, 11] = obj2[13, 9];
                    pqiobj[j, 12] = obj2[13, 10];

                    pciobj[j, 7] = Convert.ToDouble(pciobj[j, 8]) + Convert.ToDouble(pciobj[j, 10]) + Convert.ToDouble(pciobj[j, 12]) + Convert.ToDouble(pciobj[j, 14]);
                    rqiobj[j, 2] = Convert.ToDouble(rqiobj[j, 3]) + Convert.ToDouble(rqiobj[j, 5]) + Convert.ToDouble(rqiobj[j, 7]) + Convert.ToDouble(rqiobj[j, 9]);
                    tdobj[j, 2] = Convert.ToDouble(tdobj[j, 3]) + Convert.ToDouble(tdobj[j, 5]) + Convert.ToDouble(tdobj[j, 7]) + Convert.ToDouble(tdobj[j, 9]);
                    pqiobj[j, 2] = Convert.ToDouble(pqiobj[j, 3]) + Convert.ToDouble(pqiobj[j, 5]) + Convert.ToDouble(pqiobj[j, 7]) + Convert.ToDouble(pqiobj[j, 9]);

                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                }

                //路段合计
                srcsheet = destbook.Sheets["Sheet2"] as MSExcel.Worksheet;
                string roaddegreestr = ((MSExcel.Range)(srcsheet.Cells[7, 5])).Value.ToString();
                int roaddegree = 0;
                if (roaddegreestr == "快速路")
                    roaddegree = 0;
                else if (roaddegreestr == "主干路")
                    roaddegree = 1;
                else if (roaddegreestr == "次干路")
                    roaddegree = 2;
                else if (roaddegreestr == "支路")
                    roaddegree = 3;

                object[,] obj = new object[1, 8];
                obj[0, 0] = string.Format("=AVERAGE(AM3:AM{0})", srclist.m_lanelist.Count + 2);
                obj[0, 1] = string.Format("=IF(AM{3}>={0},\"A\",IF(AM{3}>={1},\"B\",IF(AM{3}>={2},\"C\",\"D\")))",
                MyExcelCity._PCIGrade[roaddegree][0], MyExcelCity._PCIGrade[roaddegree][1], MyExcelCity._PCIGrade[roaddegree][2], srclist.m_lanelist.Count + 3);

                obj[0, 2] = string.Format("=AVERAGE(AO3:AO{0})", srclist.m_lanelist.Count + 2);
                obj[0, 3] = string.Format("=IF(AO{3}>={0},\"A\",IF(AO{3}>={1},\"B\",IF(AO{3}>={2},\"C\",\"D\")))",
                MyExcelCity._RQIGrade[roaddegree][0], MyExcelCity._RQIGrade[roaddegree][1], MyExcelCity._RQIGrade[roaddegree][2], srclist.m_lanelist.Count + 3);

                obj[0, 4] = string.Format("=AVERAGE(AQ3:AQ{0})", srclist.m_lanelist.Count + 2);
                obj[0, 5] = string.Format("=IF(AQ{3}>={0},\"A\",IF(AQ{3}>={1},\"B\",IF(AQ{3}>={2},\"C\",\"D\")))",
                MyExcelCity._PQIGrade[roaddegree][0], MyExcelCity._PQIGrade[roaddegree][1], MyExcelCity._PQIGrade[roaddegree][2], srclist.m_lanelist.Count + 3);

                obj[0, 6] = string.Format("=IF(Sheet2!E7=\"支路\",\"/\",AVERAGE(AS3:AS{0}))", srclist.m_lanelist.Count + 2);
                obj[0, 7] = string.Format("=IF(Sheet2!E7=\"支路\",\"/\",IF(AS{3}>={0},\"A\",IF(AS{3}>={1},\"B\",IF(AS{3}>={2},\"C\",\"D\"))))",
                MyExcelCity._MTDGrade[roaddegree][0], MyExcelCity._MTDGrade[roaddegree][1], MyExcelCity._MTDGrade[roaddegree][2], srclist.m_lanelist.Count + 3);

                GlobalExcel.WriteExcel(srclist.m_lanelist.Count + 3, 35, 1, 4, "合计", destsheet, 63);
                destrange = destsheet.get_Range(string.Format("AM{0}:AT{0}", srclist.m_lanelist.Count + 3));
                destrange.Value2 = obj;
                GlobalExcel.SetBorderLine(destrange, 63);

                //病害面积统计
                object[,] lqdisareaobj = new object[lqdisarea.Length, 1];
                for (int i = 0; i < lqdisarea.Length; ++i)
                {
                    lqdisareaobj[i, 0] = lqdisarea[i];
                }
                destrange = destsheet.get_Range("C11:C23");
                destrange.Value2 = lqdisareaobj;

                object[,] sndisareaobj = new object[sndisarea.Length, 1];
                for (int i = 0; i < sndisarea.Length; ++i)
                {
                    sndisareaobj[i, 0] = sndisarea[i];
                }
                destrange = destsheet.get_Range("C31:C44");
                destrange.Value2 = sndisareaobj;

                //设置表格外框、更新图
                destrange = destsheet.get_Range(string.Format("K2:O{0}", srclist.m_lanelist.Count + 2));
                GlobalExcel.SetBorderLine(destrange, 63);
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(1);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);

                destrange = destsheet.get_Range(string.Format("Q2:U{0}", srclist.m_lanelist.Count + 2));
                GlobalExcel.SetBorderLine(destrange, 63);
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(2);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);

                destrange = destsheet.get_Range(string.Format("W2:AA{0}", srclist.m_lanelist.Count + 2));
                GlobalExcel.SetBorderLine(destrange, 63);
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(3);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);

                destrange = destsheet.get_Range(string.Format("AC2:AG{0}", srclist.m_lanelist.Count + 2));
                GlobalExcel.SetBorderLine(destrange, 63);
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(4);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);

                destrange = destsheet.get_Range(string.Format("AI2:AT{0}", srclist.m_lanelist.Count + 2));
                GlobalExcel.SetBorderLine(destrange, 63);

                if (lqunitnum == 0 || snunitnum == 0)
                {
                    if (lqunitnum != 0)
                    {
                        GlobalExcel.WriteExcel(2, 5, 2, 1, "沥青路面", destsheet, 63);
                        GlobalExcel.WriteExcel(2, 6, 2, 1, lqunitnum.ToString(), destsheet, 63);
                    }
                    if (snunitnum != 0)
                    {
                        GlobalExcel.WriteExcel(2, 5, 2, 1, "水泥路面", destsheet, 63);
                        GlobalExcel.WriteExcel(2, 6, 2, 1, snunitnum.ToString(), destsheet, 63);
                    }
                }
                else
                {
                    destsheet.Cells[2, 6] = lqunitnum;
                    destsheet.Cells[3, 6] = snunitnum;
                }
                destsheet.Cells[2, 7] = srclist.m_lanelist.Count;

                double sumperval = 0;
                int disperrownum = 0;
                int disperrowcnt = 0;
                object[,] dispercntobj;
                // 病害类型的饼图 沥青
                srcrange = destsheet.get_Range("F10:G16");
                obj = (object[,])srcrange.Value2;
                disperrownum = 0;
                sumperval = 0;
                for (int tt = 2; tt < 8; ++tt)
                {
                    try
                    {
                        double perval = Convert.ToDouble(obj[tt, 2]);
                        if (perval > 0.05)
                        {
                            ++disperrownum;
                        }
                        else
                        {
                            sumperval = sumperval + perval;
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
                if (sumperval > 0)
                {
                    disperrownum = disperrownum + 2;
                }
                else
                {
                    disperrownum = disperrownum + 1;
                }
                dispercntobj = new object[disperrownum, 2];
                disperrowcnt = 0;
                dispercntobj[disperrowcnt, 0] = obj[1, 1];
                dispercntobj[disperrowcnt, 1] = obj[1, 2];
                for (int tt = 2; tt < 8; ++tt)
                {
                    try
                    {
                        double perval = Convert.ToDouble(obj[tt, 2]);
                        if (perval > 0.05)
                        {
                            ++disperrowcnt;
                            dispercntobj[disperrowcnt, 0] = obj[tt, 1];
                            dispercntobj[disperrowcnt, 1] = obj[tt, 2];
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
                if (sumperval > 0)
                {
                    ++disperrowcnt;
                    dispercntobj[disperrowcnt, 0] = "其他病害类型合计";
                    dispercntobj[disperrowcnt, 1] = sumperval;
                }
                destrange = destsheet.get_Range("F18:G" + (disperrownum + 17));
                destrange.Value2 = dispercntobj;
                GlobalExcel.SetBorderLine(destrange, 63);
                destrange = destsheet.get_Range("F19:G" + (disperrownum + 17));
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(5);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);
                // 病害类型的饼图 水泥
                srcrange = destsheet.get_Range("F30:G37");
                obj = (object[,])srcrange.Value2;
                disperrownum = 0;
                sumperval = 0;
                for (int tt = 2; tt < 9; ++tt)
                {
                    try
                    {
                        double perval = Convert.ToDouble(obj[tt, 2]);
                        if (perval > 0.05)
                        {
                            ++disperrownum;
                        }
                        else
                        {
                            sumperval = sumperval + perval;
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
                if (sumperval > 0)
                {
                    disperrownum = disperrownum + 2;
                }
                else
                {
                    disperrownum = disperrownum + 1;
                }
                dispercntobj = new object[disperrownum, 2];
                disperrowcnt = 0;
                dispercntobj[disperrowcnt, 0] = obj[1, 1];
                dispercntobj[disperrowcnt, 1] = obj[1, 2];
                for (int tt = 2; tt < 9; ++tt)
                {
                    try
                    {
                        double perval = Convert.ToDouble(obj[tt, 2]);
                        if (perval > 0.05)
                        {
                            ++disperrowcnt;
                            dispercntobj[disperrowcnt, 0] = obj[tt, 1];
                            dispercntobj[disperrowcnt, 1] = obj[tt, 2];
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                }
                if (sumperval > 0)
                {
                    ++disperrowcnt;
                    dispercntobj[disperrowcnt, 0] = "其他病害类型合计";
                    dispercntobj[disperrowcnt, 1] = sumperval;
                }
                destrange = destsheet.get_Range("F39:G" + (disperrownum + 38));
                destrange.Value2 = dispercntobj;
                GlobalExcel.SetBorderLine(destrange, 63);
                destrange = destsheet.get_Range("F40:G" + (disperrownum + 38));
                chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(6);
                chart = chartobj.Chart;
                chart.SetSourceData(destrange, 2);
              
                //各个指标的数据重新写入路段的sheet3
                destsheet = destbook.Sheets["Sheet3"] as MSExcel.Worksheet;
                destrange = destsheet.get_Range(string.Format("A3:R{0}", srclist.m_lanelist.Count + 2));
                destrange.Value2 = pciobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = destsheet.get_Range(string.Format("T3:AF{0}", srclist.m_lanelist.Count + 2));
                destrange.Value2 = rqiobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = destsheet.get_Range(string.Format("AH3:AT{0}", srclist.m_lanelist.Count + 2));
                destrange.Value2 = tdobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = destsheet.get_Range(string.Format("AV3:BH{0}", srclist.m_lanelist.Count + 2));
                destrange.Value2 = pqiobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                //各个指标的数据重新写入汇总的Sheet1
                destrange = CollectDestSheet.get_Range(string.Format("A{0}:R{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                destrange.Value2 = pciobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = CollectDestSheet.get_Range(string.Format("T{0}:AF{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                destrange.Value2 = rqiobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = CollectDestSheet.get_Range(string.Format("AH{0}:AT{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                destrange.Value2 = tdobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                destrange = CollectDestSheet.get_Range(string.Format("AV{0}:BH{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                destrange.Value2 = pqiobj;
                GlobalExcel.SetBorderLine(destrange, 63);

                CollectRow = CollectRow + srclist.m_lanelist.Count;

                excelApp.ScreenUpdating = true;

                destbook.Save();
                destbook.Close(Type.Missing, Type.Missing, Type.Missing);
            }
            CollectBook.Save();
            CollectBook.Close(Type.Missing, Type.Missing, Type.Missing);

            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void PastExcelPic2Word(MSWord.Selection currentSelection, MSExcel.Worksheet srcsheet, int PicIdx, string PicName, bool IsTypeParagraph = true)
        {
            if (IsTypeParagraph)
            {
                currentSelection.TypeParagraph();
            }

            object oStyleName = "报告图与下段同页";
            currentSelection.set_Style(ref oStyleName);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            currentSelection.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            MSExcel.Shape excelshape = srcsheet.Shapes.Item(PicIdx) as Microsoft.Office.Interop.Excel.Shape;
            System.Windows.Forms.Clipboard.Clear();
            excelshape.Copy();
            currentSelection.PasteAndFormat(MSWord.WdRecoveryType.wdChartPicture);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            currentSelection.MoveRight();
            currentSelection.TypeParagraph();
            currentSelection.TypeText(PicName);
            oStyleName = "报告图标题3";
            currentSelection.set_Style(ref oStyleName);
            Thread.Sleep(GlobalWord.wd_sleep_us);
        }

        private static void WriteTablePicStr(MSWord.Selection currentSelection, object oStyleName, string str)
        {
            currentSelection.TypeText(str);
            SetStyle(currentSelection, oStyleName, false);
        }

        /// <summary>
        /// 上海浦公城镇路段主体报告模板中，养护对策的表格中的保留松、还是严，那一列，0-松，1-严
        /// </summary>
        public static int _YHTypeDoc = 0;
        // 生成路段的报告
        private static void WriteMainRoad2Docx(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclists)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\路段主体模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            MSWord.Document wordDoc = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            MSExcel.Workbook srcbook2 = null;
            MSExcel.Worksheet srcsheet2 = null;
            MSWord.Selection currentSelection = null;
            int tablecount = 1;
            object oStyleName;
            MSWord.Table curtable = null;

            bool IsHasTD = false;

            string[] IndexJudgeStrs = { "路面综合评价", "路面破损评价", "路面行驶质量评价", "路面抗滑能力评价" };
            int wordcnt = 0;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                string roadpartfname = "路段" + srclist.m_roadpart.m_id + "#"
                    + srclist.m_roadpart.m_roadinfo.m_code + "_"
                    + srclist.m_roadpart.m_roadinfo.m_name + "（"
                    + srclist.m_roadpart.m_startlocation + "-"
                    + srclist.m_roadpart.m_endlocation + "）";

                ++wordcnt;
                string destdoc = srcinfo.DirectoryName + "\\路段统计\\" + roadpartfname + ".docx";
                wordDoc = wordApp.Documents.Open(srcdoc,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
                wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wordDoc.SpellingChecked = false;
                wordDoc.ShowSpellingErrors = false;

                string srcxls = srcinfo.DirectoryName + "\\路段统计\\" + roadpartfname + ".xlsx";
                srcbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                bool IsHasStreetPic = false;
                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "道路概况表例")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        oStyleName = "报告表标题3";
                        WriteTablePicStr(currentSelection, oStyleName, "{道路名称}道路概况表");
                        break;
                    }
                }

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "地理位置示意图例")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        oStyleName = "报告图标题3";
                        WriteTablePicStr(currentSelection, oStyleName, "{道路名称}（{路段起点}～{路段终点}）地理位置示意图");
                        break;
                    }
                }

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "起点照图例")
                    {
                        IsHasStreetPic = true;
                        book.Select();
                        currentSelection = wordApp.Selection;
                        oStyleName = "报告照片标题3";
                        WriteTablePicStr(currentSelection, oStyleName, "起点照");
                        break;
                    }
                }

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "终点照图例")
                    {
                        IsHasStreetPic = true;
                        book.Select();
                        currentSelection = wordApp.Selection;
                        oStyleName = "报告照片标题3";
                        WriteTablePicStr(currentSelection, oStyleName, "终点照");
                        break;
                    }
                }

                srcsheet = srcbook.Sheets["Sheet2"] as MSExcel.Worksheet;
                srcrange = srcsheet.get_Range("E2:E12");
                object[,] srcobj = (object[,])srcrange.Value2;
                if (srcobj[6, 1] != null && srcobj[6, 1].ToString() != "支路")
                {
                    IsHasTD = true;
                }
                else
                {
                    IsHasTD = false;
                }

                srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                srcrange = srcsheet.get_Range("A1:G3");
                GlobalWord.PastExcel2Word(wordApp, srcrange, wordDoc, "道路概况表");
                srcobj = (object[,])srcrange.Value2;

                // 技术状况评定等级分布的文字描述
                // PQI、PCI、RQI、TD
                List<string>[] IndexValStrs = new List<string>[4];
                for (int i = 0; i < 4; ++i)
                {
                    IndexValStrs[i] = new List<string>();
                }

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "路面病害类型分析")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        break;
                    }
                }

                if (srcobj[2, 5] != null && srcobj[2, 5].ToString() == "沥青路面"
                    || srcobj[3, 5] != null && srcobj[3, 5].ToString() == "沥青路面")
                {
                    srcrange = srcsheet.get_Range("A10:D24");
                    PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, "路面病害类型面积统计汇总表（沥青）", "报告表标题3", false);
                }

                if (srcobj[2, 5] != null && srcobj[2, 5].ToString() == "水泥路面"
                    || srcobj[3, 5] != null && srcobj[3, 5].ToString() == "水泥路面")
                {
                    srcrange = srcsheet.get_Range("A30:D45");
                    PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, "路面病害类型面积统计汇总表（水泥）", "报告表标题3", false);
                }

                bool islq = false;
                if (srcobj[2, 5] != null && srcobj[2, 5].ToString() == "沥青路面"
                    || srcobj[3, 5] != null && srcobj[3, 5].ToString() == "沥青路面")
                {
                    double sumarea = 0.0;
                    srcrange = srcsheet.get_Range("A10:D24");
                    object[,] lqareaobj = (object[,])srcrange.Value2;
                    try
                    {
                        sumarea = Convert.ToDouble(lqareaobj[15, 3]);
                    }
                    catch (System.Exception ex) { }
                    if (sumarea > 0)
                    {
                        PastExcelPic2Word(currentSelection, srcsheet, 5, "沥青路面病害类型面积占比图", false);
                        islq = true;
                    }
                }

                if (srcobj[2, 5] != null && srcobj[2, 5].ToString() == "水泥路面"
                    || srcobj[3, 5] != null && srcobj[3, 5].ToString() == "水泥路面")
                {
                    double sumarea = 0.0;
                    srcrange = srcsheet.get_Range("A30:D45");
                    object[,] snareaobj = (object[,])srcrange.Value2;
                    try
                    {
                        sumarea = Convert.ToDouble(snareaobj[16, 3]);
                    }
                    catch (System.Exception ex) { }
                    if (sumarea > 0)
                    {
                        PastExcelPic2Word(currentSelection, srcsheet, 6, "水泥路面病害类型面积占比图", islq);
                    }
                }

                int idx = 0;
                tablecount = wordDoc.Tables.Count;
                foreach (MSWord.Table temptable in wordDoc.Tables)
                {
                    ++idx;
                    if (IsHasStreetPic)
                    {
                        if (idx != 2)
                        {
                            FromatTable(wordApp, temptable, 0.6f, 0, false, (srclist.m_lanelist.Count));
                        }
                    }
                    else
                    {
                        FromatTable(wordApp, temptable, 0.6f, 0, false, (srclist.m_lanelist.Count));
                    }
                }

                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "车道路面技术状况评定等级分布统计汇总表")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        break;
                    }
                }

                bool allA = true;
                bool allB = true;
                int tablenum1 = wordDoc.Tables.Count;
                for (int i = 0; i < srclist.m_lanelist.Count; ++i)
                {
                    srcbook2 = excelApp.Workbooks.Open(srclist.m_lanelist[i].m_xlsxpath, Type.Missing,
                                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    srcsheet2 = srcbook2.Sheets["统计"] as MSExcel.Worksheet;
                    string roadnumstr = ((MSExcel.Range)srcsheet2.Cells[14, 8]).Value.ToString();
                    string tableheader = roadnumstr + "路面技术状况评定等级分布统计汇总表";
                    if (IsHasTD)
                    {
                        srcrange = srcsheet2.get_Range("H21:P26");
                    }
                    else
                    {
                        srcrange = srcsheet2.get_Range("H21:N26");
                    }
                    curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, tableheader, "报告表标题3", true);

                    // 技术状况评定等级分布的文字描述
                    // PQI、PCI、RQI、TD
                    object[,] objsrc = (object[,])srcrange.Value2;
                    double cval, dval, aval, bval;
                    for (int k = 0; k < 4; ++k)
                    {
                        string str = roadnumstr + IndexJudgeStrs[k];
                        try
                        {
                            aval = Convert.ToDouble(objsrc[3, 3 + k * 2]);
                            if (aval < 100.0)
                            {
                                allA = false;
                            }
                        }
                        catch (System.Exception ex) { }
                        try
                        {
                            bval = Convert.ToDouble(objsrc[4, 3 + k * 2]);
                            if (bval < 100.0)
                            {
                                allB = false;
                            }
                        }
                        catch (System.Exception ex) { }

                        try
                        {
                            cval = Convert.ToDouble(objsrc[5, 3 + k * 2]);
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                        try
                        {
                            dval = Convert.ToDouble(objsrc[6, 3 + k * 2]);
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                        if (cval > 0)
                        {
                            str = str + "等级为“C”的长度占比为" + cval.ToString("0.00") + "%";
                            if (dval > 0)
                            {
                                str = str + "；等级为“D”的长度占比为" + dval.ToString("0.00") + "%。";
                                IndexValStrs[k].Add(str);
                            }
                            else
                            {
                                str = str + "。";
                                IndexValStrs[k].Add(str);
                            }
                        }
                        else
                        {
                            if (dval > 0)
                            {
                                str = str + "等级为“D”的长度占比为" + dval.ToString("0.00") + "%。";
                                IndexValStrs[k].Add(str);
                            }
                        }
                    }
                    srcbook2.Close(Type.Missing, Type.Missing, Type.Missing);
                }

                PastExcelPic2Word(currentSelection, srcsheet, 1, "各车道单元PCI等级占比分布图", false);
                PastExcelPic2Word(currentSelection, srcsheet, 2, "各车道单元RQI等级占比分布图");
                PastExcelPic2Word(currentSelection, srcsheet, 3, "各车道单元PQI等级占比分布图");
                if (IsHasTD)
                {
                    PastExcelPic2Word(currentSelection, srcsheet, 4, "各车道单元TD等级占比分布图");
                }

                srcsheet = srcbook.Sheets["Sheet2"] as MSExcel.Worksheet;
                srcrange = srcsheet.get_Range("A1:T34");
                srcobj = (object[,])srcrange.Value2;

                //写入路面技术状况评定等级的文字描述
                if (IndexValStrs[0].Count > 0 || IndexValStrs[1].Count > 0 || IndexValStrs[2].Count > 0 || IndexValStrs[3].Count > 0)
                {
                    oStyleName = "报告正文";
                    currentSelection.TypeParagraph();
                    currentSelection.set_Style(ref oStyleName);

                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，各指标评价等级为“C”或“D”的占比统计结果如下：");
                    currentSelection.TypeParagraph();
                    oStyleName = "标题 6";
                    currentSelection.set_Style(ref oStyleName);
                    if (IndexValStrs[1].Count > 0)
                    {
                        currentSelection.TypeText("路面状况指数PCI。");
                        foreach (string str in IndexValStrs[1])
                        {
                            currentSelection.TypeText(str);
                        }
                    }
                    if (IndexValStrs[2].Count > 0)
                    {
                        if (IndexValStrs[1].Count > 0)
                        {
                            currentSelection.TypeParagraph();
                        }
                        currentSelection.TypeText("路面行驶质量指数RQI。");
                        foreach (string str in IndexValStrs[2])
                        {
                            currentSelection.TypeText(str);
                        }
                    }
                    if (IndexValStrs[0].Count > 0)
                    {
                        if (IndexValStrs[2].Count > 0)
                        {
                            currentSelection.TypeParagraph();
                        }
                        currentSelection.TypeText("路面综合评价指数PQI。");
                        foreach (string str in IndexValStrs[0])
                        {
                            currentSelection.TypeText(str);
                        }
                    }
                    if (IndexValStrs[3].Count > 0)
                    {
                        if (IndexValStrs[0].Count > 0)
                        {
                            currentSelection.TypeParagraph();
                        }
                        currentSelection.TypeText("路面构造深度TD。");
                        foreach (string str in IndexValStrs[3])
                        {
                            currentSelection.TypeText(str);
                        }
                    }
                }
                else if (IndexValStrs[0].Count == 0 && IndexValStrs[1].Count == 0 && IndexValStrs[2].Count == 0 && IndexValStrs[3].Count == 0)
                {
                    oStyleName = "报告正文";
                    currentSelection.TypeParagraph();
                    currentSelection.set_Style(ref oStyleName);

                    if (allA)
                    {
                        currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“A”。");
                    }
                    else if (allB)
                    {
                        currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“B”。");
                    }
                    else if (!allA && !allB)
                    {
                        currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“A”或“B”。");
                    }
                }

                int tablenum = wordDoc.Tables.Count;
                for (int i = tablenum1 + 1; i <= tablenum; ++i)
                {
                    FromatTable(wordApp, wordDoc.Tables[i], 0.6f, 6, false, srclist.m_lanelist.Count);
                }

                srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                if (IsHasTD)
                {
                    srcrange = srcsheet.get_Range("AI1:AT" + (srclist.m_lanelist.Count + 3).ToString());
                }
                else
                {
                    srcrange = srcsheet.get_Range("AI1:AR" + (srclist.m_lanelist.Count + 3).ToString());
                }
                object[,] tjobj = (object[,])srcrange.Value2;
                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "各车道路面技术状况评价结果汇总表")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        break;
                    }
                }
                curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, "各车道路面技术状况评价结果汇总表", "报告表标题3", true);
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                FromatTable(wordApp, curtable, 0.6f, 7, true, srclist.m_lanelist.Count);

                tablenum1 = wordDoc.Tables.Count;
                for (int i = 0; i < srclist.m_lanelist.Count; ++i)
                {
                    srcbook2 = excelApp.Workbooks.Open(srclist.m_lanelist[i].m_xlsxpath, Type.Missing,
                                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // 粘贴起止点景观照片
                    if (i == 0)
                    {
                        srcsheet2 = srcbook2.Sheets["景观图像"] as MSExcel.Worksheet;

                        foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                        {
                            if (book.Name == "起止点照")
                            {
                                book.Select();
                                currentSelection = wordApp.Selection;

                                MSExcel.Shape excelshape = srcsheet2.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
                                System.Windows.Forms.Clipboard.Clear();
                                excelshape.Copy();
                                currentSelection.Paste();
                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                currentSelection.TypeText("  ");
                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                excelshape = srcsheet2.Shapes.Item(2) as Microsoft.Office.Interop.Excel.Shape;
                                System.Windows.Forms.Clipboard.Clear();
                                excelshape.Copy();
                                currentSelection.Paste();
                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                break;
                            }
                        }

                        foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                        {
                            if (book.Name == "养护对策一览表")
                            {
                                book.Select();
                                currentSelection = wordApp.Selection;
                                break;
                            }
                        }
                    }

                    srcsheet2 = srcbook2.Sheets["统计"] as MSExcel.Worksheet;
                    string tableheader = ((MSExcel.Range)srcsheet2.Cells[14, 8]).Value.ToString() + "养护对策一览表";

                    srcsheet2 = srcbook2.Sheets["单元对策"] as MSExcel.Worksheet;
                    int trownum = GlobalExcel.judegeusedrow(srcsheet2, 1);
                    srcrange = srcsheet2.get_Range("A1:N" + trownum.ToString());
                    srcrange.ColumnWidth = 9.0f;

                    curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, tableheader, "报告表标题3", true);

                    srcbook2.Close(Type.Missing, Type.Missing, Type.Missing);
                }

                tablenum = wordDoc.Tables.Count;
                for (int i = tablenum1 + 1; i <= tablenum; ++i)
                {
                    MSWord.Table temptable = wordDoc.Tables[i];
                    if (_YHTypeDoc == 1)
                    {
                        temptable.Columns[13].Delete();
                    }
                    else if (_YHTypeDoc == 0)
                    {
                        temptable.Columns[14].Delete();
                    }
                    temptable.Columns[12].Delete();
                    temptable.Columns[8].Delete();
                    temptable.Cell(1, 1).Range.Text = "序号";
                    temptable.Cell(1, 11).Range.Text = "养护对策";
                    FromatTable(wordApp, temptable, 0.6f);
                }

                // 字段替换                
                // 技术状况描述的文字
                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "技术状况描述")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        break;
                    }
                }
                currentSelection.TypeParagraph();
                oStyleName = "标题 6";
                currentSelection.set_Style(ref oStyleName);
                currentSelection.TypeText("{道路名称}PQI均值为" + Convert.ToDouble(tjobj[srclist.m_lanelist.Count + 3, 9]).ToString("0.00")
                    + "，等级为" + tjobj[srclist.m_lanelist.Count + 3, 10].ToString()
                    + "，其分项指标PCI（" + Convert.ToDouble(tjobj[srclist.m_lanelist.Count + 3, 5]).ToString("0.00")
                    + "）等级" + tjobj[srclist.m_lanelist.Count + 3, 6].ToString()
                    + "、RQI（" + Convert.ToDouble(tjobj[srclist.m_lanelist.Count + 3, 7]).ToString("0.00")
                    + "）等级" + tjobj[srclist.m_lanelist.Count + 3, 8].ToString()
                    + "。");
                if (srcobj[7, 5] != null && srcobj[7, 5].ToString() != "支路")
                {
                    currentSelection.TypeParagraph();
                    currentSelection.TypeText("{道路名称}路面抗滑TD平均值为" + Convert.ToDouble(tjobj[srclist.m_lanelist.Count + 3, 11]).ToString("0.00")
                        + "，等级为" + tjobj[srclist.m_lanelist.Count + 3, 12].ToString()
                        + "。");
                }

                Dictionary<string, string> datas = new Dictionary<string, string>();
                datas.Add("{道路名称}", srcobj[8, 2] == null ? "" : srcobj[8, 2].ToString());
                datas.Add("{路段起点}", srcobj[3, 5] == null ? "" : srcobj[3, 5].ToString());
                datas.Add("{路段终点}", srcobj[4, 5] == null ? "" : srcobj[4, 5].ToString());
                datas.Add("{县}", srcobj[5, 2] == null ? "" : srcobj[5, 2].ToString());
                datas.Add("{街道乡镇}", srcobj[6, 2] == null ? "" : srcobj[6, 2].ToString());
                datas.Add("{道路类型}", srcobj[10, 2] == null ? "" : srcobj[10, 2].ToString());
                datas.Add("{道路等级}", srcobj[7, 5] == null ? "" : srcobj[7, 5].ToString());
                datas.Add("{道路全长}", srcobj[8, 5] == null ? "" : srcobj[8, 5].ToString());
                datas.Add("{道路宽度}", srcobj[9, 5] == null ? "" : srcobj[9, 5].ToString());
                datas.Add("{路面类型}", srcobj[11, 5] == null ? "" : srcobj[11, 5].ToString());
                datas.Add("{车道情况}", srcobj[3, 8] == null ? "" : srcobj[3, 8].ToString());
                datas.Add("{检测结果附表}", "附录" + wordcnt.ToString());
                datas.Add("{车道描述}", srcobj[4, 8] == null ? "" : srcobj[4, 8].ToString() + srcobj[5, 8] == null ? "" : srcobj[5, 8].ToString() + "车道");
                datas.Add("{检测日期}", srcobj[8, 8] == null ? "" : srcobj[8, 8].ToString().Insert(6, "月").Insert(4, "年") + "日");

                //主要病害描述 
                foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                {
                    if (book.Name == "路面病害描述")
                    {
                        book.Select();
                        currentSelection = wordApp.Selection;
                        break;
                    }
                }
                WriteDiseaseStrs(currentSelection, srcsheet);

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

                object myFind = wordDoc.Content.Find;
                foreach (var item in datas)
                {
                    object findText = item.Key;
                    object replaceText = item.Value;
                    Parameters[0] = findText;
                    Parameters[9] = replaceText;
                    while (true)
                    {
                        try
                        {
                            myFind.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, myFind, Parameters);
                            break;
                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                        }
                    }
                }

                //插入交叉引用
                WriteCrossReference(wordApp, wordDoc);

                srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                wordDoc.Save();
                wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            }

            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static void WriteCrossReference(MSWord.Application wordApp, MSWord.Document wordDoc)
        {
            //病害面积表
            MSWord.Selection currentSelection = null;

            object crossReferenceItems = wordDoc.GetCrossReferenceItems(MSWord.WdReferenceType.wdRefTypeNumberedItem);
            Array crossitemarr = ((Array)(crossReferenceItems));
            int crossitemnum = crossitemarr.Length;
            string[] crossitemarrstrs = new string[crossitemnum + 1];

            int OverViewTablbeStrIdx = 0;
            int MapPicStrIdx = 0;
            int StartPicStrIdx = 0;
            int EndPicStrIdx = 0;

            int LQDisTableStrIdx = 0;
            int SNDisTableStrIdx = 0;

            int LQDisPicStrIdx = 0;
            int SNDisPicStrIdx = 0;

            int RoadHZTableStr1Idx = 0;
            int RoadHZTableStr2Idx = 0;

            int IndexPicStr1Idx = 0;
            int IndexPicStr2Idx = 0;

            int HZTableStrIdx = 0;

            int YHTableStr1Idx = 0;
            int YHTableStr2Idx = 0;

            for (int i = 1; i <= crossitemnum; ++i)
            {
                crossitemarrstrs[i] = (string)(crossitemarr.GetValue(i));

                if (crossitemarrstrs[i].Contains("道路概况表"))
                {
                    OverViewTablbeStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("地理位置示意图"))
                {
                    MapPicStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("起点照"))
                {
                    StartPicStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("终点照"))
                {
                    EndPicStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("路面病害类型面积统计汇总表（沥青）"))
                {
                    LQDisTableStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("路面病害类型面积统计汇总表（水泥）"))
                {
                    SNDisTableStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("沥青路面病害类型面积占比图"))
                {
                    LQDisPicStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("水泥路面病害类型面积占比图"))
                {
                    SNDisPicStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("技术状况评定等级分布统计汇总表"))
                {
                    if (RoadHZTableStr1Idx == 0)
                    {
                        RoadHZTableStr1Idx = i;
                    }
                    RoadHZTableStr2Idx = i;
                }
                else if (crossitemarrstrs[i].Contains("等级占比分布图"))
                {
                    if (IndexPicStr1Idx == 0)
                    {
                        IndexPicStr1Idx = i;
                    }
                    IndexPicStr2Idx = i;
                }
                else if (crossitemarrstrs[i].Contains("路面技术状况评价结果汇总表"))
                {
                    HZTableStrIdx = i;
                }
                else if (crossitemarrstrs[i].Contains("养护对策一览表"))
                {
                    if (YHTableStr1Idx == 0)
                    {
                        YHTableStr1Idx = i;
                    }
                    YHTableStr2Idx = i;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "道路概况表交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                            MSWord.WdReferenceKind.wdNumberRelativeContext, OverViewTablbeStrIdx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "地理位置示意图交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                            MSWord.WdReferenceKind.wdNumberRelativeContext, MapPicStrIdx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "起点照片交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                            MSWord.WdReferenceKind.wdNumberRelativeContext, StartPicStrIdx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "终点照片交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                            MSWord.WdReferenceKind.wdNumberRelativeContext, EndPicStrIdx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "病害面积表交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    if (LQDisTableStrIdx != 0 && SNDisTableStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisTableStrIdx, true);
                            currentSelection.TypeText("、");
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisTableStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    else if (LQDisTableStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisTableStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    else if (SNDisTableStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisTableStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "病害面积图交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    if (LQDisPicStrIdx != 0 && SNDisPicStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisPicStrIdx, true);
                            currentSelection.TypeText("、");
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisPicStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    else if (LQDisPicStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisPicStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    else if (SNDisPicStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisPicStrIdx, true);
                        }
                        catch (System.Exception ex) { }
                    }
                    else
                    {
                        for (int i = 0; i < 2; ++i)
                        {
                            currentSelection.Delete();
                        }
                        for (int i = 0; i < 9; ++i)
                        {
                            currentSelection.TypeBackspace();
                        }
                    }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "技术状况表交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, RoadHZTableStr1Idx, true);
                        if (RoadHZTableStr1Idx != RoadHZTableStr2Idx)
                        {
                            currentSelection.TypeText("～");
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                    MSWord.WdReferenceKind.wdNumberRelativeContext, RoadHZTableStr2Idx, true);
                        }
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "技术状况图交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, IndexPicStr1Idx, true);
                        currentSelection.TypeText("～");
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, IndexPicStr2Idx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "汇总表交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                            MSWord.WdReferenceKind.wdNumberRelativeContext, HZTableStrIdx, true);
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "养护建议表交叉引用")
                {
                    book.Select();
                    currentSelection = wordApp.Selection;

                    try
                    {
                        currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, YHTableStr1Idx, true);
                        if (YHTableStr1Idx != YHTableStr2Idx)
                        {
                            currentSelection.TypeText("～");
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, YHTableStr2Idx, true);
                        }
                    }
                    catch (System.Exception ex) { }
                    break;
                }
            }
        }

        private static void WriteDiseaseStrs(MSWord.Selection currentSelection, MSExcel.Worksheet srcsheet)
        {
            bool hasstr = false;
            string disstrs = "";
            List<string> disname = new List<string>();
            List<string> percent = new List<string>();
            MSExcel.Range srcrange = srcsheet.get_Range("F11:G16");
            object[,] srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < 7; ++i)
            {
                if (Convert.ToDouble(srcobj[i, 2]) > 0.2)
                {
                    disname.Add(srcobj[i, 1].ToString());
                    percent.Add((Convert.ToDouble(srcobj[i, 2]) * 100).ToString("0.00") + "%");
                }
            }
            if (disname.Count > 0)
            {
                disstrs = "本路段沥青路面主要病害为";
                for (int i = 0; i < disname.Count; ++i)
                {
                    disstrs = disstrs + disname[i];
                    if (i < disname.Count - 1)
                    {
                        disstrs = disstrs + "、";
                    }
                }
                if (disname.Count == 1)
                {
                    disstrs = disstrs + "，病害占比为";
                }
                else
                {
                    disstrs = disstrs + "，病害占比分别为";
                }
                for (int i = 0; i < percent.Count; ++i)
                {
                    disstrs = disstrs + percent[i];
                    if (i < percent.Count - 1)
                    {
                        disstrs = disstrs + "、";
                    }
                }
                disstrs = disstrs + "。";
                currentSelection.TypeText(disstrs);
                hasstr = true;
            }

            disname.Clear();
            percent.Clear();
            srcrange = srcsheet.get_Range("F31:G37");
            srcobj = (object[,])srcrange.Value2;
            for (int i = 1; i < 8; ++i)
            {
                if (Convert.ToDouble(srcobj[i, 2]) > 0.2)
                {
                    disname.Add(srcobj[i, 1].ToString());
                    percent.Add((Convert.ToDouble(srcobj[i, 2]) * 100).ToString("0.00") + "%");
                }
            }
            if (disname.Count > 0)
            {
                disstrs = "本路段水泥路面主要病害为";
                for (int i = 0; i < disname.Count; ++i)
                {
                    disstrs = disstrs + disname[i];
                    if (i < disname.Count - 1)
                    {
                        disstrs = disstrs + "、";
                    }
                }
                if (disname.Count == 1)
                {
                    disstrs = disstrs + "，病害占比为";
                }
                else
                {
                    disstrs = disstrs + "，病害占比分别为";
                }
                for (int i = 0; i < percent.Count; ++i)
                {
                    disstrs = disstrs + percent[i];
                    if (i < percent.Count - 1)
                    {
                        disstrs = disstrs + "、";
                    }
                }
                disstrs = disstrs + "。";
                currentSelection.TypeText(disstrs);
                hasstr = true;
            }

            if (!hasstr)
            {
                for (int i = 0; i < 18; ++i)
                {
                    currentSelection.TypeBackspace();
                }
                currentSelection.Delete();
            }
        }

        private static void WriteAll2Docx(MSWord.Application wordApp, string srcpath, List<RoadPartProjectClass> srclists)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\总报告模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            MSWord.Document wordDoc = null;
            MSWord.Selection currentSelection = null;

            string destdoc = srcinfo.DirectoryName + "\\总报告.docx";
            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;

            wordDoc.Paragraphs.Last.Range.Select();
            currentSelection = wordApp.Selection;
            string srcfname = null;

            srcfname = srcinfo.DirectoryName + "\\报告头.docx";
            currentSelection.InsertFile(srcfname);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            wordDoc.Fields.Update();
            Thread.Sleep(GlobalWord.wd_sleep_us);

            foreach (RoadPartProjectClass srclist in srclists)
            {
                currentSelection.TypeBackspace();
                string roadpartfname = "路段" + srclist.m_roadpart.m_id + "#"
                    + srclist.m_roadpart.m_roadinfo.m_code + "_"
                    + srclist.m_roadpart.m_roadinfo.m_name + "（"
                    + srclist.m_roadpart.m_startlocation + "-"
                    + srclist.m_roadpart.m_endlocation + "）";

                srcfname = srcinfo.DirectoryName + "\\路段统计\\" + roadpartfname + ".docx";
                currentSelection.InsertFile(srcfname);
                Thread.Sleep(GlobalWord.wd_sleep_us);

                wordDoc.Fields.Update();
                Thread.Sleep(GlobalWord.wd_sleep_us);
            }

            currentSelection.InsertNewPage();

            srcfname = srcinfo.DirectoryName + "\\结论.docx";
            currentSelection.InsertFile(srcfname);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            wordDoc.Fields.Update();
            Thread.Sleep(GlobalWord.wd_sleep_us);

            currentSelection.TypeParagraph();
            currentSelection.TypeParagraph();
            currentSelection.TypeText("（以下无正文）");
            currentSelection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // 写入 报告结论 和 养护建议
            string SuggestionStr = "本项目部本年度无大中修路段建议。";
            string ConclusionStr = "";
            MSWord.Tables ttables = wordDoc.Tables;
            MSWord.Table ttable = ttables[ttables.Count];
            if (ttable.Cell(1, 1).Range.Text.Contains("序号") && ttable.Cell(1, 2).Range.Text.Contains("路名") && ttable.Cell(1, 3).Range.Text.Contains("路段"))
            {
                string curstr = null;
                string oldstr = null;
                SuggestionStr = "本项目部建议进行中修及以上的路段为";
                for (int i = 2; i <= ttable.Rows.Count; ++i)
                {
                    curstr = ttable.Cell(i, 2).Range.Text.Replace("\r\a", "（")
                        + ttable.Cell(i, 3).Range.Text.Replace("\r\a", "），");
                    if (oldstr != curstr)
                    {
                        SuggestionStr += curstr;
                    }
                    oldstr = curstr;
                }
                SuggestionStr += "具体单元详见正文第7节";
            }

            string[] tstrs = File.ReadAllLines(srcinfo.DirectoryName + "\\结论.txt");
            for (int i = 0; i < tstrs.Length; ++i)
            {
                ConclusionStr += tstrs[i].Substring(tstrs[i].IndexOf("总体") + 2);
                if (i < tstrs.Length - 1)
                {
                    ConclusionStr += "\n";
                }
            }

            // 插入附录
            currentSelection.InsertNewPage();
            srcfname = srcinfo.DirectoryName + "\\附录.docx";
            currentSelection.InsertFile(srcfname);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            wordDoc.Fields.Update();
            Thread.Sleep(GlobalWord.wd_sleep_us);

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "报告结论")
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(ConclusionStr);
                    break;
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "报告养护建议")
                {
                    book.Range.Select();
                    currentSelection = wordApp.Selection;
                    currentSelection.TypeText(SuggestionStr);
                    break;
                }
            }

            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        // 第六章的，结论与分析评估；第七章的大中修路段建议
        private static void WriteSummary2Docx(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\路段结论模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            object[,] srcobj = null;
            MSWord.Document wordDoc = null;
            MSWord.Selection currentSelection = null;
            string tmpstr = null;

            if (File.Exists(srcinfo.DirectoryName + "\\结论.txt"))
            {
                File.Delete(srcinfo.DirectoryName + "\\结论.txt");
            }

            string destdoc = srcinfo.DirectoryName + "\\结论.docx";
            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;

            string CollectSrcXls = srcinfo.DirectoryName + "\\路段统计\\报告汇总.xlsx";
            MSExcel.Workbook CollectBook = excelApp.Workbooks.Open(CollectSrcXls, Type.Missing, true,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            MSExcel.Worksheet CollectDestSheet = CollectBook.Sheets["Sheet2"] as MSExcel.Worksheet;

            int tablecount = 0;
            bool isTypeText = true;
            string tableheader = null;
            MSExcel.Range srcxlsrange = null;

            wordDoc.Paragraphs.Last.Range.Select();
            currentSelection = wordApp.Selection;
            currentSelection.TypeText("结论与分析评估");
            object oStyleName = "标题 1";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            oStyleName = "报告表标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            //oStyleName = "报告附表1（隐藏）";
            //currentSelection.set_Style(oStyleName);
            //currentSelection.TypeParagraph();

            oStyleName = "报告图标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            tmpstr = "由PCI的统计图表可以看出，";
            srcxlsrange = CollectDestSheet.get_Range("A1:L7");
            srcobj = (object[,])srcxlsrange.Value2;
            isTypeText = GetSummaryStr(srcobj, "PCI", ref tmpstr);
            if (isTypeText)
            {
                currentSelection.TypeText("路面损坏状况（PCI）");
                oStyleName = "标题 2";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeParagraph();

                tableheader = "路面状况指数PCI及评价等级分类统计表";
                PastExcelTable2Word(wordDoc, srcxlsrange, currentSelection, ref tablecount, tableheader, "报告表标题3", false);
                PastExcelPic2Word(currentSelection, CollectDestSheet, 1, "路面损坏状况评价等级分布示意图", false);

                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr);
                currentSelection.TypeParagraph();

                tmpstr = GetSummaryPercentStr(srcobj, "路面破损状况");
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr + "\n");
                currentSelection.TypeParagraph();
            }

            tmpstr = "由RQI的统计图表可以看出，";
            srcxlsrange = CollectDestSheet.get_Range("A12:L18");
            srcobj = (object[,])srcxlsrange.Value2;
            isTypeText = GetSummaryStr(srcobj, "RQI", ref tmpstr);
            if (isTypeText)
            {
                currentSelection.TypeText("路面行驶质量（RQI）");
                oStyleName = "标题 2";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeParagraph();

                tableheader = "路面行驶质量指数RQI及评价等级分类统计表";
                PastExcelTable2Word(wordDoc, srcxlsrange, currentSelection, ref tablecount, tableheader, "报告表标题3", false);
                PastExcelPic2Word(currentSelection, CollectDestSheet, 2, "路面行驶质量评价等级分布示意图", false);

                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr);
                currentSelection.TypeParagraph();

                tmpstr = GetSummaryPercentStr(srcobj, "路面行驶质量");
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr + "\n");
                currentSelection.TypeParagraph();
            }

            tmpstr = "由PQI的统计图表可以看出，";
            srcxlsrange = CollectDestSheet.get_Range("A23:L29");
            srcobj = (object[,])srcxlsrange.Value2;
            isTypeText = GetSummaryStr(srcobj, "PQI", ref tmpstr);
            if (isTypeText)
            {
                currentSelection.TypeText("路面综合评价指标（PQI）");
                oStyleName = "标题 2";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeParagraph();

                tableheader = "路面综合评价指标PQI及评价等级分类统计表";
                PastExcelTable2Word(wordDoc, srcxlsrange, currentSelection, ref tablecount, tableheader, "报告表标题3", false);
                PastExcelPic2Word(currentSelection, CollectDestSheet, 3, "路面综合评价等级分布示意图", false);

                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr);
                currentSelection.TypeParagraph();

                tmpstr = GetSummaryPercentStr(srcobj, "路面综合评价");
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr + "\n");
                currentSelection.TypeParagraph();
            }

            tmpstr = "由TD的统计图表可以看出，";
            srcxlsrange = CollectDestSheet.get_Range("A34:L40");
            srcobj = (object[,])srcxlsrange.Value2;
            isTypeText = GetSummaryStr(srcobj, "TD", ref tmpstr);
            if (isTypeText)
            {
                currentSelection.TypeText("沥青路面抗滑能力（TD）");
                oStyleName = "标题 2";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeParagraph();

                tableheader = "路面抗滑能力TD及评价等级分类统计表";
                PastExcelTable2Word(wordDoc, srcxlsrange, currentSelection, ref tablecount, tableheader, "报告表标题3", false);
                PastExcelPic2Word(currentSelection, CollectDestSheet, 4, "路面抗滑能力评价等级分布示意图", false);

                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr);
                currentSelection.TypeParagraph();

                tmpstr = GetSummaryPercentStr(srcobj, "路面抗滑能力");
                currentSelection.TypeText(tmpstr);
                File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr + "\n");
                currentSelection.TypeParagraph();
            }

            // 第七章 大中修路段建议
            MSExcel.Worksheet YHJYSheet = CollectBook.Sheets["Sheet3"] as MSExcel.Worksheet;
            int userownum = GlobalExcel.judegeusedrow(YHJYSheet, 1, 2);
            MSExcel.Range srcrange = YHJYSheet.get_Range("A1:R" + userownum.ToString());
            List<string[]> listdest = new List<string[]>();
            srcobj = (object[,])srcrange.Value2;

            int YHCol = 0;
            if (_YHTypeDoc == 0)
                YHCol = 17;
            else if (_YHTypeDoc == 1)
                YHCol = 18;

            string[] strs = { "序号", "路名", "路段", "行车\n方向", "车道\n编号", "路面\n类型", "桩号（起）", "桩号（止）", 
                                "单元\n长度\n（m）", "PCI\n等级", "RQI\n等级", "TD\n等级", "养护对策"};
            listdest.Add(strs);

            int tmpval = 0;
            for (int i = 2; i <= userownum; ++i)
            {
                string tstr = srcobj[i, YHCol].ToString();
                if (tstr.Contains("中修") || tstr.Contains("大修"))
                {
                    string[] tstrs = new string[13];
                    tstrs[0] = listdest.Count.ToString();   //序号
                    tstrs[1] = srcobj[i, 3].ToString();     //路名
                    if (srcobj[i, 4] != null && srcobj[i, 5] != null)
                    {
                        tstrs[2] = srcobj[i, 4].ToString() + "~" + srcobj[i, 5].ToString();     //路段
                    }
                    else if (srcobj[i, 4] != null)
                    {
                        tstrs[2] = srcobj[i, 4].ToString();
                    }
                    else if (srcobj[i, 5] != null)
                    {
                        tstrs[2] = srcobj[i, 5].ToString();
                    }
                    tstrs[3] = srcobj[i, 6].ToString();     //行车方向
                    tstrs[4] = srcobj[i, 7].ToString();     //车道编号
                    tstrs[5] = srcobj[i, 8].ToString();     //路面类型

                    tmpval = Convert.ToInt32(srcobj[i, 9]);
                    tstrs[6] = "K" + (tmpval / 1000).ToString() + "+" + (tmpval % 1000).ToString("000");     //桩号-起

                    tmpval = Convert.ToInt32(srcobj[i, 10]);
                    tstrs[7] = "K" + (tmpval / 1000).ToString() + "+" + (tmpval % 1000).ToString("000");     //桩号-止

                    tstrs[8] = srcobj[i, 11].ToString();     //单元长度
                    tstrs[9] = srcobj[i, 13].ToString();     //PCI等级
                    tstrs[10] = srcobj[i, 14].ToString();    //RQI等级
                    tstrs[11] = srcobj[i, 15].ToString();    //TD等级
                    tstrs[12] = tstr;    //养护对策
                    listdest.Add(tstrs);
                }
            }

            currentSelection.TypeText("大中修路段建议");
            oStyleName = "标题 1";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            oStyleName = "报告表标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            //oStyleName = "报告附表1（隐藏）";
            //currentSelection.set_Style(oStyleName);
            //currentSelection.TypeParagraph();

            oStyleName = "报告图标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            // 存在大中修
            if (listdest.Count > 1)
            {
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText("根据本次道路检测的基础数据，结合道路现场，综合评定以下单元为中修及以上路段单元。");
                currentSelection.TypeParagraph();

                oStyleName = "报告表标题3";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText("大中修路段单元清单");
                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);

                MSWord.Table table = currentSelection.Tables.Add(currentSelection.Range, listdest.Count, 13, Type.Missing, Type.Missing);
                for (int i = 0; i < listdest.Count; ++i)
                {
                    for (int j = 0; j < 13; ++j)
                    {
                        table.Cell(i + 1, j + 1).Range.Text = listdest[i][j];
                    }
                }

                table.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
                table.AllowAutoFit = false;
                table.LeftPadding = 0.0f;
                table.RightPadding = 0.0f;
                table.TopPadding = 0.0f;
                table.BottomPadding = 0.0f;
                table.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100.0f;
                table.Columns.PreferredWidth = 6.0f;
                table.Columns[1].PreferredWidth = 5.0f;
                table.Columns[2].PreferredWidth = 9.4f;
                table.Columns[3].PreferredWidth = 13.0f;
                table.Columns[4].PreferredWidth = 6.4f;
                table.Columns[5].PreferredWidth = 4.8f;
                table.Columns[6].PreferredWidth = 4.8f;
                table.Columns[7].PreferredWidth = 8.8f;
                table.Columns[8].PreferredWidth = 8.8f;
                table.Columns[13].PreferredWidth = 15.0f;

                wordDoc.Paragraphs.Last.Range.Select();
                currentSelection = wordApp.Selection;
                oStyleName = "报告表下空行";
                currentSelection.set_Style(oStyleName);
            }
            else
            {
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText("根据本次道路检测的基础数据，结合道路现场，综合评定本项目部本年度无建议中修及以上路段单元。");
                currentSelection.TypeParagraph();
            }

            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                FromatTable(wordApp, temptable, 0.6f);
            }

            wordDoc.Save();

            CollectBook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        // 报告的开头部分
        private static void WriteHeader2Docx(MSWord.Application wordApp, string srcpath,
            ProjectProjectClass tproject, ReportProjectClass treport)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\报告头模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            MSWord.Table curtable = null;
            MSWord.Range wordrange = null;
            MSWord.Document wordDoc = null;
            string destdoc = srcinfo.DirectoryName + "\\报告头.docx";
            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wordDoc.SpellingChecked = false;
            wordDoc.ShowSpellingErrors = false;

            object oStyleName = "报告表格内容（通用居中 小五）";

            Dictionary<string, string> datas = new Dictionary<string, string>();
            datas.Add("{检测（分）报告编号}", treport.m_report.m_report_num);
            datas.Add("{检测（分）报告名称}", treport.m_report.m_report_name);
            datas.Add("{委托单位}", tproject.m_project.m_entrust_client);
            datas.Add("{项目名称}", tproject.m_project.m_project_name);
            datas.Add("{报告日期}", tproject.m_project.m_date.Insert(6, "月").Insert(4, "年") + "日");
            datas.Add("{合同/委托编号}", tproject.m_project.m_contract_num == null ? tproject.m_project.m_entrust_serial : tproject.m_project.m_contract_num);
            datas.Add("{合同编号}", tproject.m_project.m_contract_num);
            datas.Add("{委托编号}", tproject.m_project.m_entrust_serial);
            datas.Add("{委托日期}", tproject.m_project.m_entrust_date.Insert(6, "月").Insert(4, "年") + "日");
            datas.Add("{检测起始日期}", tproject.m_project.m_testing_start_date.Insert(6, "月").Insert(4, "年") + "日");
            datas.Add("{检测终止日期}", tproject.m_project.m_testing_end_date.Insert(6, "月").Insert(4, "年") + "日");

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

            object myFind = wordDoc.Content.Find;
            foreach (var item in datas)
            {
                object findText = item.Key;
                object replaceText = item.Value;
                Parameters[0] = findText;
                Parameters[9] = replaceText;
                while (true)
                {
                    try
                    {
                        myFind.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, myFind, Parameters);
                        break;
                    }
                    catch (Exception ex)
                    {
                        Thread.Sleep(GlobalWord.wd_sleep_us);
                    }
                }
            }

            #region 签字表
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "签字表")
                {
                    wordrange = book.Range;
                    curtable = wordrange.Tables.Add(wordrange, treport.m_personList.Count + 1, 5);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            curtable.Cell(1, 1).Range.Text = "岗位";
            curtable.Cell(1, 2).Range.Text = "姓名";
            curtable.Cell(1, 3).Range.Text = "职业资格证书编号";
            curtable.Cell(1, 4).Range.Text = "职称";
            curtable.Cell(1, 5).Range.Text = "签字";
            int trowcnt = 3;
            foreach (TestingPersonClass tperson in treport.m_personList)
            {
                if (tperson.m_post.Contains("负责"))
                {
                    curtable.Cell(2, 1).Range.Text = "项目负责人";
                    curtable.Cell(2, 2).Range.Text = tperson.m_name;
                    curtable.Cell(2, 3).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(2, 4).Range.Text = tperson.m_title;
                }
                else if (tperson.m_post.Contains("编写"))
                {
                    curtable.Cell(treport.m_personList.Count - 1, 1).Range.Text = "报告编写人";
                    curtable.Cell(treport.m_personList.Count - 1, 2).Range.Text = tperson.m_name;
                    curtable.Cell(treport.m_personList.Count - 1, 3).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(treport.m_personList.Count - 1, 4).Range.Text = tperson.m_title;
                }
                else if (tperson.m_post.Contains("审核"))
                {
                    curtable.Cell(treport.m_personList.Count, 1).Range.Text = "报告审核人";
                    curtable.Cell(treport.m_personList.Count, 2).Range.Text = tperson.m_name;
                    curtable.Cell(treport.m_personList.Count, 3).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(treport.m_personList.Count, 4).Range.Text = tperson.m_title;
                }
                else if (tperson.m_post.Contains("批准"))
                {
                    curtable.Cell(treport.m_personList.Count + 1, 1).Range.Text = "报告批准人";
                    curtable.Cell(treport.m_personList.Count + 1, 2).Range.Text = tperson.m_name;
                    curtable.Cell(treport.m_personList.Count + 1, 3).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(treport.m_personList.Count + 1, 4).Range.Text = tperson.m_title;
                }
                else
                {
                    curtable.Cell(trowcnt, 2).Range.Text = tperson.m_name;
                    curtable.Cell(trowcnt, 3).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(trowcnt, 4).Range.Text = tperson.m_title;
                    ++trowcnt;
                }
            }
            curtable.Cell(3, 1).Merge(curtable.Cell(treport.m_personList.Count - 2, 1));
            curtable.Cell(3, 1).Range.Text = "项目主要参加人员";

            curtable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            curtable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            curtable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            curtable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            curtable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            curtable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            curtable.AllowAutoFit = false;
            curtable.LeftPadding = 0.0f;
            curtable.RightPadding = 0.0f;
            curtable.TopPadding = 0.0f;
            curtable.BottomPadding = 0.0f;

            curtable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            curtable.Columns[1].PreferredWidth = 19.6f;
            curtable.Columns[2].PreferredWidth = 19f;
            curtable.Columns[3].PreferredWidth = 27.9f;
            curtable.Columns[4].PreferredWidth = 14.5f;
            curtable.Columns[5].PreferredWidth = 18.8f;

            curtable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            curtable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                curtable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                curtable.Cell(1, 1).Select();
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            curtable.Rows.AllowBreakAcrossPages = 0;
            curtable.ApplyStyleHeadingRows = true;

            float height = 1.5f;
            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            curtable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            #endregion

            #region 报告编制说明表
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "报告编制说明")
                {
                    wordrange = book.Range;
                    curtable = wordrange.Tables.Add(wordrange, tproject.m_reportlist.Count + 1, 4);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            trowcnt = 1;
            curtable.Cell(trowcnt, 1).Range.Text = "序号";
            curtable.Cell(trowcnt, 2).Range.Text = "报告编号";
            curtable.Cell(trowcnt, 3).Range.Text = "报告名称";
            curtable.Cell(trowcnt, 4).Range.Text = "项目部";
            foreach (ReportProjectClass ttreport in tproject.m_reportlist)
            {
                ++trowcnt;
                curtable.Cell(trowcnt, 1).Range.Text = (trowcnt - 1).ToString();
                curtable.Cell(trowcnt, 2).Range.Text = ttreport.m_report.m_report_num;
                curtable.Cell(trowcnt, 3).Range.Text = ttreport.m_report.m_report_name;
                curtable.Cell(trowcnt, 4).Range.Text = ttreport.m_report.m_project_name;
                if (ttreport.m_report.m_id == treport.m_report.m_id)
                {
                    for (int j = 1; j <= 4; ++j)
                    {
                        curtable.Cell(trowcnt, j).Range.Font.Bold = 1;
                    }
                }
            }
            curtable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            curtable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            curtable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            curtable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            curtable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            curtable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            curtable.AllowAutoFit = false;
            curtable.LeftPadding = 0.0f;
            curtable.RightPadding = 0.0f;
            curtable.TopPadding = 0.0f;
            curtable.BottomPadding = 0.0f;

            curtable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            curtable.Columns[1].PreferredWidth = 6.2f;
            curtable.Columns[2].PreferredWidth = 29.6f;
            curtable.Columns[3].PreferredWidth = 40.7f;
            curtable.Columns[4].PreferredWidth = 23.4f;

            curtable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            curtable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                curtable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                curtable.Cell(1, 1).Select();
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            curtable.Rows.AllowBreakAcrossPages = 0;
            curtable.ApplyStyleHeadingRows = true;

            height = 0.9f;
            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            curtable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            #endregion

            string[] roadpartgrade = { "快速路", "主干路", "次干路", "支路", "合计" };
            int[] roadpartnum = { 0, 0, 0, 0, 0 };
            double[] roadpartlength = { 0.0, 0.0, 0.0, 0.0, 0.0 };

            #region 道路清单
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "道路清单")
                {
                    wordrange = book.Range;
                    curtable = wordrange.Tables.Add(wordrange, treport.m_roadpartlist.Count + 1, 7);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            trowcnt = 1;
            curtable.Cell(trowcnt, 1).Range.Text = "序号";
            curtable.Cell(trowcnt, 2).Range.Text = "道路名称";
            curtable.Cell(trowcnt, 3).Range.Text = "路段";
            curtable.Cell(trowcnt, 4).Range.Text = "道路等级";
            curtable.Cell(trowcnt, 5).Range.Text = "车道";
            curtable.Cell(trowcnt, 6).Range.Text = "道路里程\n（m）";
            curtable.Cell(trowcnt, 7).Range.Text = "检测长度\n（m）";
            for (int j = 1; j <= 7; ++j)
            {
                curtable.Cell(trowcnt, j).Range.Font.Bold = 1;
            }
            foreach (RoadPartProjectClass troadpart in treport.m_roadpartlist)
            {
                ++trowcnt;
                curtable.Cell(trowcnt, 1).Range.Text = (trowcnt - 1).ToString();
                curtable.Cell(trowcnt, 2).Range.Text = troadpart.m_roadpart.m_roadinfo.m_name;
                curtable.Cell(trowcnt, 3).Range.Text = troadpart.m_roadpart.m_startlocation + "~" + troadpart.m_roadpart.m_endlocation;
                curtable.Cell(trowcnt, 4).Range.Text = troadpart.m_roadpart.m_part_grade;
                curtable.Cell(trowcnt, 5).Range.Text = troadpart.m_roadpart.m_type;
                curtable.Cell(trowcnt, 6).Range.Text = troadpart.m_roadpart.m_roadinfo.m_length;
                int sumlen = 0;
                foreach (LaneProjectClass tlane in troadpart.m_lanelist)
                {
                    sumlen += Math.Abs(Convert.ToInt16(tlane.m_lane.m_endmile) - Convert.ToInt16(tlane.m_lane.m_startmile));
                }
                curtable.Cell(trowcnt, 7).Range.Text = sumlen.ToString();

                if (troadpart.m_roadpart.m_part_grade == "快速路")
                {
                    ++roadpartnum[0];
                    roadpartlength[0] += Convert.ToDouble(troadpart.m_roadpart.m_length);
                }
                else if (troadpart.m_roadpart.m_part_grade == "主干路")
                {
                    ++roadpartnum[1];
                    roadpartlength[1] += Convert.ToDouble(troadpart.m_roadpart.m_length);
                }
                else if (troadpart.m_roadpart.m_part_grade == "次干路")
                {
                    ++roadpartnum[2];
                    roadpartlength[2] += Convert.ToDouble(troadpart.m_roadpart.m_length);
                }
                else if (troadpart.m_roadpart.m_part_grade == "支路")
                {
                    ++roadpartnum[3];
                    roadpartlength[3] += Convert.ToDouble(troadpart.m_roadpart.m_length);
                }
                ++roadpartnum[4];
                roadpartlength[4] += Convert.ToDouble(troadpart.m_roadpart.m_length);
            }

            curtable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            curtable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            curtable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            curtable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            curtable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            curtable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            curtable.AllowAutoFit = false;
            curtable.LeftPadding = 0.0f;
            curtable.RightPadding = 0.0f;
            curtable.TopPadding = 0.0f;
            curtable.BottomPadding = 0.0f;

            curtable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            curtable.Columns[1].PreferredWidth = 6.8f;
            curtable.Columns[2].PreferredWidth = 14.4f;
            curtable.Columns[3].PreferredWidth = 27.5f;
            curtable.Columns[4].PreferredWidth = 13.7f;
            curtable.Columns[5].PreferredWidth = 14.4f;
            curtable.Columns[6].PreferredWidth = 11.4f;
            curtable.Columns[7].PreferredWidth = 11.4f;

            curtable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            curtable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                curtable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                curtable.Cell(1, 1).Select();
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            curtable.Rows.AllowBreakAcrossPages = 0;
            curtable.ApplyStyleHeadingRows = true;

            height = 0.7f;
            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            curtable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            #endregion

            #region 不同道路等级检测情况汇总表
            int trownum = 0;
            for (int i = 0; i < 5; ++i)
            {
                if (roadpartnum[i] != 0)
                {
                    ++trownum;
                }
            }
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "不同道路等级检测情况汇总表")
                {
                    wordrange = book.Range;
                    curtable = wordrange.Tables.Add(wordrange, trownum + 1, 3);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            trowcnt = 1;
            curtable.Cell(trowcnt, 1).Range.Text = "道路等级";
            curtable.Cell(trowcnt, 2).Range.Text = "检测路段（个）";
            curtable.Cell(trowcnt, 3).Range.Text = "道路长度（m）";
            for (int j = 1; j <= 3; ++j)
            {
                curtable.Cell(trowcnt, j).Range.Font.Bold = 1;
            }
            for (int i = 0; i < 5; ++i)
            {
                if (roadpartnum[i] != 0)
                {
                    ++trowcnt;
                    curtable.Cell(trowcnt, 1).Range.Text = roadpartgrade[i];
                    curtable.Cell(trowcnt, 2).Range.Text = roadpartnum[i].ToString();
                    curtable.Cell(trowcnt, 3).Range.Text = roadpartlength[i].ToString();
                }
            }
            curtable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            curtable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            curtable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            curtable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            curtable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            curtable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            curtable.AllowAutoFit = false;
            curtable.LeftPadding = 0.0f;
            curtable.RightPadding = 0.0f;
            curtable.TopPadding = 0.0f;
            curtable.BottomPadding = 0.0f;

            curtable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            curtable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                curtable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                curtable.Cell(1, 1).Select();
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            curtable.Rows.AllowBreakAcrossPages = 0;
            curtable.ApplyStyleHeadingRows = true;

            height = 0.7f;
            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            curtable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            #endregion

            #region 主要检测人员及分工表
            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "主要检测人员及分工表")
                {
                    wordrange = book.Range;
                    curtable = wordrange.Tables.Add(wordrange, treport.m_personList.Count + 1, 5);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            trowcnt = 1;
            curtable.Cell(trowcnt, 1).Range.Text = "序号";
            curtable.Cell(trowcnt, 2).Range.Text = "姓名";
            curtable.Cell(trowcnt, 3).Range.Text = "职称";
            curtable.Cell(trowcnt, 4).Range.Text = "职业资格证书编号";
            curtable.Cell(trowcnt, 5).Range.Text = "项目分工";
            foreach (TestingPersonClass tperson in treport.m_personList)
            {
                ++trowcnt;
                curtable.Cell(trowcnt, 1).Range.Text = (trowcnt - 1).ToString();
                curtable.Cell(trowcnt, 2).Range.Text = tperson.m_name;
                curtable.Cell(trowcnt, 3).Range.Text = tperson.m_title;
                curtable.Cell(trowcnt, 4).Range.Text = tperson.m_CertificateNo;
                curtable.Cell(trowcnt, 5).Range.Text = tperson.m_duty;
            }
            curtable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
            curtable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            curtable.Range.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            curtable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            curtable.Range.Paragraphs.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;

            curtable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
            curtable.AllowAutoFit = false;
            curtable.LeftPadding = 0.0f;
            curtable.RightPadding = 0.0f;
            curtable.TopPadding = 0.0f;
            curtable.BottomPadding = 0.0f;

            curtable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
            curtable.Columns[1].PreferredWidth = 8.7f;
            curtable.Columns[2].PreferredWidth = 10.7f;
            curtable.Columns[3].PreferredWidth = 16.1f;
            curtable.Columns[4].PreferredWidth = 30.2f;
            curtable.Columns[5].PreferredWidth = 34f;

            curtable.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth150pt;
            curtable.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            curtable.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;

            try
            {
                curtable.Rows[1].HeadingFormat = -1;
            }
            catch
            {
                curtable.Cell(1, 1).Select();
                GlobalWord.wordAppHeadingFormat(wordApp);
            }
            curtable.Rows.AllowBreakAcrossPages = 0;
            curtable.ApplyStyleHeadingRows = true;

            height = 0.7f;
            height = (float)(height * 0.3937008 * 72); // 单位从厘米转换为磅
            curtable.Rows.SetHeight(height, Microsoft.Office.Interop.Word.WdRowHeightRule.wdRowHeightAtLeast);
            #endregion

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        private static string GetSummaryPercentStr(object[,] srcobj, string indexValStr)
        {
            string tmpstr = "总体而言，";

            double perA = Convert.ToDouble(srcobj[7, 4]);
            double perB = Convert.ToDouble(srcobj[7, 6]);
            double perC = Convert.ToDouble(srcobj[7, 8]);
            double perD = Convert.ToDouble(srcobj[7, 10]);
            double perABC = perA + perB + perC;

            tmpstr = tmpstr + perABC.ToString("0.00") + "%的检测道路的" + indexValStr + "处于C级以上";
            if (perD > 0)
            {
                tmpstr = tmpstr + "，" + perD.ToString("0.00") + "%处于D级";
            }
            tmpstr = tmpstr + "。";

            return tmpstr;
        }
        private static bool GetSummaryStr(object[,] srcobj, string indexValStr, ref string tmpstr)
        {
            bool res = true;

            List<string> tmpstrlist = new List<string>();
            List<string> tmpstrlist1 = new List<string>();
            List<string> tmpstrlist2 = new List<string>();

            if (indexValStr == "TD")
            {
                if (Convert.ToDouble(srcobj[7, 11]) <= 0.0)
                {
                    res = false;
                    return res;
                }
            }

            tmpstr = tmpstr + "检测道路总体" + indexValStr + "均值为" + Convert.ToDouble(srcobj[7, 11]).ToString("0.00") + "。其中，";
            if (Convert.ToDouble(srcobj[3, 2]) > 0)
            {
                tmpstrlist.Add("快速路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[3, 11]).ToString("0.00"));
                }
                catch (System.Exception) { }
                tmpstrlist2.Add(srcobj[3, 12].ToString() + "级");
            }

            if (Convert.ToDouble(srcobj[4, 2]) > 0)
            {
                tmpstrlist.Add("主干路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[4, 11]).ToString("0.00"));
                }
                catch (System.Exception ) { }
                tmpstrlist2.Add(srcobj[4, 12].ToString() + "级");
            }

            if (Convert.ToDouble(srcobj[5, 2]) > 0)
            {
                tmpstrlist.Add("次干路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[5, 11]).ToString("0.00"));
                }
                catch (System.Exception ) { }
                tmpstrlist2.Add(srcobj[5, 12].ToString() + "级");
            }

            if (Convert.ToDouble(srcobj[6, 2]) > 0)
            {
                tmpstrlist.Add("支路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[6, 11]).ToString("0.00"));
                }
                catch (System.Exception ex) { }
                tmpstrlist2.Add(srcobj[6, 12].ToString() + "级");
            }

            if (indexValStr == "TD")
            {
                if (tmpstrlist.Count == 1 && tmpstrlist[0] == "支路")
                {
                    res = false;
                    return res;
                }
            }

            if (tmpstrlist.Count > 0)
            {
                if (tmpstrlist.Count == 1)
                {
                    tmpstr = tmpstr + tmpstrlist[0];
                }
                else
                {
                    for (int i = 0; i < tmpstrlist.Count - 1; ++i)
                    {
                        tmpstr = tmpstr + tmpstrlist[i];
                        if (i == tmpstrlist.Count - 2)
                        {
                            tmpstr = tmpstr + "和";
                        }
                        else
                        {
                            tmpstr = tmpstr + "、";
                        }
                    }
                    tmpstr = tmpstr + tmpstrlist[tmpstrlist.Count - 1];
                }
            }

            if (tmpstrlist1.Count > 0)
            {
                if (tmpstrlist1.Count == 1)
                {
                    tmpstr = tmpstr + "的" + indexValStr + "平均值为";
                    tmpstr = tmpstr + tmpstrlist1[0];
                }
                else
                {
                    tmpstr = tmpstr + "的" + indexValStr + "平均值分别为";
                    for (int i = 0; i < tmpstrlist1.Count - 1; ++i)
                    {
                        tmpstr = tmpstr + tmpstrlist1[i];
                        if (i == tmpstrlist1.Count - 2)
                        {
                            tmpstr = tmpstr + "和";
                        }
                        else
                        {
                            tmpstr = tmpstr + "、";
                        }
                    }
                    tmpstr = tmpstr + tmpstrlist1[tmpstrlist1.Count - 1];
                }
            }

            tmpstr = tmpstr + "，";

            if (tmpstrlist2.Count > 0)
            {
                if (tmpstrlist2.Count == 1)
                {
                    tmpstr = tmpstr + "为" + tmpstrlist2[0] + "。";
                }
                else
                {
                    bool isOneGrade = true;
                    for (int i = 0; i < tmpstrlist2.Count; ++i)
                    {
                        if (tmpstrlist2[i] != tmpstrlist2[0])
                        {
                            isOneGrade = false;
                            break;
                        }
                    }
                    if (isOneGrade)
                    {
                        tmpstr = tmpstr + "均为" + tmpstrlist2[0] + "。";
                    }
                    else
                    {
                        tmpstr = tmpstr + "分别为";
                        for (int i = 0; i < tmpstrlist2.Count; ++i)
                        {
                            tmpstr = tmpstr + tmpstrlist2[i];
                            if (i == tmpstrlist2.Count - 1)
                            {
                                tmpstr = tmpstr + "。";
                            }
                            else
                            {
                                tmpstr = tmpstr + "、";
                            }
                        }
                    }
                }
            }

            return res;
        }
        #endregion
    }
}
