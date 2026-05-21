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
using DevExpress.XtraPrinting;
using System.ServiceModel.Channels;
using NPOI.OpenXmlFormats.Dml.Chart;
using ICSharpCode.SharpZipLib.Zip;
using DevExpress.XtraPrinting.Native;
using System.ServiceModel.Description;
using System.Windows.Forms;
using System.Diagnostics.Eventing.Reader;
using Spire.Pdf;
using SqlSugar;
using System.Globalization;
using Framework.Other;

namespace XRDataProcess
{
    class MyWordCity
    {
        private static string _RoadPName;//路段名
        private static string _RoadLName;//路线名
        private static string _RoadType;//路线名

        private static string _Smile;//起点桩号
        private static string _Emile;//终点桩号
        private static string _RoadToltalArea;//路面总面积
        private static string _RoadPCIdegree;//路面PCI等级
        private static string _RoadRQIdegree;//路面RQI等级
        private static string _RoadPQIdegree;//路面PQI等级
        private static string[] _xlstype = { "IRI", "MTD", "PCI", "病害统计", "PQI", "Rut" };
        private static string[] _sidetypeL = { "左一幅", "左二幅", "左三幅", "左四幅", "左五幅", "左六幅", "左七幅", "左八幅" };
        private static string[] _sidetypeR = { "右一幅", "右二幅", "右三幅", "右四幅", "右五幅", "右六幅", "右七幅", "右八幅" };
        private static string[][] _sidetype;
        private static bool[][] _sidef;
        private static string[][][] _xlsnames;
        private static int _sidenum = 8;

        public static void ExportWord(MSWord.Application wordApp, MSExcel.Application excelApp, MSWord.Document wordDoc, MSExcel.Workbook excelXls)
        {
            ExportWordString(wordDoc);
            ExportWordTable(wordApp, excelApp, wordDoc, excelXls);
            //更新目录
            int count = wordDoc.TablesOfContents.Count;
            for (int i = 0; i < count; i++)
            {
                wordDoc.TablesOfContents[i + 1].Update();
            }
        }

        //简单的字段替换
        public static void ExportWordString(MSWord.Document wordDoc)
        {
            Dictionary<string, string> datas = new Dictionary<string, string>();
            datas.Add("{路段名}", _RoadPName);
            datas.Add("{路线名}", _RoadLName);
            datas.Add("{材质}", _RoadType);

            datas.Add("{起点桩号}", _Smile);
            datas.Add("{终点桩号}", _Emile);
            datas.Add("{路段总面积}", _RoadToltalArea);
            datas.Add("{路面破损评价等级}", _RoadPCIdegree);
            datas.Add("{路面行驶质量评价等级}", _RoadRQIdegree);

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
                catch (Exception)
                {
                    Thread.Sleep(500);
                    ++kk;
                }

                if (kk > 10) break;
            }
        }

        //写入表
        public static void ExportWordTable(MSWord.Application wordApp, MSExcel.Application excelApp,
            MSWord.Document wordDoc, MSExcel.Workbook excelbook)
        {
            MSExcel.Worksheet excelsheet = null;
            MSExcel.Range excelrange = null;
            MSWord.Range wordrange = null;

            #region 各指标评价等级表
            string[] sheetnames = { "PCI统计表", "IRI统计表", "Rut统计表", "MTD统计表", "PQI统计表" };
            string[] tableheadpre = { "表2.2-", "表2.5-", "表2.6-", "表2.7-", "表2.8-" };
            string[] tableheadapp = { "PCI评价等级统计表", "路面平整度评价等级统计表", "路面车辙深度评价等级统计表", "路面构造深度统计表", "路面综合评价指数（PQI）统计表" };
            int[] colnums = { 4, 7, 7, 6, 8 };
            int excelrow = 0;
            int colval = 0;
            String wordstr = null;
            int tableidx = 0;
            for (int sidx = 0; sidx < sheetnames.Length; ++sidx)
            {
                colval = 0;
                tableidx = 1;
                excelsheet = excelbook.Sheets[sheetnames[sidx]] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 3, 3);

                wordrange = GlobalWord.GetMarkRange(wordDoc, sheetnames[sidx]);
                GlobalWord.wordAppGoTo(wordApp, wordrange);
                for (int i = 0; i < 2; ++i)
                {
                    for (int j = 0; j < _sidenum; ++j)
                    {
                        if (_sidef[i][j])
                        {
                            wordstr = String.Format("{0}2:{1}{2}",
                                GlobalExcel.GetCol((char)('A' + colval)), GlobalExcel.GetCol((char)('A' + colval + colnums[sidx] - 1)), excelrow);
                            excelrange = excelsheet.get_Range(String.Format("{0}2:{1}{2}",
                                GlobalExcel.GetCol((char)('A' + colval)), GlobalExcel.GetCol((char)('A' + colval + colnums[sidx] - 1)), excelrow));
                            System.Windows.Forms.Clipboard.Clear();
                            excelrange.Copy();

                            wordstr = string.Format("{0}{1:0} {2}({3}){4}", tableheadpre[sidx], tableidx++, _RoadLName, _sidetype[i][j], tableheadapp[sidx]);
                            GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                            GlobalWord.wordAppTypeText(wordApp, wordstr);
                            GlobalWord.wordAppSelectionPaste(wordApp);

                            wordstr = ((MSExcel.Range)excelsheet.Cells[1, colval + 1]).Text.ToString();
                            wordstr = ReSummary(wordstr);
                            GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphLeft);
                            GlobalWord.wordAppTypeText(wordApp, wordstr);
                            colval += colnums[sidx];
                        }
                    }
                }
            }
            #endregion

            #region 路段评价统计表
            string[] sheetnames2 = { "路段病害面积统计表", "路段IRI统计表", "路段Rut统计表", "路段MTD统计表", "路段PQI统计表", "路段合计" };
            string[] tableheadpre2 = { "表3.1-1 ", "表3.4-1 ", "表3.5-1 ", "表3.6-1 ", "表3.7-1 ", "表3.8-1 " };
            string[] tableheadapp2 = { "外观病害面积统计表", "道路各检测路段路面平整度评价等级表", "道路各检测路段路面车辙评价等级表", "道路各检测路段构造深度评价表", "道路各评定路段路面综合评价指数（PQI）", "路段评价汇总统计表" };
            int[] colnums2 = { 5, 7, 8, 7, 7, 7 };
            for (int sidx = 0; sidx < sheetnames2.Length; ++sidx)
            {
                excelsheet = excelbook.Sheets[sheetnames2[sidx]] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 3, 3);

                wordrange = GlobalWord.GetMarkRange(wordDoc, sheetnames2[sidx]);
                GlobalWord.wordAppGoTo(wordApp, wordrange);

                excelrange = excelsheet.get_Range(String.Format("A2:{0}{1}",
                    GlobalExcel.GetCol((char)('A' + colnums2[sidx] - 1)), excelrow));
                System.Windows.Forms.Clipboard.Clear();
                excelrange.Copy();

                wordstr = string.Format("{0}{1}{2}", tableheadpre2[sidx], _RoadPName, tableheadapp2[sidx]);
                GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                GlobalWord.wordAppTypeText(wordApp, wordstr);
                GlobalWord.wordAppSelectionPaste(wordApp);
            }
            #endregion

            #region 表格格式化
            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                temptable.Range.Font.Name = "宋体";
                temptable.Range.Font.Size = 12;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
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

        public static void ExportAppWordTable(MSWord.Application wordApp, MSExcel.Application excelApp, MSWord.Document wordDoc, int sidx)
        {
            MSWord.Range wordrange = null;

            #region 将10m的报表原始数据粘贴到word附件中
            string[] bookmarkapp = { "病害数据", "IRI数据", "Rut数据", "MTD数据", "PQI数据" };
            string[] typeapp = { "原始10米路面病害数据", "原始10米平整度数据", "原始10米车辙数据", "原始10米构造深度数据", "原始10米综合评价（PQI）数据" };
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            string fname = null, wordstr = null;
            int[] srcsheetidxval = { 3, 0, 5, 1, 4 };
            int[] srcsortrow = { 3, 4, 4, 4, 3 };
            int[] srccopyrow = { 2, 2, 2, 2, 2 };
            int srcrownum = 0;
            int srccolnum = 0;

            wordrange = GlobalWord.GetMarkRange(wordDoc, bookmarkapp[sidx]);
            GlobalWord.wordAppGoTo(wordApp, wordrange);
            for (int i = 0; i < 2; ++i)
            {
                for (int j = 0; j < _sidenum; ++j)
                {
                    if (_sidef[i][j])
                    {
                        fname = _xlsnames[i][j][srcsheetidxval[sidx]].Replace("\\100米报表\\", "\\10米报表\\").Replace("_100m.xlsx", "_10m.xlsx");
                        fname = _xlsnames[i][j][srcsheetidxval[sidx]].Replace("\\100米报表\\", "\\10米报表\\").Replace("_100m.xls", "_10m.xls");
                        srcbook = excelApp.Workbooks.Open(fname, Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        if (sidx < 1)
                        {
                            foreach (MSExcel.Worksheet tsheet in srcbook.Sheets)
                            {
                                if (tsheet.Name.Contains("病害列表"))
                                {
                                    srcsheet = tsheet;
                                }
                            }
                        }
                        else
                        {
                            srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                        }
                        srcrownum = GlobalExcel.judegeusedrow(srcsheet, 1, srcsortrow[sidx]);
                        srccolnum = GlobalExcel.judegeusedcol(srcsheet, srcrownum - 2 > srcsortrow[sidx] ? srcrownum - 2 : srcsortrow[sidx], 1);
                        srcrange = srcsheet.get_Range(String.Format("A{0}:{1}{2}", srccopyrow[sidx], GlobalExcel.GetCol((char)(srccolnum + 'A' - 1)), srcrownum));
                        System.Windows.Forms.Clipboard.Clear();
                        srcrange.Copy();

                        wordstr = string.Format("\r\n{0}{1}", _sidetype[i][j], typeapp[sidx]);
                        if (sidx < 1)
                        {
                            wordstr += "单位：㎡";
                        }
                        GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphLeft);
                        GlobalWord.wordAppTypeText(wordApp, wordstr);
                        GlobalWord.wordAppSelectionPaste(wordApp);

                        srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    }
                }
            }
            #endregion

            #region 表格格式化
            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                temptable.Range.Font.Name = "宋体";
                temptable.Range.Font.Size = 12;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
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

        public static String ReSummary(String wordstr)
        {
            string res = null;
            string[] s = wordstr.Split('，');
            foreach (string str in s)
            {
                if (str.Contains("总数0%"))
                {
                    continue;
                }
                else
                {
                    res = string.Format("{0}，{1}", res, str);
                }
            }
            res = res + "。";
            return res.Remove(0, 1);
        }

        public static void OutputDoc(MSWord.Application wordApp, MSExcel.Application excelApp, List<string> _ExcelPathList)
        {

            int generation;
            string destpath = _ExcelPathList[0].Substring(0, _ExcelPathList[0].LastIndexOf("\\"));
            destpath = destpath.Substring(0, destpath.LastIndexOf("\\")) + "\\";
            string[] srcdocname = { "附件二道路路面病害明细表（10米）", "附件三道路平整度检测数据表（10米）", "附件四道路车辙检测数据表（10米）", "附件五道路构造深度检测数据表（10米）", "附件六道路综合评价检测数据表（10米）" };
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\报告主体（100米）.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = destpath + "报告主体（100米）.docx";

            LoadIniParm();
            GetXlsNames(_ExcelPathList);

            MSExcel.Workbook destbook = null;
            MSWord.Document wordDoc = null;

            string srcxls = string.Format(@"{0}\报告模板\城镇道路\统计报表.xlsx", System.Windows.Forms.Application.StartupPath);
            destbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            destbook.SaveAs(destdoc.Replace(".docx", ".xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CreateWordTable(excelApp, destbook);

            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWord(wordApp, excelApp, wordDoc, destbook);

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            destbook.Save();
            destbook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();

            for (int i = 0; i < srcdocname.Length; ++i)
            {
                srcdoc = string.Format(@"{0}\报告模板\城镇道路\{1}.docx", System.Windows.Forms.Application.StartupPath, srcdocname[i]);
                destdoc = string.Format("{0}{1}.docx", destpath, srcdocname[i]);

                wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量

                ExportWordString(wordDoc);
                ExportAppWordTable(wordApp, excelApp, wordDoc, i);

                wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
                generation = System.GC.GetGeneration(wordApp);
                System.GC.Collect(generation);//垃圾回收
                System.GC.WaitForPendingFinalizers();
                wordDoc = null;
            }
        }

        private static void GetXlsNames(List<string> _ExcelPathList)
        {
            for (int i = 0; i < 2; ++i)
            {
                for (int j = 0; j < _sidenum; ++j)
                {
                    //if (!_sidef[i][j])
                    //{
                    //    break;
                    //}
                    if (_sidef[i][j])
                    {
                        for (int k = 0; k < _xlstype.Length; ++k)
                        {
                            foreach (string str in _ExcelPathList)
                            {
                                if (str.Contains(_sidetype[i][j]) && str.Contains(_xlstype[k]))
                                {
                                    _xlsnames[i][j][k] = str;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }

        private static void LoadIniParm()
        {
            _xlsnames = new string[2][][];//左右、第几幅、指标表
            _sidef = new bool[2][];
            _sidetype = new string[2][];
            for (int i = 0; i < 2; i++)
            {
                _xlsnames[i] = new string[_sidenum][];
                _sidef[i] = new bool[_sidenum];
                _sidetype[i] = new string[_sidenum];
                for (int j = 0; j < _sidenum; ++j)
                {
                    _xlsnames[i][j] = new string[_xlstype.Length];
                }
            }
            _sidetype[0] = _sidetypeL;
            _sidetype[1] = _sidetypeR;

            IniFiles inisetting = new IniFiles(System.Windows.Forms.Application.StartupPath + @"\DocSetting.ini");
            _RoadPName = inisetting.ReadString("Road", "RoadPName", "123").Replace("\0", "");
            _RoadLName = inisetting.ReadString("Road", "RoadLName", "123").Replace("\0", "");
            _RoadType = inisetting.ReadString("Road", "RoadType", "123").Replace("\0", "");
            try
            {
                string path = inisetting.ReadString("Road", "ExcelPath", "123").Replace("\0", "");
                string name = path.Substring(path.LastIndexOf('\\'));
                _Smile = name.Substring(name.IndexOf("K"), 9);
                _Emile = name.Substring(name.LastIndexOf("~") + 1, 9);
            }
            catch
            { }
            for (int i = 0; i < _sidenum; ++i)
            {
                _sidef[0][i] = inisetting.ReadBool("Road", "RoadLine0" + i.ToString(), false);
                _sidef[1][i] = inisetting.ReadBool("Road", "RoadLine1" + i.ToString(), false);
            }

        }

        //生成word里边所需的所有表格到一个单独的excel表单中
        private static void CreateWordTable(MSExcel.Application excelApp, MSExcel.Workbook destbook)
        {
            MSExcel.Worksheet destsheet = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            MSExcel.Range destrange = null;
            int destcolsidx = 0;//目的表单要拷贝的起始列序号，从1开始
            int destrowsidx = 0;//目的表单要拷贝的起始行序号，从1开始
            int srcrownum = 0;//源表单要拷贝的区间行数，从1开始

            string[] destsheetname = { "PCI统计表", "IRI统计表", "Rut统计表", "MTD统计表", "PQI统计表" };
            int[] srcsheetidxval = { 2, 0, 5, 1, 4 };

            //等级评价统计表
            int[] srccolnumval = { 5, 8, 8, 7, 9 };
            int[] srcsortrow = { 3, 4, 4, 4, 3 };
            //int[] srccolval = { 13, 16, 16, 15, 17 };
            //int[] srccolval = { 14, 17, 17, 16, 18 };
            int[] srccolval = { 15, 18, 18, 17, 19 };
            int[] srccolnum = { 3, 4, 4, 4, 3 };
            int[] destcoloff = { 3, 6, 6, 5, 7 };
            object[,] ohead = new object[1, 3];
            string datastr;
            for (int sidx = 0; sidx < destsheetname.Length; ++sidx)
            {
                destrowsidx = 3;
                destcolsidx = 1;
                for (int i = 0; i < 2; ++i)
                {
                    for (int j = 0; j < _sidenum; ++j)
                    {
                        //if (!_sidef[i][j])
                        //{
                        //    break;
                        //}
                        if (_sidef[i][j])
                        {
                            srcbook = excelApp.Workbooks.Open(_xlsnames[i][j][srcsheetidxval[sidx]], Type.Missing,
                                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                            GlobalExcel.Reflection(srcsheet, srcsortrow[sidx], 1, srccolnumval[sidx], true);
                            GlobalExcel.Reflection(srcsheet, srcsortrow[sidx], 1, 2, false);
                            srcbook.Save();

                            destsheet = destbook.Sheets[destsheetname[sidx]] as MSExcel.Worksheet;
                            srcrownum = GlobalExcel.judegeusedrow(srcsheet, 1, srcsortrow[sidx]);

                            //桩号
                            srcrange = srcsheet.get_Range(String.Format("A2:B{0}", srcrownum));
                            destrange = destsheet.get_Range(String.Format("{0}2", GlobalExcel.GetCol((char)('A' + destcolsidx - 1))));
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy(destrange);

                            //内容
                            srcrange = srcsheet.get_Range(String.Format("D2:{0}{1}", GlobalExcel.GetCol((char)('A' + srccolnumval[sidx] - 1)), srcrownum));
                            destrange = destsheet.get_Range(String.Format("{0}2", GlobalExcel.GetCol((char)('A' + destcolsidx + 1))));
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy(destrange);

                            //总结语
                            srcrange = srcsheet.get_Range(String.Format("{0}2", GlobalExcel.GetCol((char)('A' + srccolval[sidx] - 7))));
                            destrange = destsheet.get_Range(String.Format("{0}1", GlobalExcel.GetCol((char)('A' + destcolsidx - 1))));
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy();
                            destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);

                            //路段评价统计表
                            if (sidx > 0)
                            {
                                destsheet = destbook.Sheets["路段" + destsheetname[sidx]] as MSExcel.Worksheet;

                                //桩号
                                ohead[0, 0] = string.Format("{0:K0+000}～{1:K0+000}",
                                    ((MSExcel.Range)srcsheet.Cells[srccolnum[sidx], 1]).Value,
                                    ((MSExcel.Range)srcsheet.Cells[srcrownum, 2]).Value);
                                ohead[0, 1] = _sidetype[i][j];
                                ohead[0, 2] = "百分率（%）";
                                destrange = destsheet.get_Range(String.Format("A{0}:C{0}", destrowsidx));
                                destrange.Value2 = ohead;

                                //内容
                                srcrange = srcsheet.get_Range(String.Format("{0}4:{1}4",
                                GlobalExcel.GetCol((char)('A' + srccolval[sidx])), GlobalExcel.GetCol((char)('A' + srccolval[sidx] + srccolnum[sidx]))));
                                destrange = destsheet.get_Range("D" + destrowsidx.ToString());
                                System.Windows.Forms.Clipboard.Clear();
                                srcrange.Copy();
                                destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
                            }

                            //路段合计表
                            destsheet = destbook.Sheets["路段合计"] as MSExcel.Worksheet;
                            ohead[0, 0] = string.Format("{0:K0+000}～{1:K0+000}",
                                ((MSExcel.Range)srcsheet.Cells[srccolnum[sidx], 1]).Value,
                                ((MSExcel.Range)srcsheet.Cells[srcrownum, 2]).Value);
                            ohead[0, 1] = _sidetype[i][j];
                            destrange = destsheet.get_Range(String.Format("A{0}:B{0}", destrowsidx));
                            destrange.Value2 = ohead;

                            datastr = String.Format("=AVERAGE({0}!{1}{2}:{1}{3})",
                                destsheetname[sidx], GlobalExcel.GetCol((char)('A' + destcolsidx + destcoloff[sidx] - 2)), srcsortrow[sidx], srcrownum);
                            destsheet.Cells[destrowsidx, sidx + 3] = datastr;

                            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            destcolsidx += srccolnumval[sidx] - 1;
                            ++destrowsidx;
                        }
                    }
                }
            }
            if (destsheet == null) return;
            object[,] odata = new object[1, 5];
            GlobalExcel.WriteExcel(destrowsidx, 1, 1, 2, "总计", destsheet, 0);
            for (int i = 0; i < 5; ++i)
            {
                odata[0, i] = String.Format("=AVERAGE({0}3:{0}{1})", GlobalExcel.GetCol((char)('A' + i + 2)), destrowsidx - 1);
            }
            destrange = destsheet.get_Range(String.Format("C{0}:G{0}", destrowsidx));
            destrange.Value2 = odata;

            //路段病害面积统计表
            destsheet = destbook.Sheets["路段病害面积统计表"] as MSExcel.Worksheet;
            object[,] osumsum = new object[20, 1];
            object[,] osumper = new object[20, 1];
            double[] disareasum = new double[21];
            bool IsFirstSheet = true;
            for (int i = 0; i < 2; ++i)
            {
                for (int j = 0; j < _sidenum; ++j)
                {
                    if (_sidef[i][j])
                    {

                        srcbook = excelApp.Workbooks.Open(_xlsnames[i][j][3], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        srcsheet = srcbook.Sheets.get_Item(1) as MSExcel.Worksheet;

                        srcrownum = GlobalExcel.judegeusedrow(srcsheet, 3, 3);
                        if (IsFirstSheet)
                        {
                            srcrange = srcsheet.get_Range(String.Format("A4:B{0}", srcrownum));
                            destrange = destsheet.get_Range("A3");
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy(destrange);
                            IsFirstSheet = false;
                        }
                        srcrownum = srcrownum - 3;
                        disareasum[0] += Convert.ToDouble(((MSExcel.Range)srcsheet.Cells[4, 6]).Value);
                        for (int di = 0; di < srcrownum; di++)
                        {
                            disareasum[di + 1] += Convert.ToDouble(((MSExcel.Range)srcsheet.Cells[di + 4, 3]).Value);
                        }
                        srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    }
                }
            }
            for (int di = 0; di < srcrownum; di++)
            {
                osumsum[di, 0] = disareasum[di + 1];
                osumper[di, 0] = String.Format("=C{0}/D3", di + 3);
            }
            destrange = destsheet.get_Range("C3:C" + (srcrownum + 1));
            destrange.Value2 = osumsum;
            GlobalExcel.WriteExcel(3, 4, srcrownum, 1, disareasum[0].ToString(), destsheet, 0);
            destrange = destsheet.get_Range("E3:E" + (srcrownum + 1));
            destrange.Value2 = osumper;
            GlobalExcel.WriteExcel(srcrownum + 2, 3, 1, 1, String.Format("=SUM(C3:C{0})", srcrownum + 1), destsheet, 0);
            GlobalExcel.WriteExcel(srcrownum + 2, 5, 1, 1, String.Format("=C{0}/D3", srcrownum + 2), destsheet, 0);
        }

        #region   模板一
        public static void OutputModel1Doc(MSWord.Application wordApp, MSExcel.Application excelApp, List<string> _ExcelPathList)
        {
            int generation;
            string destpath = _ExcelPathList[0].Substring(0, _ExcelPathList[0].LastIndexOf("\\"));
            destpath = destpath.Substring(0, destpath.LastIndexOf("\\")) + "\\";
            string[] srcdocname = { "附件二道路路面病害明细表（10米）", "附件三道路平整度检测数据表（10米）", "附件四道路车辙检测数据表（10米）", "附件五道路构造深度检测数据表（10米）", "附件六道路综合评价检测数据表（10米）" };
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板一\模板一道路使用现状评价分析.docx",
                System.Windows.Forms.Application.StartupPath);
            string destdoc = destpath + "道路使用现状评价分析.docx";

            LoadIniParm();
            GetXlsNames(_ExcelPathList);

            MSExcel.Workbook destbook = null;
            MSWord.Document wordDoc = null;

            string srcxls = string.Format(@"{0}\报告模板\城镇道路\模板一\模板一统计报表.xlsx", System.Windows.Forms.Application.StartupPath);
            destbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            destbook.SaveAs(destdoc.Replace("docx", "xlsx"), Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            CreateWordTableModel1(excelApp, destbook);

            wordDoc = wordApp.Documents.Open(srcdoc,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//Word文档变量
            wordDoc.SaveAs(destdoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ExportWordModel1(wordApp, excelApp, wordDoc, destbook);

            wordDoc.Save();
            wordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(wordApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            destbook.Save();
            destbook.Close(Type.Missing, Type.Missing, Type.Missing);
            generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        //生成word里边所需的所有表格到一个单独的excel表单中
        private static void CreateWordTableModel1(MSExcel.Application excelApp, MSExcel.Workbook destbook)
        {
            MSExcel.Worksheet destsheet = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Workbook ttsrcbook = null;
            MSExcel.Worksheet ttsrcsheet = null;
            MSExcel.Range srcrange = null;
            MSExcel.Range destrange = null;
            int destcolsidx = 0;//目的表单要拷贝的起始列序号，从1开始
            int destrowsidx = 0;//目的表单要拷贝的起始行序号，从1开始
            int srcrownum = 0;//源表单要拷贝的区间行数，从1开始

            string[] destsheetname = { "PCI统计表", "IRI统计表", "Rut统计表", "MTD统计表", "PQI统计表" };
            int[] srcsheetidxval = { 2, 0, 5, 1, 4 };

            //等级评价统计表
            int[] srccolnumval = { 7, 10, 9, 9, 15 };
            int[] srcsortrow = { 2, 3, 3, 3, 2 };
            //int[] srccolval = { 13, 16, 16, 15, 17 };
            //int[] srccolval = { 14, 17, 17, 16, 18 };
            int[] srccolval = { 15, 18, 18, 17, 19 };
            int[] srccolnum = { 3, 4, 4, 4, 3 }; //复制的起始行
            int[] destcoloff = { 3, 6, 6, 5, 10 };

            object[,] ohead = new object[1, 3];
            string datastr;
            for (int sidx = 0; sidx < destsheetname.Length; ++sidx)
            {
                destrowsidx = 3;
                destcolsidx = 1;
                for (int i = 0; i < 2; ++i)
                {
                    for (int j = 0; j < _sidenum; ++j)
                    {
                        if (_sidef[i][j])
                        {
                            srcbook = excelApp.Workbooks.Open(_xlsnames[i][j][srcsheetidxval[sidx]], Type.Missing,
                                false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                            srcbook.Save();

                            destsheet = destbook.Sheets[destsheetname[sidx]] as MSExcel.Worksheet;
                            srcrownum = GlobalExcel.judegeusedrow(srcsheet, 1, 1);

                            //桩号
                            srcrange = srcsheet.get_Range(String.Format("B{0}:C{1}", srccolnum[sidx], srcrownum));
                            destrange = destsheet.get_Range(String.Format("{0}1", GlobalExcel.GetCol((char)('A' + destcolsidx - 1))));
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy(destrange);

                            //内容
                            srcrange = srcsheet.get_Range(String.Format("F{0}:{1}{2}", srccolnum[sidx], GlobalExcel.GetCol((char)('A' + srccolnumval[sidx])), srcrownum));
                            destrange = destsheet.get_Range(String.Format("{0}1", GlobalExcel.GetCol((char)('A' + destcolsidx + 1))));
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            srcrange.Copy(destrange);

                            if (_xlsnames[i][j][sidx].Contains("IRI"))  // 获取路面面积  表不对要修改
                            {
                                ttsrcbook = excelApp.Workbooks.Open(_xlsnames[i][j][sidx], Type.Missing,
                               false, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                ttsrcsheet = ttsrcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                                _RoadToltalArea = ((MSExcel.Range)ttsrcsheet.Cells[3, 9]).Text.ToString();
                                ttsrcbook.Save();
                            }

                            //路段评价统计表
                            if (sidx > 0)
                            {
                                destsheet = destbook.Sheets["路段" + destsheetname[sidx]] as MSExcel.Worksheet;

                                //桩号
                                ohead[0, 0] = string.Format("{0:K0+000}～{1:K0+000}",
                                    ((MSExcel.Range)srcsheet.Cells[srccolnum[sidx] + 1, 2]).Value,
                                    ((MSExcel.Range)srcsheet.Cells[srcrownum, 3]).Value);
                                ohead[0, 1] = _sidetype[i][j];
                                ohead[0, 2] = "百分率（%）";
                                destrange = destsheet.get_Range(String.Format("A{0}:C{0}", destrowsidx));
                                destrange.Value2 = ohead;

                                //内容
                                srcrange = srcsheet.get_Range(String.Format("{0}4:{1}4",
                                GlobalExcel.GetCol((char)('A' + srccolval[sidx])), GlobalExcel.GetCol((char)('A' + srccolval[sidx] + srccolnum[sidx]))));
                                destrange = destsheet.get_Range("D" + destrowsidx.ToString());
                                System.Windows.Forms.Clipboard.Clear();
                                srcrange.Copy();
                                destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
                            }

                            //路段合计表
                            destsheet = destbook.Sheets["路段合计"] as MSExcel.Worksheet;
                            ohead[0, 0] = string.Format("{0:K0+000}～{1:K0+000}",
                                ((MSExcel.Range)srcsheet.Cells[srccolnum[sidx], 1]).Value,
                                ((MSExcel.Range)srcsheet.Cells[srcrownum, 2]).Value);
                            ohead[0, 1] = _sidetype[i][j];
                            destrange = destsheet.get_Range(String.Format("A{0}:B{0}", destrowsidx));
                            destrange.Value2 = ohead;

                            datastr = String.Format("=AVERAGE({0}!{1}{2}:{1}{3})",
                                destsheetname[sidx], GlobalExcel.GetCol((char)('A' + destcolsidx + destcoloff[sidx] - 2)), srcsortrow[sidx], srcrownum - srccolnum[sidx] + 1);
                            destsheet.Cells[destrowsidx, sidx + 3] = datastr;
                            destsheet.Cells[3, sidx + 8] = string.Format("=IF(C{0}>=85,\"A\",IF(C{0}>=70,\"B\",IF(C{0}>=60,\"C\",\"D\")))", 3);
                            destsheet.Cells[3, sidx + 9] = string.Format("=IF(D{0}>=3.6,\"A\",IF(D{0}>=3,\"B\",IF(D{0}>=2.4,\"C\",\"D\"))) ", 3);
                            destsheet.Cells[3, sidx + 10] = string.Format("=IF(E{0}>=90,\"优\",IF(E{0}>=80,\"良\",IF(E{0}>=70,\"中\",IF(E{0}>=60,\"次\",\"差\"))))", 3);
                            destsheet.Cells[3, sidx + 11] = string.Format("=IF(F{0}>=0.45,\"A\",IF(F{0}>=0.42,\"B\",IF(F{0}>=0.4,\"C\",\"D\")))", 3);
                            destsheet.Cells[3, sidx + 12] = string.Format("=IF(G{0}>=85,\"A\",IF(G{0}>=70,\"B\",IF(G{0}>=60,\"C\",\"D\")))", 3);

                            // PCI =IF(C3>=85,"A",IF(C3>=70,"B",IF(C3>=60,"C","D")))   RQI ==IF(F3>=3.6,"A",IF(F3>=3,"B",IF(F3>=2.4,"C","D")))  mtd==IF(E3>=0.45,"A",IF(E3>=0.42,"B",IF(E3>=0.4,"C","D")))
                            //PQI==IF(J2>=85,"A",IF(J2>=70,"B",IF(J2>=60,"C","D")))  RUT= =IF(G5>=90,"优",IF(G5>=80,"良",IF(G5>=70,"中",IF(G5>=60,"次","差"))))
                            srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            destcolsidx += srccolnumval[sidx] - 1;
                            ++destrowsidx;
                        }
                    }
                }
            }
            if (destsheet == null) return;
            object[,] odata = new object[1, 5];
            GlobalExcel.WriteExcel(destrowsidx, 1, 1, 2, "总计", destsheet, 0);
            for (int i = 0; i < 5; ++i)
            {
                odata[0, i] = String.Format("=AVERAGE({0}3:{0}{1})", GlobalExcel.GetCol((char)('A' + i + 2)), destrowsidx - 1);
            }
            destrange = destsheet.get_Range(String.Format("C{0}:G{0}", destrowsidx));
            destrange.Value2 = odata;

            //路段病害面积汇总表
            destsheet = destbook.Sheets["路段病害面积汇总表"] as MSExcel.Worksheet;
            object[,] osumsum = new object[20, 1];
            object[,] osumper = new object[20, 1];
            double[] disareasum = new double[21];
            bool IsFirstSheet = true;
            for (int i = 0; i < 2; ++i)
            {
                for (int j = 0; j < _sidenum; ++j)
                {
                    if (_sidef[i][j])
                    {
                        srcbook = excelApp.Workbooks.Open(_xlsnames[i][j][3], Type.Missing,
                            true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        srcsheet = srcbook.Sheets.get_Item(1) as MSExcel.Worksheet;

                        srcrownum = GlobalExcel.judegeusedrow(srcsheet, 3, 3);
                        if (IsFirstSheet)
                        {
                            srcrange = srcsheet.get_Range(String.Format("A1:O{0}", srcrownum));
                            destrange = destsheet.get_Range("A1");
                            GlobalExcel.SetBorderLine(srcrange, 63);
                            System.Windows.Forms.Clipboard.Clear();
                            srcrange.Copy(destrange);
                            IsFirstSheet = false;
                        }
                        srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                    }
                }
            }
        }

        public static void ExportWordModel1(MSWord.Application wordApp, MSExcel.Application excelApp, MSWord.Document wordDoc, MSExcel.Workbook excelXls)
        {
            ExportWordTableModel1(wordApp, excelApp, wordDoc, excelXls);
            ExportWordString(wordDoc);
            //更新目录
            int count = wordDoc.TablesOfContents.Count;
            for (int i = 0; i < count; i++)
            {
                wordDoc.TablesOfContents[i + 1].Update();
            }
        }

        public static void ExportWordTableModel1(MSWord.Application wordApp, MSExcel.Application excelApp,
            MSWord.Document wordDoc, MSExcel.Workbook excelbook)
        {
            MSExcel.Worksheet excelsheet = null;
            MSExcel.Range excelrange = null;
            MSWord.Range wordrange = null;

            #region 各指标评价等级表

            string[] sheetnames = { "路段病害面积汇总表", "IRI统计表", "PQI统计表", "路段合计" };
            string[] tableheadpre = { "表3.2-", "表3.1-", "表3.3-", "表3.4-" };
            string[] tableheadapp = { "路段病害面积汇总表", "路面平整度评价等级统计表", "路面综合评价指数（PQI）统计表", "路段评价汇总统计表" };
            int[] colnums = { 15, 7, 11, 7 };  //表跨越的列数
            int excelrow = 0;
            int colval = 0;
            String wordstr = null;
            int tableidx = 0;

            for (int sidx = 0; sidx < sheetnames.Length; ++sidx)
            {
                colval = 0;
                tableidx = 1;
                excelsheet = excelbook.Sheets[sheetnames[sidx]] as MSExcel.Worksheet;
                excelrow = GlobalExcel.judegeusedrow(excelsheet, 3, 3);

                wordrange = GlobalWord.GetMarkRange(wordDoc, sheetnames[sidx]);
                GlobalWord.wordAppGoTo(wordApp, wordrange);
                for (int i = 0; i < 2; ++i)
                {
                    for (int j = 0; j < _sidenum; ++j)
                    {
                        if (_sidef[i][j])
                        {
                            wordstr = String.Format("{0}1:{1}{2}",
                                GlobalExcel.GetCol((char)('A' + colval)), GlobalExcel.GetCol((char)('A' + colval + colnums[sidx] - 1)), excelrow);
                            excelrange = excelsheet.get_Range(String.Format("{0}1:{1}{2}",
                                GlobalExcel.GetCol((char)('A' + colval)), GlobalExcel.GetCol((char)('A' + colval + colnums[sidx] - 1)), excelrow));
                            System.Windows.Forms.Clipboard.Clear();
                            excelrange.Copy();

                            wordstr = string.Format("{0}{1:0} {2}({3}){4}", tableheadpre[sidx], tableidx++, _RoadLName, _sidetype[i][j], tableheadapp[sidx]);
                            GlobalWord.wordAppAlignment(wordApp, MSWord.WdParagraphAlignment.wdAlignParagraphCenter);
                            GlobalWord.wordAppTypeText(wordApp, wordstr);
                            GlobalWord.wordAppSelectionPaste(wordApp);

                            if (sidx == 3)
                            {
                                _RoadPCIdegree = ((MSExcel.Range)excelsheet.Cells[3, 8]).Text.ToString();
                                _RoadRQIdegree = ((MSExcel.Range)excelsheet.Cells[3, 9]).Text.ToString();
                                _RoadPQIdegree = ((MSExcel.Range)excelsheet.Cells[3, 12]).Text.ToString();
                            }
                        }
                    }
                }
            }
            #endregion

            #region 表格格式化
            foreach (MSWord.Table temptable in wordDoc.Tables)
            {
                temptable.Range.Font.Name = "宋体";
                temptable.Range.Font.Size = 12;
                temptable.Range.ParagraphFormat.LineSpacingRule = MSWord.WdLineSpacing.wdLineSpaceSingle;
                temptable.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                temptable.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                temptable.AutoFitBehavior(MSWord.WdAutoFitBehavior.wdAutoFitWindow);
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
        #endregion

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

        public static void OutputMode7DocHeader(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath,
            ProjectProjectClass tproject, ReportProjectClass treport)
        {
            // 报告的头部
            WriteHeader2Docx(wordApp, excelApp, srcpath, tproject, treport);
        }

        public static void OutputMode7DocMerge(MSWord.Application wordApp, string srcpath, List<RoadPartProjectClass> srclist, ReportProjectClass treport)
        {
            // 合并到一起
            WriteAll2Docx(wordApp, srcpath, srclist, treport);
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
        private static float[] width_wc = { 10.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f, 8.0f };
        private static float[] width_mt = { 8.0f, 10.0f, 10.0f, 10.0f, 10.0f, 15.0f, 10.0f };
        // private static float[] width_hz = { 12.0f, 7.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f };//各车道路面技术状况评价结果汇总表
        private static float[] width_hz = { 10.0f, 7.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 5.0f, 7.0f, 10.0f };//各车道路面技术状况评价结果汇总表

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
                    catch (Exception) { }
                }
                for (int i = 5; i < 11; ++i)
                {
                    try
                    {
                        temptable.Cell(1, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception) { }
                }
                for (int i = 5; i < 15; ++i)
                {
                    try
                    {
                        temptable.Cell(2, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception) { }
                }
                for (int i = 1; i < 15; ++i)
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
                for (int i = 1; i < 12; ++i)
                {
                    try
                    {
                        temptable.Cell(roadnum + 3, i).Range.set_Style(ref oStyleName);
                    }
                    catch (Exception) { }
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
                    }
                    break;
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
                    }
                    break;
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
                    }
                    break;
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
                    }
                    break;
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
                    }
                    break;
                case 6:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = 10;
                    }
                    break;
                case 7:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_hz[1];
                        for (int i = 4; i < 14; ++i)
                        {
                            try
                            {
                                temptable.Cell(2, i + 1).PreferredWidth = width_hz[i];
                            }
                            catch (Exception) { }
                        }
                        for (int i = 0; i < 4; ++i)
                        {
                            try
                            {
                                temptable.Cell(1, i + 1).PreferredWidth = width_hz[i];
                            }
                            catch (Exception) { }
                        }

                        for (int i = 0; i < roadnum; ++i)
                        {
                            for (int j = 0; j < 14; ++j)
                            {
                                try
                                {
                                    temptable.Cell(i + 2, j + 1).PreferredWidth = width_hz[j];
                                }
                                catch (Exception) { }
                            }
                        }
                        for (int i = 0; i < 12; ++i)
                        {
                            try
                            {
                                temptable.Cell(roadnum + 3, i + 2).PreferredWidth = width_hz[i + 4];
                            }
                            catch (Exception) { }
                        }
                    }
                    break;
                case 8:
                    {
                        temptable.Range.Cells.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.Range.Cells.PreferredWidth = width_wc[1];
                        for (int i = 0; i < width_wc.Length; ++i)
                        {
                            temptable.Columns[i + 1].PreferredWidth = width_wc[i];
                        }

                        temptable.PreferredWidthType = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPercent;
                        temptable.PreferredWidth = 105.0f;
                    }
                    break;
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
                catch (Exception)
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
                    catch (Exception)
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
                catch (Exception)
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

                    Thread.Sleep(500);

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
                    catch (Exception)
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
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\04路段附录模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);

            MSWord.Document wordDoc = null;
            MSExcel.Workbook[] srcbooks = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            MSWord.Table curtable = null;
            string[] typeheaderstrs = null;
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
            // wordApp.ScreenUpdating = false;

            int[] delcol = { 8, 7, 5 };
            int wordtablecnt = 0;

            foreach (RoadPartProjectClass srclist in srclists)
            {
                typeheaderstrs = new string[] {
                                            "路面破损检测与评定结果表",
                                          "路面平整度检测与评定结果表",
                                          "路面构造深度检测与评定结果表",
                                          "路面结构强度检测与评定结果表"};

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
                        catch (Exception)
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                        }
                    }
                }
                for (int i = 0; i < typeheaderstrs.Length; ++i)
                {
                    bool IsSheetTypeHeader = true;
                    for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                    {
                        bool hasWc = srclist.m_lanelist[j].m_wcDataClasses.Count > 0;
                        if (!hasWc && i == typeheaderstrs.Length - 1)
                        {
                            continue;
                        }

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
                            catch (Exception)
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
                            catch (Exception)
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
                                    catch (Exception)
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
                            else if (i == 2)
                            {
                                srcsheet = srcbooks[j].Sheets["TD"] as MSExcel.Worksheet;
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                srcrange = srcsheet.get_Range("A1:J" + userownum.ToString());
                            }
                            else
                            {
                                srcsheet = srcbooks[j].Sheets["弯沉"] as MSExcel.Worksheet;
                                userownum = GlobalExcel.judegeusedrow(srcsheet, 1, 2);
                                srcrange = srcsheet.get_Range("A1:J" + userownum.ToString());
                            }

                            tableheader = tabelheaderapp + typeheaderstrs[i];
                            // 报告表格内容（通用居中 小五）
                            curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告附表2");
                            //curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref wordtablecnt, tableheader, "报告表格内容（通用居中 小五）");
                            for (int k = 0; k < delcol.Length; ++k)
                            {
                                while (true)
                                {
                                    try
                                    {
                                        curtable.Columns[delcol[k]].Delete();
                                        break;
                                    }
                                    catch (Exception)
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

        // 创建评价表
        private static Dictionary<string, Dictionary<string, List<EvaluationRange>>> evaluationTable
              = new Dictionary<string, Dictionary<string, List<EvaluationRange>>>();

        /// <summary>
        /// 上海惠普获得评价等级
        /// </summary>
        /// <param name="jtldj">交通量等级</param>
        /// <param name="jclx">基层类型</param>
        /// <param name="cxValue">弯沉值</param>
        /// <returns></returns>
        private static string shhpGetJudgeStr(string jtldj, string jclx, double cxValue)
        {
            initHpJudegeDictionary();
            // 检查表中是否存在指定的基层类型和交通量等级
            if (evaluationTable.ContainsKey(jclx) && evaluationTable[jclx].ContainsKey(jtldj))
            {
                string evaluation = "";
                // 遍历对应的评价区间
                foreach (var range in evaluationTable[jclx][jtldj])
                {
                    // 检查值是否在当前区间内
                    if (cxValue >= range.MinValue && cxValue <= range.MaxValue)
                    {
                        return range.Evaluation; // 返回评价结果
                    }
                }
            }
            else
            {
                throw new Exception("请检查沉陷文件\n【基层类型(粒料及沥青稳定|半刚性)】【交通量等级(很轻|轻|中|重|特重)】\n是否填写正确!");
            }

            return "未找到";


        }
        // 定义评价区间类
        class EvaluationRange
        {
            public double MinValue { get; set; }
            public double MaxValue { get; set; }
            public string Evaluation { get; set; }
        }

        private static void initHpJudegeDictionary()
        {
            if (evaluationTable.Count == 0)
            {
                // 添加基层类型和交通量等级的组合及其对应的评价区间
                evaluationTable["粒料及沥青稳定"] = new Dictionary<string, List<EvaluationRange>>
                {
                     {
                     "很轻", new List<EvaluationRange>
                         {
                             new EvaluationRange { MinValue = -10000, MaxValue = 98, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 98, MaxValue = 126, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 126, MaxValue = 10000, Evaluation = "不足" },

                         }
                     },
                    {
                    "轻", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -1000, MaxValue = 77, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 77, MaxValue = 98, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 98, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "中", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 60, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 60, MaxValue = 81, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 81, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                      {
                    "重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 46, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 46, MaxValue = 67, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 67, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "特重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 35, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 35, MaxValue = 56, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 56, MaxValue = 10000, Evaluation = "不足" },
                         }
                    }

                 };

                // 添加基层类型和交通量等级的组合及其对应的评价区间
                evaluationTable["半刚性"] = new Dictionary<string, List<EvaluationRange>>
                {
                     {
                     "很轻", new List<EvaluationRange>
                         {
                             new EvaluationRange { MinValue = -10000, MaxValue = 77, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 77, MaxValue = 98, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 98, MaxValue = 10000, Evaluation = "不足" },

                         }
                     },
                    {
                    "轻", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -1000, MaxValue = 56, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 56, MaxValue = 77, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 77, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "中", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 42, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 42, MaxValue =59, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 59, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                      {
                    "重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 31, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 31, MaxValue = 46, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 46, MaxValue = 10000, Evaluation = "不足" },
                         }
                    },
                     {
                    "特重", new List<EvaluationRange>
                        {
                             new EvaluationRange { MinValue = -10000, MaxValue = 21, Evaluation = "足够" },
                             new EvaluationRange { MinValue = 21, MaxValue = 35, Evaluation = "临界" },
                             new EvaluationRange { MinValue = 35, MaxValue = 10000, Evaluation = "不足" },
                         }
                    }

                 };
            }
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
            MSExcel.Worksheet CollectPicSheet = CollectBook.Sheets["Sheet2"] as MSExcel.Worksheet;
            int CollectRow = 3;

            int Collect_yhstartrowidx = 2;
            int Collect_yhendrowidx = 2;
            int roadpartidx = 0;

            //判断当前导入的工程是否都没有弯沉数据
            bool allNoHaveWcData = true;
            //一个路段含有弯沉工程则该路段下所有上下行都会计入该 变量
            int hasWcDataProjectCount = 0;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                for (int D = 0; D < srclist.m_lanelist.Count; ++D)
                {
                    if (srclist.m_lanelist[D].m_wcDataClasses.Count > 0)
                    {
                        allNoHaveWcData = false;

                        break;
                    }
                }
                if (!allNoHaveWcData)
                {
                    break;
                }
            }

            if (!allNoHaveWcData)
            {
                foreach (RoadPartProjectClass srclist in srclists)
                {
                    for (int D = 0; D < srclist.m_lanelist.Count; ++D)
                    {
                        if (srclist.m_lanelist[D].m_wcDataClasses.Count > 0)
                        {
                            hasWcDataProjectCount += srclist.m_lanelist.Count;
                            break;
                        }
                    }
                }
            }

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
                object[,] wcobj = new object[srclist.m_lanelist.Count, 11];

                // excelApp.ScreenUpdating = false;
                int yhstartrowidx = 2;
                int yhendrowidx = 2;

                //判断该路段的各个车道是否都没有弯沉数据

                //当前路段拥有弯沉数据的工程个数
                int hasWcDataCurProjectCount = 0;

                List<DateTime> times = new List<DateTime>();
                for (int D = 0; D < srclist.m_lanelist.Count; ++D)
                {
                    if (srclist.m_lanelist[D].m_wcDataClasses.Count > 0)
                    {
                        hasWcDataCurProjectCount++;
                    }
                }
                if (hasWcDataCurProjectCount == 0)
                {
                    //删除单元格
                    destsheet = destbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                    var deleteRange = destsheet.get_Range("AX1:BA2");
                    deleteRange.Delete();
                    deleteRange = destsheet.get_Range(string.Format("AU1:AV{0}", srclist.m_lanelist.Count + 3));
                    deleteRange.Delete();
                    //MSExcel.Shape excelshape = destsheet.Shapes.Item(7) as Microsoft.Office.Interop.Excel.Shape;
                    //excelshape.Delete();
                    chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(7);
                    chartobj.Delete();
                }
                for (int j = 0; j < srclist.m_lanelist.Count; ++j)
                {
                    bool curProjecthasWc = srclist.m_lanelist != null && srclist.m_lanelist.Count > 0 && srclist.m_lanelist[j].m_wcDataClasses.Count > 0;

                    srcbook = excelApp.Workbooks.Open(srclist.m_lanelist[j].m_xlsxpath, Type.Missing,
                                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    destsheet = destbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                    srcsheet = srcbook.Sheets["统计"] as MSExcel.Worksheet;
                    //srcrange = srcsheet.get_Range("H23:P26");
                    srcrange = srcsheet.get_Range("H23:P34"); //添加弯沉
                    object[,] obj1 = (object[,])srcrange.Value2;

                    srcrange = srcsheet.get_Range("H14:U14");
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

                    if (curProjecthasWc)
                    {
                        destsheet.Cells[3 + j, 50] = obj2[1, 1];
                        destsheet.Cells[3 + j, 51] = obj1[10, 3];
                        destsheet.Cells[3 + j, 52] = obj1[11, 3];
                        destsheet.Cells[3 + j, 53] = obj1[12, 3];
                    }
                    else
                    {
                        if (hasWcDataCurProjectCount > 0)
                        {
                            destsheet.Cells[3 + j, 47] = string.Format("/");
                            destsheet.Cells[3 + j, 48] = string.Format("/");
                            //GlobalExcel.SetBorderLine(destsheet.Cells[3 + j, 47], 63);
                            //GlobalExcel.SetBorderLine(destsheet.Cells[3 + j, 48], 63);

                           //destsheet.Cells[3 + j, 50] = obj2[1, 1];
                           //destsheet.Cells[3 + j, 51] = 0;
                           //destsheet.Cells[3 + j, 52] = 0;
                           //destsheet.Cells[3 + j, 53] = 0;

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
                    string curTimeStr = obj2[8, 8].ToString();
                     

                    DateTime dateTime = DateTime.ParseExact(curTimeStr, "yyyyMMdd",null);
                    times.Add(dateTime);
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
                    if (curProjecthasWc)
                    {
                        wcobj[j, 0] = obj2[11, 2];

                    }
                    else
                    {
                        if (!allNoHaveWcData)
                        {
                            wcobj[j, 0] = obj2[11, 2];
                        }
                    }
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
                    if (!curProjecthasWc)
                    {
                        for (int i = 1; i <= yhendrowidx - yhstartrowidx + 1; i++)
                        {
                            destrange.Cells[i, 7] = "/";

                        }
                    }
                    destrange = CollectYHSheet.get_Range(string.Format("F{0}:R{1}", Collect_yhstartrowidx, Collect_yhendrowidx));
                    destrange.PasteSpecial(MSExcel.XlPasteType.xlPasteValues);
                    if (!curProjecthasWc)
                    {
                        for (int i = 1; i <= yhendrowidx-yhstartrowidx+1; i++)
                        {
                            destrange.Cells[i, 7] = "/";

                        }
                    }
                    yhstartrowidx = yhendrowidx + 1;
                    Collect_yhstartrowidx = Collect_yhendrowidx + 1;
                    
                    // 车道统计信息
                    srcsheet = srcbook.Sheets["统计"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("H2:U34");
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
                    if (curProjecthasWc)
                    {
                        wcobj[j, 1] = obj2[13, 1];
                        wcobj[j, 3] = obj2[31, 2];
                        wcobj[j, 4] = obj2[31, 3];
                        wcobj[j, 5] = obj2[32, 2];
                        wcobj[j, 6] = obj2[32, 3];
                        wcobj[j, 7] = obj2[33, 2];
                        wcobj[j, 8] = obj2[33, 3];
                        wcobj[j, 9] = obj2[13, 13];
                        wcobj[j, 10] = obj2[13, 14];
                        wcobj[j, 2] = Convert.ToDouble(wcobj[j, 3]) + Convert.ToDouble(wcobj[j, 5]) + Convert.ToDouble(wcobj[j, 7]);
                    }
                    else
                    {
                        if (!allNoHaveWcData)
                        {
                            wcobj[j, 1] = obj2[13, 1];
                            wcobj[j, 3] = 0;
                            wcobj[j, 4] = 0;
                            wcobj[j, 5] = 0;
                            wcobj[j, 6] = 0;
                            wcobj[j, 7] = 0;
                            wcobj[j, 8] = 0;
                            wcobj[j, 9] = 0;
                            wcobj[j, 10] = 0;
                            wcobj[j, 2] = 0;
                        }
                    }
                    pciobj[j, 7] = Convert.ToDouble(pciobj[j, 8]) + Convert.ToDouble(pciobj[j, 10]) + Convert.ToDouble(pciobj[j, 12]) + Convert.ToDouble(pciobj[j, 14]);
                    rqiobj[j, 2] = Convert.ToDouble(rqiobj[j, 3]) + Convert.ToDouble(rqiobj[j, 5]) + Convert.ToDouble(rqiobj[j, 7]) + Convert.ToDouble(rqiobj[j, 9]);
                    tdobj[j, 2] = Convert.ToDouble(tdobj[j, 3]) + Convert.ToDouble(tdobj[j, 5]) + Convert.ToDouble(tdobj[j, 7]) + Convert.ToDouble(tdobj[j, 9]);
                    pqiobj[j, 2] = Convert.ToDouble(pqiobj[j, 3]) + Convert.ToDouble(pqiobj[j, 5]) + Convert.ToDouble(pqiobj[j, 7]) + Convert.ToDouble(pqiobj[j, 9]);

                    srcbook.Close(Type.Missing, Type.Missing, Type.Missing);
                }


                Microsoft.Office.Interop.Excel.Worksheet destsheetTemp = destbook.Sheets["Sheet2"] as MSExcel.Worksheet;
                times = times.Distinct().ToList();
                times.Sort();
                string timeStr = "";

                foreach (var time in times)
                {
                    timeStr += time.ToString("yyyy年MM月dd日") + ",";
                }
                // 去除最后一个逗号
                if (timeStr.EndsWith(","))
                {
                    timeStr = timeStr.Substring(0, timeStr.Length - 1);
                }
                // 设置单元格格式为文本
                destsheetTemp.Cells[8, 8].NumberFormat = "@";
                destsheetTemp.Cells[8, 8] = timeStr;

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
                if (hasWcDataCurProjectCount > 0)
                {
                    obj = new object[1, 10];
                }
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
                if (hasWcDataCurProjectCount > 0)
                {
                    destrange = destsheet.get_Range(string.Format("AM{0}:AV{0}", srclist.m_lanelist.Count + 3));
                    obj[0, 8] = string.Format("=ROUND(AVERAGE(AU3:AU{0}),1)", srclist.m_lanelist.Count + 2);


                    var wcRange = destsheet.get_Range(string.Format("AU3:AU{0}", srclist.m_lanelist.Count + 2));
                    //object[,] wcValues = wcRange.Value2;

                    if (hasWcDataCurProjectCount == srclist.m_lanelist.Count)
                    {
                        double wcValue = 0;
                        List<double> tempWcDatas = new List<double>();
                        for (int ttt = 0; ttt < srclist.m_lanelist.Count; ttt++)
                        {
                            tempWcDatas.Add(Convert.ToDouble(wcobj[ttt, 9]));
                        }
                        wcValue = tempWcDatas.Average();

                        string ljlx = srclist.m_lanelist[0].m_wcDataClasses.First().WcLjlx;

                        for (int j = 1; j < srclist.m_lanelist.Count; ++j)
                        {
                            string nowljlx = srclist.m_lanelist[j].m_wcDataClasses.First().WcLjlx;
                            if (nowljlx == "/" || nowljlx != ljlx)
                            {
                                ljlx = "/";
                            }
                        }
                        if (ljlx == "/")
                        {
                            obj[0, 9] = string.Format("/");
                        }
                        else
                        {
                            //交通量等级  
                            string jtldj = srclist.m_lanelist.First().m_wcDataClasses.First().traffic;
                            obj[0, 9] = shhpGetJudgeStr(jtldj, ljlx, wcValue);
                        }
                    }
                    else
                    {
                        obj[0, 9] = string.Format("/");
                    }

                }
                else
                {
                    destrange = destsheet.get_Range(string.Format("AM{0}:AT{0}", srclist.m_lanelist.Count + 3));

                }
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


                //后添加弯沉图
                if (hasWcDataCurProjectCount > 0)
                {
                    destrange = destsheet.get_Range(string.Format("AX2:BA{0}", srclist.m_lanelist.Count + 2));
                    GlobalExcel.SetBorderLine(destrange, 63);
                    chartobj = (MSExcel.ChartObject)destsheet.ChartObjects(7);
                    chart = chartobj.Chart;
                    chart.ChartType = MSExcel.XlChartType.xlColumnClustered;
                    chart.SetSourceData(destrange, Missing.Value);
                }
                if (hasWcDataCurProjectCount > 0)
                {
                    destrange = destsheet.get_Range(string.Format("AI2:AV{0}", srclist.m_lanelist.Count + 2));
                }
                else
                {
                    destrange = destsheet.get_Range(string.Format("AI2:AT{0}", srclist.m_lanelist.Count + 2));
                }

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
                    catch (System.Exception)
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
                    catch (System.Exception)
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
                    catch (System.Exception)
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
                    catch (System.Exception)
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

                destrange = destsheet.get_Range(string.Format("BJ3:BT{0}", srclist.m_lanelist.Count + 2));
                if (allNoHaveWcData)
                {
                    destrange = destsheet.get_Range(string.Format("BJ1:BT{0}", srclist.m_lanelist.Count + 2));
                    destrange.Delete();
                }
                else
                {
                    destrange.Value2 = wcobj;
                    GlobalExcel.SetBorderLine(destrange, 63);
                }


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

                destrange = CollectDestSheet.get_Range(string.Format("BJ{0}:BT{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                if (!allNoHaveWcData)
                {
                    destrange.Value2 = wcobj;
                    GlobalExcel.SetBorderLine(destrange, 63);

                }
                else
                {
                    try
                    {
                        destrange = CollectDestSheet.get_Range(string.Format("BJ1:BT{1}", CollectRow, CollectRow + srclist.m_lanelist.Count - 1));
                        destrange.Delete();
                        destrange = CollectPicSheet.get_Range("A45:I51");
                        destrange.Delete();
                        destrange = CollectPicSheet.get_Range("F46:I50");
                        destrange.Delete();
                        chartobj = (MSExcel.ChartObject)CollectPicSheet.ChartObjects(5);
                        chartobj.Delete();
                    }
                    catch (Exception ex)
                    {

                        
                    }
                    
                }

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

        private static void PastExcelPic2Word(MSWord.Selection currentSelection, MSExcel.Worksheet srcsheet, int PicIdx, string PicName,
            bool IsTypeParagraph = true,float height = 0.0f,float width = 0.0f )
        {
            if (IsTypeParagraph)
            {
                currentSelection.TypeParagraph();
            }

            object oStyleName = "报告图与下段同页";
            currentSelection.set_Style(ref oStyleName);
            Thread.Sleep(GlobalWord.wd_sleep_us);

            currentSelection.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

            bool notfinished = true;
            do
            {
                try
                {
                    MSExcel.Shape excelshape = srcsheet.Shapes.Item(PicIdx) as Microsoft.Office.Interop.Excel.Shape;
                    System.Windows.Forms.Clipboard.Clear();
                    if (height!=0)
                    {
                        excelshape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        float temp =   (float)srcsheet.Application.CentimetersToPoints(height); 
                        excelshape.Width = (float)srcsheet.Application.CentimetersToPoints(width);
                        excelshape.Height = (float)srcsheet.Application.CentimetersToPoints(height);
                    }
                    excelshape.Copy();
                    Thread.Sleep(500);
                    currentSelection.PasteAndFormat(MSWord.WdRecoveryType.wdChartPicture);
                    Thread.Sleep(GlobalWord.wd_sleep_us);
                    notfinished = false;
                }
                catch (System.Exception ex)
                {

                    notfinished = true;
                }
            } while (notfinished);

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

        static DateTime ParseDate(string dateString)
        {
            // 假设日期格式为 "xxxx年xx月xx日"
            string[] parts = dateString.Split(new string[] { "年", "月", "日" }, StringSplitOptions.None);
            int year = int.Parse(parts[0]);
            int month = int.Parse(parts[1]);
            int day = int.Parse(parts[2]);

            return new DateTime(year, month, day);
        }

        private static List<string>[] handleIndexStr(List<string>[] IndexValStrs )
        {
            List<string>[] result = new List<string>[4];
            for (int i = 0; i < 4; ++i)
            {
                result[i] = new List<string>();
            }

            for (int i = 0; i < IndexValStrs.Length; i++)
            {
                for (int t = 0; t < IndexValStrs[i].Count; t++)
                {
                    string curStr = IndexValStrs[i][t];
                    if (curStr.Contains("上行"))
                    { 
                        string replaceStr = curStr.Replace("上行", "下行");

                        int findIndex = IndexValStrs[i].FindIndex(d => d.Equals(replaceStr));
                        if (findIndex == t+1)
                        {
                            string newStr = curStr.Replace("上行", "上丶下行").Replace("占比为","占比均为");
                            result[i].Add(newStr);
                            t++;
                        }
                        else
                        {
                            result[i].Add(curStr);
                        }
                    }
                    else
                    {
                        result[i].Add(curStr);
                    }  
                }


            }



            return    result;
        }

        // 生成路段的报告
        private static void WriteMainRoad2Docx(MSWord.Application wordApp, MSExcel.Application excelApp, string srcpath, List<RoadPartProjectClass> srclists)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\02路段主体模板.docx",
                System.Windows.Forms.Application.StartupPath);
            FileInfo srcinfo = new FileInfo(srcpath);
            //  wordApp.ScreenUpdating = false;
            MSWord.Document wordDoc = null;
            MSExcel.Workbook srcbook = null;
            MSExcel.Worksheet srcsheet = null;
            MSExcel.Range srcrange = null;
            MSExcel.Workbook srcbook2 = null;
            MSExcel.Worksheet srcsheet2 = null;
            MSWord.Selection currentSelection = null;
            //弯沉统计表插入位置
            MSWord.Selection currentWcSelection = null;
            int tablecount = 1;
            object oStyleName;
            MSWord.Table curtable = null;

            bool IsHasTD = false;

            string[] IndexJudgeStrs = { "路面综合评价", "路面破损评价", "路面行驶质量评价", "路面抗滑能力评价" };
            int wordcnt = 0;
            //只要任一工程有弯沉数据就需要出具相应数据
            bool hasWcData = false;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                for (int i = 0; i < srclist.m_lanelist.Count; i++)
                {
                    if (srclist.m_lanelist[i].m_wcDataClasses.Count > 0)
                    {
                        hasWcData = true;
                        break;
                    }
                    if (hasWcData)
                    {
                        break;
                    }
                }
            }
            MSWord.Bookmark myBookmark = null;
            foreach (RoadPartProjectClass srclist in srclists)
            {
                bool curProjectHasWc = false;
                int hasWcProjectIndex = 0; 
                for (int i = 0; i < srclist.m_lanelist.Count; i++)
                {
                    if (srclist.m_lanelist[i].m_wcDataClasses.Count > 0)
                    {
                        curProjectHasWc = true;
                        hasWcProjectIndex = i;
                        break;
                    }
                    if (curProjectHasWc)
                    {
                        break;
                    }
                }

                string roadpartfname = "路段" + srclist.m_roadpart.m_id + "#"
                    + srclist.m_roadpart.m_roadinfo.m_code + "_"
                    + srclist.m_roadpart.m_roadinfo.m_name + "（"
                    + srclist.m_roadpart.m_startlocation + "-"
                    + srclist.m_roadpart.m_endlocation + "）";

                ++wordcnt;
                string destdoc = srcinfo.DirectoryName + "\\路段统计\\" + roadpartfname + ".docx";
                if (curProjectHasWc)
                {
                    srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\02路段主体模板_弯沉.docx",
                System.Windows.Forms.Application.StartupPath);
                }
                else
                {
                    srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\02路段主体模板.docx",
                System.Windows.Forms.Application.StartupPath);
                }
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

                if (srclist.m_MapImg != null)
                {
                    foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                    {
                        if (book.Name == "地理位置示意图")
                        {
                            book.Select();
                            currentSelection = wordApp.Selection;
                            MSWord.InlineShape tmppic = currentSelection.InlineShapes.AddPicture(srclist.m_MapImg);
                            tmppic.Width = 283.5f;
                            tmppic.Height = 212.625f;
                            break;
                        }
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
                wordDoc.Tables[1].Columns[wordDoc.Tables[1].Columns.Count].Delete();
                srcobj = (object[,])srcrange.Value2;

                // 技术状况评定等级分布的文字描述
                // PQI、PCI、RQI、TD
                List<string>[] IndexValStrs = new List<string>[4];
                //路面弯沉评价文字
                string wcJudgeStr = "路面弯沉。";
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                MSWord.Selection tempSelection = null;
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
                        catch (System.Exception) { }
                        try
                        {
                            bval = Convert.ToDouble(objsrc[4, 3 + k * 2]);
                            if (bval < 100.0)
                            {
                                allB = false;
                            }
                        }
                        catch (System.Exception) { }

                        try
                        {
                            cval = Convert.ToDouble(objsrc[5, 3 + k * 2]);
                        }
                        catch (System.Exception)
                        {
                            continue;
                        }
                        try
                        {
                            dval = Convert.ToDouble(objsrc[6, 3 + k * 2]);
                        }
                        catch (System.Exception)
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
             
                // 在当前位置插入书签
                
                //所有弯沉都达标
                bool allWcZg = true;
                if (curProjectHasWc)
                {
                    string bzWcStr = "路面弯沉代表值。";
                    List<string> allRoadMsgs = new List<string>();
                    //全部都是不足
                    List<double> allBzScore = new List<double>();
                    for (int i = 0; i < srclist.m_lanelist.Count; ++i)
                    {
                        bool curHasWc = srclist.m_lanelist[i].m_wcDataClasses.Count > 0;
                        if (curHasWc)
                        {

                            srcbook2 = excelApp.Workbooks.Open(srclist.m_lanelist[i].m_xlsxpath, Type.Missing,
                                                                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            srcsheet2 = srcbook2.Sheets["统计"] as MSExcel.Worksheet;
                            string roadnumstr = ((MSExcel.Range)srcsheet2.Cells[14, 8]).Value.ToString();
                            //string tableheader = roadnumstr + "路面强度等级分布统计汇总表";
                            srcrange = srcsheet2.get_Range("H30:J34");
                            // curtable = PastExcelTable2Word(wordDoc, srcrange, currentSelection, ref tablecount, tableheader, "报告表标题3", true);

                            // 技术状况评定等级分布的文字描述
                            // PQI、PCI、RQI、TD
                            object[,] objsrc = (object[,])srcrange.Value2;
                            string bzScore = Convert.ToDouble(objsrc[5, 3]).ToString("f2");
                            double dbzScore = double.Parse(bzScore);
                            if (dbzScore > 0)
                            {
                                wcJudgeStr += roadnumstr + $"弯沉评价等级为\"不足\"的长度占比为{bzScore}%。"; 
                            }
                            bzWcStr += roadnumstr + "丶";
                            allRoadMsgs.Add(roadnumstr);
                            allBzScore.Add(dbzScore);
                            if ( dbzScore!= 0)
                            {
                                allWcZg = false;
                            }
                            srcbook2.Close(Type.Missing, Type.Missing, Type.Missing);
                        }
                    }


                    if (allWcZg)
                    {
                        wcJudgeStr = bzWcStr.Substring(0, bzWcStr.Length - 1) + $"弯沉评价等级为\"不足\"的长度占比均为0.00%。";
                    }
                    allBzScore = allBzScore.Distinct().ToList();
                    if (allBzScore.Count == 0 )
                    {

                    }
                    if (allBzScore .Count == 1)
                    {
                        allRoadMsgs =  allRoadMsgs.Distinct().ToList();
                       string msg =  ProcessLanes(allRoadMsgs);
                       wcJudgeStr = $"路面弯沉。{msg}弯沉评价等级为\"不足\"的长度占比均为{allBzScore.First().ToString("f2")}%。";
                    }
                    tempSelection = currentSelection;
                    currentWcSelection = tempSelection;
                    myBookmark = wordDoc.Bookmarks.Add("WcMarkName", currentSelection.Range);
                    oStyleName = "报告表下空行";
                    SetStyle(currentSelection, oStyleName, true);
                }
                float heigth = 0f;
                float width = 0f;
                if (IsHasTD && curProjectHasWc)
                {
                    heigth = 5.2f;
                    width = 10.38f; 

                }
                else
                {
                    heigth = 5.52f;
                    width = 11.02f;
                }

                    PastExcelPic2Word(currentSelection, srcsheet, 1, "各车道单元PCI等级占比分布图", height: heigth, width: width);

                PastExcelPic2Word(currentSelection, srcsheet, 2, "各车道单元RQI等级占比分布图", height: heigth, width: width);
                PastExcelPic2Word(currentSelection, srcsheet, 3, "各车道单元PQI等级占比分布图", height: heigth, width: width);
                if (IsHasTD)
                {
                    PastExcelPic2Word(currentSelection, srcsheet, 4, "各车道单元TD等级占比分布图",height: heigth, width: width);

                }
                if (curProjectHasWc)
                    PastExcelPic2Word(currentSelection, srcsheet, 7, "各车道单元弯沉等级占比分布图", height: heigth, width: width);

                srcsheet = srcbook.Sheets["Sheet3"] as MSExcel.Worksheet;
                srcrange = srcsheet.get_Range(string.Format("G3:G{0}", srclist.m_lanelist.Count + 3));
                srcobj = (object[,])srcrange.Value2;
                
                string roadDescribe = "";
                string wcRoadDescribe = "";
                if (srclist.m_lanelist.Count>1)
                {
                    
                   
                        List<string> laneMsg = new List<string>();
                        for (int i = 0; i < srclist.m_lanelist.Count; i++)
                        {

                            laneMsg.Add(srcobj[i + 1, 1].ToString()); 
                        }

                        roadDescribe = ProcessLanes(laneMsg);
                     
                }
                if (curProjectHasWc)
                {
                    srcrange = srcsheet.get_Range(string.Format("BJ3:BS{0}", srclist.m_lanelist.Count + 3));
                    srcobj = (object[,])srcrange.Value2;

                    List<string> laneMsg = new List<string>();
                    for (int i = 0; i < srclist.m_lanelist.Count; i++)
                    {
                        string lanMsg = srcobj[i + 1, 2].ToString();
                        double wcValue = double.Parse( srcobj[i + 1, 3].ToString());

                        if (wcValue !=0 )
                        {
                            laneMsg.Add(lanMsg);

                        }
                    } 
                        wcRoadDescribe = ProcessLanes(laneMsg);
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

                    if (curProjectHasWc)
                    {
                        if (allWcZg)
                        {
                            currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段弯沉指标评价等级均为“足够”或“临界”；其余各指标评价等级为“C”或“D”的占比统计结果如下："); 
                        }
                        else
                        {
                            currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，各指标评价等级为“不足”、“C”或“D”的占比统计结果如下：");

                        }

                    }
                    else
                    {
                        currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，各指标评价等级为“C”或“D”的占比统计结果如下：");

                    }

                    currentSelection.TypeParagraph();
                    oStyleName = "标题 6";
                    currentSelection.set_Style(ref oStyleName);
                    IndexValStrs =  handleIndexStr(IndexValStrs);

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
                        if (curProjectHasWc )
                        {
                            if (allWcZg)
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“A”；本路段弯沉指标评价等级无“不足”。");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“A”；本路段弯沉指标评价等级无“不足”。");

                                }

                            }
                            else
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“A”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“A”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }

                            }
                        }
                        else
                        {
                            currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“A”。"); 
                        }
                    }
                    else if (allB)
                    {
                        if (curProjectHasWc)
                        {
                            if (allWcZg)
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“B”；本路段弯沉指标评价等级无“不足”。");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“B”；本路段弯沉指标评价等级无“不足”。");

                                }

                            }
                            else
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“B”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“B”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }

                            }
                        }
                        else
                        {
                            currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“B”。");

                        }
                    }
                    else if (!allA && !allB)
                    {
                        if (curProjectHasWc)
                        {
                            if (allWcZg)
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“A”或“B”；本路段弯沉指标评价等级无“不足”。");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“A”或“B”；本路段弯沉指标评价等级无“不足”。");

                                }

                            }
                            else
                            {
                                if (IsHasTD)
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI、TD指标评价等级均为“A”或“B”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }
                                else
                                {
                                    currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段PCI、RQI、PQI指标评价等级均为“A”或“B”；弯沉指标评价等级为“不足”的占比统计结果如下：");

                                }

                            }
                        }
                        else
                        {
                                 currentSelection.TypeText("由路面技术状况评定等级分布情况分析可知，本路段各指标评价等级均为“A”或“B”。");
                        }
                      
                    }
                }

                if (curProjectHasWc&&!allWcZg)
                {
                    currentSelection.TypeParagraph();
                    oStyleName = "标题 6";
                    currentSelection.set_Style(ref oStyleName);
                    currentSelection.TypeText(wcJudgeStr);
                }

                int tablenum = wordDoc.Tables.Count;
                for (int i = tablenum1 + 1; i <= tablenum; ++i)
                {
                    FromatTable(wordApp, wordDoc.Tables[i], 0.6f, 6, false, srclist.m_lanelist.Count);
                }
                bool needDeleteTdCol = false;
                srcsheet = srcbook.Sheets["Sheet1"] as MSExcel.Worksheet;
                if (curProjectHasWc)
                {

                    if (IsHasTD)
                    {
                        srcrange = srcsheet.get_Range("AI1:AV" + (srclist.m_lanelist.Count + 3).ToString());
                    }
                    else
                    {
                        srcrange = srcsheet.get_Range("AI1:AV" + (srclist.m_lanelist.Count + 3).ToString());
                        needDeleteTdCol = true;
                    }
                }
                else
                {
                    if (IsHasTD)
                    {
                        srcrange = srcsheet.get_Range("AI1:AT" + (srclist.m_lanelist.Count + 3).ToString());

                    }
                    else
                    {
                        srcrange = srcsheet.get_Range("AI1:AR" + (srclist.m_lanelist.Count + 3).ToString());

                    }

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

                if (needDeleteTdCol)
                {

                    try
                    {
                        {
                            // 删除Word表格中的最后一列
                            for (int row = 1; row <= curtable.Rows.Count; row++)
                            {
                                curtable.Cell(row, curtable.Columns.Count - 2).Delete();
                                curtable.Cell(row, curtable.Columns.Count - 3).Delete();
                            }
                        }

                    }
                    catch (Exception)
                    {

                    }
                }
                

                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                FromatTable(wordApp, curtable, 0.6f, 7, true, srclist.m_lanelist.Count);

                tablenum1 = wordDoc.Tables.Count;
                for (int i = 0; i < srclist.m_lanelist.Count; ++i)
                {
                    bool sigleHasWc = srclist.m_lanelist[i].m_wcDataClasses.Count > 0 ? true:false;
                    srcbook2 = excelApp.Workbooks.Open(srclist.m_lanelist[i].m_xlsxpath, Type.Missing,
                                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    // 粘贴起止点景观照片
                    if (i == 0)
                    {
                        srcsheet2 = srcbook2.Sheets["景观图像"] as MSExcel.Worksheet;
                        #region 粘贴可能出问题
                        foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
                        {
                            if (book.Name == "起止点照")
                            {
                                book.Select();
                                currentSelection = wordApp.Selection;
                                int temp = 10;
                                MSExcel.Shape excelshape = null;
                                try
                                {
                                    while (temp > 0)
                                    {
                                        excelshape = srcsheet2.Shapes.Item(1) as Microsoft.Office.Interop.Excel.Shape;
                                        System.Windows.Forms.Clipboard.Clear();
                                        excelshape.Copy();
                                        currentSelection.Paste();
                                        break;
                                    }

                                }
                                catch (Exception ex)
                                {
                                    temp--;
                                    if (temp == 0)
                                    {
                                        if (MessageBox.Show($"由于本机剪切板故障，本次{destdoc}文件\n照片粘贴失败 是否继续导出剩余内容？", "提示窗口", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            throw new Exception($"{destdoc}由于设备问题导出失败");
                                        }
                                    }
                                }
                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                currentSelection.TypeText("  ");
                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                excelshape = srcsheet2.Shapes.Item(2) as Microsoft.Office.Interop.Excel.Shape;
                                System.Windows.Forms.Clipboard.Clear();
                                Thread.Sleep(GlobalWord.wd_sleep_us);
                                excelshape.Copy();
                                Thread.Sleep(GlobalWord.wd_sleep_us);
                                currentSelection.Paste();

                                Thread.Sleep(GlobalWord.wd_sleep_us);

                                break;
                            }
                        }
                        #endregion
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
                    if (!sigleHasWc)
                    {

                        for (int rowIdx = 2; rowIdx <= trownum; rowIdx++)
                        {
                            curtable.Cell(rowIdx,8).Range.Text = "/";

                        }
                    }
                    srcbook2.Close(false, Type.Missing, Type.Missing);
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
                    if (!curProjectHasWc)
                    {
                        temptable.Columns[8].Delete();
                        temptable.Cell(1, 1).Range.Text = "序号";
                        temptable.Cell(1, 11).Range.Text = "养护对策";
                        FromatTable(wordApp, temptable, 0.6f);
                    }
                    else
                    {
                        temptable.Cell(1, 1).Range.Text = "序号";
                        temptable.Cell(1, 12).Range.Text = "养护对策";
                        FromatTable(wordApp, temptable, 0.6f);
                    }


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

                if (curProjectHasWc && tjobj[srclist.m_lanelist.Count + 3, 13] != null && tjobj[srclist.m_lanelist.Count + 3, 14] != null)
                {
                    string wcStr = tjobj[srclist.m_lanelist.Count + 3, 14].ToString();
                    if (wcStr == "/")
                    {
                        wcStr = "。";
                    }
                    else
                    {
                        wcStr = $"，等级为{wcStr}。";
                    }
                        currentSelection.TypeParagraph();
                    currentSelection.TypeText("{道路名称}路面弯沉代表值为" + Convert.ToDouble(tjobj[srclist.m_lanelist.Count + 3, 13]).ToString("0.0")
                        + wcStr);
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




                List<DateTime> allTimes = new List<DateTime>();
                string tempStr = srcobj[8, 8] == null ? "" : srcobj[8, 8].ToString();
                List<string> projectTimeStrs = srcobj[8, 8].ToString().Split(',').ToList();
                //添加工程采集日期
                foreach (var projectTime in projectTimeStrs)
                {
                    string format = "yyyy年MM月dd日";
                    DateTime dt =  DateTime.ParseExact(projectTime, format, null);
                    allTimes.Add(dt);
                }

                if (curProjectHasWc)
                {
                    //添加所有的弯沉日期 
                    foreach (var lane in srclist.m_lanelist)
                    {
                        bool tempHasWcData = lane.m_wcDataClasses.Count > 0;
                        if (tempHasWcData)
                        {
                            string wcTime = lane.m_wcDataClasses.First().time;
                            string[] timeSplit = wcTime.Split('/');
                            if (timeSplit.Length > 0)
                            {
                                string timeS1 = timeSplit[0] + "年" + timeSplit[1] + "月" + timeSplit[2] + "日";
                                DateTime date1 = ParseDate(timeS1);
                                allTimes.Add(date1);
                            }
                        } 
                    } 
                    
                } 
                // 使用 Distinct 方法去除重复的日期
                List<DateTime> distinctTimes = allTimes.Distinct().ToList();
                distinctTimes.Sort();
                if (distinctTimes.Count == 1)
                {
                    tempStr = distinctTimes.First().ToString("yyyy年MM月dd日");
                }
               else if (distinctTimes.Count == 2)
                {
                    tempStr = distinctTimes.First().ToString("yyyy年MM月dd日") + "、" +
                                distinctTimes[1].ToString("yyyy年MM月dd日");
                }
                else
                {
                    tempStr = distinctTimes.First().ToString("yyyy年MM月dd日") + "～" +
                                distinctTimes.Last().ToString("yyyy年MM月dd日");
                }


                datas.Add("{检测日期}", tempStr); 
                datas.Add("{车道布置}", srclist.m_LaneLayout);



                if (string.IsNullOrEmpty(roadDescribe))
                { 
                    roadDescribe = srcobj[4, 8] == null ? "" : srcobj[4, 8].ToString() + srcobj[5, 8] == null ? "" : srcobj[5, 8].ToString() + "车道";
                }
                
                datas.Add("{车道描述}", roadDescribe);
                string cdmsStr = "";
                if (IsHasTD)
                {
                    cdmsStr = "路面状况指数PCI、路面行驶质量指数RQI、路面综合评价指数PQI、构造深度TD";

                    if (curProjectHasWc)
                    {
                        if (roadDescribe.Equals(wcRoadDescribe))
                        {
                            cdmsStr += "、路面弯沉";
                        }
                        else
                        {
                            cdmsStr += $"；并对该路段{wcRoadDescribe}路面弯沉";
                        }
                    }
                    
                    
                }
                else
                {
                    cdmsStr = "路面状况指数PCI、路面行驶质量指数RQI、路面综合评价指数PQI";
                    if (curProjectHasWc)
                    {
                        if (roadDescribe.Equals(wcRoadDescribe))
                        {
                            cdmsStr += "、路面弯沉";
                        }
                        else
                        {
                            cdmsStr += $"；并对该路段{wcRoadDescribe}路面弯沉";
                        }
                    }
                    
                }
                datas.Add("{车道指标}", cdmsStr);
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
                        catch (Exception)
                        {
                            Thread.Sleep(GlobalWord.wd_sleep_us);
                        }
                    }
                }

                if (curProjectHasWc)
                {
                    myBookmark.Select();

                    srcsheet = srcbook.Sheets["Sheet3"] as MSExcel.Worksheet;
                    srcrange = srcsheet.get_Range("BK1:BR" + (srclist.m_lanelist.Count + 2).ToString());

                    //新建一个新的报表 把srcrange写入  然后获得这个区域 在接下来删除第二列

                    // 创建一个新的工作表
                    MSExcel.Worksheet newSheet = srcbook.Sheets.Add(After: srcbook.Sheets[srcbook.Sheets.Count]);

                    // 将原始范围复制到新的工作表
                    srcrange.Copy(Destination: newSheet.Range["A1"]);

                    // 获取新的范围
                    MSExcel.Range newRange = newSheet.Range["A1:S" + (srclist.m_lanelist.Count + 2).ToString()];
                    // 删除第二列
                    Microsoft.Office.Interop.Excel.Range secondColumn = newRange.Columns[2];
                    secondColumn.Delete(MSExcel.XlDeleteShiftDirection.xlShiftToLeft);

                    newRange = newSheet.Range["A1:G" + (srclist.m_lanelist.Count + 2).ToString()];

                    string tableheader = "各车道路面结构强度等级分布统计汇总表";
                    // currentSelection.Select(); 
                    curtable = PastExcelTable2Word(wordDoc, newRange, currentSelection, ref tablecount, tableheader, "报告表标题3", true);
                    FromatTable(wordApp, wordDoc.Tables[5 + srclist.m_lanelist.Count - 1], 0.6f, 6, false, srclist.m_lanelist.Count);
                    newSheet.Delete();

                }
                //插入交叉引用
                WriteCrossReference(wordApp, wordDoc, curProjectHasWc);

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



       private static string ProcessLanes(List<string> inputStrings)
        {
            // 使用字典来存储车道信息
            Dictionary<string, HashSet<string>> laneInfo = new Dictionary<string, HashSet<string>>();

            foreach (var input in inputStrings)
            {
                // 提取车道编号和方向
                int laneIndex = input.IndexOf("车道");
                string direction = input.Substring(0, 2);
                string laneNumber = input.Substring(2,input.Length-4);

                // 将信息存储到字典中
                if (!laneInfo.ContainsKey(laneNumber))
                {
                    laneInfo[laneNumber] = new HashSet<string>();
                }
                laneInfo[laneNumber].Add(direction);
            }

            // 构建结果字符串
            StringBuilder resultBuilder = new StringBuilder();
            List<string> summarizedLanes = new List<string>();
            List<string> detailedLanes = new List<string>();

            foreach (var lane in laneInfo)
            {
                string laneNumber = lane.Key;
                var directions = lane.Value;

                // 格式化输出
                if (directions.Contains("上行") && directions.Contains("下行"))
                {
                    summarizedLanes.Add($"{laneNumber}");
                }
                else if (directions.Contains("上行"))
                {
                    detailedLanes.Add($"上行{laneNumber}车道");
                }
                else if (directions.Contains("下行"))
                {
                    detailedLanes.Add($"下行{laneNumber}车道");
                }
            }

            // 将总结的车道信息添加到结果字符串
            if (summarizedLanes.Count == 0)
            {

            }
            else if (summarizedLanes.Count==1)
            {
                resultBuilder.Append(  summarizedLanes.First());
                resultBuilder.Append("车道");
            }
            else if (summarizedLanes.Count>1)
            {
                resultBuilder.Append(string.Join("、", summarizedLanes));
                resultBuilder.Append("车道");
            }
          
            // 如果有详细的车道信息，添加分隔符和详细信息
            if (detailedLanes.Count > 0)
            {
                if (resultBuilder.Length > 0)
                {
                    resultBuilder.Append("与");
                }
                if (detailedLanes.Count==1)
                {
                    resultBuilder.Append(detailedLanes.First());
                }
                else
                {
                    resultBuilder.Append(string.Join("与", detailedLanes)); 
                }
            }

            return resultBuilder.ToString();
        }

        private static void WriteCrossReference(MSWord.Application wordApp, MSWord.Document wordDoc, bool hasWc)
        {
            //病害面积表
            MSWord.Selection currentSelection = null;

            object crossReferenceItems = wordDoc.GetCrossReferenceItems(MSWord.WdReferenceType.wdRefTypeNumberedItem);//MSWord.WdReferenceType.wdRefTypeNumberedItem,
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

            int offval = -3;

            for (int i = 1; i <= crossitemnum; ++i)
            {
                crossitemarrstrs[i] = (string)(crossitemarr.GetValue(i));

                if (crossitemarrstrs[i].Contains("道路概况表"))
                {
                    OverViewTablbeStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("地理位置示意图"))
                {
                    MapPicStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("起点照"))
                {
                    StartPicStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("终点照"))
                {
                    EndPicStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("路面病害类型面积统计汇总表（沥青）"))
                {
                    LQDisTableStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("路面病害类型面积统计汇总表（水泥）"))
                {
                    SNDisTableStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("沥青路面病害类型面积占比图"))
                {
                    LQDisPicStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("水泥路面病害类型面积占比图"))
                {
                    SNDisPicStrIdx = i - offval;
                }

                else if (crossitemarrstrs[i].Contains("等级占比分布图"))
                {
                    if (IndexPicStr1Idx == 0)
                    {
                        IndexPicStr1Idx = i - offval;
                    }
                    IndexPicStr2Idx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("路面技术状况评价结果汇总表"))
                {
                    HZTableStrIdx = i - offval;
                }
                else if (crossitemarrstrs[i].Contains("养护对策一览表"))
                {
                    if (YHTableStr1Idx == 0)
                    {
                        YHTableStr1Idx = i - offval;
                    }
                    YHTableStr2Idx = i - offval;
                }
                if (!hasWc)
                {
                    //表1.1-5	各车道路面结构强度等级分布统计汇总表 
                    if (crossitemarrstrs[i].Contains("技术状况评定等级分布统计汇总表"))
                    {
                        if (RoadHZTableStr1Idx == 0)
                        {
                            RoadHZTableStr1Idx = i - offval;
                        }
                        RoadHZTableStr2Idx = i - offval;
                    }

                }
                else
                {
                    if (crossitemarrstrs[i].Contains("技术状况评定等级分布统计汇总表"))
                    {

                        if (RoadHZTableStr1Idx == 0)
                        {
                            RoadHZTableStr1Idx = i - offval;
                        }
                        RoadHZTableStr2Idx = i - offval+1;
                    }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                        catch (System.Exception) { }
                    }
                    else if (LQDisTableStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisTableStrIdx, true);
                        }
                        catch (System.Exception) { }
                    }
                    else if (SNDisTableStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisTableStrIdx, true);
                        }
                        catch (System.Exception) { }
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
                        catch (System.Exception) { }
                    }
                    else if (LQDisPicStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, LQDisPicStrIdx, true);
                        }
                        catch (System.Exception) { }
                    }
                    else if (SNDisPicStrIdx != 0)
                    {
                        try
                        {
                            currentSelection.InsertCrossReference(MSWord.WdReferenceType.wdRefTypeNumberedItem,
                                MSWord.WdReferenceKind.wdNumberRelativeContext, SNDisPicStrIdx, true);
                        }
                        catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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
                    catch (System.Exception) { }
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

        private static void WriteAll2Docx(MSWord.Application wordApp, string srcpath, List<RoadPartProjectClass> srclists, ReportProjectClass treport)
        {
            int generation;
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\00总报告模板.docx",
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

            //wordDoc.PageSetup.TopMargin = currentSelection.PageSetup.TopMargin = (float)(1.5 * 0.3937008 * 72); // 单位从厘米转换为磅
            //                                                                                                    //  currentSelection.PageSetup.
            //wordDoc.PageSetup.BottomMargin = currentSelection.PageSetup.BottomMargin = (float)(1.5 * 0.3937008 * 72); // 单位从厘米转换为磅

            // 设置第一节（报告头）边距
            MSWord.Section firstSection = wordDoc.Sections[1];
            firstSection.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f);
            firstSection.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.5f);

            currentSelection.PageSetup.HeaderDistance = (float)(1.5 * 0.3937008 * 72); // 单位从厘米转换为磅
            currentSelection.PageSetup.FooterDistance = (float)(1.3 * 0.3937008 * 72); // 单位从厘米转换为磅 

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
                int times = 1;
                string firstStr = "";
                SuggestionStr = "本项目部建议进行中修及以上的路段为";

                for (int i = 2; i <= ttable.Rows.Count; ++i)
                {
                    curstr = ttable.Cell(i, 2).Range.Text.Replace("\r\a", "（")
                        + ttable.Cell(i, 3).Range.Text.Replace("\r\a", "），");
                    if (i == 2)
                    {
                        firstStr = curstr;
                    }
                    if (oldstr != curstr)
                    {
                        SuggestionStr += curstr;
                       
                    }
                    times++;
                    oldstr = curstr;
                }
                    SuggestionStr = "本项目部建议进行中修及以上的路段为" + firstStr.Substring(0,firstStr.Length-1) + "等" + (times-1).ToString() + "个，具体单元详见正文第7节。";
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

            // 设置后续所有节边距（从第2节开始）
            for (int i = 2; i <= wordDoc.Sections.Count; i++)
            {
                MSWord.PageSetup sectionSetup = wordDoc.Sections[i].PageSetup;
                sectionSetup.TopMargin = wordApp.CentimetersToPoints(2.5f);
                sectionSetup.BottomMargin = wordApp.CentimetersToPoints(2.5f);
            }

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

            //设置页眉-第2节
            object What = MSWord.WdGoToItem.wdGoToPage;
            object Which = MSWord.WdGoToDirection.wdGoToNext;
            object Name = "3";
            wordDoc.ActiveWindow.Selection.GoTo(ref What, ref Which, ref Name);
            wordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
            MSWord.Tables headertabels = wordApp.Selection.HeaderFooter.Range.Tables;
            if (headertabels.Count > 0)
            {
                headertabels[1].Cell(1, 1).Range.Text = treport.m_report.m_report_name;
                headertabels[1].Cell(1, 2).Range.Text = treport.m_report.m_report_num;
            }
            //设置页眉-第3节
            object What2 = MSWord.WdGoToItem.wdGoToSection;
            object Which2 = MSWord.WdGoToDirection.wdGoToNext;
            object Name2 = "2";
            wordDoc.ActiveWindow.Selection.GoTo(ref What2, ref Which2, ref Name2);
            wordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
            MSWord.Tables headertabels2 = wordApp.Selection.HeaderFooter.Range.Tables;
            if (headertabels.Count > 0)
            {
                headertabels2[1].Cell(1, 1).Range.Text = treport.m_report.m_report_name;
                headertabels2[1].Cell(1, 2).Range.Text = treport.m_report.m_report_num;
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
            string srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\03路段结论模板.docx",
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
            // currentSelection.TypeParagraph();
            oStyleName = "报告表标题1（隐藏）";
            currentSelection.set_Style(oStyleName);


            currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();

            oStyleName = "报告图标题1（隐藏）";

            currentSelection.set_Style(oStyleName);

            currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();

            oStyleName = "报告照片标题1（隐藏）";
            currentSelection.set_Style(oStyleName);

            currentSelection.Font.Hidden = -1;
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
                {
                    oStyleName = "报告表标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告图标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告照片标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                }


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

                {
                    oStyleName = "报告表标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告图标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告照片标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                }
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
                {
                    oStyleName = "报告表标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告图标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告照片标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                }

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
                {
                    oStyleName = "报告表标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告图标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    oStyleName = "报告照片标题2（隐藏）";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                }
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
                // wordDoc.Fields.Update();
            }
            bool hasWcData = false;
            //添加弯沉 
            tmpstr = "由弯沉的统计图表可以看出，";
            srcxlsrange = CollectDestSheet.get_Range("A45:I51");
            srcobj = (object[,])srcxlsrange.Value2;

            if (srcobj[7, 2] == null)
            {
                hasWcData = false;
            }
            else
            {
                if (Convert.ToDouble(srcobj[7, 2]) == 0)
                {
                    hasWcData = false;
                }
                else
                {
                    hasWcData = true;
                }
            }

            if (hasWcData)
            {
                isTypeText = GetWcSummaryStr(srcobj, "弯沉代表值", ref tmpstr);
                if (isTypeText)
                {
                    currentSelection.TypeText("沥青路面结构强度（弯沉）");
                    oStyleName = "标题 2";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.TypeParagraph();
                    {
                        oStyleName = "报告表标题2（隐藏）";
                        currentSelection.set_Style(oStyleName);
                        currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                        oStyleName = "报告图标题2（隐藏）";
                        currentSelection.set_Style(oStyleName);
                        currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                        oStyleName = "报告照片标题2（隐藏）";
                        currentSelection.set_Style(oStyleName);
                        currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();
                    }
                    tableheader = "路面结构强度及评价等级分类统计表";
                    PastExcelTable2Word(wordDoc, srcxlsrange, currentSelection, ref tablecount, tableheader, "报告表标题3", false);
                    PastExcelPic2Word(currentSelection, CollectDestSheet, 5, "路面结构强度评价等级分布示意图", false);

                    currentSelection.TypeParagraph();
                    oStyleName = "报告正文";
                    currentSelection.set_Style(oStyleName);
                    currentSelection.TypeText(tmpstr);
                    File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr);
                    currentSelection.TypeParagraph();

                    tmpstr = GetSummaryWcPercentStr(srcobj, "路面结构强度");
                    currentSelection.TypeText(tmpstr);
                    File.AppendAllText(srcinfo.DirectoryName + "\\结论.txt", tmpstr + "\n");
                    currentSelection.TypeParagraph();
                    // wordDoc.Fields.Update();
                }

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
            string[] strs = null;
            if (hasWcData)
            {
                strs = new string[] {
                    "序号", "路名", "路段", "行车\n方向", "车道\n编号", "路面\n类型", "桩号（起）", "桩号（止）",
                                "单元\n长度\n（m）","弯沉等级", "PCI\n等级", "RQI\n等级", "TD\n等级", "养护对策"};
            }
            else
            {
                strs = new string[] {
                    "序号", "路名", "路段", "行车\n方向", "车道\n编号", "路面\n类型", "桩号（起）", "桩号（止）",
                                "单元\n长度\n（m）", "PCI\n等级", "RQI\n等级", "TD\n等级", "养护对策"};
            }

            listdest.Add(strs);

            int tmpval = 0;
            for (int i = 2; i <= userownum; ++i)
            {
                string tstr = srcobj[i, YHCol].ToString();
                if (tstr.Contains("中修") || tstr.Contains("大修"))
                {
                    string[] tstrs = new string[strs.Length];
                    tstrs[0] = listdest.Count.ToString();   //序号
                    tstrs[1] = srcobj[i, 3].ToString();     //路名
                    if (srcobj[i, 4] != null && srcobj[i, 5] != null)
                    {
                        tstrs[2] = srcobj[i, 4].ToString() + "～" + srcobj[i, 5].ToString();     //路段
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
                    if (hasWcData)
                    {
                        tstrs[9] = srcobj[i, 12].ToString();     //弯沉
                        tstrs[10] = srcobj[i, 13].ToString();     //PCI等级
                        tstrs[11] = srcobj[i, 14].ToString();    //RQI等级
                        tstrs[12] = srcobj[i, 15].ToString();    //TD等级
                        tstrs[13] = tstr;    //养护对策
                    }
                    else
                    {
                        tstrs[9] = srcobj[i, 13].ToString();     //PCI等级
                        tstrs[10] = srcobj[i, 14].ToString();    //RQI等级
                        tstrs[11] = srcobj[i, 15].ToString();    //TD等级
                        tstrs[12] = tstr;    //养护对策
                    }

                    listdest.Add(tstrs);
                }
            }

            currentSelection.TypeText("大中修路段建议");
            oStyleName = "标题 1";
            currentSelection.set_Style(oStyleName);
            currentSelection.TypeParagraph();

            oStyleName = "报告表标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.Font.Hidden = -1;
            currentSelection.TypeParagraph();

            oStyleName = "报告图标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.Font.Hidden = -1;
            currentSelection.TypeParagraph();

            oStyleName = "报告照片标题1（隐藏）";
            currentSelection.set_Style(oStyleName);
            currentSelection.Font.Hidden = -1; currentSelection.TypeParagraph();

            // 存在大中修
            if (listdest.Count > 1)
            {
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText("根据本次道路检测的基础数据，结合道路现场，综合评定以下单元为中修及以上路段单元。");
                currentSelection.TypeParagraph();

                oStyleName = "报告表标题2"; //修改
                currentSelection.set_Style(oStyleName);
                currentSelection.TypeText("大中修路段单元清单");
                currentSelection.TypeParagraph();
                oStyleName = "报告正文";
                currentSelection.set_Style(oStyleName);
                MSWord.Table table = null;
                if (hasWcData)
                {
                    table = currentSelection.Tables.Add(currentSelection.Range, listdest.Count, 14, Type.Missing, Type.Missing);

                }
                else
                {
                    table = currentSelection.Tables.Add(currentSelection.Range, listdest.Count, 13, Type.Missing, Type.Missing);

                }
                for (int i = 0; i < listdest.Count; ++i)
                {
                    if (hasWcData)
                    {
                        for (int j = 0; j < 14; ++j)
                        {
                            table.Cell(i + 1, j + 1).Range.Text = listdest[i][j];
                        }
                    }
                    else
                    {
                        for (int j = 0; j < 13; ++j)
                        {
                            table.Cell(i + 1, j + 1).Range.Text = listdest[i][j];
                        }
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

                if (hasWcData)
                {
                    //table.Columns[13].PreferredWidth = 4.8f;
                    table.Columns[14].PreferredWidth = 15.0f;
                }
                else
                {

                    table.Columns[13].PreferredWidth = 15.0f;
                }
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
        private static void WriteHeader2Docx(MSWord.Application wordApp, MSExcel.Application excelApp,
            string srcpath, ProjectProjectClass tproject, ReportProjectClass treport)
        {
            int generation;
            string srcdoc = "";

            //只要任一工程有弯沉数据就需要出具相应数据
            bool hasWcData = false;
            foreach (RoadPartProjectClass srclist in treport.m_roadpartlist)
            {
                for (int i = 0; i < srclist.m_lanelist.Count; i++)
                {
                    if (srclist.m_lanelist[i].m_wcDataClasses.Count > 0)
                    {
                        hasWcData = true;
                        break;
                    }
                    if (hasWcData)
                    {
                        break;
                    }
                }
            }


            if (hasWcData)
            {
                srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\01报告头模板_弯沉.docx",
                System.Windows.Forms.Application.StartupPath);
            }
            else
            {
                srcdoc = string.Format(@"{0}\报告模板\城镇道路\模板5\01报告头模板.docx",
                System.Windows.Forms.Application.StartupPath);
            }

            FileInfo srcinfo = new FileInfo(srcpath);

            MSExcel.Workbook workBook = null;
            MSExcel.Worksheet workSheet = null;
            MSWord.Table curtable = null;
            MSWord.Range wordrange = null;
            MSWord.Document wordDoc = null;

            // 读取所有的车道报表里面的实际检测长度数据
            for (int i1 = 0; i1 < treport.m_roadpartlist.Count; ++i1)
            {
                for (int i2 = 0; i2 < treport.m_roadpartlist[i1].m_lanelist.Count; ++i2)
                {
                    if (treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath != null)
                    {
                        workBook = excelApp.Workbooks.Open(treport.m_roadpartlist[i1].m_lanelist[i2].m_xlsxpath, Type.Missing,
                                    true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        workSheet = workBook.Sheets["控制信息"] as MSExcel.Worksheet;
                        int smile = Convert.ToInt32(((MSExcel.Range)workSheet.Cells[9, 8]).Value);
                        int emile = Convert.ToInt32(((MSExcel.Range)workSheet.Cells[10, 8]).Value);
                        treport.m_roadpartlist[i1].m_lanelist[i2].m_laneRealLength = Math.Abs(smile - emile);
                        workBook.Close();
                    }
                    else
                    {
                        treport.m_roadpartlist[i1].m_lanelist[i2].m_laneRealLength = 0;
                    }
                }
            }

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
            datas.Add("{检测起始日期}", tproject.m_project.m_testing_start_date.Replace("/","").Insert(6, "月").Insert(4, "年") + "日");
            datas.Add("{检测终止日期}", tproject.m_project.m_testing_end_date.Replace("/", "").Insert(6, "月").Insert(4, "年") + "日");

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
                    catch (Exception)
                    {
                        Thread.Sleep(GlobalWord.wd_sleep_us);
                    }
                }
            }

            foreach (MSWord.Bookmark book in wordDoc.Bookmarks)
            {
                if (book.Name == "检测项目")
                {
                    book.Select();
                    MSWord.Selection currentSelection = wordApp.Selection;
                    List<string> indexstrlist = new List<string>();
                    foreach (IndexInfoClass tindex in tproject.m_project.m_indexlist)
                    {
                        if (tindex.m_tesing == "是"
                            && (tindex.m_index == "RQI"
                            || tindex.m_index == "PCI"
                            || tindex.m_index == "TD"
                            || tindex.m_index == "弯沉"))
                        {
                            string tstr = "路面" + tindex.m_name;
                            if (!indexstrlist.Contains(tstr))
                            {
                                indexstrlist.Add(tstr);
                            }
                        }
                    }
                    if (indexstrlist.Count > 0)
                    {
                        if (hasWcData)
                        {
                            if (!indexstrlist.Contains("路面强度"))
                            {
                                indexstrlist.Add("路面结构强度");
                            }
                            else
                            {
                                indexstrlist.Remove("路面强度");
                                indexstrlist.Add("路面结构强度");

                            }
                        }

                        for (int i = 0; i < indexstrlist.Count; ++i)
                        {
                            currentSelection.TypeText(indexstrlist[i]);
                            if (i < indexstrlist.Count - 1)
                            {
                                currentSelection.TypeParagraph();
                            }
                        }
                    }
                    break;
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
                    if (hasWcData)
                    {
                        curtable = wordrange.Tables.Add(wordrange, treport.m_roadpartlist.Count+1, 8);

                    }
                    else
                        curtable = wordrange.Tables.Add(wordrange, treport.m_roadpartlist.Count+1 , 7);
                    break;
                }
            }
            curtable.Range.set_Style(ref oStyleName);

            trowcnt = 1;
            curtable.Cell(trowcnt, 1).Range.Text = "序号";
            curtable.Cell(trowcnt, 2).Range.Text = "道路名称";
            curtable.Cell(trowcnt, 3).Range.Text = "路段";
            curtable.Cell(trowcnt, 4).Range.Text = "道路等级";
            curtable.Cell(trowcnt, 5).Range.Text = "路面类型";
            curtable.Cell(trowcnt, 6).Range.Text = "道路里程\n（m）";

            if (hasWcData)
            {
                curtable.Cell(trowcnt, 7).Range.Text = "检测长度（PCI丶RQI丶TD）\n（m）";
                curtable.Cell(trowcnt, 8).Range.Text = "检测长度（弯沉）\n（m）";
            }
            else
            {
                curtable.Cell(trowcnt, 7).Range.Text = "检测长度\n（m）";

            }
            if (hasWcData)
            {
                for (int j = 1; j <= 8; ++j)
                {
                    curtable.Cell(trowcnt, j).Range.Font.Bold = 1;
                }
            }
            else
            {
                for (int j = 1; j <= 7; ++j)
                {
                    curtable.Cell(trowcnt, j).Range.Font.Bold = 1;
                }
            }
            foreach (RoadPartProjectClass troadpart in treport.m_roadpartlist)
            {
                ++trowcnt;
                curtable.Cell(trowcnt, 1).Range.Text = (trowcnt - 1).ToString();
                curtable.Cell(trowcnt, 2).Range.Text = troadpart.m_roadpart.m_roadinfo.m_name;
                curtable.Cell(trowcnt, 3).Range.Text = troadpart.m_roadpart.m_startlocation + "～" + troadpart.m_roadpart.m_endlocation;
                curtable.Cell(trowcnt, 4).Range.Text = troadpart.m_roadpart.m_part_grade;
                curtable.Cell(trowcnt, 5).Range.Text = troadpart.m_roadpart.m_type;
                curtable.Cell(trowcnt, 6).Range.Text = troadpart.m_roadpart.m_length;
                int sumlen = 0;
                double sumWcLen = 0;
                foreach (LaneProjectClass tlane in troadpart.m_lanelist)
                {
                    sumlen += tlane.m_laneRealLength;
                    if (tlane.m_wcDataClasses.Count > 0)
                    {
                        sumWcLen += tlane.m_wcDataClasses.First().wcLength;
                    }

                }
                if (hasWcData)
                {
                    curtable.Cell(trowcnt, 7).Range.Text = sumlen.ToString();
                    string wcleng = sumWcLen == 0 ? "/" : sumWcLen.ToString();
                    curtable.Cell(trowcnt, 8).Range.Text = wcleng;
                }
                else
                {
                    curtable.Cell(trowcnt, 7).Range.Text = sumlen.ToString();
                }
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
            if (hasWcData)
            {
                curtable.Columns[1].PreferredWidth = 6.7f;
                curtable.Columns[2].PreferredWidth = 14.4f;
                curtable.Columns[3].PreferredWidth = 22.4f;
                curtable.Columns[4].PreferredWidth = 11.2f;
                curtable.Columns[5].PreferredWidth = 11.2f;
                curtable.Columns[6].PreferredWidth = 11.2f;
                curtable.Columns[7].PreferredWidth = 11.2f;
                curtable.Columns[8].PreferredWidth = 11.3f;

            }
            else
            {
                curtable.Columns[1].PreferredWidth = 6.8f;
                curtable.Columns[2].PreferredWidth = 14.4f;
                curtable.Columns[3].PreferredWidth = 27.5f;
                curtable.Columns[4].PreferredWidth = 13.7f;
                curtable.Columns[5].PreferredWidth = 14.4f;
                curtable.Columns[6].PreferredWidth = 11.4f;
                curtable.Columns[7].PreferredWidth = 11.4f;

            }

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
            curtable.Cell(trowcnt, 3).Range.Text = "道路里程（m）";
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
                if (tperson.m_duty.Length > 0)
                {
                    ++trowcnt;
                    curtable.Cell(trowcnt, 1).Range.Text = (trowcnt - 1).ToString();
                    curtable.Cell(trowcnt, 2).Range.Text = tperson.m_name;
                    curtable.Cell(trowcnt, 3).Range.Text = tperson.m_title;
                    curtable.Cell(trowcnt, 4).Range.Text = tperson.m_CertificateNo;
                    curtable.Cell(trowcnt, 5).Range.Text = tperson.m_duty;
                }
                else
                {
                    ++trowcnt;
                    curtable.Rows[trowcnt].Delete();
                    --trowcnt;
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
        private static string GetSummaryWcPercentStr(object[,] srcobj, string indexValStr)
        {
            string tmpstr = "总体而言，";

            double perA = Convert.ToDouble(srcobj[7, 4]);
            double perB = Convert.ToDouble(srcobj[7, 6]);
            double perC = Convert.ToDouble(srcobj[7, 8]);
            double perAB = perA + perB;

            tmpstr = tmpstr + perAB.ToString("0.00") + "%的检测道路的" + indexValStr + "处于临界及以上";
            if (perC > 0)
            {
                tmpstr = tmpstr + "，" + perC.ToString("0.00") + "%处于不足";
            }
            tmpstr = tmpstr + "。";

            return tmpstr;
        }
        private static string GetSummaryPercentStr(object[,] srcobj, string indexValStr)
        {
            string tmpstr = "总体而言，";

            double perA = Convert.ToDouble(srcobj[7, 4]);
            double perB = Convert.ToDouble(srcobj[7, 6]);
            double perC = Convert.ToDouble(srcobj[7, 8]);
            double perD = Convert.ToDouble(srcobj[7, 10]);
            double perABC = perA + perB + perC;

            tmpstr = tmpstr + perABC.ToString("0.00") + "%的检测道路的" + indexValStr + "处于C级及以上";
            if (perD > 0)
            {
                tmpstr = tmpstr + "，" + perD.ToString("0.00") + "%处于D级";
            }
            tmpstr = tmpstr + "。";

            return tmpstr;
        }
        private static bool GetWcSummaryStr(object[,] srcobj, string indexValStr, ref string tmpstr)
        {
            bool res = true;
            List<string> tmpstrlist = new List<string>();
            List<string> tmpstrlist1 = new List<string>();

            tmpstr = tmpstr + "检测道路总体" + indexValStr + "均值为" + Convert.ToDouble(srcobj[7, 9]).ToString("0.0") + "。其中，";

            if (Convert.ToDouble(srcobj[3, 2]) > 0)
            {
                tmpstrlist.Add("快速路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[3, 9]).ToString("0.0"));
                }
                catch (System.Exception) { }

            }

            if (Convert.ToDouble(srcobj[4, 2]) > 0)
            {
                tmpstrlist.Add("主干路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[4, 9]).ToString("0.0"));
                }
                catch (System.Exception) { }

            }

            if (Convert.ToDouble(srcobj[5, 2]) > 0)
            {
                tmpstrlist.Add("次干路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[5, 9]).ToString("0.0"));
                }
                catch (System.Exception) { }

            }

            if (Convert.ToDouble(srcobj[6, 2]) > 0)
            {
                tmpstrlist.Add("支路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[6, 9]).ToString("0.0"));
                }
                catch (System.Exception) { }

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
                    tmpstr = tmpstr + tmpstrlist1[0] + "。";
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
                    tmpstr = tmpstr + tmpstrlist1[tmpstrlist1.Count - 1] + "。";
                }
            }

            return res;
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
                catch (System.Exception) { }
                tmpstrlist2.Add(srcobj[4, 12].ToString() + "级");
            }

            if (Convert.ToDouble(srcobj[5, 2]) > 0)
            {
                tmpstrlist.Add("次干路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[5, 11]).ToString("0.00"));
                }
                catch (System.Exception) { }
                tmpstrlist2.Add(srcobj[5, 12].ToString() + "级");
            }

            if (Convert.ToDouble(srcobj[6, 2]) > 0)
            {
                tmpstrlist.Add("支路");
                try
                {
                    tmpstrlist1.Add(Convert.ToDouble(srcobj[6, 11]).ToString("0.00"));
                }
                catch (System.Exception) { }
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
