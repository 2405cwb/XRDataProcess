using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace XRDataProcess
{
    /// <summary>
    /// 生成多车道合并报表
    /// </summary>
    class MyExcelMerge
    {
        /// <summary>
        /// 多车道合并报表模板1，四川振兴
        /// </summary>
        /// <param name="_MergeType">合并模式，0-全幅，1-上行右幅，2-下行左幅</param>
        /// <param name="_ExcelType">合并报表模板类型，0-四川振兴</param>
        /// <param name="_IsMergeIdx">指标是否需要合并：PCI\RQI\RDI\PBI\PWI\SMTD\PQI</param>
        /// <param name="_UpXlsFiles">上行报表路径，车道、指标</param>
        /// <param name="_DownXlsFiles">下行报表路径，车道、指标</param>
        /// <param name="_UpXlsFiles">上行公里报表路径，车道、指标</param>
        /// <param name="_DownXlsFiles">下行公里报表路径，车道、指标</param>
        /// <param name="_OutputPath">合并报表放置路径</param>
        public static void OutputMerge(int _MergeType, int _ExcelType, MergeIndexInfo[] MergeInfo,
            string[][] _UpXlsFiles, string[][] _DownXlsFiles,
            string[][] _UpXlsFilesKM, string[][] _DownXlsFilesKM,
            string _OutputPath)
        {
            MSExcel.Application excelApp = new MSExcel.Application()
            {
                Visible = true,
                DisplayAlerts = false /*保存Excel的时候，不弹出是否保存的窗口直接进行保存*/,
                AlertBeforeOverwriting = false
            };

            if (_ExcelType == 0)
            {
                MergeExcel1(excelApp, MergeInfo, _UpXlsFiles, _DownXlsFiles, _UpXlsFilesKM, _DownXlsFilesKM, _OutputPath, _MergeType);
            }
            
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
            excelApp.Quit();
        }

        /// <summary>
        /// 四川振兴，全幅
        /// </summary>
        /// <param name="_IsMergeIdx">指标是否需要合并：PCI\RQI\RDI\PBI\PWI\SMTD\PQI</param>
        /// <param name="_UpXlsFiles">上行报表路径，车道、指标</param>
        /// <param name="_DownXlsFiles">下行报表路径，车道、指标</param>
        /// <param name="_UpXlsFiles">上行公里报表路径，车道、指标</param>
        /// <param name="_DownXlsFiles">下行公里报表路径，车道、指标</param>
        /// <param name="_OutputPath">合并报表放置路径</param>
        /// <param name="_MergeType">合并模式，0-全幅，1-上行右幅，2-下行左幅</param>
        private static void MergeExcel1(MSExcel.Application excelApp, MergeIndexInfo[] MergeInfo,
            string[][] _UpXlsFiles, string[][] _DownXlsFiles,
            string[][] _UpXlsFilesKM, string[][] _DownXlsFilesKM,
            string _OutputPath, int MergeType)
        {
            for (int j = 0; j < MergeInfo.Length; ++j)
            {
                if (!MergeInfo[j]._IsMergeIdx)
                    continue;

                MyExcelData[] exceldatas_up = null;
                MyExcelData[] exceldatas_down = null;
                MyExcelData[] exceldatas_upKM = null;
                MyExcelData[] exceldatas_downKM = null;

                if (_UpXlsFiles != null)
                {
                    //读取所有车道的原始报表数据
                    exceldatas_up = new MyExcelData[_UpXlsFiles.Length];
                    for (int k = 0; k < _UpXlsFiles.Length; ++k)
                    {
                        exceldatas_up[k] = new MyExcelData();
                        ReadExcelData(excelApp, _UpXlsFiles[k][j], MergeInfo[j]._ExcelStartRow, ref exceldatas_up[k]);
                    }
                    //读取所有车道的原始公里报表数据
                    exceldatas_upKM = new MyExcelData[_UpXlsFilesKM.Length];
                    for (int k = 0; k < _UpXlsFilesKM.Length; ++k)
                    {
                        exceldatas_upKM[k] = new MyExcelData();
                        ReadExcelData(excelApp, _UpXlsFilesKM[k][j], MergeInfo[j]._ExcelStartRow, ref exceldatas_upKM[k]);
                    }
                }

                if (_DownXlsFiles != null)
                {
                    exceldatas_down = new MyExcelData[_DownXlsFiles.Length];
                    for (int k = 0; k < _DownXlsFiles.Length; ++k)
                    {
                        exceldatas_down[k] = new MyExcelData();
                        ReadExcelData(excelApp, _DownXlsFiles[k][j], MergeInfo[j]._ExcelStartRow, ref exceldatas_down[k]);
                    }
                    exceldatas_downKM = new MyExcelData[_DownXlsFilesKM.Length];
                    for (int k = 0; k < _DownXlsFilesKM.Length; ++k)
                    {
                        exceldatas_downKM[k] = new MyExcelData();
                        ReadExcelData(excelApp, _DownXlsFilesKM[k][j], MergeInfo[j]._ExcelStartRow, ref exceldatas_downKM[k]);
                    }
                }

                //开始合并
                switch (j)
                {
                    case 5: MergeExcel1_SMTD(excelApp, exceldatas_up, exceldatas_down, exceldatas_upKM, exceldatas_downKM,
                        _OutputPath, MergeInfo[j], MergeType); break;
                    case 1: MergeExcel1_IRI(excelApp, exceldatas_up, exceldatas_down, exceldatas_upKM, exceldatas_downKM,
                         _OutputPath, MergeInfo[j], MergeType); break;
                    default: break;
                }
            }
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

            MSExcel.Worksheet worksheet = workbook.Worksheets["Sheet1"] as MSExcel.Worksheet;
            exceldata.datarow = GlobalExcel.judegeusedrow(worksheet, 1, startrow);
            exceldata.datacol = GlobalExcel.judegeusedcol(worksheet, startrow, 1);

            MSExcel.Range workrange = worksheet.get_Range(string.Format("A{0}:{1}{2}", startrow, GlobalExcel.GetCol((char)(exceldata.datacol - 1 + 'A')), exceldata.datarow));
            exceldata.dataobj = (object[,])workrange.Value2;
            exceldata.datarow = exceldata.datarow - startrow + 1;

            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }

        class MyExcelData
        {
            public int datarow = 0;
            public int datacol = 0;
            public object[,] dataobj = null;
        }

        /// <summary>
        /// 合并SMTD数据
        /// </summary>
        /// <param name="excelApp"></param>
        /// <param name="exceldatas_up">上行车道数据</param>
        /// <param name="exceldatas_down">下行车道数据</param>
        private static void MergeExcel1_SMTD(MSExcel.Application excelApp,
            MyExcelData[] exceldatas_up, MyExcelData[] exceldatas_down,
            MyExcelData[] exceldatas_upKM, MyExcelData[] exceldatas_downKM, 
            string _OutputPath, MergeIndexInfo mergeinfo, int MergeType)
        {
            string[] MergeTypeStr = { "全幅", "上行右半幅", "下行左半幅" };
            string srcxls = string.Format(@"{0}\报表模板\多车道合并\模板1\{1}.xlsx",
                System.Windows.Forms.Application.StartupPath, mergeinfo._ExcelMergeName);
            MSExcel.Workbook workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            string destxls = string.Format(@"{0}\{1}_{2}_{3}m.xlsx", _OutputPath, mergeinfo._ExcelMergeName, MergeTypeStr[MergeType], mergeinfo._OriUnitLen);
            workbook.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet worksheet = workbook.Worksheets["数据表"] as MSExcel.Worksheet;

            //数据表
            int uprow = 0;
            if (exceldatas_up != null)
            {
                for (int i = 0; i < exceldatas_up.Length; ++i)
                {
                    uprow += exceldatas_up[i].datarow;
                }
            }

            int downrow = 0;
            if (exceldatas_down != null)
            {
                for (int i = 0; i < exceldatas_down.Length; ++i)
                {
                    downrow += exceldatas_down[i].datarow;
                }
            }
            
            int maxrow = Math.Max(uprow, downrow);
            object[,] objval = new object[maxrow, 28];

            int numval = 0;
            if (exceldatas_up != null)
            {
                for (int i = 0; i < exceldatas_up.Length; ++i)
                {
                    for (int j = 1; j <= exceldatas_up[i].datarow; ++j)
                    {
                        objval[numval, 0] = numval + 1;
                        objval[numval, 2] = "右幅" + exceldatas_up[i].dataobj[j, 3].ToString();
                        objval[numval, 6] = exceldatas_up[i].dataobj[j, 1];
                        objval[numval, 9] = "一";
                        objval[numval, 10] = exceldatas_up[i].dataobj[j, 2];
                        objval[numval, 13] = exceldatas_up[i].dataobj[j, 6];
                        ++numval;
                    }
                }
            }

            numval = 0;
            if (exceldatas_down != null)
            {
                for (int i = 0; i < exceldatas_down.Length; ++i)
                {
                    for (int j = 1; j <= exceldatas_down[i].datarow; ++j)
                    {
                        objval[numval, 14] = numval + 1;
                        objval[numval, 16] = "左幅" + exceldatas_down[i].dataobj[j, 3].ToString();
                        objval[numval, 20] = exceldatas_down[i].dataobj[j, 1];
                        objval[numval, 23] = "一";
                        objval[numval, 24] = exceldatas_down[i].dataobj[j, 2];
                        objval[numval, 27] = exceldatas_down[i].dataobj[j, 6];
                        ++numval;
                    }
                }
            }

            MSExcel.Range destrange = worksheet.get_Range(string.Format("A5:AB{0}", maxrow + 4));
            destrange.Value2 = objval;
            GlobalExcel.SetBorderLine(destrange, 63);

            int rowidx = 0;
            for (int i = 0; i < maxrow; ++i )
            {
                rowidx = i + 5;

                destrange = worksheet.get_Range(string.Format("A{0}:B{0}", rowidx));//获取单元格
                destrange.MergeCells = true; //合并单元格

                destrange = worksheet.get_Range(string.Format("C{0}:F{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("G{0}:I{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("K{0}:M{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("O{0}:P{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("Q{0}:T{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("U{0}:W{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("Y{0}:AA{0}", rowidx));
                destrange.MergeCells = true;
            }

            //公里的报告
            worksheet = workbook.Worksheets["报告"] as MSExcel.Worksheet;

            int uprowKM = 0;
            if (exceldatas_upKM != null)
            {
                for (int i = 0; i < exceldatas_upKM.Length; ++i)
                {
                    if (exceldatas_upKM[i].datarow > uprowKM)
                        uprowKM = exceldatas_upKM[i].datarow;
                }
            }

            int downrowKM = 0;
            if (exceldatas_downKM != null)
            {
                for (int i = 0; i < exceldatas_downKM.Length; ++i)
                {
                    if (exceldatas_downKM[i].datarow > downrowKM)
                        downrowKM = exceldatas_downKM[i].datarow;
                }
            }

            int maxrowKM = uprowKM + downrowKM;
            object[,] objvalKM = new object[maxrowKM, 33];
            int rownumKM = 0;
            int[] currows = null;

            if (exceldatas_up != null)
            {
                currows = new int[exceldatas_up.Length];
                for (int i = 0; i < exceldatas_upKM.Length; ++i)
                {
                    for (int j = 0; j < exceldatas_up.Length; ++j)
                    {
                        currows[j] = 1;
                    }
                    for (int j = 1; j <= exceldatas_upKM[i].datarow; ++j)
                    {
                        objvalKM[rownumKM, 0] = rownumKM + 1;
                        objvalKM[rownumKM, 2] = exceldatas_upKM[i].dataobj[j, 1];
                        objvalKM[rownumKM, 6] = exceldatas_upKM[i].dataobj[j, 2];
                        objvalKM[rownumKM, 10] = "右幅车道";

                        double sumval = 0;
                        int ptnumsum = 0;
                        int ptnum = 0;
                        int emile = Convert.ToInt32(exceldatas_upKM[i].dataobj[j, 2]);
                        for (int k = 0; k < exceldatas_up.Length; ++k)
                        {
                            while (currows[k] <= exceldatas_up[k].datarow && Convert.ToInt32(exceldatas_up[k].dataobj[currows[k], 2]) <= emile)
                            {
                                double curval = Convert.ToDouble(exceldatas_up[k].dataobj[currows[k], 6]);
                                sumval += curval;
                                ++ptnumsum;
                                ++currows[k];
                                if (mergeinfo._ThreshType == 0)
                                {
                                    if (curval >= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                                else
                                {
                                    if (curval <= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                            };
                        }
                        objvalKM[rownumKM, 13] = ptnumsum;
                        objvalKM[rownumKM, 18] = ptnum;
                        objvalKM[rownumKM, 23] = sumval / ptnumsum;
                        objvalKM[rownumKM, 28] = string.Format("=S{0}/N{0}*100", rownumKM + 13);
                        ++rownumKM;
                    }
                    break;
                }
            }

            if (exceldatas_down != null)
            {
                currows = new int[exceldatas_down.Length];
                for (int i = 0; i < exceldatas_downKM.Length; ++i)
                {
                    for (int j = 0; j < exceldatas_down.Length; ++j)
                    {
                        currows[j] = 1;
                    }
                    for (int j = 1; j <= exceldatas_downKM[i].datarow; ++j)
                    {
                        objvalKM[rownumKM, 0] = rownumKM + 1;
                        objvalKM[rownumKM, 2] = exceldatas_downKM[i].dataobj[j, 1];
                        objvalKM[rownumKM, 6] = exceldatas_downKM[i].dataobj[j, 2];
                        objvalKM[rownumKM, 10] = "左幅车道";

                        double sumval = 0;
                        int ptnumsum = 0;
                        int ptnum = 0;
                        int emile = Convert.ToInt32(exceldatas_downKM[i].dataobj[j, 2]);
                        for (int k = 0; k < exceldatas_down.Length; ++k)
                        {
                            while (currows[k] <= exceldatas_down[k].datarow && Convert.ToInt32(exceldatas_down[k].dataobj[currows[k], 2]) <= emile)
                            {
                                double curval = Convert.ToDouble(exceldatas_down[k].dataobj[currows[k], 6]);
                                sumval += curval;
                                ++ptnumsum;
                                ++currows[k];
                                if (mergeinfo._ThreshType == 0)
                                {
                                    if (curval >= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                                else
                                {
                                    if (curval <= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                            };
                        }
                        objvalKM[rownumKM, 13] = ptnumsum;
                        objvalKM[rownumKM, 18] = ptnum;
                        objvalKM[rownumKM, 23] = sumval / ptnumsum;
                        objvalKM[rownumKM, 28] = string.Format("=S{0}/N{0}*100", rownumKM+13);
                        ++rownumKM;
                    }
                    break;
                }
            }

            destrange = worksheet.get_Range(string.Format("A13:AG{0}", rownumKM+12));
            destrange.Value2 = objvalKM;
            GlobalExcel.SetBorderLine(destrange, 63);

            rowidx = 0;
            for (int i = 0; i < rownumKM; ++i)
            {
                rowidx = i + 13;

                destrange = worksheet.get_Range(string.Format("A{0}:B{0}", rowidx));//获取单元格
                destrange.MergeCells = true; //合并单元格

                destrange = worksheet.get_Range(string.Format("C{0}:F{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("G{0}:J{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("K{0}:M{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("N{0}:R{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("S{0}:W{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("X{0}:AB{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("AC{0}:AG{0}", rowidx));
                destrange.MergeCells = true;
            }

            if (mergeinfo._ThreshType == 0)
            {
                worksheet.Cells[11, 29] = "≥" + mergeinfo._ThreshVal.ToString("0.00");
            }
            else
            {
                worksheet.Cells[11, 29] = "≤" + mergeinfo._ThreshVal.ToString("0.00");
            }

            workbook.Save();
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
    
        private static void MergeExcel1_IRI(MSExcel.Application excelApp,
            MyExcelData[] exceldatas_up, MyExcelData[] exceldatas_down,
            MyExcelData[] exceldatas_upKM, MyExcelData[] exceldatas_downKM, 
            string _OutputPath, MergeIndexInfo mergeinfo, int MergeType)
        {
            string[] MergeTypeStr = {"全幅", "上行右半幅", "下行左半幅" };
            string srcxls = string.Format(@"{0}\报表模板\多车道合并\模板1\{1}.xlsx",
                System.Windows.Forms.Application.StartupPath, mergeinfo._ExcelMergeName);
            MSExcel.Workbook workbook = excelApp.Workbooks.Open(srcxls, Type.Missing,
                true, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            string destxls = string.Format(@"{0}\{1}_{2}_{3}m.xlsx", _OutputPath, mergeinfo._ExcelMergeName, MergeTypeStr[MergeType], mergeinfo._OriUnitLen);
            workbook.SaveAs(destxls, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                MSExcel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            MSExcel.Worksheet worksheet = workbook.Worksheets["数据表"] as MSExcel.Worksheet;

            //数据表
            int uprow = 0;
            if (exceldatas_up != null)
            {
                for (int i = 0; i < exceldatas_up.Length; ++i)
                {
                    uprow += exceldatas_up[i].datarow;
                }
            }

            int downrow = 0;
            if (exceldatas_down != null)
            {
                for (int i = 0; i < exceldatas_down.Length; ++i)
                {
                    downrow += exceldatas_down[i].datarow;
                }
            }
            
            int maxrow = Math.Max(uprow, downrow);
            object[,] objval = new object[maxrow, 12];

            int numval = 0;
            if (exceldatas_up != null)
            {
                for (int i = 0; i < exceldatas_up.Length; ++i)
                {
                    for (int j = 1; j <= exceldatas_up[i].datarow; ++j)
                    {
                        objval[numval, 0] = numval + 1;
                        objval[numval, 1] = exceldatas_up[i].dataobj[j, 1];
                        objval[numval, 2] = exceldatas_up[i].dataobj[j, 2];
                        objval[numval, 3] = "右幅" + exceldatas_up[i].dataobj[j, 3].ToString();
                        objval[numval, 4] = exceldatas_up[i].dataobj[j, 6];
                        objval[numval, 5] = string.Format("=E{0}*0.6", numval + 5);
                        ++numval;
                    }
                }
            }

            numval = 0;
            if (exceldatas_down != null)
            {
                for (int i = 0; i < exceldatas_down.Length; ++i)
                {
                    for (int j = 1; j <= exceldatas_down[i].datarow; ++j)
                    {
                        objval[numval, 6] = numval + 1;
                        objval[numval, 7] = exceldatas_down[i].dataobj[j, 1];
                        objval[numval, 8] = exceldatas_down[i].dataobj[j, 2];
                        objval[numval, 9] = "左幅" + exceldatas_down[i].dataobj[j, 3].ToString();
                        objval[numval, 10] = exceldatas_down[i].dataobj[j, 6];
                        objval[numval, 11] = string.Format("=E{0}*0.6", numval + 5);
                        ++numval;
                    }
                }
            }

            MSExcel.Range destrange = worksheet.get_Range(string.Format("A5:L{0}", maxrow + 4));
            destrange.Value2 = objval;
            GlobalExcel.SetBorderLine(destrange, 63);
            
            //公里的报告
            worksheet = workbook.Worksheets["报告"] as MSExcel.Worksheet;

            int uprowKM = 0;
            if (exceldatas_upKM != null)
            {
                for (int i = 0; i < exceldatas_upKM.Length; ++i)
                {
                    if (exceldatas_upKM[i].datarow > uprowKM)
                        uprowKM = exceldatas_upKM[i].datarow;
                }
            }

            int downrowKM = 0;
            if (exceldatas_downKM != null)
            {
                for (int i = 0; i < exceldatas_downKM.Length; ++i)
                {
                    if (exceldatas_downKM[i].datarow > downrowKM)
                        downrowKM = exceldatas_downKM[i].datarow;
                }
            }

            int maxrowKM = uprowKM + downrowKM;
            object[,] objvalKM = new object[maxrowKM, 33];
            int rownumKM = 0;
            int[] currows = null;
            int rownumKMUp = 0;

            if (exceldatas_up != null)
            {
                currows = new int[exceldatas_up.Length];
                for (int i = 0; i < exceldatas_upKM.Length; ++i)
                {
                    for (int j = 0; j < exceldatas_up.Length; ++j)
                    {
                        currows[j] = 1;
                    }
                    for (int j = 1; j <= exceldatas_upKM[i].datarow; ++j)
                    {
                        objvalKM[rownumKM, 0] = rownumKM + 1;
                        objvalKM[rownumKM, 2] = exceldatas_upKM[i].dataobj[j, 1];
                        objvalKM[rownumKM, 6] = exceldatas_upKM[i].dataobj[j, 2];

                        double sumval = 0;
                        int ptnumsum = 0;
                        int ptnum = 0;
                        int emile = Convert.ToInt32(exceldatas_upKM[i].dataobj[j, 2]);
                        for (int k = 0; k < exceldatas_up.Length; ++k)
                        {
                            while (currows[k] <= exceldatas_up[k].datarow && Convert.ToInt32(exceldatas_up[k].dataobj[currows[k], 2]) <= emile)
                            {
                                double curval = Convert.ToDouble(exceldatas_up[k].dataobj[currows[k], 6]);
                                sumval += curval;
                                ++ptnumsum;
                                ++currows[k];
                                if (mergeinfo._ThreshType == 0)
                                {
                                    if (curval >= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                                else
                                {
                                    if (curval <= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                            };
                        }
                        objvalKM[rownumKM, 10] = sumval / ptnumsum;
                        objvalKM[rownumKM, 14] = string.Format("=K{0}*0.6", rownumKM + 14);
                        objvalKM[rownumKM, 18] = string.Format("=AE{0}/AA{0}*100", rownumKM + 14); ;
                        objvalKM[rownumKM, 22] = string.Format("=S{0}", rownumKM + 14);
                        objvalKM[rownumKM, 26] = ptnumsum;
                        objvalKM[rownumKM, 30] = ptnum;
                        ++rownumKM;
                    }
                    break;
                }
            }
            rownumKMUp = rownumKM;

            if (exceldatas_down != null)
            {
                currows = new int[exceldatas_down.Length];
                for (int i = 0; i < exceldatas_downKM.Length; ++i)
                {
                    for (int j = 0; j < exceldatas_down.Length; ++j)
                    {
                        currows[j] = 1;
                    }
                    for (int j = 1; j <= exceldatas_downKM[i].datarow; ++j)
                    {
                        objvalKM[rownumKM, 0] = rownumKM + 1;
                        objvalKM[rownumKM, 2] = exceldatas_downKM[i].dataobj[j, 1];
                        objvalKM[rownumKM, 6] = exceldatas_downKM[i].dataobj[j, 2];

                        double sumval = 0;
                        int ptnumsum = 0;
                        int ptnum = 0;
                        int emile = Convert.ToInt32(exceldatas_downKM[i].dataobj[j, 2]);
                        for (int k = 0; k < exceldatas_down.Length; ++k)
                        {
                            while (currows[k] <= exceldatas_down[k].datarow && Convert.ToInt32(exceldatas_down[k].dataobj[currows[k], 2]) <= emile)
                            {
                                double curval = Convert.ToDouble(exceldatas_down[k].dataobj[currows[k], 6]);
                                sumval += curval;
                                ++ptnumsum;
                                ++currows[k];
                                if (mergeinfo._ThreshType == 0)
                                {
                                    if (curval >= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                                else
                                {
                                    if (curval <= mergeinfo._ThreshVal)
                                    {
                                        ++ptnum;
                                    }
                                }
                            };
                        }
                        objvalKM[rownumKM, 10] = sumval / ptnumsum;
                        objvalKM[rownumKM, 14] = string.Format("=K{0}*0.6", rownumKM + 14);
                        objvalKM[rownumKM, 18] = string.Format("=AE{0}/AA{0}*100", rownumKM + 14); ;
                        objvalKM[rownumKM, 22] = string.Format("=S{0}", rownumKM + 14);
                        objvalKM[rownumKM, 26] = ptnumsum;
                        objvalKM[rownumKM, 30] = ptnum;
                        ++rownumKM;
                    }
                    break;
                }
            }

            destrange = worksheet.get_Range(string.Format("A14:AG{0}", rownumKM + 13));
            destrange.Value2 = objvalKM;
            GlobalExcel.SetBorderLine(destrange, 63);

            destrange = worksheet.get_Range(string.Format("C14:J{0}", rownumKMUp + 13));
            destrange.NumberFormatLocal = "\"Y\"K0+000";
            destrange = worksheet.get_Range(string.Format("C{0}:J{1}", rownumKMUp + 14, rownumKM + 13));
            destrange.NumberFormatLocal = "\"Z\"K0+000";

            int rowidx = 0;
            for (int i = 0; i < rownumKM; ++i)
            {
                rowidx = i + 14;

                destrange = worksheet.get_Range(string.Format("A{0}:B{0}", rowidx));//获取单元格
                destrange.MergeCells = true; //合并单元格

                destrange = worksheet.get_Range(string.Format("C{0}:F{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("G{0}:J{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("K{0}:N{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("O{0}:R{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("S{0}:V{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("W{0}:Z{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("AA{0}:AD{0}", rowidx));
                destrange.MergeCells = true;

                destrange = worksheet.get_Range(string.Format("AE{0}:AG{0}", rowidx));
                destrange.MergeCells = true;
            }

            if (mergeinfo._ThreshType == 0)
            {
                worksheet.Cells[11, 30] = "≥" + mergeinfo._ThreshVal.ToString("0.00");
            }
            else
            {
                worksheet.Cells[11, 30] = "≤" + mergeinfo._ThreshVal.ToString("0.00");
            }

            workbook.Save();
            workbook.Close(Type.Missing, Type.Missing, Type.Missing);
            int generation = System.GC.GetGeneration(excelApp);
            System.GC.Collect(generation);//垃圾回收
            System.GC.WaitForPendingFinalizers();
        }
    
    }
}
