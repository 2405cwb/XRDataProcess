using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace Farmework.Office.Excel
{

    /// <summary>
    /// NPOI Excel 读写帮助类
    /// 支持 .xls 和 .xlsx，带进度报告（IProgress<int>），自动资源释放
    /// </summary>
    public static class ExcelHelper_NPOI
    {
        #region 写入 Excel（泛型导出）

        /// <summary>
        /// 将 List<T> 导出为 Excel 文件
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="data">要导出的数据列表</param>
        /// <param name="filePath">保存路径（.xlsx 或 .xls）</param>
        /// <param name="sheetName">工作表名称，默认为 "Sheet1"</param>
        /// <param name="progress">进度报告（0-100），可为 null</param>
        public static void ExportToExcel<T>(
            List<T> data,
            string filePath,
            string sheetName = "Sheet1",
            IProgress<int> progress = null)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("文件路径不能为空", nameof(filePath));

            IWorkbook workbook = CreateWorkbook(filePath);
            ISheet sheet = workbook.CreateSheet(sheetName);

            // 创建表头样式
            ICellStyle headerStyle = CreateHeaderStyle(workbook);

            // 获取属性（支持 public get 属性）
            PropertyInfo[] properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead)
                .ToArray();

            // 写入表头
            IRow headerRow = sheet.CreateRow(0);
            for (int i = 0; i < properties.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(properties[i].Name);
                cell.CellStyle = headerStyle;
            }

            // 写入数据行
            int total = data.Count;
            for (int i = 0; i < total; i++)
            {
                IRow row = sheet.CreateRow(i + 1);
                for (int j = 0; j < properties.Length; j++)
                {
                    var value = properties[j].GetValue(data[i])?.ToString() ?? "";
                    row.CreateCell(j).SetCellValue(value);
                }

                // 报告进度
                int percent = (int)((i + 1) * 100.0 / total);
                progress?.Report(percent);
            }

            // 自动列宽
            for (int i = 0; i < properties.Length; i++)
            {
                sheet.AutoSizeColumn(i);

                double widthDouble = sheet.GetColumnWidth(i);
                int newWidth = (int)Math.Min(widthDouble + 1000, int.MaxValue);
                sheet.SetColumnWidth(i, newWidth);
            }

            // 保存文件
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            progress?.Report(100);
        }

        #endregion

        #region 读取 Excel

        /// <summary>
        /// 读取 Excel 文件并转换为 List<T>
        /// </summary>
        /// <typeparam name="T">目标类型，属性名需与 Excel 表头一致</typeparam>
        /// <param name="filePath">Excel 文件路径</param>
        /// <param name="sheetIndex">工作表索引，默认为 0</param>
        /// <param name="hasHeader">第一行是否为表头，默认为 true</param>
        /// <param name="progress">进度报告（0-100），可为 null</param>
        /// <returns>解析后的数据列表</returns>
        public static List<T> ImportFromExcel<T>(
            string filePath,
            int sheetIndex = 0,
            int dataStartIndex = 0 ,
            bool hasHeader = true,
            IProgress<int> progress = null, 
            int processStartValue = 0 ) where T : new()
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel 文件未找到", filePath);

            var result = new List<T>();
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = CreateWorkbook(filePath, fs);
            ISheet sheet = workbook.GetSheetAt(sheetIndex) ?? throw new ArgumentException($"工作表索引 {sheetIndex} 不存在");

            int startRow = hasHeader ? 1 : 0;
            int totalRows = sheet.PhysicalNumberOfRows - startRow;
            if (totalRows <= 0) return result;

            // 获取属性映射
            var properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanWrite)
                .ToDictionary(p => p.Name, p => p, StringComparer.OrdinalIgnoreCase);

            // 读取表头（如果有）
            Dictionary<int, string> headerMap = new();
            if (hasHeader)
            {
                IRow headerRow = sheet.GetRow(0);
                for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
                {
                    var cell = headerRow.GetCell(i);
                    string header = cell?.ToString()?.Trim() ?? "";
                    headerMap[i] = header;
                }
            }

            // 读取数据行
            for (int i = dataStartIndex; i < totalRows; i++)
            {
                IRow row = sheet.GetRow(startRow + i);
                if (row == null) continue;

                T item = new T();
                bool hasData = false;

                foreach (var kv in hasHeader ? headerMap : Enumerable.Range(0, row.LastCellNum).Select(x => new KeyValuePair<int, string>(x, properties.Keys.ElementAtOrDefault(x) ?? "")))
                {
                    int colIndex = kv.Key;
                    string propName = kv.Value;
                    if (!properties.TryGetValue(propName, out PropertyInfo prop)) continue;

                    var cell = row.GetCell(colIndex);
                    object value = GetCellValue(cell, prop.PropertyType);
                    if (value != null)
                    {
                        prop.SetValue(item, value);
                        hasData = true;
                    }
                }

                if (hasData) result.Add(item);

                // 报告进度
                int percent =  (int)((i + 1) * 100.0 / totalRows);
                progress?.Report(percent+ processStartValue);
            }

            //progress?.Report(100);
            return result;
        }

        /// <summary>
        /// 读取 Excel 文件并转换为 List<T>
        /// </summary>
        /// <typeparam name="T">目标类型，属性名需与 Excel 表头一致</typeparam>
        /// <param name="filePath">Excel 文件路径</param>
        /// <param name="sheetIndex">工作表索引，默认为 0</param>
        /// <param name="hasHeader">第一行是否为表头，默认为 true</param>
        /// <param name="progress">进度报告（0-100），可为 null</param>
        /// <returns>解析后的数据列表</returns>
        public static List<T> ImportFromExcel<T>(
            string filePath,
            string sheetName = "",
            int dataStartIndex = 0,
            bool hasHeader = true,
            IProgress<int> progress = null,
            int processStartValue = 0) where T : new()
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException("Excel 文件未找到", filePath);

            var result = new List<T>();
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = CreateWorkbook(filePath, fs);
            ISheet sheet = workbook.GetSheet(sheetName) ?? throw new ArgumentException($"工作表索引 {sheetName} 不存在");

            int startRow = hasHeader ? 1 : 0;
            int totalRows = sheet.PhysicalNumberOfRows - startRow;
            if (totalRows <= 0) return result;

            // 获取属性映射
            var properties = typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanWrite)
                .ToDictionary(p => p.Name, p => p, StringComparer.OrdinalIgnoreCase);

            // 读取表头（如果有）
            Dictionary<int, string> headerMap = new();
            if (hasHeader)
            {
                IRow headerRow = sheet.GetRow(0);
                for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
                {
                    var cell = headerRow.GetCell(i);
                    string header = cell?.ToString()?.Trim() ?? "";
                    headerMap[i] = header;
                }
            }

            // 读取数据行
            for (int i = dataStartIndex; i < totalRows; i++)
            {
                IRow row = sheet.GetRow(startRow + i);
                if (row == null) continue;

                T item = new T();
                bool hasData = false;

                foreach (var kv in hasHeader ? headerMap : Enumerable.Range(0, row.LastCellNum).Select(x => new KeyValuePair<int, string>(x, properties.Keys.ElementAtOrDefault(x) ?? "")))
                {
                    int colIndex = kv.Key;
                    string propName = kv.Value;
                    if (!properties.TryGetValue(propName, out PropertyInfo prop)) continue;

                    var cell = row.GetCell(colIndex);
                    object value = GetCellValue(cell, prop.PropertyType);
                    if (value != null)
                    {
                        prop.SetValue(item, value);
                        hasData = true;
                    }
                }

                if (hasData) result.Add(item);

                // 报告进度
                int percent = (int)((i + 1) * 100.0 / totalRows);
                progress?.Report(percent + processStartValue);
            }

            //progress?.Report(100);
            return result;
        }

        #endregion

        #region 私有辅助方法

        /// <summary>
        /// 创建工作簿（根据文件扩展名）
        /// </summary>
        private static IWorkbook CreateWorkbook(string filePath, Stream stream = null)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext switch
            {
                ".xls" => stream != null ? new HSSFWorkbook(stream) : new HSSFWorkbook(),
                ".xlsx" => stream != null ? new XSSFWorkbook(stream) : new XSSFWorkbook(),
                _ => throw new NotSupportedException("不支持的文件格式，仅支持 .xls 和 .xlsx")
            };
        }

        /// <summary>
        /// 创建表头样式
        /// </summary>
        private static ICellStyle CreateHeaderStyle(IWorkbook workbook)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            style.FillPattern = FillPattern.SolidForeground;

            NPOI.SS.UserModel.IFont font = workbook.CreateFont();
            font.IsBold = true;
            font.Color = IndexedColors.White.Index;
            style.SetFont(font);

            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            return style;
        }

        /// <summary>
        /// 获取单元格值并转换类型
        /// </summary>
        private static object GetCellValue(ICell cell, Type targetType)
        {
            if (cell == null) return null;

            return cell.CellType switch
            {
                CellType.String => Convert.ChangeType(cell.StringCellValue.Trim(), targetType),
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                    ? cell.DateCellValue
                    : Convert.ChangeType(cell.NumericCellValue, targetType),
                CellType.Boolean => cell.BooleanCellValue,
                CellType.Formula => GetCellValue(cell, cell.CachedFormulaResultType, targetType),
                _ => null
            };
        }

        private static object GetCellValue(ICell cell, CellType resultType, Type targetType)
        {
            return resultType switch
            {
                CellType.String => Convert.ChangeType(cell.StringCellValue.Trim(), targetType),
                CellType.Numeric => Convert.ChangeType(cell.NumericCellValue, targetType),
                CellType.Boolean => cell.BooleanCellValue,
                _ => null
            };
        }

        #endregion


    }
}
