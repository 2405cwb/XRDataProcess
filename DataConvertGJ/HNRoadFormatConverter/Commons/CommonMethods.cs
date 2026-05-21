using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace HNRoadFormatConverter.Commons
{

    public static class CellValueParser
    {
        // 缓存公式计算器（每个 Workbook 独立）
        private static readonly ConditionalWeakTable<IWorkbook, IFormulaEvaluator> _evaluatorCache = new ConditionalWeakTable<IWorkbook, IFormulaEvaluator>();
        private static readonly object _cacheLock = new object(); // 全局缓存锁

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static object GetCellValueHighSpeed(ICell cell)
        {
            if (cell == null) return DBNull.Value;

            // 快速路径：优先处理常见类型
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return (object)cell.DateCellValue;
                    return (object)cell.NumericCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.String:
                    return cell.StringCellValue.AsSpan().Trim().ToString();

                case CellType.Formula:
                    return ProcessFormulaCell(cell);

                default:
                    return DBNull.Value;
            }
        }

        private static object ProcessFormulaCell(ICell cell)
        {
            var workbook = cell.Sheet.Workbook;
            // 第一层无锁快速检查
            if (_evaluatorCache.TryGetValue(workbook, out var evaluator))
                return EvaluateWithCached(evaluator, cell);

            // 进入锁区域
            lock (_cacheLock)
            {
                // 双重检查锁定模式
                if (!_evaluatorCache.TryGetValue(workbook, out evaluator))
                {
                    evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                    _evaluatorCache.Add(workbook, evaluator);
                }
            }

            return EvaluateWithCached(evaluator, cell);
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static object EvaluateWithCached(IFormulaEvaluator evaluator, ICell cell)
        {
            try
            {
                var value = evaluator.Evaluate(cell);
                return value.CellType switch
                {
                    CellType.Numeric => value.NumberValue,
                    CellType.Boolean => value.BooleanValue,
                    CellType.String => value.StringValue.AsSpan().Trim().ToString(),
                    _ => DBNull.Value
                };
            }
            catch (Exception ex)
            {
                return cell.ToString().AsSpan().Trim().ToString();
            }
        }
    }
    public class CommonMethods
    {
        /// <summary>
        /// 获取单元格的值（支持多种类型）。
        /// </summary>
        /// <param name="cell">NPOI 的 ICell 对象</param>
        /// <returns>单元格值的对象表示</returns>
        [MethodImpl(MethodImplOptions.AggressiveInlining)] // 内联优化
        public static object GetCellValueNomal(ICell cell)
        {
            if (cell == null) return DBNull.Value;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return (object)cell.DateCellValue;
                    return (object)cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    // 处理公式计算结果
                    IFormulaEvaluator evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                    CellValue formulaValue = evaluator.Evaluate(cell);
                    return formulaValue.FormatAsString();
                default:
                    return DBNull.Value;
            }
        }
    }




}
