using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;

namespace Hiz.Extended.Npoi
{
    partial class NpoiExtensions
    {
        #region ISheet~

        /* ISheet/Row
         * 
         * short DefaultRowHeight { get; set; }
         * float DefaultRowHeightInPoints { get; set; }
         * bool DisplayRowColHeadings { get; set; }
         * 
         * // 获取首行索引 (0 based); 如果没有数据, 依然返回零值 (不太合理).
         * int FirstRowNum { get; }
         * // 获取尾行索引 (0 based); 如果没有数据, 依然返回零值 (不太合理).
         * int LastRowNum { get; }
         * 
         * // Returns the number of physically defined rows (NOT the number of rows in the sheet).
         * int PhysicalNumberOfRows { get; }
         * 
         * CellRangeAddress RepeatingRows { get; set; }
         * int[] RowBreaks { get; }
         * bool RowSumsBelow { get; set; }
         * bool RowSumsRight { get; set; }
         * short TopRow { get; set; }
         * 
         * IRow CopyRow(int sourceIndex, int targetIndex);
         * void CreateFreezePane(int colSplit, int rowSplit);
         * void CreateFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow);
         * IRow CreateRow(int rownum);
         * IRow GetRow(int rownum);
         * IEnumerator GetRowEnumerator();
         * void GroupRow(int fromRow, int toRow);
         * bool IsRowBroken(int row);
         * void RemoveRow(IRow row);
         * void RemoveRowBreak(int row);
         * void SetRowBreak(int row);
         * void SetRowGroupCollapsed(int row, bool collapse);
         * void ShiftRows(int startRow, int endRow, int n);
         * void ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight);
         * void UngroupRow(int fromRow, int toRow);
         */
        public static int GetMaxRowCount(ISheet sheet)
        {
            if (sheet is XSSFSheet || sheet is SXSSFSheet/*流式文件*/)
                return SpreadsheetVersion.EXCEL2007.MaxRows; // 0x100000 = 1048576

            if (sheet is HSSFSheet)
                return SpreadsheetVersion.EXCEL97.MaxRows; // 0x010000 = 65536

            throw new NotSupportedException();
        }
        // public static int GetMaxColumnCount(ISheet sheet)
        // {
        //     if (sheet is XSSFSheet || sheet is SXSSFSheet/*流式文件*/)
        //         return SpreadsheetVersion.EXCEL2007.MaxColumns;
        //     if (sheet is HSSFSheet)
        //         return SpreadsheetVersion.EXCEL97.MaxColumns;
        //     throw new NotSupportedException();
        // }

        public static bool HasRows(this ISheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            return sheet.PhysicalNumberOfRows > 0;
        }

        /// <summary>
        /// 枚举当前表单所有的有效行;
        /// HSSFSheet: 按创建的顺序输出;
        /// XSSFSheet: 按索引的顺序输出;
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IEnumerable<IRow> GetRows(this ISheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            // foreach (IRow row in sheet) // 为什么这么写不会抛出编译异常 ? ISheet 没有实现 IEnumerable.
            //     yield return row;

            // HSSFSheet/XSSFSheet: GetRowEnumerator() 内部调用 GetEnumerator();
            var enumerator = sheet.GetEnumerator();
            while (enumerator.MoveNext())
                yield return (IRow)enumerator.Current;
        }

        /// <summary>
        /// 返回第一个有效行的行索引; -1: 无任何行;
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static int GetFirstRowIndex(this ISheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            /* v2.3:
             * 以下两种情况 FirstRowNum 返回零值:
             * 1. ISheet 无任何行.
             * 2. ISheet 索引零的位置有行实例.
             * 
             * 因此此处第一行的索引需要修正.
             */
            if (sheet.PhysicalNumberOfRows > 0)
                return sheet.FirstRowNum;
            return -1;
        }

        /// <summary>
        /// 返回最末个有效行的行索引; -1: 无任何行;
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static int GetLastRowIndex(this ISheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            /* v2.3:
             * 以下两种情况 LastRowNum 返回零值:
             * 1. ISheet 无任何行.
             * 2. ISheet 只有一行, 并且位于首行.
             * 
             * 因此此处最末行的索引需要修正.
             */
            if (sheet.PhysicalNumberOfRows > 0)
                return sheet.LastRowNum;
            return -1;
        }

        /// <summary>
        /// 获取指定索引的行, 若没有则新建.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static IRow GetOrAddRow(this ISheet sheet, int index)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));
            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index), "不能负数");

            var row = sheet.GetRow(index);
            if (row == null)
            {
                row = sheet.CreateRow(index);
            }
            return row;
        }

        /// <summary>
        /// 在最末行后面追加新行.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static IRow AddRow(this ISheet sheet)
        {
            var last = sheet.GetLastRowIndex();
            var row = sheet.CreateRow(++last);
            return row;
        }
        public static IEnumerable<IRow> AddRowRange(this ISheet sheet, int count)
        {
            var last = sheet.GetLastRowIndex();
            var end = last + count;
            while (last < end)
            {
                yield return sheet.CreateRow(++last);
            }
        }
        #endregion

        #region ShiftRows~

        public static IRow InsertRow(this ISheet sheet, int index)
        {
            ShiftRowsBeforeInsert(sheet, index, 1);

            var row = sheet.GetRow(index);
            if (row == null)
                row = sheet.CreateRow(index);
            return row;
        }

        /// <summary>
        /// 插入指定行数, 插入位置后面的行全部下移.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="index">插入起始索引</param>
        /// <param name="count">行数</param>
        /// <returns></returns>
        public static IEnumerable<IRow> InsertRowRange(this ISheet sheet, int index, int count)
        {
            ShiftRowsBeforeInsert(sheet, index, count);

            var end = index + count;
            while (index < end)
            {
                // 方法1: 始终创建新行实例.
                // yield return sheet.CreateRow(index++);

                /* v2.3 经过测试:
                 * 调用 ShiftRows 之后, HSSFSheet 会先创建新行实例, 然后清空原行的单元格, 但是没有删除旧的实例, 因此可以拿来复用;
                 * XSSFSheet: 只是改变原行索引, 没有新建实例;
                 * 
                 * 方法2:
                 */
                var row = sheet.GetRow(index);
                if (row == null)
                    row = sheet.CreateRow(index);
                yield return row;
                index++;
            }
        }

        /* HSSFSheet:
         * ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight, bool moveComments) // 该方法不是由接口定义.
         * |=> ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight); // moveComments = true;
         *     |=> ShiftRows(int startRow, int endRow, int n); // copyRowHeight = false; resetOriginalRowHeight = false;
         * 
         * XSSFSheet:
         * ShiftRows(int startRow, int endRow, int n, bool copyRowHeight, bool resetOriginalRowHeight);
         * |=> ShiftRows(int startRow, int endRow, int n); // copyRowHeight = false; resetOriginalRowHeight = false;
         * 
         * SXSSFSheet: NotSupported
         */
        static void ShiftRowsBeforeInsert(ISheet sheet, int index, int count)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            if (index < 0)
                throw new ArgumentOutOfRangeException(nameof(index), "不能负数");
            if (count <= 0)
                throw new ArgumentOutOfRangeException(nameof(count), "必须正数");

            var max = GetMaxRowCount(sheet);
            if (index >= max)
                throw new ArgumentOutOfRangeException(nameof(index), "超出最大索引");

            var last = sheet.GetLastRowIndex();
            // 最后一行下移 count 行数, 可能超过最大行数.
            if (Math.Max(last, index) + count >= max)
                throw new ArgumentOutOfRangeException(nameof(index), "超出最大行数"); // NpoiExtensions.InsertRow() 没有 count 参数, 因此这里使用 index;

            if (last >= index)
            {
                /* v2.3 下移逻辑(简化说明)
                 * 
                 * HSSFSheet:
                 * 创建新行, 并复制原行的数据, 然后清空原行的单元格. (原行实例依然存在);
                 * copyRowHeight: 新行是否使用原行行高. 建议为真.
                 * resetOriginalRowHeight: 是否重置原行行高. 建议为真.
                 * 
                 * XSSFSheet:
                 * 改变原行索引, 没有创建新行.
                 * copyRowHeight: true: 保留行高; false: 重置行高为默认值;
                 * resetOriginalRowHeight: 不使用该参数; 如果在原行的位置创建新行, 行高为默认值;
                 * 
                 * 为了保持逻辑一致, 上面两个参数建议都为真值.
                 */
                sheet.ShiftRows(index, last, count, true, true); // 从插入位置到最后一行, 整体向下移动插入行数.
            }
        }

        /// <summary>
        /// 删除指定范围空行, 并将后面的有效行上移;
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="start"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public static int RemoveRowRangeWithEmpty(this ISheet sheet, int start = 0, int count = -1)
        {
            return sheet.RemoveRowRange(row => row == null || row.IsEmpty(), start, count, true);
        }

        public static int RemoveRowRange(this ISheet sheet, int start, int count = -1, bool shift = true)
        {
            return sheet.RemoveRowRange(row => true, start, count, true);
        }

        /// <summary>
        /// 根据指定条件批量删除行数, 并可选择是否上移后续行数.
        /// FirstRowIndex & LastRowIndex 范围之外不会测试匹配函数, 结果直接为真;
        /// 范围之内如果指定索引位置没有实例, 依然调用匹配函数, 参数 IRow 为空;
        /// 参数 shift 作用范围根据 start & count 决定;
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="predicate">匹配条件; 如果指定索引位置没有实例, IRow 为空.</param>
        /// <param name="start">起始行的索引</param>
        /// <param name="count">检查行数; -1: 检查到最末行.</param>
        /// <param name="shift">删除位置后面的行是否上移; 也作用于空行: 例如 1-5 空行 6-8 有值, 使用 Shift 删除两行 4-5, 后面 6-8 将会上移两行覆盖.</param>
        /// <returns>实际删除行实例的数量</returns>
        public static int RemoveRowRange(this ISheet sheet, Func<IRow, bool> predicate, int start = 0, int count = -1, bool shift = true)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));
            if (predicate == null)
                throw new ArgumentNullException(nameof(predicate));
            if (start < 0)
                throw new ArgumentOutOfRangeException(nameof(start), "不能负数");
            if (count == 0 || count < -1)
                throw new ArgumentOutOfRangeException(nameof(count), "必须正数或者负一");

            var removes = 0;
            if (sheet.PhysicalNumberOfRows > 0) // 存在行数
            {
                var max = GetMaxRowCount(sheet);
                if (start >= max) // 此处使用大于等于判断
                    throw new ArgumentOutOfRangeException(nameof(start), "超出最大索引");

                int index; // 当前行的索引
                var last = sheet.GetLastRowIndex(); // 最末个有效行的行索引
                if (count > 0) // 指定行数
                {
                    // // 取消异常判断
                    // var end = start + count;
                    // if (end > max) // 此处使用大于判断
                    //     throw new ArgumentOutOfRangeException(nameof(count), "超出最大行数");
                    var end = Math.Min(start + count, max); // 测试范围截止行的索引;

                    index = Math.Min(--end, last);
                }
                else // 检查到最末行
                {
                    index = last;
                }

                var first = sheet.GetFirstRowIndex(); // 第一个有效行的行索引
                var empties = 0; // 连续匹配数量
                var aggregate = 0; // 累计匹配(上移)数量
                while (index >= start && index >= first) // 从下往上删除; 移动行数从少到多;
                {
                    var row = sheet.GetRow(index);
                    if (predicate(row))
                    {
                        empties++; // 匹配计数加一

                        if (row != null)
                        {
                            sheet.RemoveRow(row); // 如果 shift = true 可以不用移除, 等待 ShiftRows 覆盖;
                            removes++; // 实例计数加一
                        }
                    }
                    else if (empties > 0) // 如果之间存在空行
                    {
                        if (shift)
                        {
                            // 上移起始索引 = 当前行的索引 + 中间空行数量 + 往下一行;
                            var next = index + empties;
                            sheet.ShiftRows(++next, last, -empties, true, true);
                            last -= empties; // 更新末行索引
                        }
                        aggregate += empties; // 累计
                        empties = 0; // 重新计数
                    }
                    index--; // 继续往上检查
                }

                /* 循环结束 empties 可能大于 0 (最后几行全部匹配, 但是还未处理.)
                 * 继续追加 start 与 first 之间的空行;
                 */
                if (start < first)
                {
                    empties += first - start;
                }
                if (empties > 0)
                {
                    if (shift)
                    {
                        // 最后将非空行 移到 start 位置;
                        sheet.ShiftRows(start + empties, last, -empties, true, true);
                    }
                    aggregate += empties; // 累计
                }
                /* HSSFSheet 调用 ShiftRows() 移动后原行位置会产生废行(有行实例, 但是 PhysicalNumberOfCells 为零);
                 * 例如:
                 * 原行索引 8, 现在移到索引 3, 移完之后, 会有两个实例: 索引 3 存放原行数据, 索引 8 存在一个空行实例;
                 * 
                 * XSSFSheet 直接修改原行索引, 不会产生废行.
                 */
                if (sheet is HSSFSheet && shift && aggregate > 0)
                {
                    RemoveRowArrayWithEmpty(sheet); // 删除 ShiftRows 所产生的废行
                }
            }
            return removes;
        }

        static int RemoveRowArrayWithEmpty(ISheet sheet)
        {
            /* 空行的几种情况:
             * 
             * 1. 没有实例: sheet.GetRow() == null;
             * 2. 无单元格: row.PhysicalNumberOfCells == 0;
             * 3. 所有单元格值都为空白: row.GetCells().All(c => c.CellType == CellType.Blank);
             * 4. 文本单元格值都为空格: row.GetCells().All(c => c.CellType == CellType.Blank || (c.CellType == CellType.String && string.IsNullOrWhiteSpace(c.StringCellValue)));
             * 
             * 情况 4 是否合理?
             */
            var nothings = sheet.GetRows().Where(r => r.PhysicalNumberOfCells == 0).ToArray();
            foreach (var row in nothings)
            {
                sheet.RemoveRow(row);
            }
            return nothings.Length; // 空行实例数量
        }

        #endregion

        #region ColumnWidth~

        /* ISheet/ColumnWidth
         * 
         * // Get the default column width for the sheet (if the columns do not define their own width) in characters.
         * int DefaultColumnWidth { get; set; }
         * 
         * 
         * // Get the width (in units of 1/256th of a character width).
         * int GetColumnWidth(int columnIndex);
         * 
         * // Please note, that this method works correctly only for workbooks with the default font size (Arial 10pt for .xls and Calibri 11pt for .xlsx).
         * // If the default font is changed the column width can be streched.
         * float GetColumnWidthInPixels(int columnIndex);
         * 
         * // Set the width (in units of 1/256th of a character width).
         * // The maximum column width for an individual cell is 255 characters.
         * // This value represents the number of characters that can be displayed in a cell that is formatted with the standard font.
         * void SetColumnWidth(int columnIndex, int width);
         * 
         * 
         * // Adjusts the column width to fit the contents.
         * // This process can be relatively slow on large sheets, so this should normally only be called once per column, at the end of your processing.
         * void AutoSizeColumn(int column);
         * 
         * // Whether to use the contents of merged cells when calculating the width of the column.
         * // Default is to ignore merged cells.
         * void AutoSizeColumn(int column, bool useMergedCells);
         */
        const float UnitsPerCharacter = 0x100; // 256;
        const float MaximumColumnWidthInCharacters = 255f; // 可显示的字符数量; 最大数值: 255.
        const int MaximumColumnWidthInUnits = 0xFF00; // MaximumColumnWidthInCharacters * UnitsPerCharacter;
        const float PixelsPerInch = 96f; //TODO: 像素并非物理单位, 需要根据显示设备调整;
        const float PointsPerInch = 72f;
        const float CentimetersPerInch = 2.54f;
        const float MillimetersPerInch = 25.4f;

        public static void SetColumnWidth(this ISheet sheet, int index, float width, LengthUnit unit)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            int units;
            switch (unit)
            {
                case LengthUnit.Character:
                    {
                        units = (int)(width * UnitsPerCharacter);
                    }
                    break;
                case LengthUnit.Pixel:
                    {
                        units = (int)EvaluateColumnWidthInPixels(sheet, width);
                    }
                    break;
                case LengthUnit.Point:
                    {
                        var pixels = width * PixelsPerInch / PointsPerInch;
                        units = (int)EvaluateColumnWidthInPixels(sheet, pixels);
                    }
                    break;
                case LengthUnit.Inche:
                    {
                        var pixels = width * PixelsPerInch;
                        units = (int)EvaluateColumnWidthInPixels(sheet, pixels);
                    }
                    break;
                case LengthUnit.Centimeter:
                    {
                        var pixels = width * PixelsPerInch / CentimetersPerInch;
                        units = (int)EvaluateColumnWidthInPixels(sheet, pixels);
                    }
                    break;
                case LengthUnit.Millimeter:
                    {
                        var pixels = width * PixelsPerInch / MillimetersPerInch;
                        units = (int)EvaluateColumnWidthInPixels(sheet, pixels);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(unit));
            }

            if (units > MaximumColumnWidthInUnits)
                throw new ArgumentOutOfRangeException(nameof(width));
            sheet.SetColumnWidth(index, units);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="pixels"></param>
        /// <returns>In units of 1/256th of a character width</returns>
        static float EvaluateColumnWidthInPixels(ISheet sheet, float pixels)
        {
            if (sheet is HSSFSheet)
            {
                // const float PX_DEFAULT = 32f; // 使用默认列宽时 每单位的像素点?
                // var @default = (int)(width * PX_DEFAULT);

                const float PX_MODIFIED = 36.56f; // UnitsPerPixel // 自定义列宽之后 每单位的像素点?
                var modified = (pixels * PX_MODIFIED);
                return modified;
            }
            else // XSSFSheet; SXSSFSheet;
            {
                var PixelsPerCharacter = XSSFWorkbook.DEFAULT_CHARACTER_WIDTH;
                var units = pixels * UnitsPerCharacter / PixelsPerCharacter;
                return units;
            }
        }

        public static float GetColumnWidth(this ISheet sheet, int index, LengthUnit unit)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));
            switch (unit)
            {
                case LengthUnit.Character:
                    {
                        var value = sheet.GetColumnWidth(index);
                        return (float)value / UnitsPerCharacter;
                    }
                case LengthUnit.Pixel:
                    {
                        /* XSSFSheet.GetColumnWidthInPixels() {
                         *     float value = (float)this.GetColumnWidth(columnIndex);
                         *     return (float)((double)value / 256.0 * (double)XSSFWorkbook.DEFAULT_CHARACTER_WIDTH);
                         * }
                         * 
                         * HSSFSheet.GetColumnWidthInPixels() {
                         *     int value = this.GetColumnWidth(column);
                         *     int @default = this.DefaultColumnWidth * 256;
                         *     float per = (value == default) ? HSSFSheet.PX_DEFAULT : HSSFSheet.PX_MODIFIED; // UnitsPerPixel; Unit = 1/256th of a character.
                         *     return (float)value / per;
                         * }
                         */
                        return sheet.GetColumnWidthInPixels(index);
                    }
                case LengthUnit.Point:
                    {
                        var pixels = sheet.GetColumnWidthInPixels(index);
                        return pixels * PointsPerInch / PixelsPerInch;
                    }
                case LengthUnit.Inche:
                    {
                        var pixels = sheet.GetColumnWidthInPixels(index);
                        return pixels / PixelsPerInch;
                    }
                case LengthUnit.Centimeter:
                    {
                        var pixels = sheet.GetColumnWidthInPixels(index);
                        return pixels * CentimetersPerInch / PixelsPerInch;
                    }
                case LengthUnit.Millimeter:
                    {
                        var pixels = sheet.GetColumnWidthInPixels(index);
                        return pixels * MillimetersPerInch / PixelsPerInch;
                    }
                default:
                    throw new ArgumentOutOfRangeException(nameof(unit));
            }
        }

        /* v2.4 SheetUtil.GetColumnWidth() 取消 firstRow/lastRow 参数.
         * 改用: ISheet.AutoSizeColumn(int column, bool useMergedCells);
         * 
         * /// <summary>
         * /// 局部自动列宽.
         * /// </summary>
         * /// <param name="sheet"></param>
         * /// <param name="column"></param>
         * /// <param name="firstRow"></param>
         * /// <param name="lastRow"></param>
         * /// <param name="useMergedCells"></param>
         * public static void AutoSizeColumn(this ISheet sheet, int column, int firstRow, int lastRow, bool useMergedCells = false)
         * {
         *     if (sheet == null)
         *         throw new ArgumentNullException();
         *
         *     // 内部逻辑: 使用 IWorkbook.GetFontAt(0) 字体计算宽度...
         *     var width = NPOI.SS.Util.SheetUtil.GetColumnWidth(sheet, column, useMergedCells, firstRow, lastRow);
         *
         *     // 不存在有效行返回: -1.0;
         *     if (width > 0d)
         *     {
         *         width *= UnitsPerCharacter;
         *         var maximum = (double)MaximumColumnWidthInUnits;
         *
         *         if (width > maximum)
         *             width = maximum;
         *
         *         sheet.SetColumnWidth(column, (int)width);
         *     }
         * }
         */

        #endregion
    }
}