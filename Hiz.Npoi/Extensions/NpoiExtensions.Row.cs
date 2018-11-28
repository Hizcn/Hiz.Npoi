using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Hiz.Extended.Npoi
{
    partial class NpoiExtensions
    {
        #region IRow~

        /* IRow : IEnumerable<ICell>
         * 
         * // Get defined cells in the row. // Cells.Count == PhysicalNumberOfCells;
         * List<ICell> Cells { get; }
         * 
         * // Gets the number of defined cells (NOT number of cells in the actual row!).
         * // That is to say if only columns 0,4,5 have values then there would be 3.
         * int PhysicalNumberOfCells { get; }
         * 
         * // Get the number (0 based) of the first cell Contained in this row.
         * // -1 if the row does not contain any cells.
         * short FirstCellNum { get; }
         * 
         * // Gets the index of the last cell Contained in this row PLUS ONE.
         * // The result also happens to be the 1-based column number of the last cell.
         * // This value can be used as a standard upper bound when iterating over cells:
         * // short minColIx = row.GetFirstCellNum();
         * // short maxColIx = row.GetLastCellNum();
         * // for(short colIx = minColIx; colIx < maxColIx; colIx++) {
         * //     Cell cell = row.GetCell(colIx);
         * //     if (cell == null) {
         * //         continue;
         * //     }
         * //     //... do something with cell
         * // }
         * // -1 if the row does not contain any cells.
         * short LastCellNum { get; }
         * 
         * // Get the row's height measured in twips (1/20th of a point).
         * short Height { get; set; }
         * // Returns row height measured in point size.
         * // If the height is not set, the default worksheet value is returned, NPOI.SS.UserModel.ISheet.DefaultRowHeightInPoints.
         * float HeightInPoints { get; set; }
         * 
         * // Is this row formatted? Most aren't, but some rows do have whole-row styles.
         * // For those that do, you can get the formatting from NPOI.SS.UserModel.IRow.RowStyle.
         * bool IsFormatted { get; }
         * 
         * // Returns the whole-row cell styles.
         * // Most rows won't have one of these, so will return null.
         * // Call IsFormmated to check first.
         * ICellStyle RowStyle { get; set; }
         * 
         * // Returns the rows outline level.
         * // Increased as you put it into more groups (outlines), reduced as you take it out of them.
         * int OutlineLevel { get; }
         * 
         * // Returns the Sheet this row belongs to.
         * ISheet Sheet { get; }
         * 
         * // Get row number (0 based) this row represents.
         * int RowNum { get; set; }
         * 
         * // Get whether or not to display this row with 0 height.
         * bool ZeroHeight { get; set; }
         * 
         * 
         * 
         * // Copy the source cell to the target cell.
         * // If the target cell exists, the new copied cell will be inserted before the existing one.
         * ICell CopyCell(int sourceIndex, int targetIndex);
         * 
         * // Copy the current row to the target row.
         * IRow CopyRowTo(int targetIndex);
         * 
         * // Use this to create new cells within the row and return it.
         * // The cell that is returned is a NPOI.SS.UserModel.ICell/NPOI.SS.UserModel.CellType.Blank.
         * // The type can be changed either through calling SetCellValue or SetCellType.
         * ICell CreateCell(int column, CellType type);
         * |=> ICell CreateCell(int column); // type = CellType.Blank;
         * 
         * // Returns the cell at the given (0 based) index, with the specified MissingCellPolicy.
         * ICell GetCell(int cellnum, MissingCellPolicy policy);
         * |=> ICell GetCell(int cellnum); // policy = this.Sheet.Workbook.MissingCellPolicy;
         * 
         * // Moves the supplied cell to a new column, which must not already have a cell there!
         * void MoveCell(ICell cell, int newColumn);
         * 
         * // Remove the Cell from this row.
         * void RemoveCell(ICell cell);
         */

        public static ICell GetOrAddCell(this IRow row, int index)
        {
            /* MissingCellPolicy:
             * 
             * // Missing cells are returned as null, Blank cells are returned as normal.
             * // Default. 默认: 原样返回.
             * RETURN_NULL_AND_BLANK
             * 
             * // Missing cells are returned as null, as are blank cells.
             * // 如果已创建单元格, 但是单元格的数据类型为空, 则会返回空值. 
             * RETURN_BLANK_AS_NULL
             * 
             * // A new, blank cell is Created for missing cells. Blank cells are returned as normal.
             * // 如果未创建单元格, 则将新建数据类型为空的单元格, 然后返回.
             * CREATE_NULL_AS_BLANK
             */
            // var cell = row.GetCell(index, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            // if (cell == null)
            //     cell = row.CreateCell(index, CellType.Blank);
            // return cell;
            return row.GetCell(index, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        /// <summary>
        /// 遍历行单元格; 推荐改用 foreach(var cell in row) {...}
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static IEnumerable<ICell> GetCells(this IRow row)
        {
            /* NPOIv2.3 代码:
             * 
             * HSSFRow : IRow, IEnumerable<ICell>, IEnumerable, IComparable {
             *   .Cells {
             *       get {
             *           return new System.Collections.Generic.List<ICell>(this.cells.Values);
             *       }
             *   } // cells: SortedDictionary<int, ICell>
             *   .FirstCellNum {
             *       get {
             *           if (this.row.IsEmpty)
             *               return -1;
             *           return (short)this.row.FirstCol;
             *       }
             *   }
             *   .LastCellNum {
             *       get {
             *           if (this.row.IsEmpty)
             *               return -1;
             *           return (short)this.row.LastCol;
             *       }
             *   }
             *   
             *   .GetCell(int cellnum, MissingCellPolicy policy); // 调用 RetrieveCell(int cellnum);
             *   .RetrieveCell(int cellnum) {
             *       if (this._cells.ContainsKey(cellnum)) // 字典查询
             *           return this._cells[cellnum];
             *       return null;
             *   }
             *   
             *   GetEnumerator() {
             *       return this._cells.Values.GetEnumerator();
             *   }
             * }
             * 
             * XSSFRow : IRow, IEnumerable<ICell>, IEnumerable, IComparable<XSSFRow> {
             *   .Cells {
             *       get {
             *           var list = new List<ICell>();
             *           foreach (ICell current in this._cells.Values)
             *               list.Add(current);
             *           return list;
             *       }
             *   } // _cells: SortedDictionary<int, ICell>
             *   .FirstCellNum {
             *       get {
             *           this._cells.Count == 0 ? -1 : this.GetFirstKey(this._cells.Keys);
             *       }
             *   }
             *   .GetFirstKey(SortedDictionary<int, ICell>.KeyCollection keys) {
             *       int i = 0;
             *       foreach (int current in keys) {
             *           if (i == 0)
             *               return current;
             *       }
             *       throw new ArgumentOutOfRangeException();
             *   } // 逻辑实现有点冗余. (Keys 取第一个)
             *   .LastCellNum {
             *       get {
             *           this._cells.Count == 0 ? -1 : (this.GetLastKey(this._cells.Keys) + 1);
             *       }
             *   }
             *   .GetLastKey(SortedDictionary<int, ICell>.KeyCollection keys) {
             *       int i = 0;
             *       foreach (int current in keys) {
             *           if (i == keys.Count - 1)
             *               return current;
             *           i++;
             *       }
             *       throw new ArgumentOutOfRangeException();
             *   } // (Keys 取最末个)
             *   
             *   .GetCell(int cellnum, MissingCellPolicy policy); // 调用 RetrieveCell(int cellnum);
             *   .RetrieveCell(int cellnum) {
             *       if (this._cells.ContainsKey(cellnum)) // 字典查询
             *           return this._cells[cellnum];
             *       return null;
             *   }
             *   
             *   .GetEnumerator() {
             *       return this.CellIterator();
             *   }
             *   .CellIterator() {
             *       return this._cells.Values.GetEnumerator();
             *   }
             * }
             * 
             * // 综合上述, 使用 IEnumerable<ICell> 接口遍历 Cells 性能最佳.
             */
            if (row == null)
                throw new ArgumentNullException();

            foreach (var cell in row)
                yield return cell;
        }

        /// <summary>
        /// 是否空行 (至少有一个含有有效值的单元格, 才不算作空行; 对于字符串单元格, 如果文本只有空格, 也当无效处理;
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static bool IsEmpty(this IRow row)
        {
            if (row == null)
                throw new ArgumentNullException();

            if (row.PhysicalNumberOfCells > 0)
                foreach (var cell in row) // cell 不会等于 null;
                {
                    var type = cell.CellType;
                    if (type != CellType.Blank)
                    {
                        if (type != CellType.String)
                            return false;
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text)) // 如果文本只有空格, 当作无效;
                            return false;
                    }
                }
            return true;
        }

        /// <summary>
        /// 清空本行的所有单元格.
        /// </summary>
        /// <param name="row"></param>
        public static void RemoveAllCells(this IRow row)
        {
            if (row == null)
                throw new ArgumentNullException();
            // var hssf = row as HSSFRow;
            // if (hssf != null)
            //     hssf.RemoveAllCells();
            var array = row.ToList();
            foreach (var cell in array)
                row.RemoveCell(cell);
        }

        /// <summary>
        /// 清空本行的所有单元格的值.
        /// </summary>
        /// <param name="row"></param>
        public static void RemoveAllCellsValues(this IRow row)
        {
            if (row == null)
                throw new ArgumentNullException();

            foreach (var cell in row)
                cell.SetCellType(CellType.Blank);
        }

        public static float GetRowHeight(this IRow row, LengthUnit unit)
        {
            if (row == null)
                throw new ArgumentNullException(nameof(row));
            switch (unit)
            {
                case LengthUnit.Point:
                    return row.HeightInPoints;
                case LengthUnit.Inche:
                    return row.HeightInPoints / PointsPerInch;
                case LengthUnit.Centimeter:
                    return row.HeightInPoints * CentimetersPerInch / PointsPerInch;
                case LengthUnit.Millimeter:
                    return row.HeightInPoints * MillimetersPerInch / PointsPerInch;
                case LengthUnit.Character:
                    throw new NotSupportedException();
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
        public static void SetRowHeight(this IRow row, float height, LengthUnit unit)
        {
            if (row == null)
                throw new ArgumentNullException(nameof(row));
            switch (unit)
            {
                case LengthUnit.Point:
                    row.HeightInPoints = height;
                    break;
                case LengthUnit.Inche:
                    row.HeightInPoints = height * PointsPerInch;
                    break;
                case LengthUnit.Centimeter:
                    row.HeightInPoints = height * PointsPerInch / CentimetersPerInch;
                    break;
                case LengthUnit.Millimeter:
                    row.HeightInPoints = height * PointsPerInch / MillimetersPerInch;
                    break;
                case LengthUnit.Character:
                    throw new NotSupportedException();
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        #endregion

        public static T GetCellValue<T>(this IRow row, int cellIndex, T @default = default(T))
        {
            var cell = row.GetCell(cellIndex, MissingCellPolicy.RETURN_NULL_AND_BLANK);
            if (cell == null)
                return @default;
            return cell.GetCellValue<T>(@default);
        }
    }
}
