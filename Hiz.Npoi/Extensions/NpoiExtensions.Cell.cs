using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Globalization;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Runtime.Serialization;

namespace Hiz.Npoi
{
    partial class NpoiExtensions
    {
        /// <summary>
        /// 获取单元格值最终类型, 如果含有公式, 则返回计算结果的类型.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static CellType GetCellTypeFinally(this ICell cell)
        {
            if (cell == null)
                throw new ArgumentNullException();

            var type = cell.CellType;
            if (type == CellType.Formula)
                type = cell.CachedFormulaResultType;
            return type;
        }

        static NpoiValueConvert _Convert = new NpoiValueConvert();

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="cell"></param>
        /// <param name="default"></param>
        /// <returns></returns>
        public static T GetCellValue<T>(this ICell cell, T @default = default(T))
        {
            return (T)_Convert.GetCellValue(cell, typeof(T), @default);
        }

        public static string GetCellValueAsString(this ICell cell, string @default = null, bool trim = true)
        {
            return _Convert.GetCellValueAsString(cell, @default, trim);
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void SetCellValue(this ICell cell, object value)
        {
            if (cell == null)
                throw new ArgumentNullException();

            _Convert.SetCellValue(cell, value);
        }

        const int EmusPerCentimeter = 360000; // EmusPerInch / 2.54;
        const int EmusPerInch = 914400;
        const int EmusPerPoint = 12700; // EmusPerInch / 72;
        const int EmusPerPixel = 9525; // EmusPerInch / 96;

        /// <summary>
        /// 设置单元格的 注释内容;
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="string"></param>
        /// <returns></returns>
        public static IComment SetCellComment(this ICell cell, string @string)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            var comment = cell.CellComment;
            if (comment == null)
            {
                var drawing = cell.Sheet.GetOrAddDrawingPatriarch();
                /* https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Drawing.html
                 * dx1 - the x coordinate in EMU within the first cell.
                 * dy1 - the y coordinate in EMU within the first cell.
                 * dx2 - the x coordinate in EMU within the second cell.
                 * dy2 - the y coordinate in EMU within the second cell.
                 * col1 - the column (0 based) of the first cell.
                 * row1 - the row (0 based) of the first cell.
                 * col2 - the column (0 based) of the second cell.
                 * row2 - the row (0 based) of the second cell.
                 * 
                 * EMU: English Metric Unit
                 * https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
                 */
                var anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 0, 0);
                comment = drawing.CreateCellComment(anchor);
                cell.CellComment = comment;
            }
            if (comment is XSSFComment)
                comment.String = new XSSFRichTextString(@string);
            if (comment is HSSFComment)
                comment.String = new HSSFRichTextString(@string);
            return comment;
        }
    }
}
