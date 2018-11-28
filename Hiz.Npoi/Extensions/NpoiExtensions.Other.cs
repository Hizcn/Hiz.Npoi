using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    partial class NpoiExtensions
    {
        /// <summary>
        /// 获取所有数据格式.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="builtin">是否包含内建格式</param>
        /// <returns></returns>
        static IEnumerable<ExcelDataFormat> GetFormats(this IWorkbook workbook, bool builtin = false)
        {
            if (workbook == null)
                throw new ArgumentNullException();

            var formats = GetNumberFormats(workbook);

            // 是否考虑取消 BuiltinFormats 输出; Office 软件只保存文件中已使用的格式 (替换内建+新增定制);
            var all = BuiltinFormats.GetAll();
            if (builtin)
            {
                for (var key = 0; key < all.Length; key++)
                    if (!formats.ContainsKey(key)) // 如果没有存在修改
                        yield return new ExcelDataFormat(key, all[key], ExcelDataFormatSource.Builtin);
            }

            /* 第一条由用户添加的格式的索引, 不同软件可能会有不同结果.
             * WPS = 176
             * Or HSSFWorkbook: NPOI.HSSF.UserModel.HSSFDataFormat.FIRST_USER_DEFINED_FORMAT_INDEX = 164;
             * Or XSSFWorkbook: NPOI.XSSF.Model.StylesTable.FIRST_CUSTOM_STYLE_ID = 165;
             */
            var first = all.Length;

            foreach (var i in formats)
            {
                var key = i.Key;
                if (key < first)
                {
                    var value = i.Value;
                    var original = all[key];
                    yield return new ExcelDataFormat(key, value, value != original ? ExcelDataFormatSource.BuiltinLocale : ExcelDataFormatSource.Builtin);
                }
                else
                {
                    // 
                    yield return new ExcelDataFormat(i.Key, i.Value, ExcelDataFormatSource.User);
                }
            }
        }
        internal static IDictionary<int, string> GetNumberFormats(IWorkbook workbook)
        {
            var xssf = workbook as XSSFWorkbook;
            if (xssf != null)
            {
                return xssf.GetStylesSource().GetNumberFormats();
            }

            var hssf = workbook as HSSFWorkbook;
            if (hssf != null)
            {
                return hssf.Workbook.Formats.ToDictionary(f => f.IndexCode, f => f.FormatString);
            }

            var sssf = workbook as SXSSFWorkbook;
            if (sssf != null)
            {
                return sssf.XssfWorkbook.GetStylesSource().GetNumberFormats();
            }
            throw new NotSupportedException();
        }

        #region 其它功能

        // NPOI.SS.Util.CellReference.ConvertNumToColString();
        /// <summary>
        /// 根据 列的索引 获取 列的代号 (例如: 0=A; 1=B; 26=AA;)
        /// </summary>
        /// <param name="index">[0, 0x4000)</param>
        /// <returns></returns>
        static string GetColumnCode(int index)
        {
            // Excel 2003 及以下版本的 最大列数: 0x0100; 最大行数: 0x010000;
            // Excel 2007 及以上版本的 最大列数: 0x4000; 最大行数: 0x100000;
            if (index < 0 || index >= 0x4000)
                throw new ArgumentOutOfRangeException();

            // 进制
            const int CharCount = 26;

            // |==========================================================================================|
            // | 一位字母:           [A-Z] | Count: 26^1 = 26;                  | (零位字母最大 + 1) * 26; |
            // | 两位字母:      [A-Z][A-Z] | Count: 26^1 + 26^2 = 702;          | (一位字母最大 + 1) * 26; |
            // | 三位字母: [A-Z][A-Z][A-Z] | Count: 26^1 + 26^2 + 26^3 = 18278; | (二位字母最大 + 1) * 26; |
            // |==========================================================================================|

            var length = 0; // 字母数量
            var maximum = 0; // 各个位数表示的最大值
            while (index >= maximum) // IndexBase = 0
            {
                maximum += (int)Math.Pow(CharCount, ++length);
            }

            var chars = new char[length];
            while (index >= 0)
            {
                chars[--length] = (char)(index % CharCount + 'A');

                index /= CharCount;
                index--;
            }

            return new string(chars);
        }

        #endregion

        #region 边框模板 NotOK

        //TODO: 优化 CellStyle 的创建数量;
        /// <summary>
        /// 批量应用边框模板
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="template">模板配置</param>
        /// <param name="x">起始列的索引; 从零开始;</param>
        /// <param name="y">起始行的索引; 从零开始;</param>
        /// <param name="width">列数; 至少一列;</param>
        /// <param name="height">行数; 至少一行;</param>
        public static void ApplyBorderTemplate(this ISheet sheet, BorderTemplate template, int x, int y, int width, int height)
        {
            if (sheet == null)
                throw new ArgumentNullException();
            if (template == null)
                throw new ArgumentNullException();

            if (x < 0 || y < 0)
                throw new ArgumentOutOfRangeException();
            if (width <= 0 || height <= 0)
                throw new ArgumentOutOfRangeException();

            #region 定义
            /* |==========|==========|==========|
             * | 1*1 id=1 | 2*1 id=2 | 3*1 id=3 |
             * |==========|==========|==========|
             * | 1*2 id=4 | 2*2 id=5 | 3*2 id=6 |
             * |==========|==========|==========|
             * | 1*3 id=7 | 2*3 id=8 | 3*3 id=9 |
             * |==========|==========|==========|
             * 
             * Width * Height
             * 
             * 1*1 = TopLeft
             * |==============|==============|==============|
             * | TopLeft      |              |              |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * 
             * 1*2 = TopLeft + BottomLeft
             * |==============|==============|==============|
             * | TopLeft      |              |              |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * | BottomLeft   |              |              |
             * |==============|==============|==============|
             * 
             * 1*3 = TopLeft + BottomLeft + MiddleLeft
             * |==============|==============|==============|
             * | TopLeft      |              |              |
             * |==============|==============|==============|
             * | MiddleLeft   |              |              |
             * |==============|==============|==============|
             * | BottomLeft   |              |              |
             * |==============|==============|==============|
             * 
             * 2*1 = TopLeft + TopRight
             * |==============|==============|==============|
             * | TopLeft      |              | TopRight     |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * 
             * 2*2 = (TopLeft + TopRight) + (BottomLeft + BottomRight)
             * |==============|==============|==============|
             * | TopLeft      |              | TopRight     |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * | BottomLeft   |              | BottomRight  |
             * |==============|==============|==============|
             * 
             * 2*3 = (TopLeft + TopRight) + (BottomLeft + BottomRight) + (MiddleLeft + MiddleRight)
             * |==============|==============|==============|
             * | TopLeft      |              | TopRight     |
             * |==============|==============|==============|
             * | MiddleLeft   |              | MiddleRight  |
             * |==============|==============|==============|
             * | BottomLeft   |              | BottomRight  |
             * |==============|==============|==============|
             * 
             * 3*1 = TopLeft + TopRight + TopCenter
             * |==============|==============|==============|
             * | TopLeft      | TopCenter    | TopRight     |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * 
             * 3*2 = (TopLeft + TopRight) + (BottomLeft + BottomRight) + (TopCenter + BottomCenter)
             * |==============|==============|==============|
             * | TopLeft      | TopCenter    | TopRight     |
             * |==============|==============|==============|
             * |              |              |              |
             * |==============|==============|==============|
             * | BottomLeft   | BottomCenter | BottomRight  |
             * |==============|==============|==============|
             * 
             * 3*3 = All
             * |==============|==============|==============|
             * | TopLeft      | TopCenter    | TopRight     |
             * |==============|==============|==============|
             * | MiddleLeft   | MiddleCenter | MiddleRight  |
             * |==============|==============|==============|
             * | BottomLeft   | BottomCenter | BottomRight  |
             * |==============|==============|==============|
             */

            // 第一行 第一列
            MockCellBorder bTopLeft = new MockCellBorder();
            // 第一行 中间所有列
            MockCellBorder bTopCenter = null;
            // 第一行 最末列
            MockCellBorder bTopRight = null;

            // 中间所有行 第一列
            MockCellBorder bMiddleLeft = null;
            // 中间所有行 中间所有列
            MockCellBorder bMiddleCenter = null;
            // 中间所有行 最末列
            MockCellBorder bMiddleRight = null;

            // 最末行 第一列
            MockCellBorder bBottomLeft = null;
            // 最末行 中间所有列
            MockCellBorder bBottomCenter = null;
            // 最末行 最末列
            MockCellBorder bBottomRight = null;

            if (width >= 2)
            {
                bTopRight = new MockCellBorder();
            }
            if (height >= 2)
            {
                bBottomLeft = new MockCellBorder();
            }
            if (width >= 2 && height >= 2)
            {
                bBottomRight = new MockCellBorder();
            }
            if (width >= 3)
            {
                bTopCenter = new MockCellBorder();
                if (height >= 2)
                    bBottomCenter = new MockCellBorder();
            }
            if (height >= 3)
            {
                bMiddleLeft = new MockCellBorder();
                if (width >= 2)
                    bMiddleRight = new MockCellBorder();
            }
            if (width >= 3 && height >= 3)
            {
                bMiddleCenter = new MockCellBorder();
            }
            #endregion

            #region 垂直线条
            if (template.TopStyle != BorderStyle.None)
            {
                var style = template.TopStyle;
                var color = template.TopColor ?? NpoiColor.Black;

                // 1 Top
                bTopLeft.TopStyle = style;
                bTopLeft.TopColor = color;
                if (width >= 2)
                {
                    // 3 Top
                    bTopRight.TopStyle = style;
                    bTopRight.TopColor = color;
                }
                if (width >= 3)
                {
                    // 2 Top
                    bTopCenter.TopStyle = style;
                    bTopCenter.TopColor = color;
                }
            }
            if (template.BottomStyle != BorderStyle.None)
            {
                var style = template.BottomStyle;
                var color = template.BottomColor ?? NpoiColor.Black;

                if (height >= 2)
                {
                    // 7 Bottom
                    bBottomLeft.BottomStyle = style;
                    bBottomLeft.BottomColor = color;
                    if (width >= 2)
                    {
                        // 9 Bottom
                        bBottomRight.BottomStyle = style;
                        bBottomRight.BottomColor = color;
                    }
                    if (width >= 3)
                    {
                        // 8 Bottom
                        bBottomCenter.BottomStyle = style;
                        bBottomCenter.BottomColor = color;
                    }
                }
                else // 只有一行
                {
                    // 1 Bottom
                    bTopLeft.BottomStyle = style;
                    bTopLeft.BottomColor = color;
                    if (width >= 2)
                    {
                        // 3 Bottom
                        bTopRight.BottomStyle = style;
                        bTopRight.BottomColor = color;
                    }
                    if (width >= 3)
                    {
                        // 2 Bottom
                        bTopCenter.BottomStyle = style;
                        bTopCenter.BottomColor = color;
                    }
                }
            }
            if (template.MiddleStyle != BorderStyle.None)
            {
                var style = template.MiddleStyle;
                var color = template.MiddleColor ?? NpoiColor.Black;

                if (height >= 2)
                {
                    // 1 Bottom
                    bTopLeft.BottomStyle = style;
                    bTopLeft.BottomColor = color;
                    // 7 Top
                    bBottomLeft.TopStyle = style;
                    bBottomLeft.TopColor = color;
                    if (width >= 2)
                    {
                        // 3 Bottom
                        bTopRight.BottomStyle = style;
                        bTopRight.BottomColor = color;
                        // 9 Top
                        bBottomRight.TopStyle = style;
                        bBottomRight.TopColor = color;
                    }
                    if (width >= 3)
                    {
                        // 2 Bottom
                        bTopCenter.BottomStyle = style;
                        bTopCenter.BottomColor = color;
                        // 8 Top
                        bBottomCenter.TopStyle = style;
                        bBottomCenter.TopColor = color;
                    }
                }
                if (height >= 3)
                {
                    // 4 Top+Bottom
                    bMiddleLeft.TopStyle = style;
                    bMiddleLeft.TopColor = color;
                    bMiddleLeft.BottomStyle = style;
                    bMiddleLeft.BottomColor = color;
                    if (width >= 2)
                    {
                        // 6 Top+Bottom
                        bMiddleRight.TopStyle = style;
                        bMiddleRight.TopColor = color;
                        bMiddleRight.BottomStyle = style;
                        bMiddleRight.BottomColor = color;
                    }
                    if (width >= 3)
                    {
                        // 5 Top+Bottom
                        bMiddleCenter.TopStyle = style;
                        bMiddleCenter.TopColor = color;
                        bMiddleCenter.BottomStyle = style;
                        bMiddleCenter.BottomColor = color;
                    }

                }
            }
            #endregion

            #region 水平线条
            if (template.LeftStyle != BorderStyle.None)
            {
                var style = template.LeftStyle;
                var color = template.LeftColor ?? NpoiColor.Black;

                // 1 Left
                bTopLeft.LeftStyle = style;
                bTopLeft.LeftColor = color;
                if (height >= 2)
                {
                    // 7 Left
                    bBottomLeft.LeftStyle = style;
                    bBottomLeft.LeftColor = color;
                }
                if (height >= 3)
                {
                    // 4 Left
                    bMiddleLeft.LeftStyle = style;
                    bMiddleLeft.LeftColor = color;
                }

            }
            if (template.RightStyle != BorderStyle.None)
            {
                var style = template.RightStyle;
                var color = template.RightColor ?? NpoiColor.Black;

                if (width >= 2)
                {
                    // 3 Right
                    bTopRight.RightStyle = style;
                    bTopRight.RightColor = color;
                    if (height >= 2)
                    {
                        // 9 Right
                        bBottomRight.RightStyle = style;
                        bBottomRight.RightColor = color;
                    }
                    if (height >= 3)
                    {
                        // 6 Right
                        bMiddleRight.RightStyle = style;
                        bMiddleRight.RightColor = color;
                    }

                }
                else // 只有一列
                {
                    // 1 Right
                    bTopLeft.RightStyle = style;
                    bTopLeft.RightColor = color;
                    if (height >= 2)
                    {
                        // 7 Right
                        bBottomLeft.RightStyle = style;
                        bBottomLeft.RightColor = color;
                    }
                    if (height >= 3)
                    {
                        // 4 Right
                        bMiddleLeft.RightStyle = style;
                        bMiddleLeft.RightColor = color;
                    }
                }
            }
            if (template.CenterStyle != BorderStyle.None)
            {
                var style = template.CenterStyle;
                var color = template.CenterColor ?? NpoiColor.Black;
                if (width > 2)
                {
                    // 1 Right
                    bTopLeft.RightStyle = style;
                    bTopLeft.RightColor = color;
                    // 3 Left
                    bTopRight.LeftStyle = style;
                    bTopRight.LeftColor = color;
                    if (height >= 2)
                    {
                        // 7 Right
                        bBottomLeft.RightStyle = style;
                        bBottomLeft.RightColor = color;
                        // 9 Left
                        bBottomRight.LeftStyle = style;
                        bBottomRight.LeftColor = color;
                    }
                    if (height >= 3)
                    {
                        // 4 Right
                        bMiddleLeft.RightStyle = style;
                        bMiddleLeft.RightColor = color;
                        // 6 Left
                        bMiddleRight.LeftStyle = style;
                        bMiddleRight.LeftColor = color;
                    }
                }
                if (width >= 3)
                {
                    // 2 Left+Right
                    bTopCenter.LeftStyle = style;
                    bTopCenter.LeftColor = color;
                    bTopCenter.RightStyle = style;
                    bTopCenter.RightColor = color;
                    if (height >= 2)
                    {
                        // 8 Left+Right
                        bBottomCenter.LeftStyle = style;
                        bBottomCenter.LeftColor = color;
                        bBottomCenter.RightStyle = style;
                        bBottomCenter.RightColor = color;
                    }
                    if (height >= 3)
                    {
                        // 5 Left+Right
                        bMiddleCenter.LeftStyle = style;
                        bMiddleCenter.LeftColor = color;
                        bMiddleCenter.RightStyle = style;
                        bMiddleCenter.RightColor = color;
                    }
                }
            }
            #endregion

            // 用于遍历 (此处赋值 需要等待 所有条目 实例化后)
            var styles = new MockCellBorder[] { bTopLeft, bTopCenter, bTopRight, bMiddleLeft, bMiddleCenter, bMiddleRight, bBottomLeft, bBottomCenter, bBottomRight };

            #region 对角线条
            if (template.Diagonal != BorderDiagonal.None && template.DiagonalStyle != BorderStyle.None)
            {
                var diagonal = template.Diagonal;
                var style = template.DiagonalStyle;
                var color = template.DiagonalColor ?? NpoiColor.Black;
                foreach (var b in styles)
                {
                    if (b != null)
                    {
                        b.Diagonal = diagonal;
                        b.DiagonalStyle = style;
                        b.DiagonalColor = color;
                    }
                }
            }
            #endregion

            #region 应用样式

            var workbook = sheet.Workbook;

            // var sssf = workbook as SXSSFWorkbook;
            // if (sssf != null)
            // {
            //     workbook = sssf.XssfWorkbook;
            // }
            // var xssf = workbook as XSSFWorkbook;
            // if (xssf != null)
            // {
            // }

            var hssf = workbook as HSSFWorkbook;
            if (hssf != null)
            {
                var palette = hssf.GetCustomPalette();
                var caching = new Dictionary<int, short>();
                foreach (var b in styles)
                {
                    if (b != null)
                    {
                        if (b.LeftStyle != BorderStyle.None)
                        {
                            b.LeftColorIndexed = FindColorIndexedWithPalette(palette, b.LeftColor, caching);
                        }
                        if (b.RightStyle != BorderStyle.None)
                        {
                            b.RightColorIndexed = FindColorIndexedWithPalette(palette, b.RightColor, caching);
                        }
                        if (b.TopStyle != BorderStyle.None)
                        {
                            b.TopColorIndexed = FindColorIndexedWithPalette(palette, b.TopColor, caching);
                        }
                        if (b.BottomStyle != BorderStyle.None)
                        {
                            b.BottomColorIndexed = FindColorIndexedWithPalette(palette, b.BottomColor, caching);
                        }
                        if (b.Diagonal != BorderDiagonal.None && b.DiagonalStyle != BorderStyle.None)
                        {
                            b.DiagonalColorIndexed = FindColorIndexedWithPalette(palette, b.DiagonalColor, caching);
                        }
                    }
                }
            }

            var iLastRow = y + height - 1;
            var iLastColumn = x + width - 1;
            IDictionary<int, ICellStyle> makes = new Dictionary<int, ICellStyle>();

            for (var r = y; r <= iLastRow; r++)
            {
                var row = sheet.GetOrAddRow(r);

                if (r == y) // 第一行
                {
                    for (var c = x; c <= iLastColumn; c++)
                    {
                        var cell = row.GetOrAddCell(c);

                        if (c == x) // 第一列
                        {
                            if (bTopLeft == null)
                                throw new NotImplementedException();
                            if (!NpoiCompare.EqualsBorder(cell.CellStyle, bTopLeft))
                            {
                                var style = workbook.CreateCellStyle();
                                style.CloneStyleFrom(cell.CellStyle);
                                SetBorder(style, bTopLeft);
                                cell.CellStyle = style;
                            }
                        }
                        else if (c == iLastColumn) // 最末列
                        {
                            if (bTopRight == null)
                                throw new NotImplementedException();
                            if (!NpoiCompare.EqualsBorder(cell.CellStyle, bTopRight))
                            {
                                var style = workbook.CreateCellStyle();
                                style.CloneStyleFrom(cell.CellStyle);
                                SetBorder(style, bTopRight);
                                cell.CellStyle = style;
                            }
                        }
                        else // 中间列
                        {
                            if (bTopCenter == null)
                                throw new NotImplementedException();

                            var key = (0x02 << 0x10) | (ushort)cell.CellStyle.Index;
                            if (!makes.TryGetValue(key, out ICellStyle value))
                            {
                                if (!NpoiCompare.EqualsBorder(cell.CellStyle, bTopCenter))
                                {
                                    var style = workbook.CreateCellStyle();
                                    style.CloneStyleFrom(cell.CellStyle);
                                    SetBorder(style, bTopCenter);
                                    value = style;
                                }
                                else
                                {
                                    value = cell.CellStyle;
                                }
                                makes.Add(key, value);
                            }
                            cell.CellStyle = value;
                        }
                    }
                }
                else if (r == iLastRow) // 最末行 // 至少两行才会执行
                {
                    for (var c = x; c <= iLastColumn; c++)
                    {
                        var cell = row.GetOrAddCell(c);

                        if (c == x) // 第一列
                        {
                            if (bBottomLeft == null)
                                throw new NotImplementedException();
                            if (!NpoiCompare.EqualsBorder(cell.CellStyle, bBottomLeft))
                            {
                                var style = workbook.CreateCellStyle();
                                style.CloneStyleFrom(cell.CellStyle);
                                SetBorder(style, bBottomLeft);
                                cell.CellStyle = style;
                            }
                        }
                        else if (c == iLastColumn) // 最末列
                        {
                            if (bBottomRight == null)
                                throw new NotImplementedException();
                            if (!NpoiCompare.EqualsBorder(cell.CellStyle, bBottomRight))
                            {
                                var style = workbook.CreateCellStyle();
                                style.CloneStyleFrom(cell.CellStyle);
                                SetBorder(style, bBottomRight);
                                cell.CellStyle = style;
                            }
                        }
                        else // 中间列
                        {
                            if (bBottomCenter == null)
                                throw new NotImplementedException();

                            var key = (0x08 << 0x10) | (ushort)cell.CellStyle.Index;
                            if (!makes.TryGetValue(key, out ICellStyle value))
                            {
                                if (!NpoiCompare.EqualsBorder(cell.CellStyle, bBottomCenter))
                                {
                                    var style = workbook.CreateCellStyle();
                                    style.CloneStyleFrom(cell.CellStyle);
                                    SetBorder(style, bBottomCenter);
                                    value = style;
                                }
                                else
                                {
                                    value = cell.CellStyle;
                                }
                                makes.Add(key, value);
                            }
                            cell.CellStyle = value;
                        }
                    }
                }
                else // 中间行 // 至少三行才会执行
                {
                    for (var c = x; c <= iLastColumn; c++)
                    {
                        var cell = row.GetOrAddCell(c);

                        if (c == x) // 第一列
                        {
                            if (bMiddleLeft == null)
                                throw new NotImplementedException();

                            var key = (0x04 << 0x10) | (ushort)cell.CellStyle.Index;
                            if (!makes.TryGetValue(key, out ICellStyle value))
                            {
                                if (!NpoiCompare.EqualsBorder(cell.CellStyle, bMiddleLeft))
                                {
                                    var style = workbook.CreateCellStyle();
                                    style.CloneStyleFrom(cell.CellStyle);
                                    SetBorder(style, bMiddleLeft);
                                    value = style;
                                }
                                else
                                {
                                    value = cell.CellStyle;
                                }
                                makes.Add(key, value);
                            }
                            cell.CellStyle = value;
                        }
                        else if (c == iLastColumn) // 最末列
                        {
                            if (bMiddleRight == null)
                                throw new NotImplementedException();

                            var key = (0x06 << 0x10) | (ushort)cell.CellStyle.Index;
                            if (!makes.TryGetValue(key, out ICellStyle value))
                            {
                                if (!NpoiCompare.EqualsBorder(cell.CellStyle, bMiddleRight))
                                {
                                    var style = workbook.CreateCellStyle();
                                    style.CloneStyleFrom(cell.CellStyle);
                                    SetBorder(style, bMiddleRight);
                                    value = style;
                                }
                                else
                                {
                                    value = cell.CellStyle;
                                }
                                makes.Add(key, value);
                            }
                            cell.CellStyle = value;
                        }
                        else // 中间列
                        {
                            if (bMiddleCenter == null)
                                throw new NotImplementedException();

                            var key = (0x05 << 0x10) | (ushort)cell.CellStyle.Index;
                            if (!makes.TryGetValue(key, out ICellStyle value))
                            {
                                if (!NpoiCompare.EqualsBorder(cell.CellStyle, bMiddleCenter))
                                {
                                    var style = workbook.CreateCellStyle();
                                    style.CloneStyleFrom(cell.CellStyle);
                                    SetBorder(style, bMiddleCenter);
                                    value = style;
                                }
                                else
                                {
                                    value = cell.CellStyle;
                                }
                                makes.Add(key, value);
                            }
                            cell.CellStyle = value;
                        }
                    }
                }
            }
            #endregion
        }

        static void SetBorder(ICellStyle style, MockCellBorder border)
        {
            var xssf = style as XSSFCellStyle;
            if (xssf != null)
            {
                Func<NpoiColor, XSSFColor> GetColor = (color) => (color == null || color.IsBlack) ? (XSSFColor)null : (color.Indexed > 0 ? new XSSFColor() { Indexed = color.Indexed } : new XSSFColor(color.Argb));

                xssf.BorderLeft = border.LeftStyle;
                xssf.SetLeftBorderColor(GetColor(border.LeftColor));

                xssf.BorderRight = border.RightStyle;
                xssf.SetRightBorderColor(GetColor(border.RightColor));

                xssf.BorderTop = border.TopStyle;
                xssf.SetTopBorderColor(GetColor(border.TopColor));

                xssf.BorderBottom = border.BottomStyle;
                xssf.SetBottomBorderColor(GetColor(border.BottomColor));

                xssf.BorderDiagonal = border.Diagonal;
                xssf.BorderDiagonalLineStyle = border.DiagonalStyle;
                xssf.SetDiagonalBorderColor(GetColor(border.DiagonalColor));
            }
            var hssf = style as HSSFCellStyle;
            if (hssf != null)
            {
                hssf.BorderLeft = border.LeftStyle;
                if (border.LeftColorIndexed > 0)
                    hssf.LeftBorderColor = border.LeftColorIndexed;

                hssf.BorderRight = border.RightStyle;
                if (border.RightColorIndexed > 0)
                    hssf.RightBorderColor = border.RightColorIndexed;

                hssf.BorderTop = border.TopStyle;
                if (border.TopColorIndexed > 0)
                    hssf.TopBorderColor = border.TopColorIndexed;

                hssf.BorderBottom = border.BottomStyle;
                if (border.BottomColorIndexed > 0)
                    hssf.BottomBorderColor = border.BottomColorIndexed;

                hssf.BorderDiagonal = border.Diagonal;
                hssf.BorderDiagonalLineStyle = border.DiagonalStyle;
                if (border.DiagonalColorIndexed > 0)
                    hssf.BorderDiagonalColor = border.DiagonalColorIndexed;
            }
        }

        #endregion
    }
}
