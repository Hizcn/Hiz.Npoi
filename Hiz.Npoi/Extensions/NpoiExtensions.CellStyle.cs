using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace Hiz.Extended.Npoi
{
    partial class NpoiExtensions
    {
        #region IWorkbook/ICellStyle

        /* IWorkbook 样式相关 (三个成员):
         * short NumCellStyles { get; }
         * ICellStyle GetCellStyleAt(short idx);
         * ICellStyle CreateCellStyle();
         * 
         * ICellStyle
         * |=> NPOI.HSSF.UserModel.HSSFCellStyle
         * |=> NPOI.XSSF.UserModel.XSSFCellStyle
         */
        /* XSSFCellStyle 属性的默认值
         * 
         * short Index { get; }
         * 
         * short FontIndex { get; } = 0;
         * 
         * short DataFormat { get; set; } = 0;
         * 
         * HorizontalAlignment Alignment { get; set; } = General;
         * short Indention { get; set; } = 0;
         * VerticalAlignment VerticalAlignment { get; set; } = Bottom;
         * bool WrapText { get; set; } = false;
         * bool ShrinkToFit { get; set; } = false;
         * short Rotation { get; set; } = 0;
         * 
         * bool IsLocked { get; set; } = true;
         * bool IsHidden { get; set; } = false;
         * 
         * FillPattern FillPattern { get; set; } = NoFill;
         * short FillForegroundColor { get; set; } = 64;
         * IColor FillForegroundColorColor { get; set; } = null; // 接口没有 set 方法
         * short FillBackgroundColor { get; set; } = 64;
         * IColor FillBackgroundColorColor { get; set; } = null; // 接口没有 set 方法
         * 
         * BorderStyle BorderTop { get; set; } = None;
         * short TopBorderColor { get; set; } = 8;
         * BorderStyle BorderBottom { get; set; } = None;
         * short BottomBorderColor { get; set; } = 8;
         * BorderStyle BorderLeft { get; set; } = None;
         * short LeftBorderColor { get; set; } = 8;
         * BorderStyle BorderRight { get; set; } = None;
         * short RightBorderColor { get; set; } = 8;
         * BorderDiagonal BorderDiagonal { get; set; } = None;
         * BorderStyle BorderDiagonalLineStyle { get; set; } = None;
         * short BorderDiagonalColor { get; set; } = 8;
         * 
         * // XSSFCellStyle 独有属性:
         * 
         * XSSFColor FillForegroundXSSFColor { get; set; } = null;
         * XSSFColor FillBackgroundXSSFColor { get; set; } = null;
         * 
         * XSSFColor TopBorderXSSFColor { get; } = null;
         * XSSFColor BottomBorderXSSFColor { get; } = null;
         * XSSFColor LeftBorderXSSFColor { get; } = null;
         * XSSFColor RightBorderXSSFColor { get; } = null;
         * XSSFColor DiagonalBorderXSSFColor { get; } = null;
         */
        /* HSSFCellStyle 属性的默认值
         * 
         * short Index { get; }
         * 
         * short FontIndex { get; } = 0;
         * 
         * short DataFormat { get; set; } = 0;
         * 
         * HorizontalAlignment Alignment { get; set; } = General;
         * short Indention { get; set; } = 0;
         * VerticalAlignment VerticalAlignment { get; set; } = Bottom;
         * bool WrapText { get; set; } = false;
         * bool ShrinkToFit { get; set; } = false;
         * short Rotation { get; set; } = 0;
         * 
         * bool IsLocked { get; set; } = true;
         * bool IsHidden { get; set; } = false;
         * 
         * FillPattern FillPattern { get; set; } = NoFill;
         * short FillForegroundColor { get; set; } = 64;
         * IColor FillForegroundColorColor { get; } = HSSFColor.Automatic;
         * short FillBackgroundColor { get; set; } = 64;
         * IColor FillBackgroundColorColor { get; } = HSSFColor.Automatic;
         * 
         * BorderStyle BorderTop { get; set; } = None;
         * short TopBorderColor { get; set; } = 8;
         * BorderStyle BorderBottom { get; set; } = None;
         * short BottomBorderColor { get; set; } = 8;
         * BorderStyle BorderLeft { get; set; } = None;
         * short LeftBorderColor { get; set; } = 8;
         * BorderStyle BorderRight { get; set; } = None;
         * short RightBorderColor { get; set; } = 8;
         * BorderDiagonal BorderDiagonal { get; set; } = None;
         * BorderStyle BorderDiagonalLineStyle { get; set; } = None;
         * short BorderDiagonalColor { get; set; } = 0;
         * 
         * // HSSFCellStyle 独有属性
         * string UserStyleName { get; set; } = null; // 可以定义名字
         * HSSFCellStyle ParentStyle { get; } = null;
         * short ReadingOrder { get; set; } = 0;
         */

        /// <summary>
        /// 枚举当前表格文档所有内建的单元格样式.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IEnumerable<ICellStyle> GetCellStyles(this IWorkbook workbook)
        {
            var count = workbook.NumCellStyles;
            for (short i = 0; i < count; i++)
                yield return workbook.GetCellStyleAt(i);
        }

        // 等待优化
        static ICellStyle GetOrAddCellStyle(this IWorkbook workbook, MockCellStyle other)
        {
            if (workbook == null)
                throw new ArgumentNullException();
            if (other == null)
                throw new ArgumentNullException();
            //var comparer = CellStyleEqualityComparer.Default;
            //foreach (var i in workbook.GetCellStyles())
            //    if (comparer.Equals(i, other))
            //        return i;

            foreach (var s in workbook.GetCellStyles())
            {

            }

            //var style = workbook.CreateCellStyle();
            //// 格式
            //style.DataFormat = other.DataFormat;
            //// 字体
            //var font = workbook.GetFontAt(other.FontIndex);
            //style.SetFont(font);
            //// 对齐
            //style.Alignment = other.Alignment;
            //style.Indention = other.Indention;
            //style.VerticalAlignment = other.VerticalAlignment;
            //style.WrapText = other.WrapText;
            //style.ShrinkToFit = other.ShrinkToFit;
            //style.Rotation = other.Rotation;
            //// 边框
            //style.BorderLeft = other.BorderLeft;
            //style.LeftBorderColor = other.LeftBorderColor;
            //style.BorderRight = other.BorderRight;
            //style.RightBorderColor = other.RightBorderColor;
            //style.BorderTop = other.BorderTop;
            //style.TopBorderColor = other.TopBorderColor;
            //style.BorderBottom = other.BorderBottom;
            //style.BottomBorderColor = other.BottomBorderColor;
            //style.BorderDiagonal = other.BorderDiagonal;
            //style.BorderDiagonalLineStyle = other.BorderDiagonalLineStyle;
            //style.BorderDiagonalColor = other.BorderDiagonalColor;
            //// 图案
            //style.FillPattern = other.FillPattern;
            //style.FillForegroundColor = other.FillForegroundColor;
            //style.FillBackgroundColor = other.FillBackgroundColor;
            //// 保护
            //style.IsHidden = other.IsHidden;
            //style.IsLocked = other.IsLocked;

            //return style;

            throw new NotSupportedException();
        }

        #endregion

        #region ICellStyle/Reflection~

        static InternalWorkbook GetInternalWorkbook(this HSSFCellStyle style)
        {
            if (style == null)
                throw new ArgumentNullException(nameof(style));

            var getter = NpoiReflection.HSSFCellStyleGetWorkbook.Value;
            if (getter != null)
                return getter(style);
            return null;
        }

        #endregion

        #region ICellStyle/SetFill~

        /* ICellStyle/Fill
         * 
         * XSSFCellStyle {
         *     FillPattern FillPattern { get; set; }
         *     
         *     short FillForegroundColor {
         *         get {
         *             var xssfcolor = (XSSFColor)this.FillForegroundColorColor;
         *             if (xssfcolor != null)
         *                 return xssfcolor.Indexed;
         *             return IndexedColors.Automatic.Index;
         *         }
         *         set {
         *             this.SetFillForegroundColor(new XSSFColor { Indexed = value });
         *         }
         *     }
         *     IColor FillForegroundColorColor {
         *         get {
         *             return this.FillForegroundXSSFColor;
         *         }
         *         set {
         *             this.FillForegroundXSSFColor = (XSSFColor)value;
         *         }
         *     }
         *     XSSFColor FillForegroundXSSFColor { get; set; }
         *     void SetFillForegroundColor(XSSFColor color) {
         *     }
         *     
         *     short FillBackgroundColor {
         *         get {
         *             var xssfcolor = (XSSFColor)this.FillBackgroundColorColor;
         *             if (xssfcolor != null)
         *                 return xssfcolor.Indexed;
         *             return IndexedColors.Automatic.Index;
         *         }
         *         set {
         *             this.SetFillBackgroundColor(new XSSFColor { Indexed = value });
         *         }
         *     }
         *     IColor FillBackgroundColorColor {
         *         get {
         *             return this.FillBackgroundXSSFColor;
         *         }
         *         set {
         *             this.FillBackgroundXSSFColor = (XSSFColor)value;
         *         }
         *     }
         *     XSSFColor FillBackgroundXSSFColor { get; set; }
         *     void SetFillBackgroundColor(XSSFColor color) {
         *     }
         * }
         * 
         * HSSFCellStyle {
         *     FillPattern FillPattern { get; set; }
         *     
         *     FillForegroundColor { // 只能通过该属性设置前景色;
         *         get {
         *             return this._format.FillForeground;
         *         }
         *         set {
         *             this._format.FillForeground = value;
         *             this.CheckDefaultBackgroundFills();
         *         }
         *     }
         *     IColor FillForegroundColorColor { // 只读
         *         get {
         *             var hssfpalette = new HSSFPalette(this._workbook.CustomPalette);
         *             return hssfpalette.GetColor(this.FillForegroundColor);
         *         }
         *     }
         *     
         *     short FillBackgroundColor { // 只能通过该属性设置背景色;
         *         get {
         *             short fillBackground = this._format.FillBackground;
         *             if (fillBackground == 65)
         *                 return 64;
         *             return fillBackground;
         *         }
         *         set {
         *             this._format.FillBackground = value;
         *             this.CheckDefaultBackgroundFills();
         *         }
         *     }
         *     IColor FillBackgroundColorColor { // 只读
         *         get {
         *             var hssfpalette = new HSSFPalette(this._workbook.CustomPalette);
         *             return hssfpalette.GetColor(this.FillBackgroundColor);
         *         }
         *     }
         *     
         *     // Checks if the background and foreground Fills are Set correctly when one or the other is Set to the default color.
         *     // Works like the logic table below:
		 *     // BACKGROUND   FOREGROUND
		 *     // NONE         AUTOMATIC
		 *     // 0x41         0x40
		 *     // NONE         RED/ANYTHING
		 *     // 0x40         0xSOMETHING
         *     CheckDefaultBackgroundFills() {
         *         if (this._format.FillForeground == 64) {
         *             if (this._format.FillBackground != 65)
         *                 this.FillBackgroundColor = 65;
         *         }
         *         else if (this._format.FillBackground == 65) {
         *             if (this._format.FillForeground != 64)
         *                 this.FillBackgroundColor = 64;
         *         }
         *     }
         * }
         */

        /// <summary>
        /// 
        /// </summary>
        /// <param name="style"></param>
        /// <param name="pattern">图案样式; null: 不作修改; default: NoFill;</param>
        /// <param name="foreground">图案颜色; null: 不作修改; default: Automatic;</param>
        /// <param name="background">背景颜色; null: 不作修改; default: Automatic;</param>
        public static void SetFill(this ICellStyle style, FillPattern? pattern, NpoiColor foreground = null, NpoiColor background = null)
        {
            if (style == null)
                throw new ArgumentNullException(nameof(style));

            var xssf = style as XSSFCellStyle;
            if (xssf != null)
            {
                if (pattern.HasValue)
                    xssf.FillPattern = pattern.Value;

                if (foreground != null)
                {
                    var f = foreground.Indexed > 0 ? new XSSFColor() { Indexed = foreground.Indexed } : new XSSFColor(foreground.Argb);
                    xssf.SetFillForegroundColor(f);
                }
                if (background != null)
                {
                    var b = background.Indexed > 0 ? new XSSFColor() { Indexed = background.Indexed } : new XSSFColor(background.Argb);
                    xssf.SetFillBackgroundColor(b);
                }
            }

            var hssf = style as HSSFCellStyle;
            if (hssf != null)
            {
                if (pattern.HasValue)
                    hssf.FillPattern = pattern.Value;

                if (foreground != null || background != null)
                {
                    var workbook = GetInternalWorkbook(hssf);
                    var palette = new HSSFPalette(workbook.CustomPalette);

                    // Optionally a Foreground and background Fill can be applied.
                    // Note: Ensure Foreground color is Set prior to background
                    // 注意: 确保前景色设置在背景色之前.
                    if (foreground != null)
                    {
                        var f = FindColorIndexedWithPalette(palette, foreground.Indexed, foreground.Argb, true/*允许模糊查询, 保证结果有值*/);
                        hssf.FillForegroundColor = f;
                    }
                    if (background != null)
                    {
                        var b = FindColorIndexedWithPalette(palette, background.Indexed, background.Argb, true/*允许模糊查询, 保证结果有值*/);
                        hssf.FillBackgroundColor = b;
                    }
                }
            }
        }

        #endregion

        #region ICellStyle/SetTextAlignment~

        /// <summary>
        /// 设置对齐
        /// </summary>
        /// <param name="style"></param>
        /// <param name="horizontal">水平对齐; null: 不作修改; default: General;</param>
        /// <param name="indention">水平缩进; null: 不作修改; default: 0;</param>
        /// <param name="vertical">垂直对齐; null: 不作修改; default: Bottom;</param>
        /// <param name="wrap">自动换行; null: 不作修改; default: false;</param>
        /// <param name="shrink">自动缩小字体填充; null: 不作修改; default: false;</param>
        /// <param name="rotation">文字旋转角度; null: 不作修改; default: 0;</param>
        public static void SetTextAlignment(this ICellStyle style, HorizontalAlignment? horizontal = null, short? indention = null, VerticalAlignment? vertical = null, bool? wrap = null, bool? shrink = null, short? rotation = null)
        {
            if (style == null)
                throw new ArgumentNullException(nameof(style));

            if (horizontal.HasValue)
                style.Alignment = horizontal.Value;
            if (indention.HasValue)
                style.Indention = indention.Value;

            if (vertical.HasValue)
                style.VerticalAlignment = vertical.Value;

            if (wrap.HasValue)
                style.WrapText = wrap.Value;
            if (shrink.HasValue)
                style.ShrinkToFit = shrink.Value;

            if (rotation.HasValue)
                style.Rotation = rotation.Value;
        }

        #endregion

        #region ICellStyle/SetBorders~

        /// <summary>
        /// 设置边框
        /// </summary>
        /// <param name="style"></param>
        /// <param name="edges">边框组合</param>
        /// <param name="type">边框样式; null: 不作修改; default: None;</param>
        /// <param name="color">边框颜色; null: 不作修改; default: Black;</param>
        /// <example>
        /// // 添加顶边以及左边框线
        /// style.SetBorders(BorderSides.Top | BorderSides.Left, BorderStyle.Thin, "Black");
        /// // 添加四周边框
        /// style.SetBorders(BorderSides.AllAround, BorderStyle.Thin, "Red");
        /// // 添加逆对角线
        /// style.SetBorders(BorderSides.LeftTopToRightBottom, BorderStyle.Thin, "Blue");
        /// </example>
        public static void SetBorders(this ICellStyle style, BorderEdges edges, BorderStyle? type, NpoiColor color)
        {
            if (style == null)
                throw new ArgumentNullException(nameof(style));

            var xssf = style as XSSFCellStyle;
            if (xssf != null)
            {
                XSSFColor c = null;
                if (color != null)
                {
                    c = color.Indexed > 0 ? new XSSFColor() { Indexed = color.Indexed } : new XSSFColor(color.Argb);
                }

                if ((edges & BorderEdges.Top) != BorderEdges.None)
                {
                    if (type.HasValue)
                        xssf.BorderTop = type.Value;
                    if (c != null)
                        xssf.SetTopBorderColor(c);
                }
                if ((edges & BorderEdges.Right) != BorderEdges.None)
                {
                    if (type.HasValue)
                        xssf.BorderRight = type.Value;
                    if (c != null)
                        xssf.SetRightBorderColor(c);
                }
                if ((edges & BorderEdges.Bottom) != BorderEdges.None)
                {
                    if (type.HasValue)
                        xssf.BorderBottom = type.Value;
                    if (c != null)
                        xssf.SetBottomBorderColor(c);
                }
                if ((edges & BorderEdges.Left) != BorderEdges.None)
                {
                    if (type.HasValue)
                        xssf.BorderLeft = type.Value;
                    if (c != null)
                        xssf.SetLeftBorderColor(c);
                }

                var diagonal = (BorderDiagonal)((int)(edges & BorderEdges.AllDiagonal) >> 0x08);
                if (diagonal != BorderDiagonal.None)
                {
                    xssf.BorderDiagonal = diagonal;
                    if (type.HasValue)
                        xssf.BorderDiagonalLineStyle = type.Value;
                    if (c != null)
                        xssf.SetDiagonalBorderColor(c);
                }
            }

            var hssf = style as HSSFCellStyle;
            if (hssf != null)
            {
                short index = 0;
                if (color != null)
                {
                    var workbook = hssf.GetInternalWorkbook();
                    var palette = new HSSFPalette(workbook.CustomPalette);
                    index = FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, true);
                }

                if ((edges & BorderEdges.Top) != BorderEdges.None)
                {
                    if (type.HasValue)
                        hssf.BorderTop = type.Value;
                    if (index > 0)
                        hssf.TopBorderColor = index;
                }
                if ((edges & BorderEdges.Right) != BorderEdges.None)
                {
                    if (type.HasValue)
                        hssf.BorderRight = type.Value;
                    if (index > 0)
                        hssf.RightBorderColor = index;
                }
                if ((edges & BorderEdges.Bottom) != BorderEdges.None)
                {
                    if (type.HasValue)
                        hssf.BorderBottom = type.Value;
                    if (index > 0)
                        hssf.BottomBorderColor = index;
                }
                if ((edges & BorderEdges.Left) != BorderEdges.None)
                {
                    if (type.HasValue)
                        hssf.BorderLeft = type.Value;
                    if (index > 0)
                        hssf.LeftBorderColor = index;
                }

                var diagonal = (BorderDiagonal)((int)(edges & BorderEdges.AllDiagonal) >> 0x08);
                if (diagonal != BorderDiagonal.None)
                {
                    hssf.BorderDiagonal = diagonal;
                    if (type.HasValue)
                        hssf.BorderDiagonalLineStyle = type.Value;
                    if (index > 0)
                        hssf.BorderDiagonalColor = index;
                }
            }
        }

        #endregion

        #region IWorkbook/Font~

        /* public interface IFont
         * {
         *     short Index { get; } // 索引; 依赖 Workbook 实例;
         * 
         *     string FontName { get; set; } // 字体名称
         *     
         *     double FontHeight { get; set; } // 字体大小; Unit: 1/20 Point
         *     short FontHeightInPoints { get; set; }
         *     
         *     short Boldweight { get; set; } // 属性过时
         *     bool IsBold { get; set; } // 粗体
         *     bool IsItalic { get; set; } // 斜体
         *     FontUnderlineType Underline { get; set; } // 下划
         *     bool IsStrikeout { get; set; } // 删除
         *     
         *     FontSuperScript TypeOffset { get; set; } // 上标下标
         *     
         *     short Charset { get; set; } // 字符集
         *     
         *     short Color { get; set; } // 字体颜色
         * 
         *     void CloneStyleFrom(IFont src); // 克隆
         * }
         * 
         * NPOI.HSSF.UserModel.XSSFFont {
         *     string FontName { get; set; } = "Calibri";
         *     short FontHeightInPoints { get; set; } = 11.0;
         *     bool IsBold { get; set; } = false;
         *     bool IsItalic { get; set; } = false;
         *     FontUnderlineType Underline { get; set; } = None;
         *     bool IsStrikeout { get; set; } = false;
         *     FontSuperScript TypeOffset { get; set; } = None;
         *     short Charset { get; set; } = 0;
         *     short Color { get; set; } = 8; // 若是 标准颜色 返回对应 Indexed; 若是 没有指定颜色 或者 其它颜色 全部返回 8;
         * }
         * NPOI.XSSF.UserModel.HSSFFont {
         *     string FontName { get; set; } = "Arial";
         *     short FontHeightInPoints { get; set; } = 10.0;
         *     bool IsBold { get; set; } = false;
         *     bool IsItalic { get; set; } = false;
         *     FontUnderlineType Underline { get; set; } = None;
         *     bool IsStrikeout { get; set; } = false;
         *     FontSuperScript TypeOffset { get; set; } = None;
         *     short Charset { get; set; } = 0;
         *     short Color { get; set; } = 0x7FFF;
         * }
         */

        /* XSSFFont : IFont {
         *     double FontHeight {
         *         get {
         *             if (ct_FontSize != null)
         *                 return (double)(short)(ct_FontSize.val * 20.0);
         *             return DEFAULT_FONT_SIZE * 20.0;
         *         }
         *         set {
         *             ct_FontSize.val = value; // BUG: 赋值没有除于 20;
         *         }
         *     }
         *     short FontHeightInPoints {
         *         get {
         *             return (short)(this.FontHeight / 20.0);
         *         }
         *         set {
         *             ct_FontSize.val = value;
         *         }
         *     }
         * }
         * https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.fontsize.aspx
         * CT_FontSize {
         *     double val { get; set; } // Unit: Points;
         * }
         * 
         * HSSFFont : IFont {
         *     FontRecord font;
         *     double FontHeight {
         *         get {
         *             return (double)this.font.FontHeight;
         *         }
         *         set {
         *             this.font.FontHeight = (short)value;
         *         }
         *     }
         *     short FontHeightInPoints {
         *         get {
         *             this.font.FontHeight / 20;
         *         }
         *         set {
         *             this.font.FontHeight = value * 20;
         *         }
         *     }
         * }
         * FontRecord {
         *     short FontHeight { get; set; } // Unit: 1/20 Point;
         * }
         */

        /* IFont.Charset
         * 
         * NPOI.SS.UserModel.FontCharset
         *   0: ANSI
         *   1: DEFAULT        默认
         *   2: SYMBOL         符号
         *  77: MAC
         * 128: SHIFTJIS       日语
         * 129: HANGUL/HANGEUL 韩文
         * 130: JOHAB          朝鲜
         * 134: GB2312         简体中文
         * 136: CHINESEBIG5    繁体中文
         * 161: GREEK          希腊字符
         * 162: TURKISH        土耳其语
         * 163: VIETNAMESE     越南
         * 177: HEBREW         希伯来语
         * 178: ARABIC         阿拉伯语
         * 186: BALTIC         波罗的海
         * 204: RUSSIAN        俄语
         * 222: THAI           泰语
         * 238: EASTEUROPE     东欧
         * 255: OEM
         * 
         * System.Drawing.Font.GdiCharSet
         * https://msdn.microsoft.com/zh-cn/library/system.drawing.font.gdicharset(v=vs.100).aspx
         */

        /* IWorkbook 字体相关 (四个成员):
         * 
         * short NumberOfFonts { get; }
         * IFont GetFontAt(short idx);
         * IFont FindFont(...); // fontHeight (Unit: 1/20 Point)
         * IFont CreateFont();
         * 
         * XSSFWorkbook {
         *     FindFont(...) {
         *         return this.stylesSource.FindFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
         *     }
         * }
         * StylesTable {
         *     FindFont(...) {
         *         foreach (XSSFFont xssffont in this.fonts) {
         *             if (xssffont.Boldweight == boldWeight
         *              && xssffont.Color == color
         *              && xssffont.FontHeight == (double)fontHeight
         *              && xssffont.FontName.Equals(name)
         *              && xssffont.IsItalic == italic
         *              && xssffont.IsStrikeout == strikeout
         *              && xssffont.TypeOffset == typeOffset
         *              && xssffont.Underline == underline) {
         *                 return xssffont;
         *             }
         *         }
         *         return null;
         *     }
         * }
         * 
         * HSSFWorkbook {
         *     FindFont(...) {
         *         for (short index = 0; index <= this.NumberOfFonts; index++) {
         *             if (index != 4) {
         *                 IFont fontAt = this.GetFontAt(num);
         *                 if (fontAt.Boldweight == boldWeight
         *                  && fontAt.Color == color
         *                  && fontAt.FontHeight == (double)fontHeight
         *                  && fontAt.FontName.Equals(name)
         *                  && fontAt.IsItalic == italic
         *                  && fontAt.IsStrikeout == strikeout
         *                  && fontAt.TypeOffset == typeOffset
         *                  && fontAt.Underline == underline) {
         *                     return fontAt; 
         *                 }
         *             }
         *         }
         *         return null;
         *     }
         * }
         */

        /// <summary>
        /// 枚举当前表格文档所有内建字体.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        /// <remarks>
        /// http://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Font.html
        /// </remarks>
        public static IEnumerable<IFont> GetFonts(this IWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            var count = workbook.NumberOfFonts;
            for (short i = 0; i < count; i++)
                yield return workbook.GetFontAt(i);
        }

        /// <summary>
        /// 获取 或者 添加 指定属性字体.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name">字体名称; default: XSSFWorkbook=Calibri / HSSFWorkbook=Arial;</param>
        /// <param name="size">字号 FontHeightInPoints; default: XSSFWorkbook=11.0 / HSSFWorkbook=10.0;</param>
        /// <param name="bold">粗体; default: false</param>
        /// <param name="italic">斜体; default: false</param>
        /// <param name="underline">下划; default: None</param>
        /// <param name="strikeout">删除; default: false</param>
        /// <param name="offset">上标下标; default: None</param>
        /// <param name="color">字体颜色; default: XSSFWorkbook=Black / HSSFWorkbook=0x7FFF;</param>
        /// <returns></returns>
        public static IFont GetOrAddFont(this IWorkbook workbook,
            string name, float size,
            bool bold = false, bool italic = false, FontUnderlineType underline = FontUnderlineType.None, bool strikeout = false,
            FontSuperScript offset = FontSuperScript.None,
            NpoiColor color = null)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException(nameof(name));
            if (size <= 0f)
                throw new ArgumentOutOfRangeException(nameof(size));
            if (color == null)
                color = NpoiColor.Black;

            // 由于 IFont 接口 FontHeight/FontHeightInPoints 属性数据类型设计不太合理
            // 为了统一, 对于带小数部分不作支持, 因此此处取整后再转换..
            var points = (short)size;
            var height = (short)(points * 20);
            var weight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;

            var sssf = workbook as SXSSFWorkbook;
            if (sssf != null)
                workbook = sssf.XssfWorkbook;
            var xssf = workbook as XSSFWorkbook;
            if (xssf != null)
            {
                XSSFFont font = null;
                /* SXSSFWorkbook.XssfWorkbook.GetStylesSource();
                 */
                var fonts = xssf.GetStylesSource().GetFonts();
                foreach (var f in fonts)
                {
                    if (string.Equals(f.FontName, name, StringComparison.OrdinalIgnoreCase)
                        && f.FontHeightInPoints == size
                        && f.IsBold == bold && f.IsItalic == italic && f.Underline == underline && f.IsStrikeout == strikeout
                        && f.TypeOffset == offset)
                    {
                        // 如果 Font 没有指定颜色, 返回空值, 并且 IFont.Color = 8;
                        var c = f.GetXSSFColor();
                        if ((c == null && color.IsBlack) || (c != null && color.Equals(c)))
                        {
                            font = f;
                            break;
                        }
                    }
                }
                if (font == null)
                {
                    /* SXSSFWorkbook.CreateFont() {
                     *     return this.XssfWorkbook.CreateFont();
                     * }
                     */
                    font = (XSSFFont)xssf.CreateFont();

                    font.FontName = name;

                    font.FontHeightInPoints = points;

                    font.IsBold = bold;
                    font.IsItalic = italic;
                    font.Underline = underline;
                    font.IsStrikeout = strikeout;

                    font.TypeOffset = offset;

                    if (color.Indexed > 0)
                    {
                        /* XSSFFont.Color.Set {
                         *     CT_Color ct_Color = (this._ctFont.sizeOfColorArray() == 0) ? this._ctFont.AddNewColor() : this._ctFont.GetColorArray(0);
                         *     if (value == 0x7FFF) {
                         *         ct_Color.indexed = (uint)XSSFFont.DEFAULT_FONT_COLOR; // Black
                         *         ct_Color.indexedSpecified = true;
                         *     }
                         *     else {
                         *         ct_Color.indexed = (uint)value;
                         *         ct_Color.indexedSpecified = true;
                         *     }
                         * }
                         */
                        font.Color = color.Indexed;
                    }
                    else
                    {
                        /* 	XSSFFont.SetColor(XSSFColor color) {
                         * 	    CT_Color ct_Color = (this._ctFont.sizeOfColorArray() == 0) ? this._ctFont.AddNewColor() : this._ctFont.GetColorArray(0);
                         * 	    ct_Color.SetRgb(color.RGB);
                         * 	}
                         */
                        font.SetColor(new XSSFColor(color.Argb));
                    }
                }

                return font;
            }

            var hssf = workbook as HSSFWorkbook;
            if (hssf != null)
            {
                var palette = hssf.GetCustomPalette();
                var c = FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, true);
                var font = workbook.FindFont(weight, c, height, name, italic, strikeout, offset, underline);
                if (font == null)
                {
                    font = workbook.CreateFont();

                    font.FontName = name;

                    // v2.2.1.0 XSSFFont.FontHeight 存在 BUG, 改用 FontHeightInPoints.
                    font.FontHeightInPoints = points; // 此处赋值整数, 为了统一: IWorkbook.FindFont();

                    font.IsBold = bold; // Boldweight: 属性过时;
                    font.IsItalic = italic;
                    font.Underline = underline;
                    font.IsStrikeout = strikeout;

                    font.TypeOffset = offset;

                    font.Color = c;
                }
                return font;
            }
            throw new NotSupportedException();
        }

        /// <summary>
        /// 获取默认字体
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IFont GetDefaultFont(this IWorkbook workbook)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            return workbook.GetFontAt(0);
        }

        /// <summary>
        /// 修改默认字体
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name">字体名称; default: XSSFWorkbook=Calibri / HSSFWorkbook=Arial;</param>
        /// <param name="size">字号 FontHeightInPoints; default: XSSFWorkbook=11.0 / HSSFWorkbook=10.0;</param>
        /// <param name="bold">粗体; default: false</param>
        /// <param name="italic">斜体; default: false</param>
        /// <param name="underline">下划; default: None</param>
        /// <param name="strikeout">删除; default: false</param>
        /// <param name="offset">上标下标; default: None</param>
        /// <param name="color">字体颜色; null: 不作修改; default: XSSFWorkbook=Black / HSSFWorkbook=0x7FFF;</param>
        public static void SetDefaultFont(this IWorkbook workbook,
            string name, float size,
            bool bold = false, bool italic = false, FontUnderlineType underline = FontUnderlineType.None, bool strikeout = false,
            FontSuperScript offset = FontSuperScript.None,
            NpoiColor color = null)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException(nameof(name));
            if (size <= 0f)
                throw new ArgumentOutOfRangeException(nameof(size));

            var font = workbook.GetFontAt(0);
            font.FontName = name;
            font.FontHeightInPoints = (short)size;
            font.IsBold = bold;
            font.IsItalic = italic;
            font.Underline = underline;
            font.IsStrikeout = strikeout;
            font.TypeOffset = offset;
            if (color != null)
            {
                if (workbook is XSSFWorkbook || workbook is SXSSFWorkbook)
                {
                    if (color.Indexed > 0)
                    {
                        font.Color = color.Indexed;
                    }
                    else
                    {
                        ((XSSFFont)font).SetColor(new XSSFColor(color.Argb));
                    }
                }
                else if (workbook is HSSFWorkbook)
                {
                    var hssf = workbook as HSSFWorkbook;
                    var palette = hssf.GetCustomPalette();
                    var c = FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, true);
                    font.Color = c;
                }
            }
        }

        static NpoiColor GetNpoiColor(this IFont font, IWorkbook workbook = null)
        {
            var xssf = font as XSSFFont;
            if (xssf != null)
            {
                var color = xssf.GetXSSFColor();
                if (color == null)
                {
                    return NpoiColor.Black;
                }
                if (xssf.GetThemeColor() != 0)
                {
                    //TODO:
                }
                if (color.Indexed > 0)
                {
                    return (NpoiColor)color.Indexed;
                }
                else
                {
                    return (NpoiColor)NpoiColor.GetArgbValue(color.GetARgb());
                }
            }
            var hssf = font as HSSFFont;
            if (hssf != null) // 使用文档配色
            {
                var w = workbook as HSSFWorkbook;
                if (w != null)
                {
                    var c = w.GetCustomPalette().GetColor(hssf.Color);
                    var argb = NpoiColor.GetArgbValue(c.GetTriplet());
                    NpoiColor.TryGetColor(argb, out NpoiColor color);
                    return color;
                }
                else // 使用标准配色
                {
                    NpoiColor.TryGetColor(hssf.Color, out NpoiColor color); // 可能空值;
                    return color;
                }
            }
            throw new NotSupportedException();
        }

        #endregion

        #region Color

        public static XSSFColor GetColor(this XSSFWorkbook workbook, NpoiColor color)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (color == null)
                throw new ArgumentNullException(nameof(color));

            return color.Indexed > 0 ? new XSSFColor() { Indexed = color.Indexed } : new XSSFColor(color.Argb);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="color"></param>
        /// <param name="similar">如果精确查找不到, 是否查找相近颜色;</param>
        /// <returns></returns>
        public static short GetColorIndexed(this HSSFWorkbook workbook, NpoiColor color, bool similar = true)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (color == null)
                throw new ArgumentNullException(nameof(color));

            var palette = workbook.GetCustomPalette();
            return FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, similar);
        }

        static short FindColorIndexedWithPalette(HSSFPalette palette, short indexed, byte[] argb/*不能空值*/, bool similar)
        {
            if (indexed > 0) // 若是标准颜色
            {
                // 获取该色盘指定位置的颜色;
                var d = palette.GetColor(indexed);
                if (d != null)
                {
                    // 检查颜色是否默认 (该位置的颜色没修改过)
                    if (NpoiColor.Equals(d.RGB, argb, false/*排除 Alpha 通道*/))
                        return indexed;
                }
            }

            NpoiColor.GetArgb(argb, out byte a, out byte r, out byte g, out byte b);

            // 如果默认位置颜色已被修改, 则将检查其它位置是否匹配.
            var c = palette.FindColor(r, g, b);
            if (c != null)
                return c.Indexed;

            if (similar)
            {
                // 如果精确查找无法匹配, 则将模糊查找;
                c = palette.FindSimilarColor(r, g, b);
                if (c != null)
                    return c.Indexed;
            }
            return 0;
        }
        static short FindColorIndexedWithPalette(HSSFPalette palette, NpoiColor color, IDictionary<int, short> caching/*Key: ArgbValue; Value: Indexed*/)
        {
            if (palette == null)
                throw new ArgumentNullException();
            if (color == null)
                throw new NotSupportedException();

            if (caching != null)
            {
                if (!caching.TryGetValue(color.ArgbValue, out short indexed))
                {
                    indexed = FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, true);
                    caching.Add(color.ArgbValue, indexed);
                }
                return indexed;
            }
            else
            {
                var indexed = FindColorIndexedWithPalette(palette, color.Indexed, color.Argb, true);
                return indexed;
            }
        }

        #endregion

        /* NPOI.XSSF.Model.ThemesTable
         * 主题颜色定义数量: 12;
         * https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.drawing.colorscheme.aspx
         * Dark1 深色1
         * Light1 浅色1
         * Dark2 深色2
         * Light2 浅色2
         * Accent1 文字颜色1
         * Accent2 文字颜色2
         * Accent3 文字颜色3
         * Accent4 文字颜色4
         * Accent5 文字颜色5
         * Accent6 文字颜色6
         * Hyperlink 超链接
         * FollowedHyperlink 关注的超链接
         */

        #region Palette
        //public static void FixDefaultPalette(this IWorkbook workbook)
        //{
        //    var hssf = workbook as HSSFWorkbook;
        //    if (hssf != null)
        //    {
        //        var palette = hssf.GetCustomPalette();

        //        var maroon = IndexedColors.Maroon;
        //        palette.SetColorAtIndex(maroon.Index, maroon.RGB[0], maroon.RGB[1], maroon.RGB[2]);
        //    }
        //}

        //public static bool HasDefaultPalette(this IWorkbook workbook)
        //{
        //    var hssf = workbook as HSSFWorkbook;
        //    if (hssf != null)
        //    {
        //        var palette = hssf.GetCustomPalette();
        //        foreach (var p in _MappingIndexed.Values)
        //        {
        //            var c = palette.GetColor(p.Indexed);
        //            if (c == null)
        //                return false;

        //            if (!p.Equals(c.RGB))
        //                return false;
        //        }
        //        return true;
        //    }

        //    return true; // XSSFWorkbook 向后兼容, 支持 标准色盘;
        //}
        //public static bool HasDefaultColor(this IWorkbook workbook, short indexed)
        //{
        //    var hssf = workbook as HSSFWorkbook;
        //    if (hssf != null)
        //    {
        //        if (_MappingIndexed.TryGetValue(indexed, out PColor value))
        //        {
        //            var palette = hssf.GetCustomPalette();
        //            var color = palette.GetColor(indexed);
        //            if (color != null)
        //                return value.Equals(color.RGB);
        //        }
        //        return false;
        //    }

        //    return true; // XSSFWorkbook 向后兼容, 支持 标准色盘;
        //}

        //public static bool IsDefaultColor(short indexed, byte[] argb)
        //{
        //    if (_MappingIndexed.TryGetValue(indexed, out PColor value))
        //    {
        //        return value.Equals(argb);
        //    }
        //    return false;
        //}
        #endregion
    }
}
