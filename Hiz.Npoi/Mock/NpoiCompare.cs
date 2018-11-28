using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    static class NpoiCompare
    {
        public static bool EqualsAlignment(ICellStyle style, ICopyCellAlignment alignment)
        {
            return (style == null && alignment == null)
                || (style != null && alignment != null
                && style.Alignment == alignment.Alignment && style.Indention == alignment.Indention
                && style.VerticalAlignment == alignment.VerticalAlignment
                && style.WrapText == alignment.WrapText && style.ShrinkToFit == alignment.ShrinkToFit
                && style.Rotation == alignment.Rotation)
                ;
        }

        public static bool EqualsFill(ICellStyle style, IMockCellFill fill)
        {
            var xssf = style as XSSFCellStyle;
            if (xssf != null)
            {
                return xssf.FillPattern == fill.FillPattern
                    && EqualsFillColor(xssf.FillForegroundXSSFColor, fill.FillForegroundColor)
                    && EqualsFillColor(xssf.FillBackgroundXSSFColor, fill.FillBackgroundColor)
                    ;
            }
            var hssf = style as HSSFCellStyle;
            if (hssf != null)
            {
                return hssf.FillPattern == fill.FillPattern
                    && hssf.FillForegroundColor == fill.FillForegroundColorIndexed
                    && hssf.FillBackgroundColor == fill.FillBackgroundColorIndexed
                    ;
            }
            throw new NotSupportedException();
        }
        static bool EqualsFillColor(XSSFColor color/*如果空值使用黑色*/, NpoiColor other/*不会存在空值情况*/)
        {
            return color != null ? other.Equals(color) : other.IsAutomatic;
        }

        public static bool EqualsBorder(ICellStyle style, IMockCellBorder border)
        {
            var xssf = style as XSSFCellStyle;
            if (xssf != null)
            {
                return xssf.BorderLeft == border.LeftStyle
                    && (border.LeftStyle == BorderStyle.None || EqualsBorderColor(xssf.LeftBorderXSSFColor, border.LeftColor))
                    && xssf.BorderRight == border.RightStyle
                    && (border.RightStyle == BorderStyle.None || EqualsBorderColor(xssf.RightBorderXSSFColor, border.RightColor))
                    && xssf.BorderTop == border.TopStyle
                    && (border.TopStyle == BorderStyle.None || EqualsBorderColor(xssf.TopBorderXSSFColor, border.TopColor))
                    && xssf.BorderBottom == border.BottomStyle
                    && (border.BottomStyle == BorderStyle.None || EqualsBorderColor(xssf.BottomBorderXSSFColor, border.BottomColor))
                    && xssf.BorderDiagonal == border.Diagonal && xssf.BorderDiagonalLineStyle == border.DiagonalStyle
                    && (border.Diagonal == BorderDiagonal.None || border.DiagonalStyle == BorderStyle.None || EqualsBorderColor(xssf.DiagonalBorderXSSFColor, border.DiagonalColor))
                    ;
            }
            var hssf = style as HSSFCellStyle;
            if (hssf != null)
            {
                return hssf.BorderLeft == border.LeftStyle
                    && (border.LeftStyle == BorderStyle.None || hssf.LeftBorderColor == border.LeftColorIndexed/*不会存在空值情况*/)
                    && hssf.BorderRight == border.RightStyle
                    && (border.LeftStyle == BorderStyle.None || hssf.RightBorderColor == border.RightColorIndexed)
                    && hssf.BorderTop == border.TopStyle
                    && (border.LeftStyle == BorderStyle.None || hssf.TopBorderColor == border.TopColorIndexed)
                    && hssf.BorderBottom == border.BottomStyle
                    && (border.LeftStyle == BorderStyle.None || hssf.BottomBorderColor == border.BottomColorIndexed)
                    && hssf.BorderDiagonal == border.Diagonal && hssf.BorderDiagonalLineStyle == border.DiagonalStyle
                    && (border.Diagonal == BorderDiagonal.None || border.DiagonalStyle == BorderStyle.None || hssf.BorderDiagonalColor == border.DiagonalColorIndexed)
                    ;
            }
            throw new NotSupportedException();
        }
        static bool EqualsBorderColor(XSSFColor color/*如果空值使用黑色*/, NpoiColor other/*不会存在空值情况*/)
        {
            return color != null ? other.Equals(color) : other.IsBlack;
        }

        public static bool Equals(ICellStyle style, MockCellStyle other)
        {
            var b = EqualsBorder(style, other) && EqualsFill(style, other) && EqualsAlignment(style, other);
            if (!b)
                return false;

            return b && style.DataFormat == other.DataFormat && style.FontIndex == other.FontIndex && style.IsHidden == other.IsHidden && style.IsLocked == other.IsLocked;
        }
    }
}
