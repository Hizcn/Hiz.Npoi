using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    /// <summary>
    /// 单元格样式相等比较器
    /// </summary>
    class CellStyleEqualityComparer : IEqualityComparer<ICellStyle>
    {
        // static CellStyleEqualityComparer _Default;
        // public static CellStyleEqualityComparer Default
        // {
        //     get
        //     {
        //         if (_Default == null)
        //             _Default = new CellStyleEqualityComparer();
        //         return _Default;
        //     }
        // }

        public IWorkbook Workbook { get; private set; }
        public static CellStyleEqualityComparer Create(IWorkbook workbook)
        {
            return new CellStyleEqualityComparer() { Workbook = workbook };
        }

        public bool Equals(ICellStyle x, ICellStyle y)
        {
            return (x == null && y == null)
                || (x != null && y != null
                // 比较格式
                && x.DataFormat == y.DataFormat
                // 比较字体
                && x.FontIndex == y.FontIndex
                // 比较对齐
                && x.Alignment == y.Alignment && x.Indention == y.Indention
                && x.VerticalAlignment == y.VerticalAlignment
                && x.WrapText == y.WrapText && x.ShrinkToFit == y.ShrinkToFit
                && x.Rotation == y.Rotation
                // 比较边框
                && x.BorderLeft == y.BorderLeft && x.LeftBorderColor == y.LeftBorderColor
                && x.BorderTop == y.BorderTop && x.TopBorderColor == y.TopBorderColor
                && x.BorderRight == y.BorderRight && x.RightBorderColor == y.RightBorderColor
                && x.BorderBottom == y.BorderBottom && x.BottomBorderColor == y.BottomBorderColor
                && x.BorderDiagonal == y.BorderDiagonal && x.BorderDiagonalLineStyle == y.BorderDiagonalLineStyle && x.BorderDiagonalColor == y.BorderDiagonalColor
                // 比较图案
                && x.FillPattern == y.FillPattern && x.FillForegroundColor == y.FillForegroundColor && x.FillBackgroundColor == y.FillBackgroundColor
                // 比较保护
                && x.IsLocked == y.IsLocked && x.IsHidden == y.IsHidden
                );
        }

        public int GetHashCode(ICellStyle style)
        {
            return style.Index;
        }
    }
}
