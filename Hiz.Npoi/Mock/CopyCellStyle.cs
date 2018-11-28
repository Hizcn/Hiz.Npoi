using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    /// <summary>
    /// CellStyle 字段
    /// </summary>
    class CopyCellStyle : ICopyCellAlignment, ICopyCellBorder, ICopyCellFill
    {
        #region 格式

        public short DataFormat { get; set; }

        #endregion

        #region 对齐

        public HorizontalAlignment Alignment { get; set; }
        public short Indention { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public bool ShrinkToFit { get; set; }
        public short Rotation { get; set; }

        #endregion

        #region 字体

        public short FontIndex { get; set; }

        #endregion

        #region 边框

        //public BorderStyle BorderLeft { get; set; }
        //public NpoiColor LeftBorderColor { get; set; }

        //public BorderStyle BorderRight { get; set; }
        //public NpoiColor RightBorderColor { get; set; }

        //public BorderStyle BorderTop { get; set; }
        //public NpoiColor TopBorderColor { get; set; }

        //public BorderStyle BorderBottom { get; set; }
        //public NpoiColor BottomBorderColor { get; set; }

        //public BorderDiagonal BorderDiagonal { get; set; }
        //public BorderStyle BorderDiagonalLineStyle { get; set; }
        //public NpoiColor BorderDiagonalColor { get; set; }

        public BorderStyle LeftStyle { get; set; }
        public NpoiColor LeftColor { get; set; }

        public BorderStyle RightStyle { get; set; }
        public NpoiColor RightColor { get; set; }

        public BorderStyle TopStyle { get; set; }
        public NpoiColor TopColor { get; set; }

        public BorderStyle BottomStyle { get; set; }
        public NpoiColor BottomColor { get; set; }

        public BorderDiagonal Diagonal { get; set; }
        public BorderStyle DiagonalStyle { get; set; }
        public NpoiColor DiagonalColor { get; set; }
        #endregion

        #region 图案

        public FillPattern FillPattern { get; set; }

        public NpoiColor FillBackgroundColor { get; set; }

        public NpoiColor FillForegroundColor { get; set; }

        #endregion

        #region 保护

        public bool IsHidden { get; set; } = false;
        public bool IsLocked { get; set; } = true;

        #endregion
    }

    class MockCellStyle : CopyCellStyle, IMockCellBorder, IMockCellFill
    {
        public short LeftColorIndexed { get; set; }
        public short RightColorIndexed { get; set; }
        public short TopColorIndexed { get; set; }
        public short BottomColorIndexed { get; set; }
        public short DiagonalColorIndexed { get; set; }

        public short FillBackgroundColorIndexed { get; set; }
        public short FillForegroundColorIndexed { get; set; }
    }
}
