using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    class CopyCellBorder : ICopyCellBorder
    {
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
    }

    class MockCellBorder : CopyCellBorder, IMockCellBorder
    {
        public short LeftColorIndexed { get; set; } // 实例 HSSFWorkbook 调色盘的索引
        public short RightColorIndexed { get; set; }
        public short TopColorIndexed { get; set; }
        public short BottomColorIndexed { get; set; }
        public short DiagonalColorIndexed { get; set; }
    }

    interface ICopyCellBorder
    {
        BorderStyle LeftStyle { get; set; }
        NpoiColor LeftColor { get; set; }

        BorderStyle RightStyle { get; set; }
        NpoiColor RightColor { get; set; }

        BorderStyle TopStyle { get; set; }
        NpoiColor TopColor { get; set; }

        BorderStyle BottomStyle { get; set; }
        NpoiColor BottomColor { get; set; }

        BorderDiagonal Diagonal { get; set; }
        BorderStyle DiagonalStyle { get; set; }
        NpoiColor DiagonalColor { get; set; }
    }
    internal interface IMockCellBorder : ICopyCellBorder
    {
        short LeftColorIndexed { get; set; }
        short RightColorIndexed { get; set; }
        short TopColorIndexed { get; set; }
        short BottomColorIndexed { get; set; }
        short DiagonalColorIndexed { get; set; }
    }
}
