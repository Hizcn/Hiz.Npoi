using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Hiz.Npoi
{
    /// <summary>
    /// 单元格的边框配置
    /// </summary>
    class CellBorderOptions: INamed
    {
        #region NPOI

        // BorderStyle ICellStyle.BorderLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.LeftBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // BorderStyle ICellStyle.BorderRight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.RightBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // BorderStyle ICellStyle.BorderTop { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.TopBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // BorderStyle ICellStyle.BorderBottom { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.BottomBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // BorderDiagonal ICellStyle.BorderDiagonal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderStyle ICellStyle.BorderDiagonalLineStyle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.BorderDiagonalColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        #endregion

        public string Name { get; set; }

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
}
