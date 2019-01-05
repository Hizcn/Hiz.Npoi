using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public class CellFillOptions : INamed
    {
        #region NPOI

        // short ICellStyle.FillForegroundColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // IColor ICellStyle.FillForegroundColorColor => throw new NotImplementedException();
        // short ICellStyle.FillBackgroundColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // IColor ICellStyle.FillBackgroundColorColor => throw new NotImplementedException();

        #endregion

        public string Name { get; set; }

        public FillPattern FillPattern { get; set; }

        public NpoiColor FillBackgroundColor { get; set; }

        public NpoiColor FillForegroundColor { get; set; }
    }
}
