using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    [DebuggerDisplay("Row: {RowIndex} Column: {ColumnIndex} {ErrorMessage}")]
    public class ExcelCellError
    {
        public int SheetIndex { get; internal set; }
        public string SheetName { get; internal set; }

        public int RowIndex { get; internal set; }
        public string RowString
        {
            get
            {
                var row = this.RowIndex;
                return (++row).ToString();
            }
        }

        public int ColumnIndex { get; internal set; }
        public string ColumnString
        {
            get
            {
                return NPOI.SS.Util.CellReference.ConvertNumToColString(this.ColumnIndex);
            }
        }

        public string ErrorMessage { get; internal set; }
    }
}