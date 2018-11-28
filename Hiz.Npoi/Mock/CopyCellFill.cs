using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    class CopyCellFill
    {
        public FillPattern FillPattern { get; set; }

        public NpoiColor FillBackgroundColor { get; set; }

        public NpoiColor FillForegroundColor { get; set; }
    }

    class MockCellFill
    {
        public short FillBackgroundColorIndexed { get; set; }
        public short FillForegroundColorIndexed { get; set; }
    }

    interface ICopyCellFill
    {
        FillPattern FillPattern { get; set; }

        NpoiColor FillBackgroundColor { get; set; }

        NpoiColor FillForegroundColor { get; set; }
    }

    interface IMockCellFill: ICopyCellFill
    {
        short FillBackgroundColorIndexed { get; set; }
        short FillForegroundColorIndexed { get; set; }
    }
}
