using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Hiz.Npoi
{
    class CopyCellAlignment: ICopyCellAlignment
    {
        public HorizontalAlignment Alignment { get; set; }
        public short Indention { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public bool ShrinkToFit { get; set; }
        public short Rotation { get; set; }
    }

    interface ICopyCellAlignment
    {
        HorizontalAlignment Alignment { get; set; }
        short Indention { get; set; }
        VerticalAlignment VerticalAlignment { get; set; }
        bool WrapText { get; set; }
        bool ShrinkToFit { get; set; }
        short Rotation { get; set; }
    }
}
