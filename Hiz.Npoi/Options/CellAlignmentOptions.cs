using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    class CellAlignmentOptions : INamed
    {
        public string Name { get; set; }

        public HorizontalAlignment Alignment { get; set; }
        public short Indention { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public bool WrapText { get; set; }
        public bool ShrinkToFit { get; set; }
        public short Rotation { get; set; }
    }
}
