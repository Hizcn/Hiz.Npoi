using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    class FontEqualityComparer : IEqualityComparer<IFont>
    {
        //static FontEqualityComparer _Default;
        //public static FontEqualityComparer Default
        //{
        //    get
        //    {
        //        if (_Default == null)
        //            _Default = new FontEqualityComparer();
        //        return _Default;
        //    }
        //}

        public IWorkbook Workbook { get; protected set; }
        public FontEqualityComparer(IWorkbook workbook)
        {
            this.Workbook = workbook;
        }

        public bool Equals(IFont x, IFont y)
        {
            var equals = (x == null && y == null)
                || (x != null && y != null
                && string.Equals(x.FontName, y.FontName, StringComparison.OrdinalIgnoreCase)
                && x.FontHeight == y.FontHeight
                && x.IsBold == y.IsBold
                && x.IsItalic == y.IsItalic
                && x.Underline == y.Underline
                && x.IsStrikeout == y.IsStrikeout
                && x.TypeOffset == y.TypeOffset
                // && x.Charset == y.Charset
                // && x.Color == y.Color
                );
            if (equals)
            {
                var workbook = this.Workbook;
                if (workbook != null)
                {
                    //var value1 = x.GetColorArgbValue(workbook);
                    //var value2 = y.GetColorArgbValue(workbook);
                    //equals = value1 == value2;
                }
            }
            return equals;
        }

        public int GetHashCode(IFont font)
        {
            return font.Index;
        }
    }
}
