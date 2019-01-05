using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public class ColorOptions : INamed
    {
        #region NPOI
        // short IColor.Indexed => throw new NotImplementedException();
        // byte[] IColor.RGB => throw new NotImplementedException();
        #endregion

        public string Name { get; set; }
        public NpoiColor Color { get; set; }
    }
}
