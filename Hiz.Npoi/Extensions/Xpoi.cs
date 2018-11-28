using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public static class Xpoi
    {
        public static IWorkbook OpenRead(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                return WorkbookFactory.Create(stream);
            }
        }
    }
}
