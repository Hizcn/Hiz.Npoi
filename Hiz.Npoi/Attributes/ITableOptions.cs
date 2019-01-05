using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    interface ITableOptions
    {
        float RowDefaultHeight { get; set; }

        float ColumnDefaultWidth { get; set; }

        string CellDefaultStyle { get; set; }

        bool HeaderVisible { get; set; }

        float HeaderHeight { get; set; }

        string HeaderDefaultStyle { get; set; }
    }
}
