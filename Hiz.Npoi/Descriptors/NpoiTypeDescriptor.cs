using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    class NpoiTypeDescriptor<T> : ITableOptions
    {
        IList<NpoiPropertyDescriptor<T>> _Properties = null;
        public IList<NpoiPropertyDescriptor<T>> Properties
        {
            get
            {
                if (_Properties == null)
                    _Properties = new List<NpoiPropertyDescriptor<T>>();
                return this._Properties;
            }
        }

        public float RowDefaultHeight { get; set; }
        public float ColumnDefaultWidth { get; set; }
        public string CellDefaultStyle { get; set; }
        public bool HeaderVisible { get; set; } = true;
        public float HeaderHeight { get; set; }
        public string HeaderDefaultStyle { get; set; }
    }
}
