using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    /// <summary>
    /// 单元格的值转换器
    /// </summary>
    abstract class CellValueConverter
    {
        public abstract Type Type { get; set; }

        public abstract object GetValue(ICell cell);
        public abstract void SetValue(ICell cell, object value);
    }
    abstract class CellValueConverter<TValue>
    {
        public abstract Type Type { get; set; }

        public abstract TValue GetValue(ICell cell);
        public abstract void SetValue(ICell cell, TValue value);
    }
}
