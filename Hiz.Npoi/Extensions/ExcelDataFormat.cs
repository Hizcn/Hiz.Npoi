using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    class ExcelDataFormat
    {
        string _Format;
        public string Format { get { return _Format; } }

        int _Index;
        public int Index { get { return _Index; } }

        ExcelDataFormatSource _Source;
        public ExcelDataFormatSource Source { get { return _Source; } }

        public ExcelDataFormat(int index, string format, ExcelDataFormatSource source)
        {
            _Format = format;
            _Index = index;
            _Source = source;
        }

        public override string ToString()
        {
            // return base.ToString();
            return string.Format("{0:X2}; {1} {2}", this._Index, this._Source, this._Format);
        }
    }

    enum ExcelDataFormatSource
    {
        None,

        // 内建格式
        Builtin,
        // 本地化的内建格式
        BuiltinLocale,

        // 用户添加
        User,
    }
}
