using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI.SS;
using NPOI.SS.UserModel;

namespace Hiz.Npoi
{
    // 当前 Workbook 的 样式配置 的 实例情况
    class ExcelContext
    {
        const string ExtensionXls = ".xls";
        const string ExtensionXlsx = ".xlsx";

        public string Path { get; set; }

        public OfficeArchiveFormat ArchiveFormat { get; set; }

        // 样式资源
        public ExcelOptions Options { get; set; }

        IWorkbook _Workbook = null;

        public int MaxRows
        {
            get
            {
                if (_Workbook is NPOI.XSSF.UserModel.XSSFWorkbook || _Workbook is NPOI.XSSF.Streaming.SXSSFWorkbook)
                    return SpreadsheetVersion.EXCEL2007.MaxRows;
                return SpreadsheetVersion.EXCEL97.MaxRows; // _Workbook is NPOI.HSSF.UserModel.HSSFWorkbook
            }
        }

        public OfficeArchiveFormat GetAppliedArchiveVersion()
        {
            if (this.ArchiveFormat == OfficeArchiveFormat.None)
            {
                // 不能匹配: 优先使用新的文件格式;
                return ExtensionXls.Equals(System.IO.Path.GetExtension(this.Path), StringComparison.OrdinalIgnoreCase) ? OfficeArchiveFormat.Binary : OfficeArchiveFormat.OpenXml;
            }
            return this.ArchiveFormat;
        }

        public ExcelContext(IWorkbook workbook, ExcelOptions options)
        {
            _Workbook = workbook;
            //this.Options = options;

            _IndexedCellStyleArray = new Dictionary<string, short>(StringComparer.OrdinalIgnoreCase);
        }

        // 当前 Workbook 样式索引
        IDictionary<string, short> _IndexedCellStyleArray;
        public virtual ICellStyle GetCellStyle(string name)
        {
            var options = this.Options;
            if (options == null)
                throw new InvalidOperationException();

            short value;
            if (!_IndexedCellStyleArray.TryGetValue(name, out value))
            {
                var style = options.GetCellStyle(name); // 没有找到将会抛出异常.

                //var font = style.FontOptions != null ? options.GetFont(style.FontOptions) : null;
                //var color = (font != null && font.Color != null) ? options.GetColor(font.Color) : null;

                //var cs = _Workbook.CreateCellStyle();

                //var format = style.FormatOptions != null ? options.GetDataFormat(style.FormatOptions) : null;
                //if (format != null)
                //{
                //    cs.DataFormat = _Workbook.GetOrAddDataFormat(format);
                //}

                //if (font != null)
                //{
                //    short c = 8;
                //    if (color != null)
                //    {
                //        c = ((IColor)color).Indexed;
                //    }

                //    cs.SetFont(_Workbook
                //        .GetOrAddFont(font.FontName, font.FontHeightInPoints,
                //        font.IsBold, font.IsItalic, font.Underline,
                //        font.IsStrikeout, font.TypeOffset,
                //        c));
                //}

                //value = cs.Index;
                //_IndexedCellStyleArray.Add(name, value);
            }

            return _Workbook.GetCellStyleAt(value);
        }
    }


}
