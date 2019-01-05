using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace Hiz.Npoi
{
    // 当前 Workbook 的 样式配置 的 实例缓存
    class ExcelContext
    {
        const string ExtensionXls = ".xls";
        const string ExtensionXlsx = ".xlsx";

        public string Path { get; set; }

        // public OfficeArchiveFormat ArchiveFormat { get; set; }

        // 样式资源
        public ExcelOptions Styles { get; set; }

        IWorkbook _Workbook = null;

        protected bool IsOpenXml
        {
            get
            {
                return _Workbook is XSSFWorkbook || _Workbook is SXSSFWorkbook;
            }
        }

        public int MaxRows
        {
            get
            {
                if (this.IsOpenXml)
                    return SpreadsheetVersion.EXCEL2007.MaxRows;
                return SpreadsheetVersion.EXCEL97.MaxRows; // _Workbook is NPOI.HSSF.UserModel.HSSFWorkbook
            }
        }

        //public OfficeArchiveFormat GetAppliedArchiveVersion()
        //{
        //    if (this.ArchiveFormat == OfficeArchiveFormat.None)
        //    {
        //        // 不能匹配: 优先使用新的文件格式;
        //        return ExtensionXls.Equals(System.IO.Path.GetExtension(this.Path), StringComparison.OrdinalIgnoreCase) ? OfficeArchiveFormat.Binary : OfficeArchiveFormat.OpenXml;
        //    }
        //    return this.ArchiveFormat;
        //}

        public ExcelContext(IWorkbook workbook, ExcelOptions options)
        {
            _Workbook = workbook;
            this.Styles = options;

            _IndexedCellStyleArray = new Dictionary<string, ICellStyle>(StringComparer.OrdinalIgnoreCase);
            _IndexedFontArray = new Dictionary<string, IFont>(StringComparer.OrdinalIgnoreCase);
        }

        // 当前 Workbook 样式索引
        IDictionary<string, ICellStyle> _IndexedCellStyleArray;
        public virtual ICellStyle GetCellStyle(string name)
        {
            var styles = this.Styles;
            if (styles == null)
                throw new InvalidOperationException();

            if (!_IndexedCellStyleArray.TryGetValue(name, out ICellStyle style))
            {
                var config = styles.GetCellStyle(name); // 没有找到将会抛出异常.


                //var color = (font != null && font.Color != null) ? options.GetColor(font.Color) : null;

                style = _Workbook.CreateCellStyle();

                if (config.HasDataFormat)
                {
                    style.DataFormat = _Workbook.GetOrAddDataFormat(config.DataFormat);
                }

                if (config.HasTextAlignment)
                {
                    var alignment = styles.GetCellAlignment(config.TextAlignment);
                    style.Alignment = alignment.Alignment;
                    style.Indention = alignment.Indention;
                    style.VerticalAlignment = alignment.VerticalAlignment;
                    style.WrapText = alignment.WrapText;
                    style.ShrinkToFit = alignment.ShrinkToFit;
                    style.Rotation = alignment.Rotation;
                }

                if (config.HasFont)
                {
                    var font = GetOrAddFont(config.Font);
                    style.SetFont(font);
                }

                if (config.HasBorder)
                {
                    var border = styles.GetCellBorder(config.Border);

                    style.SetBorders(BorderEdges.Left, border.LeftStyle, border.LeftColor);
                    style.SetBorders(BorderEdges.Right, border.RightStyle, border.RightColor);
                    style.SetBorders(BorderEdges.Top, border.TopStyle, border.TopColor);
                    style.SetBorders(BorderEdges.Bottom, border.BottomStyle, border.BottomColor);

                    var diagonal = (BorderEdges)((int)border.Diagonal << 0x08);
                    style.SetBorders(diagonal, border.DiagonalStyle, border.DiagonalColor);
                }

                if (config.HasFill)
                {
                    var fill = Styles.GetCellFill(config.Fill);
                    style.SetFill(fill.FillPattern, fill.FillForegroundColor, fill.FillBackgroundColor);
                }

                _IndexedCellStyleArray.Add(name, style);
            }
            return style;
        }

        IDictionary<string, IFont> _IndexedFontArray;
        protected virtual IFont GetOrAddFont(string name)
        {
            var styles = this.Styles;
            if (styles == null)
                throw new InvalidOperationException();

            if (!_IndexedFontArray.TryGetValue(name, out IFont font))
            {
                var options = styles.GetFont(name);
                if (options == null)
                    throw new InvalidOperationException();

                font = _Workbook.GetOrAddFont(options.FontName, options.FontHeightInPoints,
                    options.IsBold, options.IsItalic, options.Underline, options.IsStrikeout, options.TypeOffset,
                    options.Color);
                _IndexedFontArray.Add(name, font);
            }

            return font;
        }

        //IDictionary<string, IColor> _IndexedColorArray;
        //protected virtual IFont GetOrAddColor(string name)
        //{
        //    var styles = this.Styles;
        //    if (styles == null)
        //        throw new InvalidOperationException();

        //    if (!_IndexedColorArray.TryGetValue(name,out IColor color))
        //    {

        //    }
        //}
    }
}
