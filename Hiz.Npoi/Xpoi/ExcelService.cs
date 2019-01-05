using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace Hiz.Npoi
{
    class ExcelService
    {
        // 默认报告频率
        const int DefaultReportFrequency = 10000;

        // 报告频率
        int _ReportFrequency = DefaultReportFrequency;

        AnnotationProvider _AnnotationService = new NpoiAnnotationProvider();

        #region 集合导出

        // 导出: IEnumerable<T> => IWorkbook
        public virtual IWorkbook ExportMany<T>(IEnumerable<T> datas, RuntimeOptions options, CancellationToken cancel = default(CancellationToken), IProgress<int> progress = null)
        {
            if (datas == null)
                throw new ArgumentNullException();

            // 获取数据注解;
            var descriptor = _AnnotationService.GetExport<T>();
            var annotations = descriptor.Properties;

            // 创建文档实例;
            IWorkbook workbook = options.FileFormat == OfficeArchiveFormat.Binary ? (IWorkbook)new HSSFWorkbook() : (IWorkbook)new XSSFWorkbook();

            // 创建实例的上下文;
            var context = new ExcelContext(workbook, options.ExcelStyles);

            // 更改默认字体;
            var font = options.ExcelStyles.DefaultFont;
            if (font != null)
                workbook.SetDefaultFont(font.FontName, font.FontHeightInPoints, font.IsBold, font.IsItalic, font.Underline, font.IsStrikeout, font.TypeOffset, 0);

            // 添加样式资源;
            //foreach (var a in annotations)
            //    if (!string.IsNullOrEmpty(a.CellStyleName))
            //        a.CellStyleIndexed = context.GetCellStyle(a.CellStyleName);

            // 当前表单;
            ISheet sheet = null;

            // 报告频率;
            var frequency = this._ReportFrequency;
            // 每个频率计数;
            var p = 0;
            // 数据集合索引;
            var i = 0;

            // 数据区域允许最大行数 (不含区域表尾) // 包含: 表单标题/区域表头
            var gridRowMax = context.MaxRows;
            //if (options.HasGridFooter)
            //    gridRowMax--;

            var gridRowFirst = 0; // 数据区域首行索引
            var r = 0; // 当前表单行的索引

            // 数据区域(表格)
            foreach (var d in datas)
            {
                if (sheet == null)
                {
                    sheet = CreateSheet<T>(workbook, options, annotations);
                    r = sheet.LastRowNum;
                    r++;

                    if (descriptor.HeaderVisible)
                    {
                        CreateGridHeader(sheet, r, descriptor);
                        r++;
                    }

                    gridRowFirst = r;
                }

                var row = sheet.CreateRow(r++);
                for (var c = 0; c < annotations.Count; c++)
                {
                    var cell = row.CreateCell(c);

                    var t = annotations[c];
                    var value = t.Getter(d);
                    if (value != null)
                    {
                        cell.SetCellValue(value);
                    }
                    if (!string.IsNullOrEmpty(t.CellStyle))
                    {
                        cell.CellStyle = context.GetCellStyle(t.CellStyle);
                    }
                }

                i++;
                if (++p == frequency) // 求余性能较差, 因此此处使用递增比较
                {
                    // 每一万条检测一次取消指令, 以及报告进度一次...
                    cancel.ThrowIfCancellationRequested();
                    if (progress != null)
                        progress.Report(i);
                    p = 0;
                }

                if (r == gridRowMax)
                {
                    //if (options.HasGridFooter)
                    //{
                    //    CreateGridFooter(sheet, r);
                    //}

                    EndOfSheet(sheet, annotations, gridRowFirst, r - gridRowFirst);
                    sheet = null;
                    r = 0;
                }
            }
            if (sheet != null)
            {
                //if (options.HasGridFooter)
                //{
                //    CreateGridFooter(sheet, r);
                //}

                EndOfSheet(sheet, annotations, gridRowFirst, r - gridRowFirst);
            }

            return workbook;
        }

        // 创建表单
        ISheet CreateSheet<T>(IWorkbook workbook, RuntimeOptions options, IList<NpoiPropertyDescriptor<T>> annotations)
        {
            ISheet sheet = workbook.CreateSheet();
            var r = 0; // 当前表单 行的索引

            // 创建表单标题
            if (options.HasTitle)
                CreateTitleRow(sheet, r++, annotations.Count, options.Title);

            //// 创建(表格)表头
            //if (options.HasGridHeader)
            //    CreateGridHeader(sheet, r++, annotations);

            return sheet;
        }

        // 创建标题
        void CreateTitleRow(ISheet sheet, int rowIndex, int columnSpan, string title)
        {
            var columnIndex = 0;

            var row = sheet.CreateRow(rowIndex);
            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(title);

            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, columnIndex, columnIndex + columnSpan - 1));
        }

        // 创建表头
        void CreateGridHeader<T>(ISheet sheet, int rowIndex, NpoiTypeDescriptor<T> descriptor)
        {
            // 表头
            var header = sheet.CreateRow(rowIndex);
            var annotations = descriptor.Properties;

            for (var i = 0; i < annotations.Count; i++)
            {
                var a = annotations[i];

                var cell = header.CreateCell(i);

                cell.SetCellValue(a.GetActualColumnHeader());

                // 如果指定手动设置列宽
                if (a.Width > 0f)
                    sheet.SetColumnWidthInCharacters(i, a.Width);
            }

            // 固定表头
            sheet.CreateFreezePane(0, rowIndex);
        }

        // TODO: 创建表尾
        void CreateGridFooter(ISheet sheet, int rowIndex)
        {

        }

        // 完结表单
        void EndOfSheet<T>(ISheet sheet, IList<NpoiPropertyDescriptor<T>> annotations, int gridRowFirst, int gridRowCount)
        {
            // 自动大小表列
            //for (var i = 0; i < annotations.Count; i++)
            //{
            //    if (annotations[i].ColumnWidth < 0)
            //        sheet.AutoSizeColumn(i, gridRowFirst, gridRowFirst + gridRowCount);
            //}
        }

        #endregion
    }
}
