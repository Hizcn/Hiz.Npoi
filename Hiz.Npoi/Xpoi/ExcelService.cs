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
                    sheet = CreateSheet<T>(workbook, options, descriptor);
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

                    EndOfSheet(sheet, descriptor, gridRowFirst, r - gridRowFirst);
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

                EndOfSheet(sheet, descriptor, gridRowFirst, r - gridRowFirst);
            }

            return workbook;
        }

        // 创建表单
        ISheet CreateSheet<T>(IWorkbook workbook, RuntimeOptions options, NpoiTypeDescriptor<T> descriptor)
        {
            ISheet sheet = workbook.CreateSheet();
            var r = 0; // 当前表单 行的索引

            var annotations = descriptor.Properties;

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
        void EndOfSheet<T>(ISheet sheet, NpoiTypeDescriptor<T> descriptor, int gridRowFirst, int gridRowCount)
        {
            // 自动大小表列
            //for (var i = 0; i < annotations.Count; i++)
            //{
            //    if (annotations[i].ColumnWidth < 0)
            //        sheet.AutoSizeColumn(i, gridRowFirst, gridRowFirst + gridRowCount);
            //}
        }

        #endregion

        #region 集合导入

        // 导入: IWorkbook => IEnumerable<T>
        public virtual IEnumerable<T> ImportMany<T>(IWorkbook workbook, ExcelOptions options, CancellationToken cancel = default(CancellationToken), IProgress<int> progress = null) where T : new()
        {
            //IWorkbook workbook;
            //using (var stream = File.OpenRead(path))
            //{
            //    workbook = WorkbookFactory.Create(stream);
            //}

            cancel.ThrowIfCancellationRequested();

            // 获取数据注解;
            var root = _AnnotationService.GetImport<T>();

            // 导入错误信息
            var errors = new List<ExcelCellError>();

            var k = 0;
            var p = 0;
            for (int s = 0; s < workbook.NumberOfSheets; s++)
            {
                // 当前表单;
                var sheet = workbook.GetSheetAt(s);

                var r = 0; // 当前行的索引;

                IList<NpoiPropertyDescriptor<T>> annotations = null;
                if (root.HeaderVisible)
                {
                    var headers = LoadGridHeaders(sheet, r++);

                    IList<string> requires;
                    IList<string> optionals;
                    IList<string> surpluses;
                    annotations = root.TryCheckGridHeader(headers, out requires, out optionals, out surpluses);

                    if (requires.Count > 0)
                    {
                        var message = new StringBuilder();
                        message.AppendLine("缺少以下字段: ");
                        foreach (var q in requires)
                        {
                            message.AppendLine(q);
                        }

                        errors.Add(new ExcelCellError()
                        {
                            SheetIndex = s,
                            SheetName = sheet.SheetName,
                            ErrorMessage = message.ToString(),
                        });

                        continue; // 下一表单;
                    }

                    annotations = root.MakeArrayWithOrder(annotations).ToList();
                }
                else
                {
                    annotations = root.Properties;
                }

                var first = r;
                var last = sheet.LastRowNum;
                var count = annotations.Count;
                var cells = new ICell[count];

                while (r <= last)
                {
                    var row = sheet.GetRow(r);
                    if (row != null)
                    {
                        // 一次性获取当前行的所有单元格.
                        for (var i = 0; i < count; i++)
                            cells[i] = row.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK);

                        if (cells.Any(c => c != null))
                        {
                            var data = Activator.CreateInstance<T>(); // 创建实例

                            var had = 0; // 有值数量
                            var missed = 0; // 必需字段缺少数量
                            for (var c = 0; c < count; c++)
                            {
                                var t = annotations[c];

                                var cell = cells[c];
                                object value = null;
                                if (cell != null)
                                {
                                    try
                                    {
                                        value = cell.GetCellValue(t.PropertyType);
                                    }
                                    catch (Exception exception)
                                    {
                                        errors.Add(new ExcelCellError()
                                        {
                                            SheetIndex = s,
                                            RowIndex = r,
                                            ColumnIndex = c,
                                            ErrorMessage = exception.Message,
                                        });
                                    }
                                }

                                if (value != null)
                                {
                                    t.Setter(data, value);
                                    had++;
                                }
                                else if (t.Required)
                                {
                                    missed++;
                                    errors.Add(new ExcelCellError()
                                    {
                                        SheetIndex = s,
                                        RowIndex = r,
                                        ColumnIndex = c,
                                        ErrorMessage = "缺少必需字段",
                                    });
                                }
                            }

                            // 只要没有缺失任何必需字段, 并且至少有一个单元格有值, 那就输出;
                            if (missed == 0 && had > 0)
                                yield return data;

                            //k++;
                            //if (++p == 10000) // 求余性能较差
                            //{
                            //    // 每一万条检测一次取消指令, 以及报告进度一次...
                            //    cancel.ThrowIfCancellationRequested();
                            //    if (progress != null)
                            //        progress.Report(k);
                            //    p = 0;
                            //}
                        }
                    }
                    r++;
                }
            }
        }
        // 读取表头
        KeyValuePair<int, string>[] LoadGridHeaders(ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row != null)
                return row.GetCells().Select(c => new KeyValuePair<int, string>(c.ColumnIndex, c.GetCellValue<string>())).ToArray();
            return null;
        }
        #endregion
    }
}
