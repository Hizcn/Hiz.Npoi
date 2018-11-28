using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;

namespace Hiz.Npoi
{
    partial class NpoiExtensions
    {
        #region IWorkbook

        public static void Write(this IWorkbook workbook, string path)
        {
            if (workbook == null)
                throw new ArgumentNullException();

            using (var stream = File.Create(path))
            {
                workbook.Write(stream);
            }
        }

        #endregion

        #region IWorkbook/DataFormat~

        /* IWorkbook 数据格式相关 (一个成员):
         * IDataFormat CreateDataFormat();
         * 
         * XSSFWorkbook.CreateDataFormat() {
         *     if (this.formatter == null) {
         *         this.formatter = new XSSFDataFormat(this.stylesSource);
         *     }
         *     return this.formatter;
         * }
         * 
         * short XSSFDataFormat.GetFormat(string format) {
         *     var index = BuiltinFormats.GetBuiltinFormat(format);
         *     if (index == -1)
         *         index = this.stylesSource.PutNumberFormat(format);
         *     return index;
         * }
         * 
         * 
         * HSSFWorkbook.CreateDataFormat() {
         *     if (this.formatter == null) {
         *         this.formatter = new HSSFDataFormat(this.workbook);
         *     }
         *     return this.formatter;
         * }
         * 
         * // 存在修改内建格式情况; 比如本地货币格式: 将美元符号改为人民币符号.
         * // 构造函数添加 InternalWorkbook 中修改的(内建)格式.
         * HSSFDataFormat(InternalWorkbook workbook) {
         *     foreach(var record in workbook.Formats) {
         *         while (this.formats.Count <= record.IndexCode)
         *             this.formats.Add(null);
         *         this.formats[record.IndexCode] = record.FormatString;
         *     }
         * }
         * short HSSFDataFormat.GetFormat(string format) {
         *     if (format.ToUpper().Equals("TEXT"))
         *         format = "@";
         *     
         *     // 如果尚未添加内建格式, 那么复制进来.
         *     if (!this.movedBuiltins) {
         *         var builtin = BuiltinFormats.GetAll();
         *         for(var i = 0; i < builtin.Length; i++)
         *             this.formats[i] = builtin[i]; // v2.2.1.0 BUG: 此处将会覆盖构造函数添加的修改的内建格式(InternalWorkbook).
         *         this.movedBuiltins = true;
         *     }
         *     
         *     // 查找当前已经存在格式 (内建格式 以及 用户添加格式);
         *     for(var i = 0; i < this.formats.Length; i++)
         *         if (format.Equals(this.formats[i]))
         *             return i;
         *     
         *     // 调用 InternalWorkbook 创建格式返回索引, 起始用户索引 164.
         *     var index = (int)this.workbook.GetFormat(text, true);
         *     
         *     // 索引 50-163 之间填充空值.
         *     while (this.formats.Count <= index)
         *         this.formats.Add(null);
         *     
         *     // 记录新增用户格式以及索引..
         *     this.formats[index] = text;
         *     return (short)index;
         * }
         * short InternalWorkbook.GetFormat(string format, bool CreateIfNotFound) {
         *     foreach(var record in this.formats)
         *         if (record.FormatString.Equals(format))
         *             return (short)record.IndexCode;
         *     if (CreateIfNotFound)
         *         return (short)this.CreateFormat(format);
         *     return -1;
         * }
         */

        /// <summary>
        /// 查找或者新增格式, 然后返回格式索引.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="format">数据格式</param>
        /// <returns>DataFormat.Index</returns>
        public static short GetOrAddDataFormat(this IWorkbook workbook, string format)
        {
            // HSSFWorkbook/XSSFWorkbook: 多次调用只会返回同个实例;
            var formatter = workbook.CreateDataFormat();

            // 内建格式: NPOI.SS.UserModel.BuiltinFormats;
            // 用户自定格式起始索引: HSSFWorkbook = 164; XSSFWorkbook = 165;
            // 优先查找内嵌格式, 接着查找用户自定格式, 如果均未命中, 新增返回;
            return formatter.GetFormat(format);
        }

        #endregion

        #region IWorkbook/Sheet~

        /* IWorkbook 表单相关:
         * 
         * int NumberOfSheets { get; } // 获取表单数量.
         * 
         * ISheet GetSheetAt(int index); // 获取指定索引表单.
         * ISheet GetSheet(string name); // 获取指定名称表单.
         * int GetSheetIndex(string name); // 获取指定名称表单索引.
         * int GetSheetIndex(ISheet sheet); // 获取指定表单索引.
         * 
         * string GetSheetName(int sheet); // 获取指定索引表单名称.
         * void SetSheetName(int sheet, string name); // 设置指定索引表单名称.
         * void SetSheetOrder(string sheetname, int pos); // 设置指定名称表单顺序.
         * 
         * ISheet CreateSheet(); // 创建默认名称表单.
         * ISheet CreateSheet(string sheetname); // 创建指定名称表单.
         * void RemoveSheetAt(int index); // 移除指定索引表单.
         * 
         * ISheet CloneSheet(int sheetNum); // 克隆指定索引表单
         * 
         * // 隐藏表单
         * bool IsSheetHidden(int sheetIx);
         * bool IsSheetVeryHidden(int sheetIx);
         * void SetSheetHidden(int sheetIx, int hidden);
         * void SetSheetHidden(int sheetIx, SheetState hidden);
         * 
         * // 激活表单
         * int ActiveSheetIndex { get; }
         * void SetActiveSheet(int sheetIndex);
         * 
         * // 打印区域
         * string GetPrintArea(int sheetIndex);
         * void SetPrintArea(int sheetIndex, string reference);
         * void SetPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);
         * void RemovePrintArea(int sheetIndex);
         * // 打印标题
         * void SetRepeatingRowsAndColumns(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow);
         * 
         * ISheet
         * |=> NPOI.HSSF.UserModel.HSSFSheet
         * |=> NPOI.XSSF.UserModel.XSSFSheet
         *     |=> NPOI.XSSF.UserModel.XSSFChartSheet
         *     |=> NPOI.XSSF.UserModel.XSSFDialogsheet
         */

        public static IEnumerable<ISheet> GetSheets(this IWorkbook workbook)
        {
            var count = workbook.NumberOfSheets;
            for (var i = 0; i < count; i++)
                yield return workbook.GetSheetAt(i);
        }

        /// <summary>
        /// 获取或添加指定名称的标签
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ISheet GetOrAddSheet(this IWorkbook workbook, string name)
        {
            var sheet = workbook.GetSheet(name);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(name);
            }
            return sheet;
        }

        #endregion
    }
}
