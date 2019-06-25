using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Hiz.Npoi
{
    /* v1.1 2018-12-07:
     * 改进: OpenRead() 增加参数 relaxed.
     * 
     * v1.0
     */
    public static class Xpoi
    {
        /// <summary>
        /// 只读方式打开 Excel 文档.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="relaxed">是否允许其它程序正在修改文档</param>
        /// <returns></returns>
        public static IWorkbook OpenRead(string path, bool relaxed = true)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException(nameof(path));

            var share = relaxed ? FileShare.ReadWrite : FileShare.Read;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, share))
            {
                return WorkbookFactory.Create(stream);
            }
        }

        #region 集合导出

        public static IWorkbook ExportMany<T>(IEnumerable<T> datas, RuntimeOptions options, CancellationToken cancel = default(CancellationToken), IProgress<int> progress = null)
        {
            if (options == null)
                throw new ArgumentNullException();

            if (options.FileFormat == OfficeArchiveFormat.None)
            {
                var hssf = string.Equals(Path.GetExtension(options.FilePath), ".xls", StringComparison.OrdinalIgnoreCase) ? true : false;

                options.FileFormat = hssf ? OfficeArchiveFormat.Binary : OfficeArchiveFormat.OpenXml;
            }

            var service = new ExcelService();
            var workbook = service.ExportMany<T>(datas, options, cancel, progress);
            workbook.Write(options.FilePath);

            return workbook;
        }

        #endregion


        #region 集合导入

        // 导入: IWorkbook => IEnumerable<T>
        //public virtual IEnumerable<T> ImportMany<T>(IWorkbook workbook, RuntimeOptions options, CancellationToken cancel = default(CancellationToken), IProgress<int> progress = null)
        //{
        //}
        #endregion
    }
}
