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

        //public virtual IWorkbook ExportMany<T>(IEnumerable<T> datas, ExcelOptions options, CancellationToken cancel = default(CancellationToken), IProgress<int> progress = null)
        //{
        //}

        #endregion
    }
}
