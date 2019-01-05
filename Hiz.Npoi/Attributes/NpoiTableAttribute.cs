using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class NpoiTableAttribute : NpoiAttribute
    {
        /// <summary>
        /// 默认行高; 0: 不作修改; Default: 0; (仅限导出有效)
        /// </summary>
        public float RowDefaultHeight { get; set; }

        /// <summary>
        /// 默认列宽; 0: 不作修改; Default: 0; (仅限导出有效)
        /// </summary>
        public float ColumnDefaultWidth { get; set; }

        /// <summary>
        /// 默认的单元格样式 (仅限导出有效)
        /// </summary>
        public string CellDefaultStyle { get; set; }

        /// <summary>
        /// 是否显示列头; Default: true;
        /// </summary>
        public bool HeaderVisible { get; set; } = true;

        /// <summary>
        /// 列头高度; 如果没有设置, 则将使用 this.RowHeight; (仅限导出有效)
        /// </summary>
        public float HeaderHeight { get; set; }

        /// <summary>
        /// 列头默认样式; 如果没有设置, 则将使用 this.CellDefaultStyle; (仅限导出有效)
        /// </summary>
        public string HeaderDefaultStyle { get; set; }
    }
}
