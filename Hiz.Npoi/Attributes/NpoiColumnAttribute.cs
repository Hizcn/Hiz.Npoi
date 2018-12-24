using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class NpoiColumnAttribute : Attribute
    {
        /// <summary>
        /// 列头文本; 如果没有设置, 则将使用属性名称;
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// 列宽; 覆盖 NpoiTableAttribute.ColumnWidth; (仅限导出)
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// 列头注释 (仅限导出)
        /// </summary>
        public string HeaderComment { get; set; }

        /// <summary>
        /// 列头样式; 覆盖 NpoiTableAttribute.DefaultHeaderStyle; (仅限导出)
        /// </summary>
        public string HeaderStyle { get; set; }

        /// <summary>
        /// 单元格的样式; 覆盖 NpoiTableAttribute.DefaultCellStyle; (仅限导出)
        /// </summary>
        public string CellStyle { get; set; }
    }
}