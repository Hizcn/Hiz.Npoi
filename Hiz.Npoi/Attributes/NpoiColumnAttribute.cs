using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class NpoiColumnAttribute : NpoiAttribute
    {
        // public Type ResourceType { get; set; }

        /// <summary>
        /// 列头文本; 如果没有设置, 并且需要显示列头, 则将使用属性名称;
        /// 
        /// 导出: 用于输出表头文本;
        /// 导入: 用于匹配;
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// 列头注释 (仅限导出有效)
        /// </summary>
        public string HeaderComment { get; set; }

        /// <summary>
        /// 列宽; 覆盖 NpoiTableAttribute.ColumnDefaultWidth; (仅限导出有效)
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// 列头样式; 覆盖 NpoiTableAttribute.DefaultHeaderStyle; (仅限导出有效)
        /// </summary>
        public string HeaderStyle { get; set; }

        /// <summary>
        /// 单元格的样式; 覆盖 NpoiTableAttribute.DefaultCellStyle; (仅限导出有效)
        /// </summary>
        public string CellStyle { get; set; }

        CellType _CellType = CellType.Unknown;
        /// <summary>
        /// 重载单元格值类型; Unknown: 自动;
        /// </summary>
        public CellType CellType { get { return _CellType; } set { _CellType = value; } }

        /// <summary>
        /// 排序权重; 重载列的位置; (仅限导出有效)
        /// </summary>
        public int ColumnOrder { get; set; }
    }
}