﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Hiz.Npoi
{
    [DebuggerDisplay("{PropertyName}")]
    class NpoiPropertyDescriptor<T>
    {
        public string PropertyName { get; set; } // 属性名称

        public Type PropertyType { get; set; } // 属性类型

        public Func<T, object> Getter { get; set; } // 属性取值

        public Action<T, object> Setter { get; set; } // 属性存值

        // 导入时 用于匹配
        public string HeaderText { get; set; } // 列头文本

        public int SourceOrder { get; set; } // 原始属性顺序(反射出现顺序) [ 0, +N]
        public int ColumnOrder { get; set; } // 导出排序权重 [-N, +N] 从小到大排序 支持负数 // 通过配置指定
        public int ActualOrder { get; set; } // 导入实际索引 [ 0, +N]

        public string HeaderComment { get; set; } // 列头注释

        public CellType CellType { get; set; }

        public float Width { get; set; }

        public string CellStyle { get; set; }
        public string HeaderStyle { get; set; }

    }
}
