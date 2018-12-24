using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    [DebuggerDisplay("{PropertyName}")]
    class NpoiPropertyDescriptor<T>
    {
        public string PropertyName { get; set; } // 属性名称

        public Type PropertyType { get; set; } // 属性类型

        public Func<T, object> Getter { get; set; } // 属性取值

        public Action<T, object> Setter { get; set; } // 属性存值
    }
}
