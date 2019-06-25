using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    class NpoiTypeDescriptor<T> : ITableOptions
    {
        IList<NpoiPropertyDescriptor<T>> _Properties = null;
        public IList<NpoiPropertyDescriptor<T>> Properties
        {
            get
            {
                if (_Properties == null)
                    _Properties = new List<NpoiPropertyDescriptor<T>>();
                return this._Properties;
            }
        }

        public float RowDefaultHeight { get; set; }
        public float ColumnDefaultWidth { get; set; }
        public string CellDefaultStyle { get; set; }
        public bool HeaderVisible { get; set; } = true;
        public float HeaderHeight { get; set; }
        public string HeaderDefaultStyle { get; set; }

        /* 根据实际顺序重新生成属性集合
         * 集合数量 = 最大索引加一
         * 例如: 结果集合: { 1, 3, 7 }
         * 生成:
         * var array = new T[8];
         * array[1] = ...;
         * array[3] = ...;
         * array[7] = ...;
         * 
         * 其它为空;
         */
        internal NpoiPropertyDescriptor<T>[] MakeArrayWithOrder(IEnumerable<NpoiPropertyDescriptor<T>> annotations)
        {
            var max = annotations.Max(a => a.ActualOrder);
            var array = new NpoiPropertyDescriptor<T>[++max];

            foreach (var a in annotations)
                array[a.ActualOrder] = a;

            return array;
        }

        IDictionary<string, string> _Maps = null;
        /// <summary>
        /// 存在表头的情况下, 检查表头字段集合是否满足条件;
        /// </summary>
        /// <param name="headers">表头文本</param>
        /// <param name="requires">缺失的必备字段集</param>
        /// <param name="optionals">缺失的可选字段集</param>
        /// <param name="surpluses">多余的额外字段集</param>
        /// <returns></returns>
        internal IList<NpoiPropertyDescriptor<T>> TryCheckGridHeader(IEnumerable<KeyValuePair<int, string>> headers, out IList<string> requires, out IList<string> optionals, out IList<string> surpluses)
        {
            if (_Maps == null)
            {
                _Maps = new Dictionary<string, string>();
                foreach (var p in this.Properties)
                    _Maps.Add(p.GetActualColumnHeader(), p.PropertyName);
            }

            var properties = this._Properties;
            var all = properties.Select(p => new { p.PropertyName, p.Required }).ToList();

            var matches = new List<NpoiPropertyDescriptor<T>>(); // 匹配
            surpluses = new List<string>(); // 多余
            foreach (var pair in headers)
            {
                var text = pair.Value; // 表头文本
                string value;
                if (_Maps.TryGetValue(text, out value))
                {
                    var a = properties.Where(t => t.PropertyName == value).SingleOrDefault();
                    a.ActualOrder = pair.Key; // 更新实际位置;

                    matches.Add(a); // 添加匹配
                    all.RemoveAll(m => m.PropertyName == value);
                }
                else
                {
                    surpluses.Add(text); // 记录多余
                }
            }

            requires = all.Where(m => m.Required).Select(m => m.PropertyName).ToArray();
            optionals = all.Where(m => !m.Required).Select(m => m.PropertyName).ToArray();

            //


            return matches;
        }
    }
}
