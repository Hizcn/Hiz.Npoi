using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Hiz.Reflection;
using Hiz.Npoi.Attributes;

namespace Hiz.Npoi
{
    abstract class AnnotationProvider
    {
        [Flags]
        protected enum AccessorMode { None = 0, Getter = 0x01, Setter = 0x02, Both = Getter | Setter };

        // 导出
        public virtual NpoiTypeDescriptor<T> GetExport<T>(string group = null)
        {
            return InternalGetTypeAnnotation<T>(AccessorMode.Getter, group);
        }

        // 导入
        public virtual NpoiTypeDescriptor<T> GetImport<T>(string group = null)
        {
            return InternalGetTypeAnnotation<T>(AccessorMode.Setter, group);
        }

        protected virtual NpoiTypeDescriptor<T> InternalGetTypeAnnotation<T>(AccessorMode mode, string group)
        {
            var annotation = new NpoiTypeDescriptor<T>();

            // 重新排序 (最终输出顺序)
            var array = InternalGetProperties<T>(mode, group).OrderBy(i => i.ColumnOrder).ThenBy(i => i.SourceOrder).ToArray();

            var properties = annotation.Properties;
            foreach (var i in array)
                properties.Add(i);

            return annotation;
        }

        protected abstract IEnumerable<NpoiPropertyDescriptor<T>> InternalGetProperties<T>(AccessorMode mode, string group);
    }

    class NpoiAnnotationProvider : AnnotationProvider
    {
        static NpoiAnnotationProvider _Default = null;
        public static NpoiAnnotationProvider Default
        {
            get
            {
                if (_Default == null)
                    _Default = new NpoiAnnotationProvider();
                return _Default;
            }
        }

        /* NpoiTableAttribute
         * NpoiColumnAttribute
         */
        protected override IEnumerable<NpoiPropertyDescriptor<T>> InternalGetProperties<T>(AccessorMode mode, string group)
        {
            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var s = 0; // 反射属性出现顺序
            foreach (var p in properties)
            {
                if ((mode & AccessorMode.Getter) != AccessorMode.None && !p.CanRead)
                    continue;
                if ((mode & AccessorMode.Setter) != AccessorMode.None && !p.CanWrite)
                    continue;

                // 获取所有特性
                var attributes = (NpoiAttribute[])p.GetCustomAttributes(typeof(NpoiAttribute), true);

                // 是否忽略属性
                var ignore = attributes.OfType<NpoiIgnoreAttribute>().FirstOrDefault();
                if (ignore != null)
                    continue;

                var a = new NpoiPropertyDescriptor<T>();
                a.PropertyName = p.Name;
                a.PropertyType = p.PropertyType;
                a.SourceOrder = s++;

                var column = attributes.OfType<NpoiColumnAttribute>().FirstOrDefault();
                if (column != null)
                {
                    a.HeaderText = column.HeaderText;
                    a.HeaderComment = column.HeaderComment;
                    a.ColumnOrder = column.ColumnOrder;
                    a.CellType = column.CellType;
                    a.Width = column.Width;

                    a.HeaderStyle = column.HeaderStyle;
                    a.CellStyle = column.CellStyle;
                }

                if ((mode & AccessorMode.Getter) != AccessorMode.None)
                {
                    // 导出
                    a.Getter = p.MakeGetter<T, object>();
                }
                if ((mode & AccessorMode.Setter) != AccessorMode.None)
                {
                    // 导入
                    a.Setter = p.MakeSetter<T, object>();
                }

                yield return a;
            }
        }

        protected override NpoiTypeDescriptor<T> InternalGetTypeAnnotation<T>(AccessorMode mode, string group)
        {
            var descriptor = base.InternalGetTypeAnnotation<T>(mode, group);

            var type = typeof(T);

            return descriptor;
        }
    }

    #region System.ComponentModel.DataAnnotations
    //class DataAnnotationProvider : AnnotationProvider
    //{
    //    /* System.ComponentModel.DataAnnotations.DisplayAttribute
    //     * System.ComponentModel.DataAnnotations.RequiredAttribute
    //     */
    //    protected override IEnumerable<NpoiPropertyDescriptor<T>> InternalGetProperties<T>(AccessorMode mode, string group)
    //    {
    //        var type = typeof(T);
    //        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
    //        var s = 0; // 反射属性出现顺序
    //        foreach (var p in properties)
    //        {
    //            if ((mode & AccessorMode.Getter) != AccessorMode.None && !p.CanRead)
    //                continue;
    //            if ((mode & AccessorMode.Setter) != AccessorMode.None && !p.CanWrite)
    //                continue;

    //            // 获取所有特性
    //            var attributes = (Attribute[])p.GetCustomAttributes(typeof(Attribute), true);

    //            var display = attributes.OfType<DisplayAttribute>().FirstOrDefault();
    //            if (display != null)
    //            {
    //                // 如果显式指定自动生成为否, 则将跳过;
    //                if (display.GetAutoGenerateField() == false)
    //                {
    //                    continue;
    //                }
    //                // 如果指定分组标识, 并且不等, 则将跳过;
    //                if (!string.IsNullOrEmpty(group))
    //                {
    //                    /* if (group) == null; 返回全部;
    //                     * if (group) != null; 返回: GroupName == null 或者 GroupName == group;
    //                     */
    //                    var g = display.GetGroupName();
    //                    if (!string.IsNullOrEmpty(g) && group != g)
    //                        continue;
    //                }
    //            }

    //            var a = new NpoiPropertyDescriptor<T>();
    //            a.PropertyName = p.Name;
    //            a.PropertyType = p.PropertyType;
    //            a.SourceOrder = s++;

    //            var required = attributes.OfType<RequiredAttribute>().FirstOrDefault();
    //            if (required != null)
    //            {
    //                a.Required = true;
    //                a.AllowEmptyStrings = required.AllowEmptyStrings;
    //            }
    //            if (display != null)
    //            {
    //                a.ColumnHeader = display.GetName();
    //                a.ColumnOrder = display.GetOrder() ?? 0;
    //            }

    //            if ((mode & AccessorMode.Getter) != AccessorMode.None)
    //            {
    //                a.Getter = p.MakeGetter<T, object>();
    //            }
    //            if ((mode & AccessorMode.Setter) != AccessorMode.None)
    //            {
    //                a.Setter = p.MakeSetter<T, object>();
    //            }

    //            yield return a;
    //        }
    //    }
    //}
    #endregion

    #region System.ComponentModel
    //class DescriptorAnnotationProvider : AnnotationProvider
    //{
    //    /* System.ComponentModel.BrowsableAttribute
    //     * System.ComponentModel.CategoryAttribute
    //     * System.ComponentModel.DisplayNameAttribute
    //     */
    //    protected override IEnumerable<NpoiPropertyDescriptor<T>> InternalGetProperties<T>(AccessorMode mode, string group)
    //    {
    //        var type = typeof(T);
    //        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
    //        var s = 0; // 反射属性出现顺序
    //        foreach (var p in properties)
    //        {
    //            if ((mode & AccessorMode.Getter) != AccessorMode.None && !p.CanRead)
    //                continue;
    //            if ((mode & AccessorMode.Setter) != AccessorMode.None && !p.CanWrite)
    //                continue;

    //            // 获取所有特性
    //            var attributes = (Attribute[])p.GetCustomAttributes(typeof(Attribute), true);

    //            var browsable = attributes.OfType<BrowsableAttribute>().FirstOrDefault();
    //            if (browsable != null && !browsable.Browsable)
    //            {
    //                // 如果隐藏
    //                continue;
    //            }
    //            if (!string.IsNullOrEmpty(group))
    //            {
    //                var category = attributes.OfType<CategoryAttribute>().FirstOrDefault();
    //                if (!string.IsNullOrEmpty(category.Category) && group != category.Category)
    //                    continue;
    //            }

    //            var a = new NpoiPropertyDescriptor<T>();
    //            a.PropertyName = p.Name;
    //            a.PropertyType = p.PropertyType;
    //            a.SourceOrder = s++;

    //            var display = attributes.OfType<DisplayNameAttribute>().FirstOrDefault();
    //            if (display != null)
    //            {
    //                a.ColumnHeader = display.DisplayName;
    //            }

    //            if ((mode & AccessorMode.Getter) != AccessorMode.None)
    //            {
    //                a.Getter = p.MakeGetter<T, object>();
    //            }
    //            if ((mode & AccessorMode.Setter) != AccessorMode.None)
    //            {
    //                a.Setter = p.MakeSetter<T, object>();
    //            }

    //            yield return a;
    //        }
    //    }
    //}
    #endregion
}
