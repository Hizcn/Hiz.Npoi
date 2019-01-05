using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Hiz.Reflection
{
    static class FastHelper
    {
        static readonly ReflectionWithExpression _Service = new ReflectionWithExpression();

        #region Property.Predefined/T1-T2

        /// <summary>
        /// 用于静态属性 (支持 实例/属性 类型转换)
        /// </summary>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="member"></param>
        /// <returns></returns>
        public static Func<TProperty> MakeGetter<TProperty>(this PropertyInfo member)
        {
            return _Service.MakeGetter<TProperty>(member);
        }

        /// <summary>
        /// 用于实例属性 (支持 实例/属性 类型转换)
        /// </summary>
        /// <typeparam name="TInstance"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="member"></param>
        /// <returns></returns>
        public static Func<TInstance, TProperty> MakeGetter<TInstance, TProperty>(this PropertyInfo member)
        {
            return _Service.MakeGetter<TInstance, TProperty>(member);
        }

        /// <summary>
        /// 用于静态属性 (支持 实例/属性 类型转换)
        /// </summary>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="member"></param>
        /// <returns></returns>
        public static Action<TProperty> MakeSetter<TProperty>(this PropertyInfo member)
        {
            return _Service.MakeSetter<TProperty>(member);
        }

        /// <summary>
        /// 用于实例属性 (支持 实例/属性 类型转换)
        /// </summary>
        /// <typeparam name="TInstance"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="member"></param>
        /// <returns></returns>
        public static Action<TInstance, TProperty> MakeSetter<TInstance, TProperty>(this PropertyInfo member)
        {
            return _Service.MakeSetter<TInstance, TProperty>(member);
        }

        #endregion
    }
}