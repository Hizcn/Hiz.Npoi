﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Hiz.Reflection
{
    partial class ReflectionWithExpression // : ReflectionServiceBase
    {
        #region Constants

        //TODO: 检查表达式的参数命名

        // ParameterName // LambdaExpression: 参数名称允许重复, 不会自动修正, 但是编译之后运行正常;
        protected const string NameInstance = "instance";
        protected const string NameValue = "value";
        protected const string NameIndexes = "indexes";
        protected const string DelegateInvoke = "Invoke";

        // static readonly Type TypeDelegate = typeof(Delegate);
        protected static readonly Type TypeVoid = typeof(void);
        protected static readonly Type TypeObject = typeof(object);
        protected static readonly Type TypeObjectArray = typeof(object[]);

        // 预设方法最大参数数量 (不算返回结果类型)
        protected const int MaximumArity = 16;
        protected const int MaximumArityRef = 5;
        protected static readonly Type[] TypeActions = new[] {
            /* 00 */typeof(Action), // 零个参数; 无返回值;
            /* 01 */typeof(Action<>),
            /* 02 */typeof(Action<,>),
            /* 03 */typeof(Action<,,>),
            /* 04 */typeof(Action<,,,>),
            /* 05 */typeof(Action<,,,,>),
            /* 06 */typeof(Action<,,,,,>),
            /* 07 */typeof(Action<,,,,,,>),
            /* 08 */typeof(Action<,,,,,,,>),
            /* 09 */typeof(Action<,,,,,,,,>),
            /* 10 */typeof(Action<,,,,,,,,,>),
            /* 11 */typeof(Action<,,,,,,,,,,>),
            /* 12 */typeof(Action<,,,,,,,,,,,>),
            /* 13 */typeof(Action<,,,,,,,,,,,,>),
            /* 14 */typeof(Action<,,,,,,,,,,,,,>),
            /* 15 */typeof(Action<,,,,,,,,,,,,,,>),
            /* 16 */typeof(Action<,,,,,,,,,,,,,,,>)
        };
        protected static readonly Type[] TypeFunctions = new[] {
            /* 00 */typeof(Func<>), // 零个参数; 带返回值;
            /* 01 */typeof(Func<,>),
            /* 02 */typeof(Func<,,>),
            /* 03 */typeof(Func<,,,>),
            /* 04 */typeof(Func<,,,,>),
            /* 05 */typeof(Func<,,,,,>),
            /* 06 */typeof(Func<,,,,,,>),
            /* 07 */typeof(Func<,,,,,,,>),
            /* 08 */typeof(Func<,,,,,,,,>),
            /* 09 */typeof(Func<,,,,,,,,,>),
            /* 10 */typeof(Func<,,,,,,,,,,>),
            /* 11 */typeof(Func<,,,,,,,,,,,>),
            /* 12 */typeof(Func<,,,,,,,,,,,,>),
            /* 13 */typeof(Func<,,,,,,,,,,,,,>),
            /* 14 */typeof(Func<,,,,,,,,,,,,,,>),
            /* 15 */typeof(Func<,,,,,,,,,,,,,,,>),
            /* 16 */typeof(Func<,,,,,,,,,,,,,,,,>)
        };

        protected const int MaximumIndexes = 14;

        // void Action(TPropertyOrField value);
        protected static readonly Type TypeMemberSetStatic = TypeActions[1];
        // TPropertyOrField Function();
        protected static readonly Type TypeMemberGetStatic = TypeFunctions[0];

        // void Action(TInstance instance, TPropertyOrField value);
        protected static readonly Type TypeMemberSetInstance = TypeActions[2];
        // TPropertyOrField Function(TInstance instance);
        protected static readonly Type TypeMemberGetInstance = TypeFunctions[1];

        // void Action(TInstance instance, Object[] indexes, TProperty value);
        protected static readonly Type TypeIndexerSetInstance = TypeActions[3];
        // TProperty Function(TInstance instance, Object[] indexes);
        protected static readonly Type TypeIndexerGetInstance = TypeFunctions[2];

        #endregion

        /// <summary>
        /// 转至指定类型
        /// </summary>
        /// <param name="target">赋值目标类型</param>
        /// <param name="expression"></param>
        /// <returns></returns>
        static Expression ConvertIfNeeded(Type target, Expression expression)
        {
            /* System.Linq.Expressions.Expression
             * static BinaryExpression Assign(Expression left, Expression right) {
             *     Expression.RequiresCanWrite(left);
             *     Expression.RequiresCanRead(right);
             *     ValidateType(left.Type);
             *     ValidateType(right.Type);
             *     if (!AreReferenceAssignable(left.Type, right.Type)) {
             *         throw Error.ExpressionTypeDoesNotMatchAssignment(right.Type, left.Type);
             *     }
             *     return new AssignBinaryExpression(left, right);
             * }
             * 
             * static void ValidateType(Type type) {
             *     if (type.IsGenericTypeDefinition) {
             *         throw Error.TypeIsGeneric(type);
             *     }
             *     if (type.ContainsGenericParameters) {
             *         throw Error.TypeContainsGenericParameters(type);
             *     }
             * }
             * 
             * static bool AreReferenceAssignable(Type target, Type source) {
             *     if (AreEquivalent(target, source))
             *         return true;
             *     if (!target.IsValueType && !source.IsValueType && target.IsAssignableFrom(source))
             *         return true;
             *     return false;
             * }
             * 
             * static bool AreEquivalent(Type type, Type other) {
             *     return type == other || type.IsEquivalentTo(other); // 仅仅针对 COM 对象;
             * }
             * 
             * 
             * Test:
             * var value = Expression.Variable(typeof(object));
             * if (typeof(object).IsAssignableFrom(typeof(Int32))) { // 虽然可以分配实例
             *     // 抛出异常: "System.Int32" 类型的表达式不能用于类型 "System.Object" 的赋值.
             *     Expression.Assign(value, Expression.Constant(0xFF));
             * }
             */
            if (expression == null)
                throw new ArgumentNullException();
            if (target == null)
                throw new ArgumentNullException();

            var source = expression.Type;
            if (source != target) // 忽略 COM 对象
            {
                // 判断两个类型是否可以引用赋值: 派生类型赋给基类 / 引用类型赋给接口;
                if (target.IsValueType || source.IsValueType || !target.IsAssignableFrom(source))
                {
                    // 内部实现: 自动支持用户定义转换
                    return Expression.Convert(expression, target, null);
                }
            }
            return expression;
        }
    }
}
