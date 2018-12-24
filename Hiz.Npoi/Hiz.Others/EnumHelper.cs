using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;

namespace Hiz.Npoi
{
    /* v1.0 2018-12-10
     */
    public static class EnumHelper
    {
        const string EnumSeperator = ", ";
        static readonly char[] EnumSeperatorCharArray = new char[] { ',' };

        static readonly Type TypeEnumMemberAttribute = typeof(EnumMemberAttribute);
        static readonly Type TypeFlagsAttribute = typeof(FlagsAttribute);

        static readonly IDictionary<Type, EnumMapping> _Mappings; // Key: EnumType;

        static EnumHelper()
        {
            _Mappings = new Dictionary<Type, EnumMapping>();
        }

        #region Mapping

        class EnumMapping
        {
            // 不含 EnumMemberAttribute 特性, 使用原生实现..
            public readonly bool HasCustomAttributes;

            public EnumMapping(bool has)
            {
                this.HasCustomAttributes = has;
            }

            /* 同一类型, 一个文本对应一个枚举, 但是多个枚举的值可能相同, 例如:
             * enum TestEnum {
             *     None = 0,
             *     Unknown = 0,
             * }
             * "None" => TestEnum.None;
             * "Unknown" => TestEnum.Unknown;
             */
            public IDictionary<string, object> MappingNameValue; // Key: EnumString; Value: Enum;

            /* 同一数值对应多个枚举, 则取最后出现的枚举的文本(.Net 官方逻辑), 例如(同上面的定义):
             * TestEnum.None.ToString() => "Unknown";
             * TestEnum.Unknown.ToString() => "Unknown";
             */
            public IDictionary<ulong, string> MappingValueName; // Key: (ulong)Enum; Value: EnumString

            /* UInt64 Convert(object @enum) {
             *     return (UInt64)(...)@enum;
             * }
             */
            public Func<object, UInt64> ObjectToValue;
            public Func<UInt64, object> ValueToObject;

            public bool HasFlags;

            public object DefaultValue; // = (object)default(TEnum);
        }
        // struct EnumModel
        // {
        //     public string Name;
        //     public UInt64 Value; // @enum => ulong;
        //     public object Object; // (object)@enum;
        // }

        static readonly object _LockMappings = new object();
        static EnumMapping TryGetMapping(Type type)
        {
            if (type == null)
                throw new ArgumentNullException(nameof(type));
            if (!type.IsEnum)
                throw new ArgumentException(nameof(type));

            EnumMapping mapping;
            if (!_Mappings.TryGetValue(type, out mapping)) // 首次查询
            {
                lock (_LockMappings) // 加锁
                {
                    if (!_Mappings.TryGetValue(type, out mapping)) // 再次查询
                    {
                        var mNameValue = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                        var mValueName = new Dictionary<ulong, string>();

                        var count = 0; // 枚举文本特性数量
                        var convert = GetUInt64WithAction(type);

                        var fields = type.GetFields(BindingFlags.Static | BindingFlags.Public);
                        foreach (var f in fields)
                        {
                            string text;
                            var attribute = (EnumMemberAttribute)f.GetCustomAttributes(TypeEnumMemberAttribute, false).FirstOrDefault();
                            if (attribute != null)
                            {
                                text = attribute.Value;
                                count++;
                            }
                            else
                            {
                                text = f.Name;
                            }

                            var value = f.GetValue(null);

                            if (!mNameValue.ContainsKey(text))
                            {
                                mNameValue.Add(text, value); // 则取最先出现的枚举的文本.
                            }

                            mValueName[convert(value)] = text; // 则取最后出现的枚举的文本(.Net 官方逻辑)
                        }

                        mapping = new EnumMapping(count > 0);

                        if (mapping.HasCustomAttributes)
                        {
                            mapping.MappingNameValue = mNameValue;
                            mapping.MappingValueName = mValueName;
                            mapping.ObjectToValue = convert;
                            mapping.HasFlags = type.GetCustomAttributes(TypeFlagsAttribute, false).Any();
                            mapping.DefaultValue = Activator.CreateInstance(type);
                        }
                        _Mappings.Add(type, mapping);
                    }
                }
            }

            return mapping;
        }

        static Func<object, UInt64> GetUInt64WithAction(Type type)
        {
            var u = type.GetEnumUnderlyingType();

            if (u == typeof(Int32))
                return GetUInt64WithEnumOfInt32;
            if (u == typeof(Int16))
                return GetUInt64WithEnumOfInt16;
            if (u == typeof(SByte))
                return GetUInt64WithEnumOfSByte;
            if (u == typeof(Int64))
                return GetUInt64WithEnumOfInt64;

            if (u == typeof(UInt32))
                return GetUInt64WithEnumOfUInt32;
            if (u == typeof(UInt16))
                return GetUInt64WithEnumOfUInt16;
            if (u == typeof(Byte))
                return GetUInt64WithEnumOfByte;
            if (u == typeof(UInt64))
                return GetUInt64WithEnumOfUInt64;

            throw new NotSupportedException();
        }
        static UInt64 GetUInt64WithEnumOfUInt64(object @enum)
        {
            return (UInt64)@enum;
        }
        static UInt64 GetUInt64WithEnumOfUInt32(object @enum)
        {
            return (UInt64)(UInt32)@enum;
        }
        static UInt64 GetUInt64WithEnumOfUInt16(object @enum)
        {
            return (UInt64)(UInt16)@enum;
        }
        static UInt64 GetUInt64WithEnumOfByte(object @enum)
        {
            return (UInt64)(Byte)@enum;
        }
        static UInt64 GetUInt64WithEnumOfInt64(object @enum)
        {
            return (UInt64)(Int64)@enum;
        }
        static UInt64 GetUInt64WithEnumOfInt32(object @enum)
        {
            return (UInt64)(UInt32)(Int32)@enum;
        }
        static UInt64 GetUInt64WithEnumOfInt16(object @enum)
        {
            return (UInt64)(UInt16)(Int16)@enum;
        }
        static UInt64 GetUInt64WithEnumOfSByte(object @enum)
        {
            return (UInt64)(Byte)(SByte)@enum;
        }

        //static object GetEnumOfUInt64WithValue(UInt64 value)
        //{

        //}
        #endregion

        /* System.Enum
         * object Parse(Type enumType, string value);
         * object Parse(Type enumType, string value, bool ignoreCase);
         * bool TryParse<TEnum>(string value, out TEnum result) where TEnum : struct;
         * bool TryParse<TEnum>(string value, bool ignoreCase, out TEnum result) where TEnum : struct;
         */
        public static bool TryParse<TEnum>(string value, out TEnum result) where TEnum : struct
        {
            var type = typeof(TEnum);

            if (type != null && type.IsEnum && !string.IsNullOrWhiteSpace(value))
            {
                if (InternalTryParse(type, value, out object general))
                {
                    result = (TEnum)general;
                    return true;
                }
            }

            result = default(TEnum);
            return false;
        }
        public static object Parse(Type type, string value)
        {
            if (type == null)
                throw new ArgumentNullException(nameof(type));
            if (!type.IsEnum)
                throw new ArgumentException(nameof(type));
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException(nameof(value));

            if (!InternalTryParse(type, value, out object general))
                throw new ArgumentException(nameof(value));
            return general;
        }

        // 内部调用;
        // 必须保证: type != null && type.IsEnum && !string.IsNullOrWhiteSpace(value)
        static bool InternalTryParse(Type type, string value, out object result)
        {
            var mapping = TryGetMapping(type);
            if (mapping == null)
                throw new InvalidProgramException();

            if (!mapping.HasCustomAttributes)
            {
                try
                {
                    result = Enum.Parse(type, value, true);
                    return true;
                }
                catch
                {
                    result = null;
                    return false;
                }
            }
            else
            {
                var array = mapping.MappingNameValue;

                if (!mapping.HasFlags)
                    return array.TryGetValue(value.Trim(), out result);

                if (value.IndexOf(EnumSeperatorCharArray[0]) < 0)
                    return array.TryGetValue(value.Trim(), out result);

                var texts = value.Split(EnumSeperatorCharArray);
                var length = texts.Length;
                var count = 0; // 解析成功次数
                var aggregate = 0ul;
                var convert = mapping.ObjectToValue;
                for (var i = 0; i < length; i++)
                {
                    var t = texts[i].Trim();
                    if (!string.IsNullOrEmpty(t) && !array.TryGetValue(t, out object v))
                    {
                        aggregate |= convert(v);
                    }
                }

                result = Enum.ToObject(type, aggregate);
                return count > 0;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="enum"></param>
        /// <param name="flags">是否强制使用 FlagsAttribute 规则</param>
        /// <returns></returns>
        public static string GetString(object @enum, bool flags = false/*TODO: 待实现*/)
        {
            if (@enum == null)
                throw new ArgumentNullException();

            var type = @enum.GetType();
            if (!type.IsEnum)
                throw new ArgumentException();

            var mapping = TryGetMapping(type);
            if (mapping == null)
                throw new InvalidProgramException();

            if (!mapping.HasCustomAttributes)
            {
                // 输出枚举名称; 支持 Flags 特性; 如果没有对应名称, 则将输出 十进制文本值.
                return @enum.ToString();
            }
            else
            {
                var array = mapping.MappingValueName;
                var value = mapping.ObjectToValue(@enum);

                if (!mapping.HasFlags)
                {
                    if (!array.TryGetValue(value, out string result))
                        throw new ArgumentException();
                    return result;
                }

                //TODO: 待改进;
                /* 官方算法:
                 * [Flags]
                 * enum TestEnum : short {
                 *     One = 0x01,
                 *     Two = 0x02,
                 *     Three = One | Two,
                 *     Four = 0x04,
                 * }
                 * (TestEnum.One | TestEnum.Two).ToString() == "Three";
                 * (TestEnum.One | TestEnum.Four).ToString() == "One, Four";
                 */
                var keys = Bits.GetBits(value);
                var length = keys.Length;
                var texts = new string[length];
                for (var i = 0; i < length; i++)
                {
                    var k = keys[i];
                    if (!array.TryGetValue(k, out string t))
                    {
                        t = k.ToString(); // 值若无效, 直接转为数字文本.
                    }
                    texts[i] = t;
                }

                return string.Join(EnumSeperator, texts);
            }
        }
    }

    /* System.Enum
     * 
     * static string Format(Type enumType, object value, string format);
     * 确保:
     * enumType != null && enumType.IsEnum = true
     * value != null && (value.GetType() == enumType || value.GetType() == enumType.GetEnumUnderlyingType())
     * format != null && "GgXxDdFf".Contains(format)
     * 
     * G/g: 优先转成名称 不能则为 十进制文本值.
     * X/x: 十六进制文本 // enumValue.ToInteger().ToString("X")
     * D/d: 十进制文本值 // enumValue.ToInteger().ToString("D")
     * F/f: 强制使用 FlagsAttribute 转换文本; 等同 G/g, 只是 FlagsAttribute 不需要存在于 Enum 声明.
     * 
     * 
     * static string GetName(Type enumType, object value);
     * 确保:
     * enumType != null && enumType.IsEnum = true
     * value != null && (value.GetType() == enumType || value.GetType() == enumType.GetEnumUnderlyingType())
     * 结果:
     * 返回 value 所对应的常数名称; 如果无法找到对应, 返回空值 null.
     * 
     * 
     * static Type GetUnderlyingType(Type enumType);
     * 结果: 返回 enumType.GetEnumUnderlyingType();
     * 
     * 
     * static string[] GetNames(Type enumType);
     * 确保:
     * enumType != null && enumType.IsEnum = true
     * 结果:
     * 返回 枚举名称的字符串数组.
     * 
     * 
     * static Array GetValues(Type enumType);
     * 类似: GetNames();
     * 条目类型 == enumType;
     * 
     * 
     * static bool IsDefined(Type enumType, object value);
     * 确保:
     * enumType != null && enumType.IsEnum = true
     * value != null && (value.GetType() == enumType || value.GetType() == enumType.GetEnumUnderlyingType() || value.GetType() == typeof(string))
     * 备注:
     * value 支持字符串值, 但区分大小写.
     * 对于 Flags 枚举, 如果 value 包含多个枚举标识, 并且 enumType 没有定义该组合值, 则将返回 false.
     * 例如:
     * [Flags]
     * enum TestEnum : short {
     *     One = 0x01,
     *     Two = 0x02,
     *     Three = One | Two,
     *     Four = 0x04,
     * }
     * Enum.IsDefined(typeof(TestEnum), "Three"); // true;
     * Enum.IsDefined(typeof(TestEnum), (short)3); // true
     * Enum.IsDefined(typeof(TestEnum), TestEnum.Three); // true;
     * Enum.IsDefined(typeof(TestEnum), TestEnum.One | TestEnum.Two); // true
     * Enum.IsDefined(typeof(TestEnum), "One, Two"); // false
     * Enum.IsDefined(typeof(TestEnum), TestEnum.Four| TestEnum.One); // false
     * 
     * 
     * static object ToObject(Type enumType, Object value);
     * 确保:
     * enumType != null && enumType.IsEnum = true
     * value != null && value.GetType() == (SByte/Int16/Int32/Int64/Byte/UInt16/UInt32/UInt64)
     * 允许: value.GetType() != enumType.GetEnumUnderlyingType(); // 方法自动转换整型类型;
     * 结果: 返回枚举成员; 支持 Flags 特性;
     * 
     * static object ToObject(Type enumType, SByte value);
     * static object ToObject(Type enumType, Int16 value);
     * static object ToObject(Type enumType, Int32 value);
     * static object ToObject(Type enumType, Int64 value);
     * static object ToObject(Type enumType, Byte value);
     * static object ToObject(Type enumType, UInt16 value);
     * static object ToObject(Type enumType, UInt32 value);
     * static object ToObject(Type enumType, UInt64 value);
     * 等同: ToObject(Type, Object) 方法;
     * 
     */
}