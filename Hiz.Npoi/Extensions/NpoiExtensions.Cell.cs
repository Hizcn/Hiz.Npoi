using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Globalization;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;

namespace Hiz.Npoi
{
    partial class NpoiExtensions
    {
        static readonly Type TypeObject = typeof(Object);
        static readonly Type TypeNullable = typeof(Nullable<>);

        #region Declared Scalar Types

        // 字节集合
        static readonly Type TypeByteArray = typeof(byte[]);

        // 文本
        static readonly Type TypeString = typeof(String);
        // 富文本字符串 (NPOI: CellType.String)
        static readonly Type TypeRichTextString = typeof(IRichTextString);

        // 字符
        static readonly Type TypeChar = typeof(Char);
        static readonly Type TypeCharNullable = typeof(Nullable<Char>);

        // 布尔
        static readonly Type TypeBoolean = typeof(Boolean);
        static readonly Type TypeBooleanNullable = typeof(Nullable<Boolean>);

        // 整型
        static readonly Type TypeByte = typeof(Byte);
        static readonly Type TypeByteNullable = typeof(Nullable<Byte>);
        static readonly Type TypeSByte = typeof(SByte);
        static readonly Type TypeSByteNullable = typeof(Nullable<SByte>);

        static readonly Type TypeInt16 = typeof(Int16);
        static readonly Type TypeInt16Nullable = typeof(Nullable<Int16>);
        static readonly Type TypeInt32 = typeof(Int32);
        static readonly Type TypeInt32Nullable = typeof(Nullable<Int32>);
        static readonly Type TypeInt64 = typeof(Int64);
        static readonly Type TypeInt64Nullable = typeof(Nullable<Int64>);

        static readonly Type TypeUInt16 = typeof(UInt16);
        static readonly Type TypeUInt16Nullable = typeof(Nullable<UInt16>);
        static readonly Type TypeUInt32 = typeof(UInt32);
        static readonly Type TypeUInt32Nullable = typeof(Nullable<UInt32>);
        static readonly Type TypeUInt64 = typeof(UInt64);
        static readonly Type TypeUInt64Nullable = typeof(Nullable<UInt64>);

        // 浮点
        static readonly Type TypeSingle = typeof(Single);
        static readonly Type TypeSingleNullable = typeof(Nullable<Single>);
        static readonly Type TypeDouble = typeof(Double); // NPOI: CellType.Numeric
        static readonly Type TypeDoubleNullable = typeof(Nullable<Double>);

        // 小数
        static readonly Type TypeDecimal = typeof(Decimal);
        static readonly Type TypeDecimalNullable = typeof(Nullable<Decimal>);

        // 日期
        static readonly Type TypeDateTime = typeof(DateTime);
        static readonly Type TypeDateTimeNullable = typeof(Nullable<DateTime>);
        static readonly Type TypeDateTimeOffset = typeof(DateTimeOffset);
        static readonly Type TypeDateTimeOffsetNullable = typeof(Nullable<DateTimeOffset>);

        // 时间
        static readonly Type TypeTimeSpan = typeof(TimeSpan);
        static readonly Type TypeTimeSpanNullable = typeof(Nullable<TimeSpan>);

        // 全局标识
        static readonly Type TypeGuid = typeof(Guid);
        static readonly Type TypeGuidNullable = typeof(Nullable<Guid>);

        #endregion

        /* NPOI: CellType
         * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/CellType.html
         * 
         * // Unknown type, used to represent a state prior to initialization or the lack of a concrete type.
         * // For internal use only. // 仅限内部使用
         * Unknown = -1,
         * 
         * Numeric = 0, // 可存 DateTime // WPS: 不支持毫秒部分 ?
         * String = 1, // 文本
         * Formula = 2, // 公式
         * Blank = 3, // 空白 (单元格未赋值)
         * Boolean = 4, // 布尔
         * 
         * Error = 5, // 公式计算错误或者直接赋值错误;
         * 例如:
         * 方法1. 调用 ICell.SetCellErrorValue();
         * 方法2. 复制结果为错的公式单元格, 选中其它单元格, 右键 > 选择性粘贴 > 选中数值 > 确定.
         */

        #region Constants

        const Byte BlankValueOfErrorCode = 0; // "#NULL!"
        const Double BlankValueOfDouble = 0.0;
        const String BlankValueOfString = null; // NPOI: ""
        const Boolean BlankValueOfBoolean = false;

        // v2.2.1.0: XSSFCell = DateTime.MinValue; HSSFCell = DateTime.MaxValue (问过 Tony Qu, 是为 BUG.);
        static readonly DateTime BlankValueOfDateTime = DateTime.MinValue;

        // Excel: N(FALSE) = 0; N(TRUE) = 1;
        static readonly string[] FalseAsStringArray = new[] { "FALSE", "0" };
        static readonly string[] TrueAsStringArray = new[] { "TRUE", "1" };
        const Double FalseAsNumeric = 0d;
        const Double TrueAsNumeric = 1d;

        #endregion

        /// <summary>
        /// 获取单元格值最终类型, 如果含有公式, 则返回计算结果的类型.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static CellType GetCellTypeFinally(this ICell cell)
        {
            if (cell == null)
                throw new ArgumentNullException();

            var type = cell.CellType;
            if (type == CellType.Formula)
                type = cell.CachedFormulaResultType;
            return type;
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="cell"></param>
        /// <param name="default"></param>
        /// <returns></returns>
        public static T GetCellValue<T>(this ICell cell, T @default = default(T))
        {
            return (T)GetCellValue(cell, typeof(T), @default);
        }
        static object GetCellValue(this ICell cell, Type type, object @default = null)
        {
            if (cell == null)
                throw new ArgumentNullException();

            if (type == TypeString)
                return cell.GetCellValueAsString((String)@default);

            // 如果公开此方法, 需处理 @default 为空值的情况: GetCellValueAsDouble(@default != null ? (Double)@default : default(Double))
            if (type == TypeDouble)
                return cell.GetCellValueAsDoubleNullable((Double?)@default) ?? BlankValueOfDouble;
            if (type == TypeDoubleNullable)
                return cell.GetCellValueAsDoubleNullable((Double?)@default);

            if (type == TypeSingle)
                return cell.GetCellValueAsSingleNullable((Single?)@default) ?? 0f;
            if (type == TypeSingleNullable)
                return cell.GetCellValueAsSingleNullable((Single?)@default);

            if (type == TypeInt32)
                return cell.GetCellValueAsInt32Nullable((Int32?)@default) ?? 0;
            if (type == TypeInt32Nullable)
                return cell.GetCellValueAsInt32Nullable((Int32?)@default);
            if (type == TypeInt64)
                return cell.GetCellValueAsInt64Nullable((Int64?)@default) ?? 0;
            if (type == TypeInt64Nullable)
                return cell.GetCellValueAsInt64Nullable((Int64?)@default);
            if (type == TypeInt16)
                return cell.GetCellValueAsInt16Nullable((Int16?)@default) ?? 0;
            if (type == TypeInt16Nullable)
                return cell.GetCellValueAsInt16Nullable((Int16?)@default);

            if (type == TypeUInt32)
                return cell.GetCellValueAsUInt32Nullable((UInt32?)@default) ?? 0U;
            if (type == TypeUInt32Nullable)
                return cell.GetCellValueAsUInt32Nullable((UInt32?)@default);
            if (type == TypeUInt64)
                return cell.GetCellValueAsUInt64Nullable((UInt64?)@default) ?? 0UL;
            if (type == TypeUInt64Nullable)
                return cell.GetCellValueAsUInt64Nullable((UInt64?)@default);
            if (type == TypeUInt16)
                return cell.GetCellValueAsUInt16Nullable((UInt16?)@default) ?? 0;
            if (type == TypeUInt16Nullable)
                return cell.GetCellValueAsUInt16Nullable((UInt16?)@default);

            if (type == TypeDecimal)
                return cell.GetCellValueAsDecimalNullable((Decimal?)@default) ?? Decimal.Zero;
            if (type == TypeDecimalNullable)
                return cell.GetCellValueAsDecimalNullable((Decimal?)@default);

            if (type == TypeByte)
                return cell.GetCellValueAsByteNullable((Byte?)@default) ?? 0;
            if (type == TypeByteNullable)
                return cell.GetCellValueAsByteNullable((Byte?)@default);

            if (type == TypeSByte)
                return cell.GetCellValueAsSByteNullable((SByte?)@default) ?? 0;
            if (type == TypeSByteNullable)
                return cell.GetCellValueAsSByteNullable((SByte?)@default);

            if (type == TypeDateTime)
                return cell.GetCellValueAsDateTimeNullable((DateTime?)@default) ?? default(DateTime);
            if (type == TypeDateTimeNullable)
                return cell.GetCellValueAsDateTimeNullable((DateTime?)@default);

            if (type == TypeBoolean)
                return cell.GetCellValueAsBooleanNullable((Boolean?)@default) ?? BlankValueOfBoolean;
            if (type == TypeBooleanNullable)
                return cell.GetCellValueAsBooleanNullable((Boolean?)@default);

            if (type == TypeChar)
                return cell.GetCellValueAsCharNullable((Char?)@default) ?? Char.MinValue;
            if (type == TypeCharNullable)
                return cell.GetCellValueAsCharNullable((Char?)@default);

            if (type == TypeTimeSpan)
                return cell.GetCellValueAsTimeSpanNullable((TimeSpan?)@default) ?? TimeSpan.Zero;
            if (type == TypeTimeSpanNullable)
                return cell.GetCellValueAsTimeSpanNullable((TimeSpan?)@default);

            if (type == TypeGuid)
                return cell.GetCellValueAsGuidNullable((Guid?)@default) ?? Guid.Empty;
            if (type == TypeGuidNullable)
                return cell.GetCellValueAsGuidNullable((Guid?)@default);

            throw new NotSupportedException();
        }

        #region GetCellValue: Excel Scalar Types

        public static String GetCellValueAsString(this ICell cell, String @default = BlankValueOfString, bool trim = true)
        {
            var value = GetCellValueAsString(cell, null);
            if (value == null)
                return @default;

            if (trim)
                value = value.Trim();
            return value;
        }

        static String GetCellValueAsString(this ICell cell, String @default)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    //TODO: 空字符串是否返回 @default ?
                    return cell.StringCellValue;
                case CellType.Numeric:
                    {
                        if (DateUtil.IsCellDateFormatted(cell)) // 判断是否日期格式
                        {
                            var date = cell.DateCellValue;
                            //TODO: 应用单元格的格式, 然后输出.
                            return date.ToString(CultureInfo.InvariantCulture);
                        }
                        else
                        {
                            //TODO: 应用单元格的格式, 然后输出.
                            return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    {
                        if (cell is XSSFCell)
                        {
                            return ((XSSFCell)cell).ErrorCellString;
                        }
                        if (cell is HSSFCell)
                        {
                            HSSFErrorConstants.GetText(cell.ErrorCellValue);
                        }
                        return cell.ErrorCellValue.ToString();
                    }
                default:
                    throw new InvalidCastException();
            }
        }

        // static IRichTextString GetCellValueAsRichString(this ICell cell, string @default = null)
        // {
        //     switch (cell.GetCellTypeFinally())
        //     {
        //         case CellType.String:
        //             return cell.RichStringCellValue;
        //         case CellType.Blank:
        //         case CellType.Numeric:
        //         case CellType.Boolean:
        //         case CellType.Error:
        //             {
        //                 var value = cell.GetCellValueAsString(@default);
        //                 if (cell is XSSFCell || cell is SXSSFCell)
        //                 {
        //                     return new XSSFRichTextString(value);
        //                 }
        //                 if (cell is HSSFCell)
        //                 {
        //                     return new HSSFRichTextString(value);
        //                 }
        //                 throw new NotSupportedException();
        //             }
        //         default:
        //             throw new InvalidCastException();
        //     }
        // }

        // static Double GetCellValueAsDouble(this ICell cell, Double @default = BlankValueOfDouble)
        // {
        //     return GetCellValueAsDoubleNullable(cell) ?? @default;
        // }
        static Nullable<Double> GetCellValueAsDoubleNullable(this ICell cell, Double? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Double.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return cell.NumericCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue ? TrueAsNumeric : FalseAsNumeric;
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        // static DateTime GetCellValueAsDateTime(this ICell cell, DateTime @default = default(DateTime)/*DateTime.MinValue*/)
        // {
        //     return GetCellValueAsDateTimeNullable(cell) ?? @default;
        // }
        static Nullable<DateTime> GetCellValueAsDateTimeNullable(this ICell cell, DateTime? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            // DateTime.Parse 自动忽略首尾空白字符;
                            return DateTime.Parse(cell.StringCellValue, CultureInfo.InvariantCulture);
                        }
                        return @default;
                    }
                case CellType.Numeric:
                    {
                        //TODO: 是否检查单元格的格式设置 ? // DateUtil.IsCellDateFormatted(cell)
                        return cell.DateCellValue;
                    }
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        // static Boolean GetCellValueAsBoolean(this ICell cell, Boolean @default = BlankValueOfBoolean)
        // {
        //     return GetCellValueAsBooleanNullable(cell) ?? @default;
        // }
        static Nullable<Boolean> GetCellValueAsBooleanNullable(this ICell cell, Boolean? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            text = text.Trim(); // 去掉首尾空格

                            // 对于字符串的转换, 优先判断真值. // 匹配为真; 无效为假;
                            foreach (var t in TrueAsStringArray)
                                if (string.Equals(t, text, StringComparison.OrdinalIgnoreCase))
                                    return true;
                            //TODO: 是否严谨判断? {
                            foreach (var f in FalseAsStringArray)
                                if (string.Equals(f, text, StringComparison.OrdinalIgnoreCase))
                                    return false;
                            throw new FormatException();
                            // } End
                        }
                        return @default;
                    }
                case CellType.Numeric:
                    {
                        // 对于数字转换, 优先判断假值. // 零值为假; 非零为真;
                        // 0 => false; other => true;
                        var numeric = cell.NumericCellValue;
                        if (numeric == FalseAsNumeric)
                            return false;
                        //if (numeric == TrueAsNumeric)
                        return true;
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        #endregion

        #region GetCellValue: Other .Net Scalar Types

        // static Byte GetCellValueAsByte(this ICell cell, Byte @default = 0)
        // {
        //     return cell.GetCellValueAsByteNullable() ?? @default;
        // }
        static Nullable<Byte> GetCellValueAsByteNullable(this ICell cell, Byte? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Byte.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    // 检查溢出; 超出范围将会抛出异常;
                    // 若不检查, 超出范围将被截断: 0x1234 => 0x34, 结果将不准确;
                    return checked((Byte)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static SByte GetCellValueAsSByte(this ICell cell, SByte @default = 0)
        // {
        //     return cell.GetCellValueAsSByteNullable() ?? @default;
        // }
        static Nullable<SByte> GetCellValueAsSByteNullable(this ICell cell, SByte? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return SByte.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((SByte)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        // static Int16 GetCellValueAsInt16(this ICell cell, Int16 @default = 0)
        // {
        //     return cell.GetCellValueAsInt16Nullable() ?? @default;
        // }
        static Nullable<Int16> GetCellValueAsInt16Nullable(this ICell cell, Int16? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Int16.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((Int16)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static Int32 GetCellValueAsInt32(this ICell cell, Int32 @default = 0)
        // {
        //     return cell.GetCellValueAsInt32Nullable() ?? @default;
        // }
        static Nullable<Int32> GetCellValueAsInt32Nullable(this ICell cell, Int32? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Int32.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((Int32)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static Int64 GetCellValueAsInt64(this ICell cell, Int64 @default = 0L)
        // {
        //     return cell.GetCellValueAsInt64Nullable() ?? @default;
        // }
        static Nullable<Int64> GetCellValueAsInt64Nullable(this ICell cell, Int64? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Int64.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((Int64)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        // static UInt16 GetCellValueAsUInt16(this ICell cell, UInt16 @default = 0)
        // {
        //     return cell.GetCellValueAsUInt16Nullable() ?? @default;
        // }
        static Nullable<UInt16> GetCellValueAsUInt16Nullable(this ICell cell, UInt16? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return UInt16.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((UInt16)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static UInt32 GetCellValueAsUInt32(this ICell cell, UInt32 @default = 0U)
        // {
        //     return cell.GetCellValueAsUInt32Nullable() ?? @default;
        // }
        static Nullable<UInt32> GetCellValueAsUInt32Nullable(this ICell cell, UInt32? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return UInt32.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((UInt32)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static UInt64 GetCellValueAsUInt64(this ICell cell, UInt64 @default = 0UL)
        // {
        //     return cell.GetCellValueAsUInt64Nullable() ?? @default;
        // }
        static Nullable<UInt64> GetCellValueAsUInt64Nullable(this ICell cell, UInt64? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return UInt64.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    return checked((UInt64)cell.NumericCellValue);
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static Single GetCellValueAsSingle(this ICell cell, Single @default = 0f)
        // {
        //     return cell.GetCellValueAsSingleNullable() ?? @default;
        // }
        static Nullable<Single> GetCellValueAsSingleNullable(this ICell cell, Single? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Single.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    {
                        // Double => Single 转换使用 checked 无效;
                        // 如果浮点运算结果的数值对于目标格式来说太大，则运算的结果为 PositiveInfinity 或 NegativeInfinity（具体取决于结果的符号）。
                        // 例如: Single.MaxValue + 0.000001e+038f = Single.PositiveInfinity;
                        var result = (Single)cell.NumericCellValue;
                        if (Single.IsInfinity(result))
                            throw new OverflowException();
                        return result;
                    }
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static Decimal GetCellValueAsDecimal(this ICell cell, Decimal @default = Decimal.Zero)
        // {
        //     return cell.GetCellValueAsDecimalNullable() ?? @default;
        // }
        static Nullable<Decimal> GetCellValueAsDecimalNullable(this ICell cell, Decimal? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            return Decimal.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                    // Double 最大精度位数: 15 (内部维护位数: 17);
                    // Decimal 最大位数: 29;
                    // Single/Double => Decimal 始终检查溢出, 无需添加 checked;
                    return (Decimal)cell.NumericCellValue;
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        // static Char GetCellValueAsChar(this ICell cell, Char @default = char.MinValue/*(Char)0*/)
        // {
        //     return cell.GetCellValueAsCharNullable() ?? @default;
        // }
        static Nullable<Char> GetCellValueAsCharNullable(this ICell cell, Char? @default = null)
        {
            var value = cell.GetCellValueAsString();
            if (value != null)
            {
                var length = value.Length;
                if (length > 0)
                {
                    if (--length > 0)
                        throw new InvalidCastException(); // 不止一个字符将抛异常;
                    return value[0];
                }
            }
            return @default;
        }

        // static Nullable<Guid> GetCellValueAsGuid(this ICell cell, Guid @default = default(Guid)/*Guid.Empty*/)
        // {
        //     return cell.GetCellValueAsGuidNullable() ?? @default;
        // }
        static Nullable<Guid> GetCellValueAsGuidNullable(this ICell cell, Guid? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            // Guid.Parse() 内部将会调用 Trim 去掉首尾空白;
                            return Guid.Parse(text);
                        return @default;
                    }
                case CellType.Numeric:
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }
        // static Nullable<TimeSpan> GetCellValueAsTimeSpan(this ICell cell, TimeSpan @default = default(TimeSpan)/*TimeSpan.Zero*/)
        // {
        //     return cell.GetCellValueAsTimeSpanNullable() ?? @default;
        // }
        static Nullable<TimeSpan> GetCellValueAsTimeSpanNullable(this ICell cell, TimeSpan? @default = null)
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                            // TimeSpan.Parse() 内部将会调用 Trim 去掉首尾空白;
                            return TimeSpan.Parse(text, CultureInfo.InvariantCulture);
                        return @default;
                    }
                case CellType.Numeric:
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        #endregion

        //TODO:
        static TEnum GetCellValueAsEnum<TEnum>(this ICell cell, TEnum @default = default(TEnum)) // where TEnum : struct
        {
            switch (cell.GetCellTypeFinally())
            {
                case CellType.Blank:
                    return @default;
                case CellType.String:
                    {
                        var text = cell.StringCellValue;
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            return (TEnum)Enum.Parse(typeof(TEnum), text);
                        }
                        return @default;
                    }
                case CellType.Numeric:
                    {
                        var type = typeof(TEnum);
                        if (!type.IsEnum)
                            throw new ArgumentException();
                        var underlying = type.GetEnumUnderlyingType();
                        var value = cell.NumericCellValue;
                        var convertible = (IConvertible)value;

                        throw new NotSupportedException();
                    }
                case CellType.Boolean:
                case CellType.Error:
                default:
                    throw new InvalidCastException();
            }
        }

        #region SetCellValue

        public static void SetCellValue(this ICell cell, object value)
        {
            if (cell == null)
                throw new ArgumentNullException();

            if (value != null)
            {
                if (value is String)
                {
                    cell.SetCellValue((String)value);
                }
                #region 数值转换
                // 数值转换效率: (Double)(T)value 快过 ((IConvertible)value).ToDouble(null);
                else if (value is Double)
                {
                    // Double 最大精度位数: 15
                    cell.SetCellValue((Double)value);
                }
                else if (value is Single)
                {
                    // Single 最大精度位数: 07
                    cell.SetCellValue((Double)(Single)value);
                }
                else if (value is Decimal)
                {
                    // Decimal 最大位数: 29 // 转换损失精度 (15 位之后全为零)
                    cell.SetCellValue((Double)(Decimal)value);
                }
                else if (value is Int32)
                {
                    // Int32 最大位数: 10
                    cell.SetCellValue((Double)(Int32)value);
                }
                else if (value is Int64)
                {
                    // Int64 最大位数: 19 // 转换损失精度 (15 位之后全为零)
                    cell.SetCellValue((Double)(Int64)value);
                }
                else if (value is Int16)
                {
                    // Int16 最大位数: 5
                    cell.SetCellValue((Double)(Int16)value);
                }
                else if (value is UInt32)
                {
                    // UInt32 最大位数: 10
                    cell.SetCellValue((Double)(UInt32)value);
                }
                else if (value is UInt64)
                {
                    // UInt64 最大位数: 20 // 转换损失精度 (15 位之后全为零)
                    cell.SetCellValue((Double)(UInt64)value);
                }
                else if (value is UInt16)
                {
                    // UInt16 最大位数: 5
                    cell.SetCellValue((Double)(UInt16)value);
                }
                else if (value is Byte)
                {
                    // Byte 最大位数: 3
                    cell.SetCellValue((Double)(Byte)value);
                }
                else if (value is SByte)
                {
                    // SByte 最大位数: 3
                    cell.SetCellValue((Double)(SByte)value);
                }
                #endregion
                else if (value is DateTime)
                {
                    if (cell is SXSSFCell)
                    {
                        // object i = DateTime.Today;
                        // var k = (DateTime?)i; // 转换不会失败;
                        // NPOIv2.3 SXSSFCell BUG: 不能设置时间; SetCellValue(DateTime/DateTime?);
                        // 临时过渡方法
                        var b = ((SXSSFWorkbook)cell.Sheet.Workbook).XssfWorkbook.IsDate1904();
                        cell.SetCellValue(DateUtil.GetExcelDate((DateTime)value, b));
                    }
                    else
                    {
                        // 如果没有设置日期格式, 显示值为 Double.
                        cell.SetCellValue((DateTime)value);
                    }
                }
                else if (value is Boolean)
                {
                    cell.SetCellValue((Boolean)value);
                }
                else
                {
                    // Char
                    // Guid
                    // TimeSpan
                    // Others
                    cell.SetCellValue(value.ToString());
                }
            }
            else
            {
                cell.SetCellType(CellType.Blank);
            }
        }

        // 对于 Decimal 类型特殊处理: 如果超出 Double 范围 那么保存值为文本..
        static void SetCellValueAsString(this ICell cell, Decimal value)
        {
        }
        static void SetCellValueAsString(this ICell cell, UInt64 value)
        {
        }
        static void SetCellValueAsString(this ICell cell, Int64 value)
        {
        }

        // static void SetCellValue<T>(this ICell cell, T value) where T : struct
        // {
        // }
        // static void SetCellValue<T>(this ICell cell, Nullable<T> value) where T : struct
        // {
        //     if (!value.HasValue)
        //     {
        //         cell.SetCellType(CellType.Blank);
        //     }
        //     else
        //     {
        //         cell.SetCellValue(value.Value);
        //     }
        // }

        #endregion

        public static IDrawing GetOrAddDrawingPatriarch(this ISheet sheet)
        {
            if (sheet == null)
                throw new ArgumentNullException(nameof(sheet));

            if (sheet.DrawingPatriarch != null)
                return sheet.DrawingPatriarch;
            return sheet.CreateDrawingPatriarch();
        }

        const int EmusPerCentimeter = 360000; // EmusPerInch / 2.54;
        const int EmusPerInch = 914400;
        const int EmusPerPoint = 12700; // EmusPerInch / 72;
        const int EmusPerPixel = 9525; // EmusPerInch / 96;
        public static IComment SetCellComment(this ICell cell, string @string)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            var comment = cell.CellComment;
            if (comment == null)
            {
                var drawing = cell.Sheet.GetOrAddDrawingPatriarch();
                /* https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Drawing.html
                 * dx1 - the x coordinate in EMU within the first cell.
                 * dy1 - the y coordinate in EMU within the first cell.
                 * dx2 - the x coordinate in EMU within the second cell.
                 * dy2 - the y coordinate in EMU within the second cell.
                 * col1 - the column (0 based) of the first cell.
                 * row1 - the row (0 based) of the first cell.
                 * col2 - the column (0 based) of the second cell.
                 * row2 - the row (0 based) of the second cell.
                 * 
                 * EMU: English Metric Unit
                 * https://poi.apache.org/apidocs/dev/org/apache/poi/util/Units.html
                 */
                var anchor = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 0, 0);
                comment = drawing.CreateCellComment(anchor);
                cell.CellComment = comment;
            }
            if (comment is XSSFComment)
                comment.String = new XSSFRichTextString(@string);
            if (comment is HSSFComment)
                comment.String = new HSSFRichTextString(@string);
            return comment;
        }
    }
}
