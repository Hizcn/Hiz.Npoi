using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//namespace Hiz.Common
//{

/* v1.1 2018-12-09
 * v1.0 2012-01-12
 */
public static class Bits
{
    #region GetBits & GetBitCount
    public static UInt64[] GetBits(UInt64 value)
    {
        UInt64[] array;
        var count = TryGetBits(value, out array);

        var result = new UInt64[count];
        for (var i = 0; i < count; i++)
            result[i] = array[i];
        return result;
    }
    public static UInt32[] GetBits(UInt32 value)
    {
        UInt32[] array;
        var count = TryGetBits(value, out array);

        var result = new UInt32[count];
        for (var i = 0; i < count; i++)
            result[i] = array[i];
        return result;
    }
    public static UInt16[] GetBits(UInt16 value)
    {
        UInt32[] array;
        var count = TryGetBits((UInt32)value, out array);

        var result = new UInt16[count];
        for (var i = 0; i < count; i++)
            result[i] = (UInt16)array[i];
        return result;
    }
    public static Byte[] GetBits(Byte value)
    {
        UInt32[] array;
        var count = TryGetBits((UInt32)value, out array);

        var result = new Byte[count];
        for (var i = 0; i < count; i++)
            result[i] = (Byte)array[i];
        return result;
    }

    public static Int64[] GetBits(Int64 value)
    {
        UInt64[] array;
        var count = TryGetBits((UInt64)value, out array);

        var result = new Int64[count];
        for (var i = 0; i < count; i++)
            result[i] = (Int64)array[i];
        return result;
    }
    public static Int32[] GetBits(Int32 value)
    {
        UInt32[] array;
        var count = TryGetBits((UInt32)value, out array);

        var result = new Int32[count];
        for (var i = 0; i < count; i++)
            result[i] = (Int32)array[i];
        return result;
    }
    public static Int16[] GetBits(Int16 value)
    {
        UInt32[] array;
        var count = TryGetBits((UInt32)value, out array);

        var result = new Int16[count];
        for (var i = 0; i < count; i++)
            result[i] = (Int16)array[i];
        return result;
    }
    public static SByte[] GetBits(SByte value)
    {
        UInt32[] array;
        var count = TryGetBits((UInt32)value, out array);

        var result = new SByte[count];
        for (var i = 0; i < count; i++)
            result[i] = (SByte)array[i];
        return result;
    }

    static int TryGetBits(UInt64 value, out UInt64[] array)
    {
        var index = 0;
        UInt64 before;
        array = new UInt64[0x40];
        while (value != 0ul)
        {
            before = value;
            // 消掉二进制最右边的一
            value = value & (value - 1ul);
            // 记录差值
            array[index++] = before - value;
        }
        return index;
    }
    static int TryGetBits(UInt32 value, out UInt32[] array)
    {
        var index = 0;
        UInt32 before;
        array = new UInt32[0x20];
        while (value != 0u)
        {
            before = value;
            value = value & (value - 1u);
            array[index++] = before - value;
        }
        return index;
    }

    public static int GetBitCount(UInt64 value)
    {
        var count = 0;
        while (value != 0ul)
        {
            value = value & (value - 1ul);
            count++;
        }
        return count;
    }
    public static int GetBitCount(UInt32 value)
    {
        var count = 0;
        while (value != 0u)
        {
            value = value & (value - 1u);
            count++;
        }
        return count;
    }
    public static int GetBitCount(UInt16 value)
    {
        return GetBitCount((UInt32)value);
    }
    public static int GetBitCount(Byte value)
    {
        return GetBitCount((UInt32)value);
    }

    public static int GetBitCount(Int64 value)
    {
        return GetBitCount((UInt64)value);
    }
    public static int GetBitCount(Int32 value)
    {
        return GetBitCount((UInt32)value);
    }
    public static int GetBitCount(Int16 value)
    {
        return GetBitCount((UInt32)(UInt16)value); // 当值为负数时 需要先转换成 同长度无符号整型
    }
    public static int GetBitCount(SByte value)
    {
        return GetBitCount((UInt32)(Byte)value); // 当值为负数时 需要先转换成 同长度无符号整型
    }
    #endregion

    #region OnlySingle & MoreSingle

    /* 为什么要转为 32 位整型?
     * 系统会将小于 32 位整型 转为 32 位再计算..大概是这样的
     */

    public static bool OnlySingle(UInt64 value)
    {
        return (value != 0ul) && ((value & (value - 1ul)) == 0ul);
    }
    public static bool OnlySingle(UInt32 value)
    {
        return (value != 0u) && ((value & (value - 1u)) == 0u);
    }
    public static bool OnlySingle(UInt16 value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) == 0u);
    }
    public static bool OnlySingle(Byte value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) == 0u);
    }

    public static bool OnlySingle(Int64 value)
    {
        var v = (UInt64)value;
        return (v != 0ul) && ((v & (v - 1ul)) == 0ul);
    }
    public static bool OnlySingle(Int32 value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) == 0u);
    }
    public static bool OnlySingle(Int16 value)
    {
        var v = (UInt32)(UInt16)value; // 当值为负数时 需要先转换成 同长度无符号整型
        return (v != 0u) && ((v & (v - 1u)) == 0u);
    }
    public static bool OnlySingle(SByte value)
    {
        var v = (UInt32)(Byte)value; // 当值为负数时 需要先转换成 同长度无符号整型
        return (v != 0u) && ((v & (v - 1u)) == 0u);
    }

    public static bool MoreSingle(UInt64 value)
    {
        return (value != 0ul) && ((value & (value - 1ul)) != 0ul);
    }
    public static bool MoreSingle(UInt32 value)
    {
        return (value != 0u) && ((value & (value - 1u)) != 0u);
    }
    public static bool MoreSingle(UInt16 value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) != 0u);
    }
    public static bool MoreSingle(Byte value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) != 0u);
    }

    public static bool MoreSingle(Int64 value)
    {
        var v = (UInt64)value;
        return (v != 0ul) && ((v & (v - 1ul)) != 0ul);
    }
    public static bool MoreSingle(Int32 value)
    {
        var v = (UInt32)value;
        return (v != 0u) && ((v & (v - 1u)) != 0u);
    }
    public static bool MoreSingle(Int16 value)
    {
        var v = (UInt32)(UInt16)value; // 当值为负数时 需要先转换成 同长度无符号整型
        return (v != 0u) && ((v & (v - 1u)) != 0u);
    }
    public static bool MoreSingle(SByte value)
    {
        var v = (UInt32)(Byte)value; // 当值为负数时 需要先转换成 同长度无符号整型
        return (v != 0u) && ((v & (v - 1u)) != 0u);
    }

    #endregion

    static UInt64 RemoveLast(UInt64 value)
    {
        return value & (value - 1ul);
    }
    static UInt32 RemoveLast(UInt32 value)
    {
        return value & (value - 1u);
    }
}

//}
