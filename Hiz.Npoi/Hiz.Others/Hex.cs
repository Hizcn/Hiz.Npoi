using System;
using System.Collections.Generic;
using System.Linq;

//namespace Hiz.Common
//{

/* 十六进制文本转换
 * 
 * Author: Hiz
 * 
 * v1.4 2018-11-22:
 * 1. 部分方法增加 reverse 参数;
 * 2. 支持小写输出;
 * 
 * v1.3 2016-12-31: 微调
 * v1.2 2016-06-22: 优化性能
 * v1.1 2012-01-11
 * v1.0 2011-08-16
 */
public static class Hex
{
    const int CharsPerByte = 2;

    // 访问 Chars[index] 效率高于访问 String[index]; (Release 模式)
    /* 0-9: 0x30-0x39
     * A-F: 0x41-0x46
     * a-f: 0x61-0x66
     */
    static readonly char[] _HexUppers = new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F' };
    static readonly char[] _HexLowers = new[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

    /* 参考:
     * System.Security.Util.Hex
     * EncodeHexString() // 正序
     * EncodeHexStringFromInt() // 倒序
     * 
     * BitConverter.ToString() // 正序
     */
    /// <summary>
    /// 编码
    /// reverse = false: new byte[] { 0x01, 0x02, 0x03, 0x04 } => "01020304"; 等效 System.BitConverter.ToString().Replace("-", null);
    /// reverse = true : new byte[] { 0x01, 0x02, 0x03, 0x04 } => "04030201";
    /// 例如:
    /// int value = 0x01020304;
    /// var bytes = BitConverter.GetBytes(value); // new byte[] { 0x04, 0x03, 0x02, 0x01 }; // LittleEndian;
    /// var hex1 = BitConverter.ToString(bytes).Replace("-", null); // "04030201"; // 正序
    /// var hex2 = value.ToString("X"); // "01020304"; // 倒序: BitConverter.ToString(BitConverter.GetBytes(value).Reverse()).Replace("-", null);
    /// </summary>
    /// <param name="bytes"></param>
    /// <param name="start">起始索引</param>
    /// <param name="length">字节长度; -1: 从起始索引至数组末尾(length = bytes.Length - start)</param>
    /// <param name="lower">true: 输出大写; false: 输出小写.</param>
    /// <param name="reverse">是否倒序输出结果</param>
    /// <returns>byte[0] = String.Empty;</returns>
    public static string EncodeHexString(byte[] bytes, int start = 0, int length = -1, bool lower = false, bool reverse = false)
    {
        if (bytes == null)
            throw new ArgumentNullException("value");
        if (start < 0)
            throw new ArgumentOutOfRangeException("start");
        // if (length < 0)
        //     throw new ArgumentOutOfRangeException("length");
        if (length < 0 && length != -1)
            throw new ArgumentOutOfRangeException("length");

        int end;
        if (length > 0)
        {
            end = checked(start + length);
            if (end > bytes.Length)
                throw new ArgumentOutOfRangeException("length");
        }
        else // if (length <= 0)
        {
            end = bytes.Length;
            length = end - start;
        }

        // 字节数组转换之后 超过 字符串的最大长度
        // if (length > int.MaxValue / CharsPerByte)
        //     throw new ArgumentOutOfRangeException();

        if (length > 0)
        {
            var chars = new char[length * CharsPerByte];
            var c = 0; // 字符索引
            var digits = lower ? _HexLowers : _HexUppers;
            if (!reverse)
            {
                for (var i = start; i < end; i++)
                {
                    var item = (int)bytes[i];
                    chars[c++] = digits[item >> 0x4]; // item / 0x10;
                    chars[c++] = digits[item & 0x0F]; // item % 0x10;
                }
            }
            else
            {
                while (end > start)
                {
                    var item = (int)bytes[--end];
                    chars[c++] = digits[item >> 0x4]; // item / 0x10;
                    chars[c++] = digits[item & 0x0F]; // item % 0x10;
                }
            }

            return new string(chars);
        }
        return string.Empty;
    }

    /* 参考
     * System.Security.Util.Hex.HexDecodeHexString(); // 正序
     */
    /// <summary>
    /// 解码
    /// </summary>
    /// <param name="value"></param>
    /// <param name="reverse">是否倒序输出结果</param>
    /// <returns></returns>
    public static byte[] DecodeHexString(string value, bool reverse = false)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentNullException(nameof(value));

        value = value.Trim();
        var length = value.Length;
        if (length % CharsPerByte != 0)
            throw new ArgumentException("字符串的长度不是偶数", nameof(value));

        length /= CharsPerByte;
        var buffer = new byte[length];
        var c = 0; // 字符索引
        if (!reverse)
        {
            for (var b = 0; b < length; b++)
            {
                buffer[b] = (byte)(ParseHexDigit(value[c++]) << 0x04 | ParseHexDigit(value[c++]));
            }
        }
        else
        {
            while (length > 0)
            {
                buffer[--length] = (byte)(ParseHexDigit(value[c++]) << 0x04 | ParseHexDigit(value[c++]));
            }
        }
        return buffer;
    }

    public static bool TryDecodeHexString(string value, out byte[] bytes, bool reverse = false)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            bytes = null;
            return false;
        }

        value = value.Trim();
        var length = value.Length;
        if ((length & 0x01) != 0) // 不是偶数
        {
            bytes = null;
            return false;
        }

        length /= CharsPerByte; // 字节集合大小
        var buffer = new byte[length];
        // var digits = value.ToCharArray(); // 转成字符数组访问性能较快 但是增加内存消耗 以及垃圾回收.
        var c = 0; // 字符索引
        if (!reverse)
        {
            for (var b = 0; b < length; b++) // 字节索引
            {
                int high;
                if (!TryParseHexDigit(value[c++], out high))
                {
                    bytes = null;
                    return false;
                }
                int low;
                if (!TryParseHexDigit(value[c++], out low))
                {
                    bytes = null;
                    return false;
                }

                buffer[b] = (byte)((high << 0x04) | low);
            }
        }
        else
        {
            while (length > 0)
            {
                int high;
                if (!TryParseHexDigit(value[c++], out high))
                {
                    bytes = null;
                    return false;
                }
                int low;
                if (!TryParseHexDigit(value[c++], out low))
                {
                    bytes = null;
                    return false;
                }

                buffer[--length] = (byte)((high << 0x04) | low);
            }
        }
        bytes = buffer;
        return true;
    }

    // 范围: [0x00~0x0F]
    static int ParseHexDigit(char digit)
    {
        // 0x30~0x39
        if (digit <= '9' && digit >= '0')
            return digit - '0';
        // 0x41~0x5A
        if (digit <= 'F' && digit >= 'A')
            return digit - 'A' + '\n'; // '\n' = 10
        // 0x61~0x7A
        if (digit <= 'f' && digit >= 'a')
            return digit - 'a' + '\n'; // '\n' = 10
        throw new ArgumentException();
    }
    static bool TryParseHexDigit(char digit, out int value)
    {
        // 0x30~0x39
        if (digit <= '9' && digit >= '0')
        {
            value = digit - '0';
            return true;
        }
        // 0x41~0x5A
        if (digit <= 'F' && digit >= 'A')
        {
            value = digit - 'A' + '\n'; // '\n' = 10
            return true;
        }
        // 0x61~0x7A
        if (digit <= 'f' && digit >= 'a')
        {
            value = digit - 'a' + '\n'; // '\n' = 10
            return true;
        }
        value = -1;
        return false;
    }

    public static bool IsHexChar(char value)
    {
        return (value <= '9' && value >= '0') || (value <= 'F' && value >= 'A') || (value <= 'f' && value >= 'a');
    }

    public static bool IsHexChar(string value, int index)
    {
        // if (value == null)
        //     throw new ArgumentNullException();
        if (string.IsNullOrEmpty(value))
            return false;
        if (index < 0 || index >= value.Length)
            throw new ArgumentOutOfRangeException(nameof(index));

        return IsHexChar(value[index]);
    }

    /// <summary>
    /// 是否 十六进制的字符串
    /// </summary>
    /// <param name="value"></param>
    /// <returns>
    /// String.Empty 返回假值.
    /// </returns>
    public static bool IsHexString(string value)
    {
        if (string.IsNullOrEmpty(value))
            return false;

        foreach (var c in value)
            if (!IsHexChar(c))
                return false;

        return true;
    }

    public static bool EqualsValue(byte[] x, byte[] y)
    {
        return (x == null && y == null) || (x != null && y != null && x.Length == y.Length && x.SequenceEqual(y, EqualityComparer<byte>.Default));
    }
    public static bool EqualsValue(string x, string y)
    {
        return string.Equals(x, y, StringComparison.OrdinalIgnoreCase);
    }
    public static bool EqualsValue(string hex, byte[] bytes, bool reverse = false)
    {
        return (hex == null && bytes == null)
            || (hex != null && bytes != null && hex.Length == bytes.Length * CharsPerByte && string.Equals(hex, EncodeHexString(bytes, reverse: reverse), StringComparison.OrdinalIgnoreCase));
    }
}
//}