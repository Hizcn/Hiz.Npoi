using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Hiz.Npoi
{
    partial class NpoiColor
    {
        const short FirstIndexed = 0x08;
        public static readonly NpoiColor Automatic;
        public static readonly NpoiColor Black;

        // 只含 Npoi 标准颜色 (IndexedColors)
        static IDictionary<short, NpoiColor> _MappingIndexed; // Key: Color.Indexed; 数量: 47(不含备用索引);
        static IDictionary<string, NpoiColor> _MappingNamed; // Key: Color.Name; 不区分大小写;
        static IDictionary<int, NpoiColor> _MappingValued; // Key: Color.ArgbValue; 包含 Alpha 通道;
        static NpoiColor()
        {
            // NpoiColor 定义 Automatic = "#00000000" (Alpha = 0);
            Automatic = new NpoiColor() { Indexed = IndexedColors.Automatic.Index, Name = "Automatic", ArgbValue = 0, _Argb = new byte[4] };

            _MappingIndexed = typeof(IndexedColors)
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .Where(f => f.Name != Automatic.Name) // 因为 IndexedColors.Automatic.HexString(#FF000000) == Black.HexString 因此这里排除 Automatic;
                .Select(f => new { Name = f.Name, Color = (IndexedColors)f.GetValue(null) })
                .Select(c => new NpoiColor() { Name = c.Name, Indexed = c.Color.Index, _Argb = c.Color.RGB, ArgbValue = GetArgbValue(c.Color.RGB) })
                .ToDictionary(p => p.Indexed);
            _MappingIndexed.Add(Automatic.Indexed, Automatic);

            Black = _MappingIndexed[IndexedColors.Black.Index];

            _MappingNamed = _MappingIndexed.Values.ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);
            _MappingValued = _MappingIndexed.Values.ToDictionary(p => p.ArgbValue);
        }

        #region TryGetColor
        public static short TryGetColor(int argb, out NpoiColor color)
        {
            if (_MappingValued.TryGetValue(argb, out NpoiColor valued))
            {
                color = valued;
                return valued.Indexed;
            }
            color = new NpoiColor() { ArgbValue = argb, };
            return 0;
        }
        static short TryGetColor(short indexed, out NpoiColor color)
        {
            if (_MappingIndexed.TryGetValue(indexed, out NpoiColor value))
            {
                color = value;
                return value.Indexed;
            }

            // 不是有效 StandardColor;
            color = null;
            return -1;
        }
        #endregion

        #region TryParse
        /// <summary>
        /// 解析颜色名字 或者 十六进制文本
        /// </summary>
        /// <param name="value">颜色名称(不区分大小写) 或者 十六进制文本(RRGGBB/AARRGGBB/#RRGGBB/#AARRGGBB; 不区分大小写)</param>
        /// <param name="color"></param>
        /// <returns>如果该颜色是标准颜色, 返回色盘索引 IndexedColors.Index; 0: Parse successful but not IndexedColor; -1: Invalid String;</returns>
        public static short TryParse(string value, out NpoiColor color)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                if (value[0] == '#')
                {
                    value = value.TrimStart('#'); // 移除十六进制文本可选前缀;
                }
                if (value.Length == 0x06 || value.Length == 0x08) // "RRGGBB" / "AARRGGBB"
                {
                    if (Hex.IsHexString(value)) // 是否十六进制文本
                    {
                        var bytes = Hex.DecodeHexString(value, false);
                        var argb = GetArgbValue(bytes);

                        if (_MappingValued.TryGetValue(argb, out NpoiColor valued))
                        {
                            color = valued; // 该文本匹配到标准颜色;
                            return valued.Indexed;
                        }
                        else
                        {
                            color = new NpoiColor() { ArgbValue = argb, _Argb = bytes };
                            return 0;
                        }
                    }
                }
                if (_MappingNamed.TryGetValue(value, out NpoiColor named)) // 匹配名称
                {
                    color = named;
                    return named.Indexed;
                }
            }

            /* 1. 空白文本;
             * 2. 不是有效十六进制文本;
             * 3. 不是标准颜色名称;
             */
            color = null;
            return -1;
        }
        static short TryParse(string value, out byte[] argb)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                if (string.Equals(value, Automatic.Name, StringComparison.OrdinalIgnoreCase))
                {
                    argb = Automatic.Argb;
                    return Automatic.Indexed;
                }
                if (value[0] == '#')
                {
                    value = value.TrimStart('#'); // 移除十六进制文本可选前缀;
                }
                if (value.Length == 0x06 || value.Length == 0x08) // "RRGGBB" / "AARRGGBB"
                {
                    if (Hex.IsHexString(value)) // 是否十六进制文本
                    {
                        argb = Hex.DecodeHexString(value, false);

                        if (_MappingValued.TryGetValue(GetArgbValue(argb), out NpoiColor p))
                            return p.Indexed; // 该文本匹配到标准颜色;
                        return 0;
                    }
                }
                if (_MappingNamed.TryGetValue(value, out NpoiColor named)) // 匹配名称
                {
                    argb = named.Argb;
                    return named.Indexed;
                }
            }

            argb = null;
            return -1;
        }
        static short TryParse(string value, out byte alpha, out byte red, out byte green, out byte blue)
        {
            var indexed = TryParse(value, out byte[] argb);
            if (argb != null)
            {
                GetArgb(argb, out alpha, out red, out green, out blue);
            }
            else
            {
                alpha = 0;
                red = 0;
                green = 0;
                blue = 0;
            }
            return indexed;
        }
        #endregion

        #region Argb

        /* Length == 3: [0]=R; [1]=G; [2]=B
         * Length == 4: [0]=A; [1]=R; [2]=G; [3]=B
         */
        const int OffsetAlpha4 = 0;
        const int OffsetRed4 = 1;
        const int OffsetGreen4 = 2;
        const int OffsetBlue4 = 3;
        const int OffsetRed3 = 0;
        const int OffsetGreen3 = 1;
        const int OffsetBlue3 = 2;
        const int ArgbLength4 = 4;
        const int ArgbLength3 = 3;

        /// <summary>
        /// 等效 System.Drawing.Color.ToArgb();
        /// </summary>
        /// <param name="argb">argb 字节数组仅限 3/4 长度</param>
        /// <param name="alpha">是否包含 Alpha 通道; 如果 alpha = true 但是 argb 参数不含 Alpha 通道, 则将 Alpha 当作 0xFF;</param>
        /// <returns>alpha = true: 0xAARRGGBB; alpha = false: 0x00RRGGBB.</returns>
        public static int GetArgbValue(byte[] argb, bool alpha = true)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));
            if (argb.Length == 3)
            {
                // BigEndian
                return ((alpha ? byte.MaxValue : byte.MinValue) << 0x18) | (argb[OffsetRed3] << 0x10) | (argb[OffsetGreen3] << 0x08) | argb[OffsetBlue3];
            }
            if (argb.Length == 4)
            {
                return ((alpha ? argb[OffsetAlpha4] : byte.MinValue) << 0x18) | (argb[OffsetRed4] << 0x10) | (argb[OffsetGreen4] << 0x08) | argb[OffsetBlue4];
            }
            throw new ArgumentException(nameof(argb));
        }
        public static byte GetAlpha(byte[] argb)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));
            if (argb.Length == 3)
                return byte.MaxValue;
            if (argb.Length == 4)
                return argb[OffsetAlpha4];
            throw new ArgumentException(nameof(argb));
        }
        public static byte GetRed(byte[] argb)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));
            if (argb.Length == 3)
                return argb[OffsetRed3];
            if (argb.Length == 4)
                return argb[OffsetRed4];
            throw new ArgumentException(nameof(argb));
        }
        public static byte GetGreen(byte[] argb)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));
            if (argb.Length == 3)
                return argb[OffsetGreen3];
            if (argb.Length == 4)
                return argb[OffsetGreen4];
            throw new ArgumentException(nameof(argb));
        }
        public static byte GetBlue(byte[] argb)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));
            if (argb.Length == 3)
                return argb[OffsetBlue3];
            if (argb.Length == 4)
                return argb[OffsetBlue4];
            throw new ArgumentException(nameof(argb));
        }
        public static void GetArgb(byte[] argb, out byte alpha, out byte red, out byte green, out byte blue)
        {
            if (argb == null)
                throw new ArgumentNullException(nameof(argb));

            if (argb.Length == 3)
            {
                alpha = byte.MaxValue;
                red = argb[OffsetRed3];
                green = argb[OffsetGreen3];
                blue = argb[OffsetBlue3];
            }
            if (argb.Length == 4)
            {
                alpha = argb[OffsetAlpha4];
                red = argb[OffsetRed4];
                green = argb[OffsetGreen4];
                blue = argb[OffsetBlue4];
            }
            throw new ArgumentException(nameof(argb));
        }

        #endregion

        /// <summary>
        /// 比较颜色
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="alpha">是否比较 Alpha 通道</param>
        /// <returns></returns>
        public static bool Equals(byte[] x, byte[] y, bool alpha = false)
        {
            // if (x != null && y != null)
            // {
            //     if (x.Length != 3 && x.Length != 4)
            //         throw new ArgumentException(nameof(x));
            //     if (y.Length != 3 && y.Length != 4)
            //         throw new ArgumentException(nameof(y));
            //     return GetArgbValue(x, alpha) == GetArgbValue(y, alpha);
            // }
            // if (x == null && y == null)
            //     return true;
            // return false;
            return (x == null && y == null) || (x != null && y != null && GetArgbValue(x, alpha) == GetArgbValue(y, alpha));
        }
    }
}
