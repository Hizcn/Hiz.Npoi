using Hiz.Extended.Npoi;
using NPOI.HSSF.Record;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Sample.ConsoleApp
{
    public static class TestColor
    {
        #region NpoiBug
        class PColor
        {
            public short Indexed { get; set; }
            public int ArgbValue { get; set; }
            public string Name { get; set; }
        }
        static int GetArgbValue(byte[] argb)
        {
            if (argb.Length == 3)
            {
                return (byte.MaxValue << 0x18) | (argb[0] << 0x10) | (argb[1] << 0x08) | argb[2];
            }
            if (argb.Length == 4)
            {
                return (argb[0] << 0x18) | (argb[1] << 0x10) | (argb[2] << 0x08) | argb[3];
            }
            throw new NotSupportedException();
        }
        public static void NpoiBug()
        {
            var indexedColors = typeof(IndexedColors)
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .Where(f => f.Name != "Automatic")
                .Select(f => new { Name = f.Name, Color = (IndexedColors)f.GetValue(null), })
                .Select(c => new PColor() { Name = c.Name, Indexed = c.Color.Index, ArgbValue = GetArgbValue(c.Color.RGB) })
                .ToDictionary(c => c.ArgbValue);

            var palette = new PaletteRecord();
            var paletteColors = new PColor[palette.NumColors];
            for (var i = 0; i < palette.NumColors; i++)
            {
                var index = (short)(i + 8);
                var bytes = palette.GetColor(index);
                var argb = GetArgbValue(bytes);
                var color = indexedColors[argb];
                paletteColors[i] = new PColor()
                {
                    Indexed = index,
                    Name = color.Name,
                    ArgbValue = argb,
                };
            }

            var set = new HashSet<string>(indexedColors.Values.Select(i => i.Name)); // 47 颜色;
            var paletteNames = paletteColors.Select(i => i.Name).Distinct().ToArray(); // 46 不重复的颜色;
            var paletteLost = set.Except(paletteNames).ToArray();
            if (paletteLost.Length > 0)
            {
                Console.WriteLine("CreateDefaultPalette() 相对 IndexedColors 缺失颜色:");
                foreach (var n in paletteLost)
                {
                    Console.WriteLine(n);
                }
                Console.WriteLine();
            }

            Console.WriteLine("HSSFColor.GetIndexHash() 相对 IndexedColors 错误颜色:");
            var hssf = HSSFColor.GetIndexHash();
            foreach (var c in indexedColors.Values)
            {
                var hssfColor = (HSSFColor)hssf[(int)c.Indexed];

                var hssfName = hssfColor.GetType().Name;
                if (hssfName != c.Name)
                    Console.WriteLine("IndexedColors.ValueOf({0}).Name={1}; HSSFColor.GetIndexHash()[{0}].Name={2}", c.Indexed, c.Name, hssfName);

                var hssfValue = GetArgbValue(hssfColor.RGB);
                if (hssfValue != c.ArgbValue)
                    Console.WriteLine("IndexedColors.ValueOf({0}).RGB={1}; HSSFColor.GetIndexHash()[{0}].RGB={2}", c.Indexed, c.ArgbValue.ToString("X8"), hssfValue.ToString("X8"));
            }
            Console.WriteLine();


            //Console.WriteLine("HSSFColor.GetIndexHash() 与 CreateDefaultPalette() 的差异");
            //var hssfColors = HSSFColor.GetIndexHash().Cast<DictionaryEntry>()
            //    .Select(p => new PColor() { Indexed = (short)(int)p.Key, ArgbValue = GetArgbValue(((HSSFColor)p.Value).RGB), Name = p.Value.GetType().Name })
            //    .ToDictionary(c => c.Indexed);
            //for (var i = 0; i < palette.NumColors; i++)
            //{
            //    var p = paletteColors[i];
            //    var h = hssfColors[p.Indexed];

            //    if (p.Name != h.Name)
            //        Console.WriteLine("DefaultPalette[{0}].Name={1}; HSSFColor.GetIndexHash()[{0}].Name={2}", p.Indexed, p.Name, h.Name);
            //    if (p.ArgbValue != h.ArgbValue)
            //        Console.WriteLine("DefaultPalette[{0}].RGB={1}; HSSFColor.GetIndexHash()[{0}].RGB={2}", p.Indexed, p.ArgbValue.ToString("X8"), h.ArgbValue.ToString("X8"));
            //}
        }

        #endregion

        public static void TestHexString()
        {
            var v1 = 0x01020304;
            var bytes1 = BitConverter.GetBytes(v1); // 正序; LittleEndian;
            var hex1 = v1.ToString("X8"); // 倒序文本;
            var hex2 = BitConverter.ToString(bytes1).Replace("-", null); // 正序文本;
            var hex3 = Hex.EncodeHexString(bytes1, reverse: true); // 倒序文本
            var hex4 = Hex.EncodeHexString(bytes1, reverse: false); // 正序文本
            Debug.Assert(hex1 == hex3);
            Debug.Assert(hex2 == hex4);
            var bytes1a = Hex.DecodeHexString(hex4, reverse: false);
            var bytes1b = Hex.DecodeHexString(hex3, reverse: true);
            var v1a = BitConverter.ToInt32(bytes1a, 0);
            var v1b = BitConverter.ToInt32(bytes1a, 0);
            Debug.Assert(v1 == v1a);
            Debug.Assert(v1 == v1b);

            var v2 = 0x0102030405060708;
            var bytes2 = BitConverter.GetBytes(v2);
            var hex5 = v2.ToString("X16"); // 倒序文本;
            var hex6 = BitConverter.ToString(bytes2).Replace("-", null); // 正序文本;
            var hex7 = Hex.EncodeHexString(bytes2, reverse: true); // 倒序文本
            var hex8 = Hex.EncodeHexString(bytes2, reverse: false); // 正序文本
            Debug.Assert(hex5 == hex7);
            Debug.Assert(hex6 == hex8);
            var bytes2a = Hex.DecodeHexString(hex8, reverse: false);
            var bytes2b = Hex.DecodeHexString(hex7, reverse: true);
            var v2a = BitConverter.ToInt64(bytes2a, 0);
            var v2b = BitConverter.ToInt64(bytes2b, 0);
            Debug.Assert(v2 == v2a);
            Debug.Assert(v2 == v2b);
        }

        public static void TestColorOperators()
        {
            NpoiColor c1 = "Red";
            NpoiColor c2 = "#ff0000";
            NpoiColor c3 = 0xFFFF0000;
            NpoiColor c4 = IndexedColors.Red;
            //NpoiColor c5 = NpoiColorIndex.Red;

            // NpoiColor c6 = new HSSFColor.Red();
            // NpoiColor c7 = new XSSFColor(new byte[] { 0xFF, 0, 0 });

        }

        public static void TestColorHexString()
        {
            var c1 = IndexedColors.Plum;

            var hex0 = c1.HexString.TrimStart('#'); // 993366

            var bytes1 = Hex.DecodeHexString(hex0); // new [] { 0x99, 0x33, 0x66 };
            var value1 = NpoiColor.GetArgbValue(bytes1, false).ToString("X6"); // "993366"
            var b1 = NpoiColor.Equals(bytes1, c1.RGB, false);

            var c2 = System.Drawing.Color.FromArgb(c1.RGB[0], c1.RGB[1], c1.RGB[2]);
            var value2 = c2.ToArgb().ToString("X6"); // "FF993366"

            var indexed1 = NpoiColor.TryParse(hex0, out NpoiColor argb1);
            Debug.Assert(c1.Index == argb1.Indexed);
            Debug.Assert(c1.RGB.SequenceEqual(argb1.Argb));
            Debug.Assert(argb1.A == c2.A);
            Debug.Assert(argb1.R == c2.R);
            Debug.Assert(argb1.G == c2.G);
            Debug.Assert(argb1.B == c2.B);
        }
    }
}
