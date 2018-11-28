using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    //TODO: 颜色: 表示不作修改的值; 表示恢复默认的值; 表示清除属性的值;

    /// <summary>
    /// Excel 标准调色盘的颜色
    /// </summary>
    enum NpoiColorIndex : short
    {
        None = 0,

        Black = 8,
        White = 9,
        Red = 10,
        BrightGreen = 11,
        Blue = 12,
        Yellow = 13,
        Pink = 14,
        Turquoise = 15,
        DarkRed = 16,
        Green = 17,
        DarkBlue = 18,
        DarkYellow = 19,
        Violet = 20,
        Teal = 21,
        Grey25Percent = 22,
        Grey50Percent = 23,
        CornflowerBlue = 24,
        Maroon = 25,
        LemonChiffon = 26,

        Orchid = 28,
        Coral = 29,
        RoyalBlue = 30,
        LightCornflowerBlue = 31,

        SkyBlue = 40,
        LightTurquoise = 41,
        LightGreen = 42,
        LightYellow = 43,
        PaleBlue = 44,
        Rose = 45,
        Lavender = 46,
        Tan = 47,
        LightBlue = 48,
        Aqua = 49,
        Lime = 50,
        Gold = 51,
        LightOrange = 52,
        Orange = 53,
        BlueGrey = 54,
        Grey40Percent = 55,
        DarkTeal = 56,
        SeaGreen = 57,
        DarkGreen = 58,
        OliveGreen = 59,
        Brown = 60,
        Plum = 61,
        Indigo = 62,
        Grey80Percent = 63,

        Automatic = 64,
    }

    // 没有存在必要, 如需通过名称文本指定颜色, 直接使用枚举类型;
    static class NpoiColorNames
    {
        // public const string None = string.Empty;

        public const string Black = "Black";
        public const string White = "White";
        public const string Red = "Red";
        public const string BrightGreen = "BrightGreen";
        public const string Blue = "Blue";
        public const string Yellow = "Yellow";
        public const string Pink = "Pink";
        public const string Turquoise = "Turquoise";
        public const string DarkRed = "DarkRed";
        public const string Green = "Green";
        public const string DarkBlue = "DarkBlue";
        public const string DarkYellow = "DarkYellow";
        public const string Violet = "Violet";
        public const string Teal = "Teal";
        public const string Grey25Percent = "Grey25Percent";
        public const string Grey50Percent = "Grey50Percent";
        public const string CornflowerBlue = "CornflowerBlue";
        public const string Maroon = "Maroon";
        public const string LemonChiffon = "LemonChiffon";

        public const string Orchid = "Orchid";
        public const string Coral = "Coral";
        public const string RoyalBlue = "RoyalBlue";
        public const string LightCornflowerBlue = "LightCornflowerBlue";

        public const string SkyBlue = "SkyBlue";
        public const string LightTurquoise = "LightTurquoise";
        public const string LightGreen = "LightGreen";
        public const string LightYellow = "LightYellow";
        public const string PaleBlue = "PaleBlue";
        public const string Rose = "Rose";
        public const string Lavender = "Lavender";
        public const string Tan = "Tan";
        public const string LightBlue = "LightBlue";
        public const string Aqua = "Aqua";
        public const string Lime = "Lime";
        public const string Gold = "Gold";
        public const string LightOrange = "LightOrange";
        public const string Orange = "Orange";
        public const string BlueGrey = "BlueGrey";
        public const string Grey40Percent = "Grey40Percent";
        public const string DarkTeal = "DarkTeal";
        public const string SeaGreen = "SeaGreen";
        public const string DarkGreen = "DarkGreen";
        public const string OliveGreen = "OliveGreen";
        public const string Brown = "Brown";
        public const string Plum = "Plum";
        public const string Indigo = "Indigo";
        public const string Grey80Percent = "Grey80Percent";

        public const string Automatic = "Automatic";
    }
}