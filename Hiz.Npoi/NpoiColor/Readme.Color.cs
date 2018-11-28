using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    /* POI.Color
     * http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/Color.html
     * Color (Interface)
     * |=> HSSFColor
     * |=> ExtendedColor
     *     |=> XSSFColor
     *     |=> HSSFExtendedColor
     * 
     * POI.HSSFColorPredefined(Enum)
     * http://poi.apache.org/apidocs/dev/org/apache/poi/hssf/util/HSSFColor.HSSFColorPredefined.html
     * 
     * POI.IndexedColors(Enum)
     * https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/IndexedColors.html
     * 含有备用索引:
     * Black
     * Blue
     * BrightGreen
     * LightTurquoise
     * Pink
     * Red
     * Turquoise
     * White
     * Yellow
     */

    /* IWorkbook/HSSFWorkbook/XSSFWorkbook 没有颜色相关方法.
     * 
     * HSSFWorkbook 通过 HSSFPalette 添加颜色: {
     *     HSSFPalette GetCustomPalette(); // 
     * }
     * HSSFPalette {
     *     PaletteRecord palette;
     *     // 查找指定颜色
     *     HSSFColor FindColor(byte red, byte green, byte blue) {
     *         for(var i = 8; i < 64; i++) {
     *             var color = this.palette.GetColor(i);
     *             if (color != null) {
     *                 if (color[0] == red && color[1] == green && color[2] == blue)
     *                     return new HSSFPalette.CustomColor(i, color);
     *             }
     *         }
     *         return null;
     *     }
     *     // 查找相近颜色: Min(Abs(red - color[0]) + Abs(green - color[1]) + Abs(blue - color[2]));
     *     HSSFColor FindSimilarColor(byte red, byte green, byte blue);
     *     
     *     // 获取颜色
     *     HSSFColor GetColor(short index) {
     *         if (index == 64)
     *             return HSSFColor.Automatic.GetInstance();
     *         byte[] color = this.palette.GetColor(index);
     *         if (color != null)
     *             return new HSSFPalette.CustomColor(index, color);
     *         return null;
     *     }
     *     byte[] PaletteRecord.GetColor(short byteIndex) {
     *         var index = byteIndex - 8;
     *         if (index < 0 || index >= this.Count)
     *             return null;
     *         PColor color = this[index];
     *         return new byte[] { color._red, color._green, color._blue };
     *     }
     *     
     *     // 设置颜色
     *     void SetColorAtIndex(short index, byte red, byte green, byte blue) {
     *         this.palette.SetColor(index, red, green, blue);
     *     }
     *     // 可以看出, 不按顺序设置颜色, 将会造成空间浪费...
     *     void PaletteRecord.SetColor(short byteIndex, byte red, byte green, byte blue) {
     *         var index = byteIndex - 8;
     *         while (this.Count <= index) {
     *             this.Add(new PColor(0, 0, 0));
     *         }
     *         this[index] = new PColor(red, green, blue);
     *     }
     *     
     *     // 添加颜色
     *     HSSFColor AddColor(byte red, byte green, byte blue) {
     *         for(var i = 8; i < 64; i++) {
     *             bytes color = this.palette.GetColor(i);
     *             if (color == null) { // 找到第一个颜色为空的位置
     *                 this.SetColorAtIndex(i, red, green, blue);
     *                 return this.GetColor(i);
     *             }
     *         }
     *         // 只能有 56 个颜色, 超出将抛异常..
     *         throw new System.Exception("Could not Find free color index");
     *     }
     * }
     * PaletteRecord {
     *     STANDARD_PALETTE_SIZE = 56; // 只能存放 56 个颜色值;
     *     FIRST_COLOR_INDEX = 8; // 第一个颜色的索引; 即: 有效索引范围 [8-63].
     * }
     * 
     * XSSFWorkbook 颜色随意添加
     */

    /* 关于颜色:
     * NPOI.SS.UserModel.IndexedColors 预设了 48 种颜色 (包含一个自动颜色 Automatic.Index = 64));
     * IndexedColors.HexString = "#RRGGBB";
     * 
     * NPOI.HSSF.Util.HSSFColor 定义等同 IndexedColors, 但是一些颜色含有备用索引(Index2), 如下:
     * Index/Index2 : Name
     *    11/35     : BrightGreen // BUG: 35 重复定义;
     *    12/39     : Blue
     *    13/34     : Yellow
     *    14/33     : Pink
     *    15/35     : Turquoise
     *    16/37     : DarkRed
     *    18/32     : DarkBlue
     *    20/36     : Violet
     *    21/38     : Teal
     *    41/27     : LightTurquoise
     *    61/25     : Plum // BUG: Maroon(25).
     * NPOIv2.4 BUG: 目前 Index2 的定义会造成 Maroon 颜色丢失;
     * 使用 IndexedColors.ValueOf(Index2) 是无法获得颜色的.
     * 
     * HSSFColor 不支持透明色; HSSFColor.HexString = "RRRR:GGGG:BBBB"; 例如 "#200080" => "2020:0:8080"; 有点冗余;
     * 
     * HSSFPalette 只能存放 56 个颜色值, 起始索引: 8, 截至索引 63; 以及一个自动颜色索引 Automatic.Index = 64 (不占存储空间).
     * 默认已经填满预设颜色(IndexedColors), 共 56 个, 其中九个颜色重复(含备用索引的颜色), 如需自定颜色只能替换原值...
     * 源码详见: PaletteRecord.CreateDefaultPalette() // Standard
     * 
     * NPOI.PaletteRecord.CreateDefaultPalette()
     * 含有重复颜色: (比 HSSFColor 少一对: 11/35)
     * Index     : Name
     * 12/39     : Blue
     * 13/34     : Yellow
     * 14/33     : Pink
     * 15/35     : Turquoise
     * 16/37     : DarkRed
     * 18/32     : DarkBlue
     * 20/36     : Violet
     * 21/38     : Teal
     * 41/27     : LightTurquoise
     * 61/25     : Plum
     * NPOIv2.4 BUG: Maroon 颜色丢失, 实际不重复颜色数: 46;
     * 
     * XSSFColor 支持透明? XSSFColor.GetARGBHex() = "AARRGGBB";
     */

    /* NPOI.SS.UserModel.IColor
     * |=> NPOI.HSSF.Util.HSSFColor
     * |=> NPOI.XSSF.UserModel.XSSFColor // 支持透明通道: Alpha
     * 
     * HSSFColor {
     *     public const short Index = x;
     *     public const short Index2 = x; // 部分颜色拥有两个索引;
     *     public virtual short Indexed { get { return Index; } }
     *     public byte[] RGB { get { return this.GetTriplet(); } }
     *     public override byte[] GetTriplet() { return this.Triplet; }
     *     public static readonly byte[] Triplet = new byte[] { 0, 0, 0 };
     *     public const string HexString = ""; // 十六进制大写格式 "RRRR:GGGG:BBBB" 例如绿色: 0x008000 => "0:8080:0" (零:0, 其它数值重复输出 80: 8080)
     *     public override string GetHexString() { return HexString; }
     *     
     *     // 自动颜色
     *     // Special Default/Normal/Automatic color.
     *     // Note:
     *     // This class Is NOT in the default HashTables returned by HSSFColor.
     *     // The index Is a special case which Is interpreted in the various SetXXXColor calls.
     *     public class Automatic : HSSFColor {
     *         short Index = 64;
     *         byte[] GetTriplet() { return HSSFColor.Black.Triplet; }
     *         string GetHexString() { return "0:0:0" };
     *     }
     * }
     * 
     * XSSFColor {
     *     内部数据存储: CT_Color {
     *         // auto(bool)/autoSpecified 是否自动颜色?
     *         // indexed(uint)/indexedSpecified 预设颜色索引?
     *         // rgb(byte[])/rgbSpecified 定义颜色? // 包含透明通道: Alpha
     *         // theme(uint)/themeSpecified 主题颜色索引?
     *         // tint(double)/tintSpecified 颜色明暗
     *         void SetRgb(byte R, byte G, byte B) {
	 *             this.rgbField = new byte[4]; // 字节长度始终为四?
	 *             this.rgbField[0] = 0; // Alpha = 0; (BUG? 等于 255 才对.)
	 *             this.rgbField[1] = R;
	 *             this.rgbField[2] = G;
	 *             this.rgbField[3] = B;
	 *             this.rgbSpecified = true;
     *         }
     *         void SetRgb(byte[] rgb) {
     *             this.rgbField = new byte[rgb.Length]; // 字节长度可以为三为四?
     *             Array.Copy(rgb, this.rgbField, rgb.Length);
     *             this.rgbSpecified = true;
     *         }
     *         byte[] GetRgb() {
     *             if (this.rgbField == null)
     *                 return null;
     *             byte[] array = new byte[this.rgbField.Length];
     *             Array.Copy(this.rgbField, array, this.rgbField.Length);
     *             return array;
     *         }
     *     }
     *     
     *     public bool IsAuto { get; set; }
     *     public short Indexed  { get; set; } // Get: 未指定返回零.
     *     public byte[] RGB { get; } // 返回字节长度始终为三, 不含透明通道.
     *     public int Theme { get; set; } // Get: 未指定返回零.
     *     public double Tint { get; set; }
     *     
     *     public XSSFColor(System.Drawing.Color clr) { this.ctColor.SetRgb(clr.R, clr.G, clr.B); } // 忽略透明通道
     *     public XSSFColor(byte[] rgb) { this.ctColor.SetRgb(rgb); }  // 字节长度可以为三为四?
     *     
     *     CT_Color ctColor;
     *     private byte[] GetRGBOrARGB() {
     *         if (this.ctColor.indexedSpecified && this.ctColor.indexed > 0u) {
     *             var hssfColor = (HSSFColor)HSSFColor.GetIndexHash()[(int)this.ctColor.indexed];
     *             if (hssfColor != null)
     *                 return new byte[] { hssfColor.GetTriplet()[0], hssfColor.GetTriplet()[1], hssfColor.GetTriplet()[2] };
     *         }
     *         if (this.ctColor.IsSetRgb())
     *             return this.ctColor.GetRgb();
     *         return null;
     *     }
     *     public byte[] GetARgb() { } 调用 GetRGBOrARGB() 如果字节为三, 扩充为四, 设置透明通道 = 255;
     *     public string GetARGBHex() { } 调用 GetARgb(), 结果不含井号, 十六进制大写, 格式: AARRGGBB;
     *     public byte[] GetRgb() { return this.RGB; }
     *     public void SetRgb(byte[] rgb) { this.ctColor.SetRgb(rgb); } // 字节长度可以为三为四?
     *     public byte[] GetRgbWithTint() 调用 this.ctColor.GetRgb(); 如果字节为四, 缩减为三, 调用 ApplyTint(), 返回结果.
     * }
     * 
     * NPOI.SS.UserModel.IndexedColors {
     *     private IndexedColors(int idx, HSSFColor color) { }
     *     public short Index { get; }
     *     public byte[] RGB { get { return this.hssfColor.RGB; } }
     *     public string HexString { get; } // 十六进制小写格式: "#rrggbb";
     * }
     */
}
