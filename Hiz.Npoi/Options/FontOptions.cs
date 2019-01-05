using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    /* NPOI.SS.UserModel.IFont
     * NPOI.SS.UserModel.IFontFormatting
     * 
     * Java POI: Interface Font
     * https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Font.html
     * 
     * System.Drawing.Font
     * https://msdn.microsoft.com/zh-cn/library/system.drawing.font(v=vs.100).aspx
     */

    /// <summary>
    /// 字体配置
    /// </summary>
    public class FontOptions : INamed
    {
        #region NPOI

        // double IFont.FontHeight { get => this.FontHeightInPoints * 20f; set => this.FontHeightInPoints = (float)value / 20f; }
        // short IFont.FontHeightInPoints { get => (short)this.FontHeightInPoints; set => this.FontHeightInPoints = (float)value; }
           
        // // NPOI.SS.UserModel.FontBoldWeight
        // // Java POI: Scheduled for removal in 3.17. Use setBold(boolean).
        // short IFont.Boldweight { get => this.IsBold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal; set => this.IsBold = value == (short)FontBoldWeight.Bold; }
           
        // // GDI 字符集 (byte System.Drawing.Font.GdiCharSet)
        // short IFont.Charset { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
           
        // short IFont.Index => throw new NotSupportedException();
           
        // void IFont.CloneStyleFrom(IFont font) { throw new NotSupportedException(); }
           
        // short IFont.Color { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

        #endregion

        public string Name { get; set; }

        string _FontName = null;
        /// <summary>
        /// 字体名称 (System.Drawing.Font.Name)
        /// </summary>
        public string FontName
        {
            get
            {
                return this._FontName;
            }
            set
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException(nameof(value));
                this._FontName = value;
            }
        }

        /* 接口 FontHeightInPoints 定义:
         * Java POI: Interface Font {
         *   short getFontHeight();
         *   void setFontHeight(short height);
         *   short getFontHeightInPoints();
         *   void setFontHeightInPoints(short height);
         * }
         * 
         * NPOI: IFont {
         *   double FontHeight { get; set; }
         *   short FontHeightInPoints { get; set; }
         * }
         * 
         * 类型定义:
         * NPOI.HSSF.UserModel.HSSFFont {
         *   内部数据类型: FontRecord;
         *   字体大小存储 short: FontRecord.FontHeight; (对应: IFont.FontHeight);
         * }
         * 
         * NPOI.XSSF.UserModel.XSSFFont {
         *   内部数据类型: CT_Font;
         *   字体大小存储 double: ((CT_FontSize)CT_Font.GetSzArray(0)).val; (对应: IFont.FontHeightInPoints)
         * }
         * 
         * XSSFFont.FontHeightInPoints 默认值: 11.0 磅;
         * 
         * NPOI v2.2.1 BUG:
         * NPOI.XSSF.UserModel.XSSFFont.FontHeight: set 逻辑错误...
         * 因此使用 FontHeightInPoints 修改字体大小;
         * 
         * 综合上诉做了如下修改:
         * FontHeight 数据类型改为 short;
         * FontHeightInPoints 数据类型改为 float; 对应 System.Drawing.Font.SizeInPoints 数据类型;
         */
        float _FontHeightInPoints = 0;
        /// <summary>
        /// 字体大小 (float System.Drawing.Font.SizeInPoints)
        /// </summary>
        public float FontHeightInPoints
        {
            get
            {
                return this._FontHeightInPoints;
            }
            set
            {
                if (value <= 0f)
                    throw new ArgumentOutOfRangeException(nameof(value));
                this._FontHeightInPoints = value;
            }
        }

        /// <summary>
        /// 粗体 (bool System.Drawing.Font.Bold)
        /// </summary>
        public bool IsBold { get; set; } = false;

        /// <summary>
        /// 斜体 (bool System.Drawing.Font.Italic)
        /// </summary>
        public bool IsItalic { get; set; } = false;

        /// <summary>
        /// 删除线条 (bool System.Drawing.Font.Strikeout)
        /// </summary>
        public bool IsStrikeout { get; set; } = false;

        /// <summary>
        /// 下划线条 (bool System.Drawing.Font.Underline)
        /// </summary>
        public FontUnderlineType Underline { get; set; } = FontUnderlineType.None;

        /// <summary>
        /// 上标下标 (System.Drawing.Font NotSupported)
        /// </summary>
        public FontSuperScript TypeOffset { get; set; } = FontSuperScript.None;

        /// <summary>
        /// 字体颜色
        /// </summary>
        public NpoiColor Color { get; set; } = NpoiColor.Black;

        /// <summary>
        /// 字体配置
        /// </summary>
        /// <param name="name">字体名称</param>
        /// <param name="size">字号 (FontHeightInPoints)</param>
        public FontOptions(string name, float size)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException(nameof(name));
            if (size <= 0f)
                throw new ArgumentOutOfRangeException(nameof(name));

            this._FontName = name;
            this._FontHeightInPoints = size;
        }
    }
}
