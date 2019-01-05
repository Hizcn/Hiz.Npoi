using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public partial class NpoiColor
    {
        /// <summary>
        /// Excel 标准调色盘的颜色索引
        /// </summary>
        public short Indexed { get; private set; }

        /// <summary>
        /// 颜色名称
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 颜色的值; 等效 System.Drawing.Color.ToArgb();
        /// </summary>
        public int ArgbValue { get; private set; }

        byte[] _Argb;
        /// <summary>
        /// 如果不含透明通道(Alpha = 0xFF) 那么 Argb.Length = 3 (等效 IndexedColors.RGB);
        /// 否则 = 4; Automatic 始终 = 4;
        /// </summary>
        public byte[] Argb
        {
            get
            {
                if (_Argb == null)
                {
                    var argb = this.ArgbValue;
                    var alpha = (argb >> 0x18) & 0xFF;
                    if (alpha != 0xFF)
                    {
                        _Argb = new byte[] { (byte)alpha, (byte)((argb >> 0x10) & 0xFF), (byte)((argb >> 0x08) & 0xFF), (byte)(argb & 0xFF) };
                    }
                    else
                    {
                        _Argb = new byte[] { (byte)((argb >> 0x10) & 0xFF), (byte)((argb >> 0x08) & 0xFF), (byte)(argb & 0xFF) };
                    }
                }
                return _Argb;
            }
        }

        /// <summary>
        /// 是否含有 透明 通道.
        /// </summary>
        public bool HasAlpha
        {
            get
            {
                return ((this.ArgbValue >> 0x18) & 0xFF) != 0xFF;
            }
        }

        /// <summary>
        /// 是否 自动颜色
        /// </summary>
        public bool IsAutomatic
        {
            get
            {
                return this.Indexed == IndexedColors.Automatic.Index;
            }
        }

        /// <summary>
        /// 是否 黑色
        /// </summary>
        public bool IsBlack
        {
            get
            {
                return this.Indexed == IndexedColors.Black.Index;
            }
        }

        /// <summary>
        /// Alpha
        /// </summary>
        public byte A
        {
            get
            {
                // return (byte)((this.ArgbValue >> 0x18) & 0xFF);
                return GetAlpha(this.Argb);
            }
        }
        /// <summary>
        /// Red
        /// </summary>
        public byte R
        {
            get
            {
                // return (byte)((this.ArgbValue >> 0x10) & 0xFF);
                return GetRed(this.Argb);
            }
        }
        /// <summary>
        /// Green
        /// </summary>
        public byte G
        {
            get
            {
                // return (byte)((this.ArgbValue >> 0x08) & 0xFF);
                return GetGreen(this.Argb);
            }
        }
        /// <summary>
        /// Blue
        /// </summary>
        public byte B
        {
            get
            {
                // return (byte)(this.ArgbValue & 0xFF);
                return GetBlue(this.Argb);
            }
        }

        /// <summary>
        /// 获取 十六进制 文本
        /// </summary>
        /// <param name="lower"></param>
        /// <param name="alpha"></param>
        /// <returns></returns>
        public string GetHexString(bool lower = false, bool? alpha = null)
        {
            const string FormatUpper8 = "X8";
            const string FormatUpper6 = "X6";
            const string FormatLower8 = "x8";
            const string FormatLower6 = "x6";

            string format;
            if (alpha == true)
            {
                format = lower ? FormatLower8 : FormatUpper8;
            }
            else if (alpha == false)
            {
                format = lower ? FormatLower6 : FormatUpper6;
            }
            else
            {
                format = this.HasAlpha ? (lower ? FormatLower8 : FormatUpper8) : (lower ? FormatLower6 : FormatUpper6);
            }

            return "#" + this.ArgbValue.ToString(format);
        }

        public override string ToString()
        {
            // return base.ToString();
            if (!string.IsNullOrEmpty(this.Name))
                return this.Name;
            return this.GetHexString();
        }

        #region 类型转换

        /* 隐式转换
         * NpoiColor c1 = "Red";
         * NpoiColor c2 = "#ff0000";
         * NpoiColor c3 = 0xFFFF0000;
         * NpoiColor c4 = IndexedColors.Red;
         */
        public static implicit operator NpoiColor(string hex)
        {
            // if (NpoiColor.TryParse(hex, out NpoiColor color) < 0)
            //     throw new ArgumentException();

            // 不抛异常 转换失败 返回空值;
            NpoiColor.TryParse(hex, out NpoiColor color);
            return color;
        }

        public static implicit operator NpoiColor(int argb)
        {
            NpoiColor.TryGetColor(argb, out NpoiColor color);
            return color;
        }
        public static implicit operator NpoiColor(uint argb) // [0x80000000 - 0xFFFFFFFF]
        {
            NpoiColor.TryGetColor((int)argb, out NpoiColor color);
            return color;
        }

        //public static implicit operator NpoiColor(NpoiColorIndex value)
        //{
        //    if (NpoiColor.TryGetColor(value, out NpoiColor color) < 0)
        //        throw new ArgumentException();
        //    return color;
        //}
        public static implicit operator NpoiColor(IndexedColors value)
        {
            _MappingIndexed.TryGetValue(value.Index, out NpoiColor color);
            return color;
        }

        #endregion

        #region Equals

        public override int GetHashCode()
        {
            // return base.GetHashCode();
            return this.ArgbValue;
        }

        public override bool Equals(object other)
        {
            if (other != null)
            {
                if (other is NpoiColor)
                    return this.Equals((NpoiColor)other);
                if (other is byte[])
                    return this.Equals((byte[])other);
                if (other is XSSFColor)
                    return this.Equals((XSSFColor)other);
            }
            return false;
        }

        public virtual bool Equals(NpoiColor other)
        {
            return other != null && this.ArgbValue == other.ArgbValue;
        }
        public virtual bool Equals(byte[] argb)
        {
            return argb != null && this.ArgbValue == GetArgbValue(argb);
        }
        public virtual bool Equals(XSSFColor xssf)
        {
            return xssf != null
                && ((this.Indexed > 0 && this.Indexed == xssf.Indexed/*0: Not Specified Indexed*/) || this.ArgbValue == GetArgbValue(xssf.GetARgb()));
        }

        #endregion
    }
}
