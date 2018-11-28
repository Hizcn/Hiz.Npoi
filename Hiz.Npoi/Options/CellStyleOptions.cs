using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    /* 引用资源:
     * ICellStyle: HSSFCellStyle/XSSFCellStyle
     * IDataFormat: HSSFDataFormat/XSSFDataFormat
     * IColor: HSSFColor/XSSFColor
     * IFont: HSSFFont/XSSFFont
     * 
     * xlsx: 新增
     * XSSFCellFill (PatternType; ForegroundColor; BackgroundColor;)
     * XSSFCellBorder
     * XSSFCellAlignment (Alignment; Indention; VerticalAlignment; WrapText; ShrinkToFit; Rotation;)
     * 
     * CT_CellProtection (IsLocked; IsHidden;)
     */

    /* ICellStyle
     * 
     * 派生:
     * NPOI.HSSF.UserModel.HSSFCellStyle (xls)
     * NPOI.XSSF.UserModel.XSSFCellStyle (xlsx)
     * 
     * 成员:
     * // 样式索引
     * short Index { get; }
     * // 克隆样式
     * void CloneStyleFrom(ICellStyle source);
     * 
     * 
     * 格式:
     * short DataFormat { get; set; }
     * string GetDataFormatString();
     * 
     * 
     * 对齐:
     * // 垂直对齐
     * VerticalAlignment VerticalAlignment { get; set; }
     * // 水平对齐
     * HorizontalAlignment Alignment { get; set; }
     * // 首行缩进 (搭配水平对齐属性)
     * short Indention { get; set; }
     * // 自动换行
     * bool WrapText { get; set; }
     * // 缩小字体填充
     * bool ShrinkToFit { get; set; }
     * // 文字旋转角度
     * short Rotation { get; set; }
     * 
     * 
     * 字体:
     * short FontIndex { get; }
     * IFont GetFont(IWorkbook parentWorkbook);
     * void SetFont(IFont font);
     * 
     * 
     * 边框:
     * // 左边框
     * BorderStyle BorderLeft { get; set; }
     * short LeftBorderColor { get; set; }
     * // 上边框
     * BorderStyle BorderTop { get; set; }
     * short TopBorderColor { get; set; }
     * // 右边框
     * BorderStyle BorderRight { get; set; }
     * short RightBorderColor { get; set; }
     * // 下边框
     * BorderStyle BorderBottom { get; set; }
     * short BottomBorderColor { get; set; }
     * // 对角线
     * BorderDiagonal BorderDiagonal { get; set; }
     * BorderStyle BorderDiagonalLineStyle { get; set; }
     * short BorderDiagonalColor { get; set; }
     * 
     * 
     * 图案:
     * FillPattern FillPattern { get; set; }
     * short FillForegroundColor { get; set; }
     * IColor FillForegroundColorColor { get; }
     * short FillBackgroundColor { get; set; }
     * IColor FillBackgroundColorColor { get; }
     * 
     * 
     * 保护:
     * // 是否锁定
     * bool IsLocked { get; set; }
     * // 是否隐藏
     * bool IsHidden { get; set; }
     */

    class CellStyleOptions : INamed
    {
        #region NPOI

        // // 格式相关
        // short ICellStyle.DataFormat { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // string ICellStyle.GetDataFormatString()
        // {
        //     throw new NotImplementedException();
        // }

        // // 字体相关
        // short ICellStyle.FontIndex => throw new NotImplementedException();
        // IFont ICellStyle.GetFont(IWorkbook parentWorkbook)
        // {
        //     throw new NotImplementedException();
        // }
        // void ICellStyle.SetFont(IFont font)
        // {
        //     throw new NotImplementedException();
        // }

        // // 对齐相关
        // HorizontalAlignment ICellStyle.Alignment { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.Indention { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // VerticalAlignment ICellStyle.VerticalAlignment { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // bool ICellStyle.WrapText { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // bool ICellStyle.ShrinkToFit { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.Rotation { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // // 边框相关
        // BorderStyle ICellStyle.BorderLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.LeftBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderStyle ICellStyle.BorderRight { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.RightBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderStyle ICellStyle.BorderTop { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.TopBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderStyle ICellStyle.BorderBottom { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.BottomBorderColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderDiagonal ICellStyle.BorderDiagonal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // BorderStyle ICellStyle.BorderDiagonalLineStyle { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.BorderDiagonalColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // // 图案相关
        // FillPattern ICellStyle.FillPattern { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // short ICellStyle.FillForegroundColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // IColor ICellStyle.FillForegroundColorColor => throw new NotImplementedException();
        // short ICellStyle.FillBackgroundColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // IColor ICellStyle.FillBackgroundColorColor => throw new NotImplementedException();

        // // 保护相关
        // bool ICellStyle.IsHidden { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        // bool ICellStyle.IsLocked { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        // // 其它
        // short ICellStyle.Index => throw new NotImplementedException();
        // void ICellStyle.CloneStyleFrom(ICellStyle source)
        // {
        //     throw new NotImplementedException();
        // }
        #endregion

        public string Name { get; set; }

        /// <summary>
        /// 格式
        /// </summary>
        public string DataFormat { get; set; }

        /// <summary>
        /// 对齐
        /// </summary>
        public string TextAlignment { get; set; }

        /// <summary>
        /// 字体
        /// </summary>
        public string Font { get; set; }

        /// <summary>
        /// 边框
        /// </summary>
        public string Border { get; set; }

        /// <summary>
        /// 图案
        /// </summary>
        public string Fill { get; set; }

        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool IsHidden { get; set; } = false;
        /// <summary>
        /// 是否锁定
        /// </summary>
        public bool IsLocked { get; set; } = true;
    }
}
