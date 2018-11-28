using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    [Flags]
    public enum BorderEdges
    {
        None = 0,

        Top = 0x01,
        Right = 0x02,
        Bottom = 0x04,
        Left = 0x08,

        /// <summary>
        /// 所有四周边框
        /// </summary>
        AllAround = Top | Right | Bottom | Left,

        /// <summary>
        /// 逆对角线 (左上角 => 右下角) (left-top to right-bottom)
        /// NPOI.SS.UserModel.BorderDiagonal.Backward
        /// </summary>
        LeftTopToRightBottom = 0x0100,
        /// <summary>
        /// 正对角线 (右上角 => 左下角) (right-top to left-bottom)
        /// NPOI.SS.UserModel.BorderDiagonal.Forward
        /// </summary>
        RightTopToLeftBottom = 0x0200,

        /// <summary>
        /// NPOI.SS.UserModel.BorderDiagonal.Both
        /// </summary>
        AllDiagonal = RightTopToLeftBottom | LeftTopToRightBottom,
    }

    [Flags]
    public enum BorderTemplateEdges
    {
        None = 0,

        /// <summary>
        /// 垂直顶边
        /// </summary>
        Top = 0x01,
        /// <summary>
        /// 垂直中线
        /// </summary>
        Middle = 0x010000,
        /// <summary>
        /// 垂直底边
        /// </summary>
        Bottom = 0x04,

        /// <summary>
        /// 水平左边
        /// </summary>
        Left = 0x08,
        /// <summary>
        /// 水平中线
        /// </summary>
        Center = 0x020000,
        /// <summary>
        /// 水平右边
        /// </summary>
        Right = 0x02,

        /// <summary>
        /// 外部四边
        /// </summary>
        AllAround = Top | Right | Bottom | Left,
        /// <summary>
        /// 内部中线
        /// </summary>
        AllInside = Center | Middle,

        /// <summary>
        /// 逆对角线 (左上角 => 右下角) (left-top to right-bottom)
        /// NPOI.SS.UserModel.BorderDiagonal.Backward
        /// </summary>
        LeftTopToRightBottom = 0x0100,
        /// <summary>
        /// 正对角线 (右上角 => 左下角) (right-top to left-bottom)
        /// NPOI.SS.UserModel.BorderDiagonal.Forward
        /// </summary>
        RightTopToLeftBottom = 0x0200,
        /// <summary>
        /// NPOI.SS.UserModel.BorderDiagonal.Both
        /// </summary>
        AllDiagonal = RightTopToLeftBottom | LeftTopToRightBottom,
    }
}
