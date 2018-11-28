using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI.SS.UserModel;

namespace Hiz.Npoi
{
    /// <summary>
    /// 边框批量设置模板
    /// </summary>
    public class BorderTemplate
    {
        /// <summary>
        /// 垂直顶边样式
        /// </summary>
        public BorderStyle TopStyle { get; set; }
        /// <summary>
        /// 垂直顶边颜色
        /// </summary>
        public NpoiColor TopColor { get; set; }

        /// <summary>
        /// 垂直中线样式
        /// </summary>
        public BorderStyle MiddleStyle { get; set; }
        /// <summary>
        /// 垂直中线颜色
        /// </summary>
        public NpoiColor MiddleColor { get; set; }

        /// <summary>
        /// 垂直底边样式
        /// </summary>
        public BorderStyle BottomStyle { get; set; }
        /// <summary>
        /// 垂直底边颜色
        /// </summary>
        public NpoiColor BottomColor { get; set; }

        /// <summary>
        /// 水平左边样式
        /// </summary>
        public BorderStyle LeftStyle { get; set; }
        /// <summary>
        /// 水平左边颜色
        /// </summary>
        public NpoiColor LeftColor { get; set; }

        /// <summary>
        /// 水平中线样式
        /// </summary>
        public BorderStyle CenterStyle { get; set; }
        /// <summary>
        /// 水平中线颜色
        /// </summary>
        public NpoiColor CenterColor { get; set; }

        /// <summary>
        /// 水平右边样式
        /// </summary>
        public BorderStyle RightStyle { get; set; }
        /// <summary>
        /// 水平右边颜色
        /// </summary>
        public NpoiColor RightColor { get; set; }

        /// <summary>
        /// 对角线条组合
        /// </summary>
        public BorderDiagonal Diagonal { get; set; }
        /// <summary>
        /// 对角线条样式
        /// </summary>
        public BorderStyle DiagonalStyle { get; set; }
        /// <summary>
        /// 对角线条颜色
        /// </summary>
        public NpoiColor DiagonalColor { get; set; }

        public void RemoveBorders(BorderTemplateEdges edges)
        {
            AddBorders(edges, BorderStyle.None, NpoiColor.Black);
        }

        public void AddBorders(BorderTemplateEdges edges, BorderStyle style, NpoiColor color)
        {
            if (style == BorderStyle.None)
                throw new ArgumentException("style");
            if (color == null)
                throw new ArgumentException("color");

            if ((edges & BorderTemplateEdges.Top) != BorderTemplateEdges.None)
            {
                this.TopStyle = style;
                this.TopColor = color;
            }
            if ((edges & BorderTemplateEdges.Middle) != BorderTemplateEdges.None)
            {
                this.MiddleStyle = style;
                this.MiddleColor = color;
            }
            if ((edges & BorderTemplateEdges.Bottom) != BorderTemplateEdges.None)
            {
                this.BottomStyle = style;
                this.BottomColor = color;
            }
            if ((edges & BorderTemplateEdges.Left) != BorderTemplateEdges.None)
            {
                this.LeftStyle = style;
                this.LeftColor = color;
            }
            if ((edges & BorderTemplateEdges.Center) != BorderTemplateEdges.None)
            {
                this.CenterStyle = style;
                this.CenterColor = color;
            }
            if ((edges & BorderTemplateEdges.Right) != BorderTemplateEdges.None)
            {
                this.RightStyle = style;
                this.RightColor = color;
            }

            var diagonal = edges & BorderTemplateEdges.AllDiagonal;
            if (diagonal != BorderTemplateEdges.None)
            {
                this.Diagonal = (BorderDiagonal)((int)diagonal >> 0x08);
                this.DiagonalStyle = style;
                this.DiagonalColor = color;
            }
        }

        public static BorderTemplate GetSimple()
        {
            var template = new BorderTemplate();
            template.AddBorders(BorderTemplateEdges.AllAround | BorderTemplateEdges.AllInside, BorderStyle.Thin, NpoiColor.Black);
            return template;
        }
    }
}
