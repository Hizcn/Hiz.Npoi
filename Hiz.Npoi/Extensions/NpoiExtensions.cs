
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Reflection;
using System.Linq.Expressions;
using System.Globalization;
using NPOI.SS;

namespace Hiz.Npoi
{
    /* v1.2 2018-11-26
     * 增加 NpoiColor 以及相关扩展方法;
     * 
     * v1.1 2018-11-17
     * 改进 ICell.GetCellValue() 增加 @default 参数.
     * 增加 ICell.SetCellComment()
     * 增加 IRow.GetRowHeight()/SetRowHeight();
     * 改进 ISheet.GetColumnWidth()/SetColumnWidth() 增加 LengthUnit 参数类型;
     * 改进 ISheet,RemoveRowRangeWithEmpty() 增加行数范围 start & count 参数;
     * 修正 ISheet.RemoveRowRange() 逻辑错误; 改进 count 参数支持 -1(移除到最末行); 增加 predicate 参数;
     * 
     * v1.0 2017-04-26
     * 初始版本
     */
    public static partial class NpoiExtensions
    {
        static NpoiExtensions()
        {
        }
    }
}