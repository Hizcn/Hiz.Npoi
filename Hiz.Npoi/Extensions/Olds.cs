using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    partial class NpoiExtensions
    {
        ///// <summary>
        ///// 删除指定范围空行, 并将后面的有效行上移;
        ///// 例如: 原来有 5 行, 第 2/4 为空行, 删除之后剩下紧挨着的三行, 为原来的第 1/3/5 行.
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <param name="start">起始行的索引</param>
        ///// <param name="count">检查行数; -1: 检查到最末行.</param>
        ///// <returns>实际删除行实例的数量</returns>
        //public static int RemoveRowRangeWithEmpty(this ISheet sheet, int start = 0, int count = -1)
        //{
        //    if (sheet == null)
        //        throw new ArgumentNullException(nameof(sheet));
        //    if (start < 0)
        //        throw new ArgumentOutOfRangeException(nameof(start), "不能负数");
        //    if (count == 0 || count < -1)
        //        throw new ArgumentOutOfRangeException(nameof(count), "必须正数或者负一");

        //    var removes = 0;
        //    if (sheet.PhysicalNumberOfRows > 0) // 存在行数
        //    {
        //        var max = GetMaxRows(sheet);
        //        if (start >= max) // 此处使用大于等于判断
        //            throw new ArgumentOutOfRangeException(nameof(start), "超出最大索引");

        //        int end; // 检查范围截止行的索引;
        //        var last = sheet.GetLastRowIndex(); // 最末个有效行的行索引
        //        if (count > 0) // 指定行数
        //        {
        //            // // 取消异常判断
        //            // var end = start + count;
        //            // if (end > max) // 此处使用大于判断
        //            //     throw new ArgumentOutOfRangeException(nameof(count), "超出最大行数");
        //            end = Math.Min(start + count, max);

        //            --end;
        //        }
        //        else // 检查到最末行
        //        {
        //            end = last;
        //        }

        //        var first = sheet.GetFirstRowIndex(); // 第一个有效行的行索引
        //        var index = Math.Min(end, last); // 当前行的索引
        //        var empties = 0; // 连续空行数量
        //        while (index >= start && index >= first) // 从下往上删除; 这样上移部分不会包含空行;
        //        {
        //            var row = sheet.GetRow(index);
        //            if (row == null || row.IsEmpty())
        //            {
        //                empties++; // 空行计数加一

        //                if (row != null)
        //                    removes++; // 实例计数加一
        //            }
        //            else if (empties > 0) // 如果之间存在空行
        //            {
        //                // 上移起始索引 = 当前行的索引 + 中间空行数量 + 往下一行;
        //                var next = index + empties;
        //                sheet.ShiftRows(++next, last, -empties, true, true);
        //                last -= empties; // 更新末行索引
        //                empties = 0; // 重新计数
        //            }
        //            index--; // 继续往上检查
        //        }

        //        /* 循环结束 empties 可能大于 0 (最后几行全都为空)
        //         * 然后追加 start 与 first 之间的空行;
        //         */
        //        if (start < first)
        //        {
        //            empties += first - start;
        //        }
        //        if (empties > 0)
        //        {
        //            // 最后将非空行 移到 start 位置;
        //            sheet.ShiftRows(start + empties, last, -empties, true, true);
        //        }
        //        /* HSSFSheet 调用 ShiftRows() 移动后原行位置会产生废行(有行实例, 但是 PhysicalNumberOfCells 为零);
        //         * 例如:
        //         * 原行索引 8, 现在移到索引 3, 移完之后, 会有两个实例: 索引 3 存放原行数据, 索引 8 存在一个空行实例;
        //         * 
        //         * XSSFSheet 直接修改原行索引, 不会产生废行.
        //         */
        //        if (sheet is HSSFSheet)
        //        {
        //            InternalRemoveEmptyRows(sheet); // 删除 ShiftRows 所产生的废行
        //        }
        //    }
        //    return removes;
        //}

        ///// <summary>
        ///// 删除指定范围行数.
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <param name="start">起始行的索引</param>
        ///// <param name="count">移除行数; -1: 移除到最末行.</param>
        ///// <param name="shift">后面的行是否上移</param>
        ///// <returns>实际删除行实例的数量</returns>
        //public static int RemoveRowRange(this ISheet sheet, int start, int count = -1, bool shift = true)
        //{
        //    if (sheet == null)
        //        throw new ArgumentNullException(nameof(sheet));
        //    if (start < 0)
        //        throw new ArgumentOutOfRangeException(nameof(start), "不能负数");
        //    if (count == 0 || count < -1)
        //        throw new ArgumentOutOfRangeException(nameof(count), "必须正数或者负一");

        //    var removes = 0;
        //    if (sheet.PhysicalNumberOfRows > 0) // 存在行数
        //    {
        //        var max = GetMaxRows(sheet);
        //        if (start >= max) // 此处使用大于等于判断
        //            throw new ArgumentOutOfRangeException(nameof(start), "超出最大索引");

        //        int end; // 移除范围截止行的索引;
        //        var last = sheet.GetLastRowIndex(); // 最末个有效行的行索引
        //        if (count > 0) // 指定行数
        //        {
        //            // // 取消异常判断
        //            // var end = start + count;
        //            // if (end > max) // 此处使用大于判断
        //            //     throw new ArgumentOutOfRangeException(nameof(count), "超出最大行数");
        //            end = Math.Min(start + count, max);

        //            // 取消修正, 后面逻辑还会用到原值;
        //            // end = Math.Min(--end, last); // 修正截止行的索引
        //            --end;
        //        }
        //        else // 检查到最末行
        //        {
        //            end = last;
        //        }

        //        var index = Math.Max(start, sheet.GetFirstRowIndex()); // 移除范围起始行的索引;
        //        while (index <= end && index <= last/*取消修正 end 所增加的逻辑*/) // 从上往下删除
        //        {
        //            var row = sheet.GetRow(index++);
        //            if (row != null)
        //            {
        //                sheet.RemoveRow(row);
        //                removes++; // 实例计数加一
        //            }
        //        }

        //        if (shift && (end < last)) // 如果后续还有行数
        //        {
        //            /* 从截止的下一行到最后一行, 整体上移删除行数.
        //             * 例如:
        //             * 原来最后一行 为第八行, 前三行为空行, 后续五行有值.
        //             * 现在删除 从第一行(index=0)到第五行(count=5)
        //             * 那么实际删除两行(第四第五), 这样 前五行都为空白行.
        //             * 后续的行 就可以往上移五行.
        //             * 最终结果: 一至三行有值;
        //             */
        //            sheet.ShiftRows(++end/*截止的下一行*/, last, start - end, true, true);

        //            if (sheet is HSSFSheet)
        //            {
        //                InternalRemoveEmptyRows(sheet); // 删除 ShiftRows 所产生的废行
        //            }
        //        }
        //    }
        //    return removes;
        //}
    }
}
