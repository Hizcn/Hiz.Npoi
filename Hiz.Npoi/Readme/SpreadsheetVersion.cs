using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    /* NPOI.SS.SpreadsheetVersion
     * 
     * xls: {
     *     DefaultExtension: xls
     *     MaxRows: 65536 = 0x10000; Index: [0, 0x00FFFF]
     *     MaxColumns: 256 = 0x100; Index: [0, 0xFF]
     *     MaxFunctionArgs: 30
     *     MaxConditionalFormats: 3
     *     MaxCellStyles: 4000
     *     MaxTextLength: 32767 = 7FFF (Int16.MaxValue)
     * }
     * 
     * xlsx: {
     *     DefaultExtension: xlsx
     *     MaxRows: 1048576 = 0x100000; Index: [0, 0x0FFFFF]
     *     MaxColumns: 16384 = 0x4000; Index: [0, 0x3FFF]
     *     MaxFunctionArgs: 255;
     *     MaxConditionalFormats: 2147483647 = 0x7FFFFFFF (Int32.MaxValue)
     *     MaxCellStyles: 64000
     *     MaxTextLength: 32767 = 7FFF (Int16.MaxValue)
     * }
     */
}
