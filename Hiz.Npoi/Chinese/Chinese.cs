using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public static class Chinese
    {
        public static ExcelOptions GetExcelOptions()
        {
            var options = new ExcelOptions();

            // 默认字体
            options.DefaultFont = new FontOptions("宋体", 10f);

            // 其它字体
            options.Fonts.Add("Song10", new FontOptions("宋体", 10f));
            options.Fonts.Add("Song10b", new FontOptions("宋体", 10f) { IsBold = true });
            options.Fonts.Add("Song12b", new FontOptions("宋体", 10f) { IsBold = true });
            options.Fonts.Add("YaHei10", new FontOptions("雅黑", 10f));
            options.Fonts.Add("YaHei10b", new FontOptions("雅黑", 10f) { IsBold = true });

            options.CellStyles.Add("Title", new CellStyleOptions() { Font = "Song12b" });
            options.CellStyles.Add("Cell.Header", new CellStyleOptions() { Font = "Song10b" });
            options.CellStyles.Add("Cell", new CellStyleOptions() { Font = "Song10" });
            options.CellStyles.Add("Cell.Date", new CellStyleOptions() { Font = "Song10", DataFormat = "Date" });
            options.CellStyles.Add("Cell.P2", new CellStyleOptions() { Font = "Song10", DataFormat = "P2" });
            options.CellStyles.Add("Cell.C2", new CellStyleOptions() { Font = "YaHei10", DataFormat = "C2/Red/CNY" });

            AddFormats(options);

            return options;
        }

        // 内建格式: NPOI.SS.UserModel.BuiltinFormats;
        // 索引范围 Index: [0, 49]; 合计数量: 50;
        // [0x17-0x24] Reserved for international and undocumented.
        static void AddFormats(ExcelOptions options)
        {
            /* Office 2013 (*.xlsx)预设货币格式:
             * 修改内建格式:
             * 0x05: "¥"#,##0;"¥"\-#,##0
             * 0x06: "¥"#,##0;[Red]"¥"\-#,##0
             * 0x07: "¥"#,##0.00;"¥"\-#,##0.00 // 等同原始内建格式 0x07: 去掉负数括号改用负号+替换本地货币符号;
             * 0x08: "¥"#,##0.00;[Red]"¥"\-#,##0.00 // 等同原始内建格式 0x08: 去掉负数括号改用负号+替换本地货币符号;
             * 自定义的格式 (保留两位情况):
             * "¥"#,##0.00_);\("¥"#,##0.00\) // 等同内建格式 0x07 替换本地货币符号;
             * "¥"#,##0.00_);[Red]\("¥"#,##0.00\) // 等同内建格式 0x08 替换本地货币符号;
             * "¥"#,##0.00;[Red]"¥"#,##0.00 // 货币符号+使用千位分隔+保留两位小数(四舍五入)+负数显示红色;
             * 
             * ¥/￥ 两个符号 WPS/Office 都不能够完美统一:
             * ¥: WPS: xls 可以; xlsx 不能; Office 2013: xls/xlsx 都能识为货币格式;
             * ￥: WPS: xls/xlsx 都能识为货币格式; Office2013: xls 可以; xlsx 不能;
             * WPS 默认使用全角; Office 默认使用半角;
             * 
             * 决定还是使用全角: Shift+$ 可以按出.
             * 
             * 负数使用负号较比括号直观.
             * 
             * 预设财务格式:
             * 0x2A: [_ "¥"* #,##0_ ;_ "¥"* \-#,##0_ ;_ "¥"* "-"_ ;_ @_ ] // 不含括号 // 等同内建格式 0x2A: 去掉负数括号改用负号+替换本地货币符号;
             * 0x2C: [_ "¥"* #,##0.00_ ;_ "¥"* \-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ] // 不含括号 // 等同内建格式 0x2C: 去掉负数括号改用负号+替换本地货币符号;
             */
            var formats = options.DataFormats;
            formats.Add("F0", BuiltinFormats.GetBuiltinFormat(0x01));
            formats.Add("F2", BuiltinFormats.GetBuiltinFormat(0x02));
            formats.Add("N0", BuiltinFormats.GetBuiltinFormat(0x03));
            formats.Add("N2", BuiltinFormats.GetBuiltinFormat(0x04));
            //
            formats.Add("C0", BuiltinFormats.GetBuiltinFormat(0x05));
            formats.Add("C0/Red", BuiltinFormats.GetBuiltinFormat(0x06));
            formats.Add("C2", BuiltinFormats.GetBuiltinFormat(0x07));
            formats.Add("C2/Red", BuiltinFormats.GetBuiltinFormat(0x08));
            // 中国支持 // xlsx 格式包含 Locale ? 经测试会自动替换货币符号.
            // formats.Add("C0/CNY", BuiltinFormats.GetBuiltinFormat(0x05).Replace('$', '￥'));
            // formats.Add("C0/Red/CNY", BuiltinFormats.GetBuiltinFormat(0x06).Replace('$', '￥'));
            // formats.Add("C2/CNY", BuiltinFormats.GetBuiltinFormat(0x07).Replace('$', '￥'));
            // formats.Add("C2/Red/CNY", BuiltinFormats.GetBuiltinFormat(0x08).Replace('$', '￥'));
            formats.Add("C0/CNY", "\"￥\"#,##0;\"￥\"\\-#,##0");
            formats.Add("C0/Red/CNY", "\"￥\"#,##0;[Red]\"￥\"\\-#,##0");
            formats.Add("C2/CNY", "\"￥\"#,##0.00;\"￥\"\\-#,##0.00");
            formats.Add("C2/Red/CNY", "\"￥\"#,##0.00;[Red]\"￥\"\\-#,##0.00");
            //
            formats.Add("A0/CNY", "_ \"￥\"* #,##0_ ;_ \"￥\"* \\-#,##0_ ;_ \"￥\"* \"-\"_ ;_ @_ ");
            formats.Add("A2/Red/CNY", "_ \"￥\"* #,##0.00_ ;_ \"￥\"* \\-#,##0.00_ ;_ \"￥\"* \"-\"??_ ;_ @_ ");
            //
            formats.Add("P0", BuiltinFormats.GetBuiltinFormat(0x09));
            formats.Add("P2", BuiltinFormats.GetBuiltinFormat(0x0A));
            //
            formats.Add("E2", BuiltinFormats.GetBuiltinFormat(0x0B));
            //
            formats.Add("Text", "@"); // 0x31
            //
            formats.Add("Date", options.DefaultDateTimeFormat);
        }
    }
}