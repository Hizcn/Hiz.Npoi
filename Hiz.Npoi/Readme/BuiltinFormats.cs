using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    /* 格式种类
     * 
     * 将数字设置为货币格式
     * https://support.office.com/en-us/article/Display-numbers-as-currency-8acb42f4-cd90-4e27-8f3e-5b8e7b4473a5
     * 
     * General: 常规
     * Number: 数值
     * Currency: 货币
     * Accounting: 会计专用
     * Date: 日期
     * Time: 时间
     * Percentage: 百分比
     * Fraction: 分数
     * Scientific: 科学计数
     * Text: 文本
     * Special: 特殊
     * Custom: 自定义
     */

    /* Custom number format
     * 
     * Create or delete a custom number format
     * https://support.office.com/en-us/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4
     * 
     * Number format codes
     * https://support.office.com/en-us/article/Number-format-codes-5026BBD6-04BC-48CD-BF33-80F18B4EAE68
     * 
     * 
     * Number Format:
     * <Positive>;<Negative>;<Zero>;<Text>
     * Positive: 正数格式
     * Negative: 负数格式
     * Zero: 零值格式
     * Text: 文本格式
     * 
     * Specify colors: [Black][White][Red][Yellow][Green][Cyan][Blue][Magenta]
     * 
     * To display both text and numbers in a cell, enclose the text characters in double quotation marks (" ") or precede a single character with a backslash (\).
     * // 若要显示在单元格中的文本和数字，将文本字符括在双引号 (" ") 或在前一个字符以反斜杠 (\)。
     * 
     * The following characters are displayed without the use of quotation marks:
     * // 不必使用引号显示下表中列出的字符:
     * $ Dollar sign
     * + Plus sign
     * ( Left parenthesis
     * : Colon
     * ^ Circumflex accent (caret)
     * ' Apostrophe
     * { Left curly bracket
     * < Less-than sign
     * = Equal sign
     * - Minus sign
     * / Slash mark
     * ) Right parenthesis
     * ! Exclamation point
     * & Ampersand
     * ~ Tilde
     * } Right curly bracket
     * > Greater-than sign
     *   Space character 空格
     *   
     * Include currency symbols:
     * ¢ ALT+0162
     * £ ALT+0163
     * ¥ ALT+0165
     * € ALT+0128
     */

    /* NPOI.SS.UserModel.BuiltinFormats
     * 
     * Utility to identify built-in formats. The following is a list of the formats as returned by this class.
     * |======================================================================================================
     * | 0x00 | General                                                  | 常规
     * | 0x01 | 0                                                        | 数值: 舍去小数 (四舍五入)
     * | 0x02 | 0.00                                                     | 数值: 保留两位小数 (四舍五入)
     * | 0x03 | #,##0                                                    | 数值: 使用千位分隔 舍去小数 (四舍五入)
     * | 0x04 | #,##0.00                                                 | 数值: 使用千位分隔 保留两位小数 (四舍五入)
     * | 0x05 | "$"#,##0_);("$"#,##0)                                    | 货币: 货币符号 使用千位分隔 舍去小数 (四舍五入); 负数: 去掉负号 增加左右括号; 正数: 右边空出一个括号宽度 对齐负数;
     * | 0x06 | "$"#,##0_);[Red]("$"#,##0)                               | 货币: 负数显示红色; 其它等同 0x05;
     * | 0x07 | "$"#,##0.00_);("$"#,##0.00)                              | 货币: 保留两位小数; 其它等同 0x05;
     * | 0x08 | "$"#,##0.00_);[Red]("$"#,##0.00)                         | 货币: 负数显示红色; 其它等同 0x07;
     * | 0x09 | 0%                                                       | 百分比: 舍去小数 (四舍五入)
     * | 0x0A | 0.00%                                                    | 百分比: 保留两位小数 (四舍五入)
     * | 0x0B | 0.00E+00                                                 | 科学计数: 保留两位小数 (四舍五入); 计数不足两位 前导补零;
     * | 0x0C | # ?/?                                                    | 分数: 分母为一位数
     * | 0x0D | # ??/??                                                  | 分数: 分母为两位数; 分子不足两位 前面补一空格; 分母不足两位 后面补一空格; 使其对齐其它相同格式数值;
     * | 0x0E | m/d/yy                                                   | 日期: m: 不带前导零的数字; d: 不带前导零的数字; yy: 年份两位数字;
     * | 0x0F | d-mmm-yy                                                 | 日期: mmm: 月份简写 Jan/Feb/Mar/Apr/May/Jun/Jul/Aug/Sep/Oct/Nov/Dec;
     * | 0x10 | d-mmm                                                    | 日期: 
     * | 0x11 | mmm-yy                                                   | 日期: 
     * | 0x12 | h:mm AM/PM                                               | 时间: AM/PM; am/pm; A/P; a/p; 使用 12 小时制;
     * | 0x13 | h:mm:ss AM/PM                                            | 时间: m/mm 必须紧跟 h/hh 代码之后 或者 s/ss 代码之前; 否则表示月份;
     * | 0x14 | h:mm                                                     | 时间: h: 不带前导零的数字; mm: 带有前导零的数字;
     * | 0x15 | h:mm:ss                                                  | 时间: ss: 带有前导零的数字;
     * | 0x16 | m/d/yy h:mm                                              | 日期: 
     * |=================================================================|===========================================================
     * | 0x17 - 0x24: Reserved for international and undocumented.       |
     * |=================================================================|=============================================================
     * | 0x25 | #,##0_);(#,##0)                                          | 数值: 使用千位分隔 舍去小数 (四舍五入); 负数: 去掉负号 增加左右括号; 正数: 右边空出一个括号宽度 对齐负数;
     * | 0x26 | #,##0_);[Red](#,##0)                                     | 数值: 负数显示红色; 其它等同 0x25;
     * | 0x27 | #,##0.00_);(#,##0.00)                                    | 数值: 保留两位小数; 其它等同 0x25;
     * | 0x28 | #,##0.00_);[Red](#,##0.00)                               | 数值: 负数显示红色; 其它等同 0x27;
     * | 0x29 | _(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)                  | 会计专用: 去掉货币符号; 其它等同 0x2A;
     * | 0x2A | _("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)         | 会计专用: 货币符号与数值之间使用空格填满单元格; 例如: 正数: " $   80 "; 负数: " $  (80)"; 零值: " $    - ";
     * | 0x2B | _(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)          | 会计专用: 保留两位小数; 零值杠对齐小数点; 其它等同 0x29;
     * | 0x2C | _("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_) | 会计专用: 增加货币符号; 其它等同 0x2B;
     * | 0x2D | mm:ss                                                    | 时间: 
     * | 0x2E | [h]:mm:ss                                                | 时间: [h]: 小时为单位显示经过的时间;
     * | 0x2F | mm:ss.0                                                  | 时间: WPS 不太兼容;
     * | 0x30 | ##0.0E+0                                                 | 科学计数: 保留一位小数 (四舍五入); WPS 不太兼容;
     * | 0x31 | @                                                        | This is text format; "text" - Alias for "@";  // 输出文本形式
     * |===================================================================================================================================
     * mm 和 hh 或 ss 一起, 表示分钟; 其它情况表示月份;
     * 
     * v2.2.1.0 BUG: 29/2A 位置反了. 上面显示正确.
     */
}
