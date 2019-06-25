using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    //TODO: 增加各个数据类型的默认单元格样式.
    //TODO: 增加一个全局默认 ExcelOptions.

    public class ExcelOptions
    {
        #region 引用资源

        // Attribute 配置使用 String Key, 应用 IWorkbook 使用 Short Index;

        // 颜色
        IDictionary<string, ColorOptions> _Colors;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, ColorOptions> Colors
        {
            get
            {
                if (_Colors == null)
                    _Colors = new Dictionary<string, ColorOptions>(StringComparer.OrdinalIgnoreCase);
                return _Colors;
            }
        }
        public virtual ColorOptions GetColor(string name)
        {
            return _Colors[name];
        }

        // 格式
        IDictionary<string, string> _DataFormats;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, string> DataFormats
        {
            get
            {
                if (_DataFormats == null)
                    _DataFormats = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                return _DataFormats;
            }
        }
        public virtual string GetDataFormat(string name)
        {
            return _DataFormats[name];
        }

        // 字体
        IDictionary<string, FontOptions> _Fonts;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, FontOptions> Fonts
        {
            get
            {
                if (_Fonts == null)
                    _Fonts = new Dictionary<string, FontOptions>(StringComparer.OrdinalIgnoreCase);
                return _Fonts;
            }
        }
        public virtual FontOptions GetFont(string name)
        {
            return _Fonts[name];
        }

        // 填充 (xlsx 新增)
        IDictionary<string, CellFillOptions> _CellFills;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, CellFillOptions> CellFills
        {
            get
            {
                if (_CellFills == null)
                    _CellFills = new Dictionary<string, CellFillOptions>(StringComparer.OrdinalIgnoreCase);
                return _CellFills;
            }
        }
        public virtual CellFillOptions GetCellFill(string name)
        {
            return _CellFills[name];
        }

        // 边框 (xlsx 新增)
        IDictionary<string, CellBorderOptions> _CellBorders;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, CellBorderOptions> CellBorders
        {
            get
            {
                if (_CellBorders == null)
                    _CellBorders = new Dictionary<string, CellBorderOptions>(StringComparer.OrdinalIgnoreCase);
                return _CellBorders;
            }
        }
        public virtual CellBorderOptions GetCellBorder(string name)
        {
            return _CellBorders[name];
        }

        // 对齐 (xlsx 新增)
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        IDictionary<string, CellAlignmentOptions> _CellAlignments;
        public IDictionary<string, CellAlignmentOptions> CellAlignments
        {
            get
            {
                if (_CellAlignments == null)
                    _CellAlignments = new Dictionary<string, CellAlignmentOptions>(StringComparer.OrdinalIgnoreCase);
                return _CellAlignments;
            }
        }
        public virtual CellAlignmentOptions GetCellAlignment(string name)
        {
            return _CellAlignments[name];
        }

        // 样式
        IDictionary<string, CellStyleOptions> _CellStyles;
        /// <summary>
        /// Key: 不区分大小写
        /// </summary>
        public IDictionary<string, CellStyleOptions> CellStyles
        {
            get
            {
                if (_CellStyles == null)
                    _CellStyles = new Dictionary<string, CellStyleOptions>(StringComparer.OrdinalIgnoreCase);
                return _CellStyles;
            }
        }
        public virtual CellStyleOptions GetCellStyle(string name)
        {
            return _CellStyles[name];
        }

        #endregion

        string _DefaultDateTimeFormat = "yyyy-mm-dd hh:mm:ss";
        /// <summary>
        /// 默认时间格式 (用于显示时间单元格值)
        /// </summary>
        public virtual string DefaultDateTimeFormat { get => _DefaultDateTimeFormat; set => _DefaultDateTimeFormat = value; }

        FontOptions _DefaultFontOptions = new FontOptions(XSSFFont.DEFAULT_FONT_NAME, XSSFFont.DEFAULT_FONT_SIZE);
        /// <summary>
        /// 默认字体
        /// </summary>
        public virtual FontOptions DefaultFont { get => _DefaultFontOptions; set => _DefaultFontOptions = value; }

        public ExcelOptions()
        {
        }
    }
}
