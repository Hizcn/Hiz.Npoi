using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public enum OfficeArchiveFormat
    {
        /// <summary>
        /// 未知 / 不作指定 / 按后缀名自动判断;
        /// </summary>
        None = 0,

        /// <summary>
        /// Office 2003 以之前的版本
        /// </summary>
        Binary,

        /// <summary>
        /// Office 2007 及之后的版本
        /// </summary>
        OpenXml,
    }
}
