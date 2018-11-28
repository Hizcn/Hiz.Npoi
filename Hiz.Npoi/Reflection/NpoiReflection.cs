using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Linq.Expressions;
using NPOI.HSSF.Model;

namespace Hiz.Npoi
{
    static class NpoiReflection
    {
        public static Lazy<Func<HSSFCellStyle, InternalWorkbook>> HSSFCellStyleGetWorkbook = new Lazy<Func<HSSFCellStyle, InternalWorkbook>>(MakeHSSFCellStyleGetWorkbook);
        static Func<HSSFCellStyle, InternalWorkbook> MakeHSSFCellStyleGetWorkbook()
        {
            var type = typeof(HSSFCellStyle);
            var _workbook = type.GetField("_workbook", BindingFlags.Instance | BindingFlags.NonPublic);
            if (_workbook != null)
            {
                var expInstance = Expression.Parameter(type);
                var expWorkbook = Expression.Field(expInstance, _workbook);
                var lambda = Expression.Lambda<Func<HSSFCellStyle, InternalWorkbook>>(expWorkbook, new ParameterExpression[] { expInstance });
                var getter = lambda.Compile();
                return getter;
            }
            else
            {
            }
            return null;
        }
    }
}
