using Hiz.Npoi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.ConsoleApp
{
    class TestExportImport
    {
        public static void ExportMany()
        {
            var data = GetTestData().ToArray();

            var styles = Chinese.GetExcelOptions();

            var options = new RuntimeOptions()
            {
                FilePath = "test190105.xlsx",
                ExcelStyles = styles,
                Title = "测试标题",
            };

            var workbook = Xpoi.ExportMany<TestModel>(data, options);
            workbook.Write(options.FilePath);
        }

        static IEnumerable<TestModel> GetTestData()
        {
            for (var i = 0; i < 10; i++)
            {
                var even = (i & 0x01) == 0;

                yield return new TestModel()
                {
                    String = i.ToString(),

                    Decimal = i,
                    DecimalNullable = even ? null : (decimal?)i,

                    Single = i / 10f,
                    SingleNullable = even ? null : (float?)(i / 10f),

                    Double = i / 10d,
                    DoubleNullable = even ? null : (double?)(i / 10d),

                    SByte = (SByte)i,
                    SByteNullable = even ? null : (SByte?)i,

                    Int16 = (Int16)i,
                    Int16Nullable = even ? null : (Int16?)i,

                    Int32 = i,
                    Int32Nullable = even ? null : (Int32?)i,

                    Int64 = i,
                    Int64Nullable = even ? null : (Int64?)i,

                    Byte = (Byte)i,
                    ByteNullable = even ? null : (Byte?)i,

                    UInt16 = (UInt16)i,
                    UInt16Nullable = even ? null : (UInt16?)i,

                    UInt32 = (UInt32)i,
                    UInt32Nullable = even ? null : (UInt32?)i,

                    UInt64 = (UInt64)i,
                    UInt64Nullable = even ? null : (UInt64?)i,

                    Boolean = even,
                    BooleanNullable = even ? null : (Boolean?)even,

                    Char = even ? 'T' : 'F',
                    CharNullable = even ? null : (Char?)'F',

                    Guid = Guid.NewGuid(),
                    GuidNullable = even ? null : (Guid?)Guid.NewGuid(),

                    DateTime = DateTime.Now,
                    DateTimeNullable = even ? null : (DateTime?)DateTime.Now,

                    TimeSpan = DateTime.Now - DateTime.Today,
                    TimeSpanNullable = even ? null : (TimeSpan?)(DateTime.Now - DateTime.Today),

                    Enum = (SampleEnum)i,
                    EnumNullable = even ? null : (SampleEnum?)i,

                    EnumAsInteger = (SampleEnum)i,
                    UInt64AsString = UInt64.MaxValue,
                };
            }
        }
    }
}
