using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.ConsoleApp
{
    /* 静态元素
     * 
     * 简单表头: 每列一个列头
     * 
     * 复杂表头: 列头存在合并单元格等情况...
     */

    /* 动态元素 (数据)
     */
    class TestModel : IEquatable<TestModel>
    {
        public string String { get; set; }

        public decimal Decimal { get; set; }
        public decimal? DecimalNullable { get; set; }

        public float Single { get; set; }
        public float? SingleNullable { get; set; }

        public double Double { get; set; }
        public double? DoubleNullable { get; set; }

        public short Int16 { get; set; }
        public short? Int16Nullable { get; set; }

        public int Int32 { get; set; }
        public int? Int32Nullable { get; set; }

        public long Int64 { get; set; }
        public long? Int64Nullable { get; set; }

        public ushort UInt16 { get; set; }
        public ushort? UInt16Nullable { get; set; }

        public uint UInt32 { get; set; }
        public uint? UInt32Nullable { get; set; }

        public ulong UInt64 { get; set; }
        public ulong? UInt64Nullable { get; set; }

        public byte Byte { get; set; }
        public byte? ByteNullable { get; set; }

        public sbyte SByte { get; set; }
        public sbyte? SByteNullable { get; set; }

        public bool Boolean { get; set; }
        public bool? BooleanNullable { get; set; }

        public char Char { get; set; }
        public char? CharNullable { get; set; }

        public Guid Guid { get; set; }
        public Guid? GuidNullable { get; set; }

        public DateTime DateTime { get; set; }
        public DateTime? DateTimeNullable { get; set; }

        public TimeSpan TimeSpan { get; set; }
        public TimeSpan? TimeSpanNullable { get; set; }

        public bool Equals(TestModel other)
        {
            return other != null
                && this.String == other.String
                && this.Decimal == other.Decimal
                && this.DecimalNullable == other.DecimalNullable
                && this.Single == other.Single
                && this.SingleNullable == other.SingleNullable
                && this.Double == other.Double
                && this.DoubleNullable == other.DoubleNullable
                && this.Int16 == other.Int16
                && this.Int16Nullable == other.Int16Nullable
                && this.Int32 == other.Int32
                && this.Int32Nullable == other.Int32Nullable
                && this.Int64 == other.Int64
                && this.Int64Nullable == other.Int64Nullable
                && this.UInt16 == other.UInt16
                && this.UInt16Nullable == other.UInt16Nullable
                && this.UInt32 == other.UInt32
                && this.UInt32Nullable == other.UInt32Nullable
                && this.UInt64 == other.UInt64
                && this.UInt64Nullable == other.UInt64Nullable
                && this.Byte == other.Byte
                && this.ByteNullable == other.ByteNullable
                && this.SByte == other.SByte
                && this.SByteNullable == other.SByteNullable
                && this.Boolean == other.Boolean
                && this.BooleanNullable == other.BooleanNullable
                && this.Char == other.Char
                && this.CharNullable == other.CharNullable
                && this.Guid == other.Guid
                && this.GuidNullable == other.GuidNullable
                && this.DateTime == other.DateTime
                && this.DateTimeNullable == other.DateTimeNullable
                && this.TimeSpan == other.TimeSpan
                && this.TimeSpanNullable == other.TimeSpanNullable
                ;
        }
        public override bool Equals(object other)
        {
            return base.Equals(other as TestModel);
        }
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
