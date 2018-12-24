using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class NpoiValueAsAttribute : Attribute
    {
        CellType _CellType = CellType.Unknown;
        public CellType CellType { get => _CellType; }

        public NpoiValueAsAttribute(CellType type)
        {
            this._CellType = type;
        }
    }

    // [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    // public class ValueAsNumericAttribute : Attribute
    // {
    // }
    // [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    // public class ValueAsStringAttribute : Attribute
    // {
    // }
}
