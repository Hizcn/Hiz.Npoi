using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.ConsoleApp
{
    class TestEnum
    {
        public static void Test()
        {
            object v1 = null;
            var v2 = (SampleEnum?)v1;
            // var v3 = (SampleEnum)v1; // 将会抛出异常 NullReferenceException.

            object a1 = SampleEnum.Two;
            var a2 = (SampleEnum)a1;
            var a3 = (SampleEnum?)a1;
            var a4 = (SampleEnum?)(SampleEnum)a1;

            object c1 = "Two";
            var c2 = Enum.ToObject(typeof(SampleEnum), 1);

            var t1 = typeof(Nullable<>).MakeGenericType(typeof(SampleEnum));

            var a5 = Activator.CreateInstance(t1, a1);
        }


    }
}
