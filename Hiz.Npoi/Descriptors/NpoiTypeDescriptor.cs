using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    class NpoiTypeDescriptor<T>
    {

        #region 属性

        IList<NpoiPropertyDescriptor<T>> _Properties = null;
        public IList<NpoiPropertyDescriptor<T>> Properties
        {
            get
            {
                if (_Properties == null)
                    _Properties = new List<NpoiPropertyDescriptor<T>>();
                return this._Properties;
            }
        }

        #endregion
    }
}
