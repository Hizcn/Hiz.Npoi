﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Extended.Npoi
{
    class PaletteOptions : INamed
    {
        public string Name { get; set; }

        IList<NpoiColor> Colors { get; set; }
    }
}