using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Hiz.Npoi
{
    public class RuntimeOptions
    {
        public OfficeArchiveFormat FileFormat { get; set; }

        public string FilePath { get; set; }

        public ExcelOptions ExcelStyles { get; set; }

        string _Title;
        public string Title
        {
            get
            {
                return _Title;
            }
            set
            {
                _Title = value;
            }
        }

        bool? _HasTitle;
        public bool HasTitle
        {
            get
            {
                return _HasTitle == true || !string.IsNullOrEmpty(this.Title);
            }
            set
            {
                _HasTitle = value;
            }
        }
    }
}
