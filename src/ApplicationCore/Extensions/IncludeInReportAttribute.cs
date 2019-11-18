using System;
using System.Collections.Generic;
using System.Text;

namespace Fingers10.ExcelExport.Extensions
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class IncludeInReportAttribute : Attribute
    {
        public IncludeInReportAttribute(bool includeInReport)
        {
            IncludeInReport = includeInReport;
        }

        public bool IncludeInReport { get; }
    }
}
