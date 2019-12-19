using System;

namespace Fingers10.ExcelExport.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IncludeInReportAttribute : Attribute
    {
        public int Order { get; set; }
    }
}
