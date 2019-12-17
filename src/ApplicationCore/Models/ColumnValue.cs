using System.ComponentModel;

namespace Fingers10.ExcelExport.Models
{
    public class ColumnValue
    {
        public int Order { get; set; }
        public string Path { get; set; }
        public PropertyDescriptor PropertyDescriptor { get; set; }
    }
}
