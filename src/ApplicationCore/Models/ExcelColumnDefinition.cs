using System;
using System.Collections.Generic;
using System.Text;

namespace Fingers10.ExcelExport.Models
{
    public class ExcelColumnDefinition
    {
        public string Name { get; set; }
        public string Label { get; set; }
        public int Order { get; set; }

        public ExcelColumnDefinition(string name, string label, int order)
        {
            Name = name;
            Label = label;
            Order = order;
        }

        public ExcelColumnDefinition()
        {

        }
    }
}
