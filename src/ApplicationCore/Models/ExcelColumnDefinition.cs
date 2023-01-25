namespace Fingers10.ExcelExport.Models
{
    public class ExcelColumnDefinition
    {
        public string Name { get; set; }
        public string Label { get; set; }

        public ExcelColumnDefinition(string name, string label)
        {
            Name = name;
            Label = label;
        }

        protected ExcelColumnDefinition()
        {
        }
    }
}
