using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Fingers10.ExcelExport.Extensions
{
    public static class PropertyExtensions
    {
        public static PropertyDescriptor GetPropertyDescriptor(this PropertyInfo propertyInfo)
        {
            return TypeDescriptor.GetProperties(propertyInfo.DeclaringType).Find(propertyInfo.Name, false);
        }

        public static string GetPropertyDisplayName(this PropertyInfo propertyInfo)
        {
            var propertyDescriptor = propertyInfo.GetPropertyDescriptor();
            var displayName = propertyInfo.IsDefined(typeof(DisplayAttribute), false) ? propertyInfo.GetCustomAttributes(typeof(DisplayAttribute),
                false).Cast<DisplayAttribute>().Single().Name : null;

            return displayName ?? propertyDescriptor.DisplayName ?? propertyDescriptor.Name;
        }

        public static bool IncludePropertyInTable(this PropertyInfo propertyInfo)
        {
            var includeProperty = propertyInfo.IsDefined(typeof(IncludeInReportAttribute), false) ? propertyInfo.GetCustomAttributes(typeof(IncludeInReportAttribute),
                false).Cast<IncludeInReportAttribute>().Single().IncludeInReport : true;

            return includeProperty;
        }
    }
}
