using System;
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

        public static object GetPropertyValue(object src, string propName)
        {
            // https://stackoverflow.com/questions/1954746/using-reflection-in-c-sharp-to-get-properties-of-a-nested-object/29823444#29823444
            if (src == null) throw new ArgumentException("Value cannot be null.", "src");
            if (propName == null) throw new ArgumentException("Value cannot be null.", "propName");

            if (propName.Contains("."))//complex type nested
            {
                var temp = propName.Split(new char[] { '.' }, 2);
                return GetPropertyValue(GetPropertyValue(src, temp[0]), temp[1]);
            }
            else
            {
                var prop = src.GetType().GetProperty(propName);
                return prop != null ? prop.GetValue(src, null) : null;
            }
        }
    }
}
