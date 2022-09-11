using ClosedXML.Excel;
using Fingers10.ExcelExport.Attributes;
using Fingers10.ExcelExport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Fingers10.ExcelExport.Extensions
{
    public static class ReportExtensions
    {
        public static async Task<DataTable> ToDataTableAsync<T>(this IEnumerable<T> data, string name)
        {
            var columns = GetColumnsFromModel(typeof(T)).ToDictionary(x => x.Name, x => x.Value).OrderBy(x => x.Value.Order);
            var table = new DataTable(name ?? typeof(T).Name);

            await Task.Run(() =>
            {
                foreach (var column in columns)
                {
                    table.Columns.Add(column.Key, Nullable.GetUnderlyingType(column.Value.PropertyDescriptor.PropertyType) ?? column.Value.PropertyDescriptor.PropertyType);
                }

                foreach (T item in data)
                {
                    var row = table.NewRow();
                    foreach (var prop in columns)
                    {
                        row[prop.Key] = PropertyExtensions.GetPropertyValue(item, prop.Value.Path) ?? DBNull.Value;
                    }

                    table.Rows.Add(row);
                }
            });

            return table;
        }

        public static IEnumerable<Column> GetColumnsFromModel(Type parentClass, string parentName = null)
        {
            var complexReportProperties = parentClass.GetProperties()
                       .Where(p => p.GetCustomAttributes<NestedIncludeInReportAttribute>().Any());

            var properties = parentClass.GetProperties()
                       .Where(p => p.GetCustomAttributes<IncludeInReportAttribute>().Any());

            foreach (var prop in properties.Except(complexReportProperties))
            {
                var attribute = prop.GetCustomAttribute<IncludeInReportAttribute>();

                yield return new Column
                {
                    Name = prop.GetPropertyDisplayName(),
                    Value = new ColumnValue
                    {
                        Order = attribute.Order,
                        Path = string.IsNullOrWhiteSpace(parentName) ? prop.Name : $"{parentName}.{prop.Name}",
                        PropertyDescriptor = prop.GetPropertyDescriptor()
                    }
                };
            }

            if (complexReportProperties.Any())
            {
                foreach (var parentProperty in complexReportProperties)
                {
                    var parentType = parentProperty.PropertyType;
                    var parentAttribute = parentProperty.GetCustomAttribute<NestedIncludeInReportAttribute>();

                    var complexProperties = GetColumnsFromModel(parentType, string.IsNullOrWhiteSpace(parentName) ? parentProperty.Name : $"{parentName}.{parentProperty.Name}");

                    foreach (var complexProperty in complexProperties)
                    {
                        yield return complexProperty;
                    }
                }
            }
        }

        //public static DataTable ToFastDataTable<T>(this IEnumerable<T> data, string name)
        //{
        //    //For Excel reference
        //    //https://stackoverflow.com/questions/564366/convert-generic-list-enumerable-to-datatable
        //    // To restrict order or specific property
        //    //ObjectReader.Create(data, "Id", "Name", "Description")
        //    using (var reader = ObjectReader.Create(data))
        //    {
        //        var table = new DataTable(name ?? typeof(T).Name);
        //        table.Load(reader);
        //        return table;
        //    }
        //}

        public static async Task<byte[]> GenerateExcelForDataTableAsync<T>(this IEnumerable<T> data, string name)
        {
            var table = await data.ToDataTableAsync(name);

            using (var wb = new XLWorkbook(XLEventTracking.Disabled))
            {
                wb.Worksheets.Add(table).ColumnsUsed().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        public static async Task<byte[]> GenerateCSVForDataAsync<T>(this IEnumerable<T> data, bool spaceAfterComma) 
        {
            var builder = new StringBuilder();
            var stringWriter = new StringWriter(builder);

            string comma = spaceAfterComma ? " " : string.Empty;

            await Task.Run(() =>
            {
                var columns = GetColumnsFromModel(typeof(T)).ToDictionary(x => x.Name, x => x.Value).OrderBy(x => x.Value.Order);

                foreach (var column in columns)
                {
                    stringWriter.Write(column.Key);
                    stringWriter.Write(","+ comma);
                }
                stringWriter.WriteLine();

                foreach (T item in data)
                {
                    var properties = item.GetType().GetProperties();
                    foreach (var prop in columns)
                    {
                        stringWriter.Write(PropertyExtensions.GetPropertyValue(item, prop.Value.Path));
                        stringWriter.Write("," + comma);
                    }
                    stringWriter.WriteLine();
                }
            });

            return Encoding.UTF8.GetBytes(builder.ToString());
        }
    }
}
