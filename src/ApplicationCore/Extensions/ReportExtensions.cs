using ClosedXML.Excel;
using FastMember;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace Fingers10.ExcelExport.Extensions
{
    public static class ReportExtensions
    {
        public static async Task<DataTable> ToDataTableAsync<T>(this IEnumerable<T> data, string name)
        {
            var properties = typeof(T).GetProperties();
            var table = new DataTable(name ?? typeof(T).Name);

            await Task.Run(() =>
            {
                foreach (var prop in properties)
                {
                    if (prop.IncludePropertyInTable())
                    {
                        var propertyDescriptor = prop.GetPropertyDescriptor();
                        var displayName = prop.GetPropertyDisplayName();
                        table.Columns.Add(displayName, Nullable.GetUnderlyingType(propertyDescriptor.PropertyType) ?? propertyDescriptor.PropertyType);
                    }
                }

                foreach (T item in data)
                {
                    var row = table.NewRow();
                    foreach (var prop in properties)
                    {
                        if (prop.IncludePropertyInTable())
                        {
                            var propertyDescriptor = prop.GetPropertyDescriptor();
                            var displayName = prop.GetPropertyDisplayName();
                            row[displayName] = propertyDescriptor.GetValue(item) ?? DBNull.Value;
                        }
                    }

                    table.Rows.Add(row);
                }
            });

            return table;
        }

        public static DataTable ToFastDataTable<T>(this IEnumerable<T> data, string name)
        {
            //For Excel reference
            //https://stackoverflow.com/questions/564366/convert-generic-list-enumerable-to-datatable
            // To restrict order or specific property
            //ObjectReader.Create(data, "Id", "Name", "Description")
            using (var reader = ObjectReader.Create(data))
            {
                var table = new DataTable(name ?? typeof(T).Name);
                table.Load(reader);
                return table;
            }
        }

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
    }
}
