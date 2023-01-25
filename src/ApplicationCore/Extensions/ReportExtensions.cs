﻿using ClosedXML.Excel;
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
            var columns = typeof(T).GetCustomAttributes<IncludeAllInReportAttribute>().Any()
                ? GetAllColumnsFromModel(typeof(T)).ToDictionary(x => x.Name, x => x.Value).OrderBy(x => x.Value.Order).ToList()
                : GetColumnsFromModel(typeof(T)).ToDictionary(x => x.Name, x => x.Value).OrderBy(x => x.Value.Order).ToList();

            var table = new DataTable(name ?? typeof(T).Name);

            if (columns.Count == 0)
            {
                return table;
            }

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

        public static async Task<DataTable> ToDataTableAsync<T>(this IEnumerable<T> data, string name, List<ExcelColumnDefinition> cols = null)
        {
            var table = new DataTable(name ?? typeof(T).Name);

            if (cols == null)
                return table;

            var properties = typeof(T).GetProperties().ToList();

            if (cols.Count == 0)
                return table;

            await Task.Run(() =>
            {
                foreach (var column in cols)
                {
                    var prop = properties.FirstOrDefault(x => x.Name == column.Name);
                    table.Columns.Add(column.Label, Nullable.GetUnderlyingType(prop.GetPropertyDescriptor().PropertyType) ?? prop.GetPropertyDescriptor().PropertyType);
                }

                foreach (T item in data)
                {
                    var row = table.NewRow();

                    foreach (var prop in cols)
                    {
                        row[prop.Label] = PropertyExtensions.GetPropertyValue(item, prop.Name) ?? DBNull.Value;
                    }

                    table.Rows.Add(row);
                }
            });

            return table;
        }

        public static IEnumerable<Column> GetAllColumnsFromModel(Type parentClass, string parentName = null)
        {
            var properties = parentClass.GetProperties()
                .Where(p => !p.GetCustomAttributes<ExcludeFromReportAttribute>().Any());

            var order = 0;

            foreach (var prop in properties)
            {
                order++;
                yield return new Column
                {
                    Name = prop.GetPropertyDisplayName(),
                    Value = new ColumnValue
                    {
                        Order = order,
                        Path = string.IsNullOrWhiteSpace(parentName) ? prop.Name : $"{parentName}.{prop.Name}",
                        PropertyDescriptor = prop.GetPropertyDescriptor()
                    }
                };
            }
        }

        public static IEnumerable<Column> GetColumnsFromModel(Type parentClass, string parentName = null)
        {
            var complexReportProperties = parentClass.GetProperties()
                       .Where(p => p.GetCustomAttributes<NestedIncludeInReportAttribute>().Any()).ToList();

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

                    var complexProperties = GetColumnsFromModel(parentType, string.IsNullOrWhiteSpace(parentName) ? parentProperty.Name : $"{parentName}.{parentProperty.Name}");

                    foreach (var complexProperty in complexProperties)
                    {
                        yield return complexProperty;
                    }
                }
            }
        }

        public static async Task<byte[]> GenerateExcelForDataTableAsync<T>(this IEnumerable<T> data, string name, List<ExcelColumnDefinition> columns = null)
        {
            var table = await data.ToDataTableAsync(name, columns);

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

        public static async Task<byte[]> GenerateCSVForDataAsync<T>(this IEnumerable<T> data)
        {
            var builder = new StringBuilder();
            var stringWriter = new StringWriter(builder);

            await Task.Run(() =>
            {
                var columns = GetColumnsFromModel(typeof(T)).ToDictionary(x => x.Name, x => x.Value).OrderBy(x => x.Value.Order).ToList();

                foreach (var column in columns)
                {
                    stringWriter.Write(column.Key);
                    stringWriter.Write(", ");
                }
                stringWriter.WriteLine();

                foreach (T item in data)
                {
                    foreach (var prop in columns)
                    {
                        stringWriter.Write(PropertyExtensions.GetPropertyValue(item, prop.Value.Path));
                        stringWriter.Write(", ");
                    }
                    stringWriter.WriteLine();
                }
            });

            return Encoding.UTF8.GetBytes(builder.ToString());
        }

        public static async Task<byte[]> GenerateCSVForDataAsync<T>(this IEnumerable<T> data, List<ExcelColumnDefinition> cols = null)
        {
            var builder = new StringBuilder();
            var stringWriter = new StringWriter(builder);

            await Task.Run(() =>
            {
                cols = cols ?? new List<ExcelColumnDefinition>();

                foreach (var column in cols)
                {
                    stringWriter.Write(column.Label);
                    stringWriter.Write(", ");
                }
                stringWriter.WriteLine();

                foreach (T item in data)
                {
                    foreach (var prop in cols)
                    {
                        stringWriter.Write(PropertyExtensions.GetPropertyValue(item, prop.Name));
                        stringWriter.Write(", ");
                    }
                    stringWriter.WriteLine();
                }
            });

            return Encoding.UTF8.GetBytes(builder.ToString());
        }


        public static List<ExcelColumnDefinition> SetDefinitions(this List<ExcelColumnDefinition> columns, List<ExcelColumnDefinition> definitions)
        {
            if (definitions == null || definitions.Count == 0)
            {
                return new List<ExcelColumnDefinition>();
            }

            columns.AddRange(definitions);

            return columns;
        }

        public static List<ExcelColumnDefinition> SetDefinitions(this List<ExcelColumnDefinition> columns, params (string, string, int)[] definitions)
        {
            if (definitions == null || definitions.Length == 0)
            {
                return columns;
            }


            columns.AddRange(definitions.OrderBy(x => x.Item3).ToList().Select(item => new ExcelColumnDefinition(item.Item1, item.Item2)));

            return columns;
        }
    }
}
