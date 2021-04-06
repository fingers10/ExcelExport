using Fingers10.ExcelExport.Extensions;
using Fingers10.ExcelExport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Fingers10.ExcelExport.ActionResults
{
    public class CSVResult<T> : IActionResult where T : class
    {
        private readonly IEnumerable<T> _data;

        public CSVResult(IEnumerable<T> data, string fileName)
        {
            _data = data;
            FileName = fileName;
        }

        public string FileName { get; }
        List<ExcelColumnDefinition> Columns { get; set; }

        /// <summary>
        /// this func to set columns definitions
        /// (column name, label, order)
        /// </summary>
        /// <param name="definitions"></param>
        public void SetDefinitions(params (string, string, int)[] definitions)
        {
            if (definitions == null || definitions.Length == 0)
            {
                return;
            }

            Columns = new List<ExcelColumnDefinition>();

            foreach (var item in definitions)
            {
                Columns.Add(new ExcelColumnDefinition(item.Item1, item.Item2, item.Item3));
            }
        }

        public CSVResult(IEnumerable<T> data, string fileName, params (string, string, int)[] definitions)
        {
            _data = data;
            FileName = fileName;

            SetDefinitions(definitions);
        }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            try
            {
                byte[] csvBytes;

                if (Columns != null && Columns.Count > 0)
                    csvBytes = await _data.GenerateCSVForDataAsync(Columns);
                else
                    csvBytes = await _data.GenerateCSVForDataAsync();

                WriteExcelFileAsync(context.HttpContext, csvBytes);

            }
            catch (Exception e)
            {
                Console.WriteLine(e);

                var errorBytes = await new List<T>().GenerateCSVForDataAsync();
                WriteExcelFileAsync(context.HttpContext, errorBytes);
            }
        }

        private async void WriteExcelFileAsync(HttpContext context, byte[] bytes)
        {
            context.Response.Headers["content-disposition"] = $"attachment; filename={FileName}.csv";
            await context.Response.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
}
