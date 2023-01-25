using Fingers10.ExcelExport.Extensions;
using Fingers10.ExcelExport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Fingers10.ExcelExport.ActionResults
{
    public class CSVResult<T> : IActionResult where T : class
    {
        private readonly IEnumerable<T> _data;
        public string FileName { get; }

        private List<ExcelColumnDefinition> Columns { get; set; } = new List<ExcelColumnDefinition>();

        public CSVResult(IEnumerable<T> data, string fileName)
        {
            _data = data;
            FileName = fileName;
        }

        public CSVResult(IEnumerable<T> data, string fileName, params (string name, string label, int order)[] definitions)
        {
            _data = data;
            FileName = fileName;

            Columns.SetDefinitions(definitions);
        }

        public CSVResult(IEnumerable<T> data, string fileName, List<ExcelColumnDefinition> definitions)
        {
            _data = data;
            FileName = fileName;

            Columns.SetDefinitions(definitions);
        }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            byte[] csvBytes;

            if (Columns != null && Columns.Count > 0)
                csvBytes = await _data.GenerateCSVForDataAsync(Columns);
            else
                csvBytes = await _data.GenerateCSVForDataAsync();

            WriteExcelFileAsync(context.HttpContext, csvBytes);
        }

        private async void WriteExcelFileAsync(HttpContext context, byte[] bytes)
        {
            context.Response.Headers["content-disposition"] = $"attachment; filename={FileName}.csv";
            await context.Response.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
}
