using Fingers10.ExcelExport.Extensions;
using Fingers10.ExcelExport.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Fingers10.ExcelExport.ActionResults
{
    public class ExcelResult<T> : IActionResult where T : class
    {
        private readonly IEnumerable<T> _data;
        public string SheetName { get; }
        public string FileName { get; }
        private List<ExcelColumnDefinition> Columns { get; set; } = new List<ExcelColumnDefinition>();

        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;
        }

        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName, List<ExcelColumnDefinition> definitions)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;

            Columns.SetDefinitions(definitions);
        }

        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName, params (string name, string label, int order)[] definitions)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;

            Columns.SetDefinitions(definitions);
        }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            byte[] excelBytes;

            if (Columns != null && Columns.Count > 0)
                excelBytes = await _data.GenerateExcelForDataTableAsync(SheetName, Columns);
            else
                excelBytes = await _data.GenerateExcelForDataTableAsync(SheetName);

            WriteExcelFileAsync(context.HttpContext, excelBytes);
        }

        private async void WriteExcelFileAsync(HttpContext context, byte[] bytes)
        {
            context.Response.Headers["content-disposition"] = $"attachment; filename={FileName}.xlsx";
            await context.Response.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
}
