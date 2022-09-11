using Fingers10.ExcelExport.Extensions;
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

        public CSVResult(IEnumerable<T> data, string fileName, bool spaceAfterComma)
        {
            _data = data;
            FileName = fileName;
            SpaceAfterComma = spaceAfterComma;
        }

        public string FileName { get; }
        public bool SpaceAfterComma { get; }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            try
            {
                var csvBytes = await _data.GenerateCSVForDataAsync(SpaceAfterComma);
                WriteExcelFileAsync(context.HttpContext, csvBytes);

            }
            catch (Exception e)
            {
                Console.WriteLine(e);

                var errorBytes = await new List<T>().GenerateCSVForDataAsync(SpaceAfterComma);
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
