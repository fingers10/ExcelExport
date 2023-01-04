using Fingers10.ExcelExport.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Fingers10.ExcelExport.Models;

namespace Fingers10.ExcelExport.ActionResults
{
    public class ExcelResult<T> : IActionResult where T : class
    {
        private readonly IEnumerable<T> _data;

        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;
        }

        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName, List<ExcelColumnDefinition> columns)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;
            Columns = columns == null ? new List<ExcelColumnDefinition>() : columns;
        }

        /// <summary>
        /// definitions (column name, label, order)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        /// <param name="definitions"></param>
        public ExcelResult(IEnumerable<T> data, string sheetName, string fileName, params (string, string, int)[] definitions)
        {
            _data = data;
            SheetName = sheetName;
            FileName = fileName;
            
            SetDefinitions(definitions);
        }

        public string SheetName { get; }
        public string FileName { get; }
        private List<ExcelColumnDefinition> Columns { get; set; }

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
        public async Task ExecuteResultAsync(ActionContext context)
        {
            try
            {
                byte[] excelBytes;

                if (Columns != null && Columns.Count > 0)
                    excelBytes = await _data.GenerateExcelForDataTableAsync(SheetName, Columns);
                else
                    excelBytes = await _data.GenerateExcelForDataTableAsync(SheetName);

                WriteExcelFileAsync(context.HttpContext, excelBytes);

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private async void WriteExcelFileAsync(HttpContext context, byte[] bytes)
        {
            context.Response.Headers["content-disposition"] = $"attachment; filename={FileName}.xlsx";
            await context.Response.Body.WriteAsync(bytes, 0, bytes.Length);
        }
    }
}
