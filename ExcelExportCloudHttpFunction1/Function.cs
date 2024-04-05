using Google.Cloud.Functions.Framework;
using GrapeCity.Documents.Excel;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Threading.Tasks;

namespace ExcelExportCloudHttpFunction1
{
    public class Function : IHttpFunction
    {
        /// <summary>
        /// Logic for your function goes here.
        /// </summary>
        /// <param name="context">The HTTP context, containing the request and the response.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task HandleAsync(HttpContext context)
        {
            HttpRequest request = context.Request;
            string name = request.Query["name"].ToString();

            string Message = string.IsNullOrEmpty(name)
                ? "Hello, World!!"
                : $"Hello, {name}!!";

            //Workbook.SetLicenseKey("");

            Workbook workbook = new();
            workbook.Worksheets[0].Range["A1"].Value = Message;

            byte[] output;

            using (MemoryStream ms = new())
            {
                workbook.Save(ms, SaveFileFormat.Xlsx);
                output = ms.ToArray();
            }

            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            context.Response.Headers.Add("Content-Disposition", "attachment;filename=Result.xlsx");
            await context.Response.Body.WriteAsync(output);
        }
    }
}
