using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Timesheets;
using System.Linq;

namespace TimesheetFunction
{
    public static class ConvertTimesheet
    {
        [FunctionName("ConvertTimesheet")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req, ExecutionContext context)
        {
            var file = req.Form.Files[0];

            using var template = File.OpenRead(Path.Combine(context.FunctionAppDirectory, "generali výkaz èinnosti.xlsx"));

            var stream = Converter.GetTimesheet(file.OpenReadStream(), template);
            stream.Position = 0;

            return new FileStreamResult(stream, "application/octet-stream");
        }
    }
}