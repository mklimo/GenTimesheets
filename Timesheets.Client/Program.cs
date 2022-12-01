using System.Diagnostics;
using System.IO;
using System.Net.Http.Headers;

namespace Timesheets.Client
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            using (var multipartFormContent = new MultipartFormDataContent())
            {
                using var fileStreamContent = new StreamContent(File.OpenRead(args[0]));
                fileStreamContent.Headers.ContentType = new MediaTypeHeaderValue(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                multipartFormContent.Add(fileStreamContent, name: "file", fileName: Path.GetFileName(args[0]));
                Console.ReadKey();
                using var httpClient = new HttpClient();
            
                  //var response = await httpClient.PostAsync("http://localhost:7071/api/ConvertTimesheet/", multipartFormContent);
                var response = await httpClient.PostAsync("http://gentimesheets.azurewebsites.net/api/ConvertTimesheet/", multipartFormContent);
                var responseContentStream = response.Content.ReadAsStream();
                string timesheetName = $"výkaz činnosti {DateTime.Now.AddMonths(-1).Month}_{DateTime.Now.Year}.xlsx";

                var resultFile = File.Create(timesheetName);
                responseContentStream.Position = 0;
                responseContentStream.CopyTo(resultFile);
                resultFile.Close();
                var p = new Process();
                p.StartInfo = new ProcessStartInfo { UseShellExecute = true, FileName = timesheetName };

                p.Start();
                
            }
        }
    }
}
