using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Timesheets
{
    public static class Converter
    {
        public static Stream GetTimesheet(Stream input, Stream template)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var sourcePackage = new ExcelPackage(input))
            using (var templatePackage = new ExcelPackage(template))
            {

                var sourceSheet = sourcePackage.Workbook.Worksheets["Worklogs"];
                var start = sourceSheet.Dimension.Start;
                var end = sourceSheet.Dimension.End;

                var destinationSheet = templatePackage.Workbook.Worksheets.First();

                for (int i = 2; i <= end.Row; i++)
                {
                    var activity = sourceSheet.Cells[i, 2].Value + " - " + sourceSheet.Cells[i, 23].Value;
                    var time = sourceSheet.Cells[i, 3].Value;
                    var excelDate = sourceSheet.Cells[i, 4].Value;
                    var date2 = DateTime.FromOADate((double)excelDate);

                    destinationSheet.Cells[i + 11, 1].Value = date2.ToShortDateString();
                    destinationSheet.Cells[i + 11, 2].Value = time;
                    destinationSheet.Cells[i + 11, 4].Value = activity;

                }

                templatePackage.Save();
                return templatePackage.Stream;

            }
        }
    }
}
