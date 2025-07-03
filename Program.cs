using ExcelTool.Services.Implementations;
using System.Threading.Tasks;

namespace ExcelTool
{
    internal class Program
    {
        static async Task Main()
        {
            // Setup excel service
            var excelService = await ExcelService.CreateAsync();

            // Get CSV data
            var exampleCSV = "Name,Age,Location\nJohn Doe,30,New York\nJane Smith,25,Los Angeles\n";
            var headerMapping = new[] { "Location", "Name", "Age" };
            var formattedCSV = excelService.FormatCSVData(exampleCSV, headerMapping, 0);

            // Save to workbook
            using var workbook = excelService.GetWorkbookAtPath("example.xlsx");
            var worksheet = workbook.GetWorksheet("ExampleSheet");
            worksheet.ImportCSVData(formattedCSV);
            workbook.Save();

            // Import CSV from workbook
            var exampleCSVData2 = worksheet.GetCSVData();
            var formattedCSV2 = excelService.FormatCSVData(exampleCSVData2, headerMapping, 0);
        }
    }
}
