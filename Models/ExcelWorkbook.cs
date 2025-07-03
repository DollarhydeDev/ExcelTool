using ExcelTool.Services.Interfaces;
using ClosedXML.Excel;

namespace ExcelTool.Models.Excel
{
    public class ExcelWorkbook : IDisposable
    {
        private readonly IExcelService _excelService;
        private readonly XLWorkbook _xlWorkbook;

        public ExcelWorkbook(IExcelService excelService, XLWorkbook xlWorkbook)
        {
            _excelService = excelService;
            _xlWorkbook = xlWorkbook;
        }

        public ExcelWorksheet GetWorksheet(string sheetName)
        {
            var worksheet = _xlWorkbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name?.Trim(), sheetName?.Trim(), StringComparison.OrdinalIgnoreCase));
            if (worksheet == null)
            {
                worksheet = _xlWorkbook.Worksheets.Add(sheetName);
            }

            return new ExcelWorksheet(_excelService, worksheet);
        }
        public ExcelWorksheet GetWorksheetAt(int index)
        {
            if (index < 0 || index >= _xlWorkbook.Worksheets.Count) throw new ArgumentOutOfRangeException(nameof(index), $"Worksheet index {index} is out of range.");

            var worksheet = _xlWorkbook.Worksheets.ElementAt(index);
            return new ExcelWorksheet(_excelService, worksheet);
        }

        public void SaveAs(string filepath)
        {
            _xlWorkbook.SaveAs(filepath);
        }

        public void Save()
        {
            _xlWorkbook.Save();
        }

        public void Dispose()
        {
            _xlWorkbook.Dispose();
        }
    }
}