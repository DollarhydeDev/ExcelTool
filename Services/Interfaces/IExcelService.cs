using ExcelTool.Models.Excel;

namespace ExcelTool.Services.Interfaces
{
    public interface IExcelService
    {
        ExcelWorkbook GetWorkbookAtPath(string filePath);
        void ExportWorkbook(ExcelWorkbook workbook, string filePath);

        string GetCSVDataFromWorkbook(ExcelWorkbook workbook, string sheetName, string filePath);

        string FormatCSVData(string csvData, string[] headerMapping, int headerStartPosition = 0);

        string GetRowFromCells(string[] cells);

        string[] GetRowsFromCSVData(string csvData, int rowStartPosition = 0);
        string[] GetCellsFromRow(string row);
    }
}