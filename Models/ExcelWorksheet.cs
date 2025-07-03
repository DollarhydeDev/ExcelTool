using ExcelTool.Services.Interfaces;
using ClosedXML.Excel;
using System.Text;

namespace ExcelTool.Models.Excel
{
    public class ExcelWorksheet
    {
        private readonly IExcelService _excelService;
        private readonly IXLWorksheet _xlWorkSheet;

        public string SheetName => _xlWorkSheet.Name;

        public ExcelWorksheet(IExcelService excelService, IXLWorksheet xlWorkSheet)
        {
            _excelService = excelService;
            _xlWorkSheet = xlWorkSheet;
        }

        public string GetCSVData()
        {
            var csvData = new StringBuilder();

            var lastColumn = _xlWorkSheet.LastColumnUsed();
            var lastRow = _xlWorkSheet.LastRowUsed();

            if (lastColumn == null || lastRow == null)
            {
                return string.Empty;
            }

            for (int i = 0; i < lastRow.RowNumber(); i++)
            {
                var cellValues = new List<string>();

                for (int j = 0; j < lastColumn.ColumnNumber(); j++)
                {
                    var cell = _xlWorkSheet.Cell(i + 1, j + 1);
                    var cellValue = cell.GetFormattedString();
                    cellValues.Add(cellValue);
                }

                csvData.Append(_excelService.GetRowFromCells(cellValues.ToArray()));
            }

            return csvData.ToString();
        }

        public void ImportCSVData(string csvData, bool startFromLowestRow = false)
        {
            if (string.IsNullOrWhiteSpace(csvData)) return;

            var csvRows = _excelService.GetRowsFromCSVData(csvData);
            int startingRow;

            if (startFromLowestRow)
            {
                startingRow = _xlWorkSheet.LastRowUsed()?.RowNumber() ?? 0;
                startingRow += 3;
            }
            else
            {
                startingRow = 0;
            }

            for (int i = 0; i < csvRows.Length; i++)
            {
                var rowCells = _excelService.GetCellsFromRow(csvRows[i]);

                for (int j = 0; j < rowCells.Length; j++)
                {
                    _xlWorkSheet.Cell(startingRow + i + 1, j + 1).Value = rowCells[j];
                }
            }
        }
    }
}