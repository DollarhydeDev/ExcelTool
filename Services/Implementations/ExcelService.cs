using ExcelTool.Models.Excel;
using ExcelTool.Services.Interfaces;
using ClosedXML.Excel;
using System.Text;

namespace ExcelTool.Services.Implementations
{
    public class ExcelService : IExcelService
    {
        private string RebuildRowFromHeaderMappings(Dictionary<int, int> headerMappings, string[] rowToRebuild)
        {
            var newCSVData = new StringBuilder();

            for (int i = 0; i < rowToRebuild.Length; i++)
            {
                var currentRowCells = GetCellsFromRow(rowToRebuild[i]);
                var newRowCells = new string[headerMappings.Count];

                for (int j = 0; j < headerMappings.Count; j++)
                {
                    newRowCells[j] = currentRowCells[headerMappings[j]];
                }

                newCSVData.Append(GetRowFromCells(newRowCells));
            }

            return newCSVData.ToString();
        }
        private Dictionary<int, int> GenerateHeaderMapping(string[] currentHeaders, string[] headerMapping)
        {
            if (currentHeaders.Length == 0)
            {
                throw new Exception("CSV headers are missing");
            }
            if (headerMapping.Length == 0)
            {
                throw new Exception("Mapping headers are missing");
            }

            var mappingDictionary = new Dictionary<int, int>();
            for (int i = 0; i < headerMapping.Length; i++)
            {
                for (int j = 0; j < currentHeaders.Length; j++)
                {
                    if (headerMapping[i] == currentHeaders[j])
                    {
                        mappingDictionary.Add(i, j);
                        break;
                    }
                }

                if (!mappingDictionary.ContainsKey(i)) throw new Exception($"No header found for mapping {i}: '{headerMapping[i]}'");
            }

            return mappingDictionary;
        }
        private ExcelService() { }

        public static async Task<IExcelService> CreateAsync()
        {
            await Task.CompletedTask;
            return new ExcelService();
        }

        public ExcelWorkbook GetWorkbookAtPath(string filePath)
        {
            return new ExcelWorkbook(this, File.Exists(filePath) ? new XLWorkbook(filePath) : new XLWorkbook());
        }
        public void ExportWorkbook(ExcelWorkbook workbook, string filePath)
        {
            workbook.SaveAs(filePath);
        }

        public string GetCSVDataFromWorkbook(ExcelWorkbook workbook, string sheetName, string filePath)
        {
            var worksheet = workbook.GetWorksheet(sheetName);
            return worksheet.GetCSVData();
        }

        public string FormatCSVData(string csvData, string[] headerMapping, int headerStartPosition)
        {
            if (string.IsNullOrWhiteSpace(csvData))
            {
                throw new Exception("CSV data is empty.");
            }

            var csvRows = GetRowsFromCSVData(csvData, headerStartPosition);
            var csvHeaders = GetCellsFromRow(csvRows[0]);
            var headerMappings = GenerateHeaderMapping(csvHeaders, headerMapping);

            // Maybe loop through each row until we find header row in the future?
            return RebuildRowFromHeaderMappings(headerMappings, csvRows);
        }

        public string GetRowFromCells(string[] cells)
        {
            var newCSVRow = new StringBuilder();
            var formattedCells = new string[cells.Length];

            for (int i = 0; i < cells.Length; i++)
            {
                var cellValue = cells[i];
                if (cellValue.Contains(",") || cellValue.Contains("\"") || cellValue.Contains("\n"))
                {
                    cellValue = "\"" + cellValue.Replace("\"", "\"\"") + "\"";
                }

                formattedCells[i] = cellValue;
            }

            newCSVRow.AppendLine(string.Join(',', formattedCells));
            return newCSVRow.ToString();
        }

        public string[] GetRowsFromCSVData(string csvData, int rowStartPosition)
        {
            var csvRows = csvData.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (csvRows.Length == 0)
            {
                throw new Exception("No rows found in CSV.");
            }

            if (rowStartPosition > 0)
            {
                List<string> formattedRows = new List<string>();
                for (int i = rowStartPosition; i < csvRows.Length; i++)
                {
                    formattedRows.Add(csvRows[i]);
                }

                if (formattedRows.Count == 0)
                {
                    throw new Exception("No rows found in CSV.");
                }

                return formattedRows.ToArray();
            }

            return csvRows;
        }
        public string[] GetCellsFromRow(string row)
        {
            List<string> cells = new List<string>();
            StringBuilder currentCell = new StringBuilder();
            bool insideQuotes = false;

            for (int i = 0; i < row.Length; i++)
            {
                char character = row[i];

                if (character == '"')
                {
                    if (insideQuotes && i + 1 < row.Length && row[i + 1] == '"')
                    {
                        currentCell.Append('"');
                        i++;
                    }
                    else
                    {
                        insideQuotes = !insideQuotes;
                    }
                }
                else if (character == ',' && !insideQuotes)
                {
                    cells.Add(currentCell.ToString());
                    currentCell.Clear();
                }
                else
                {
                    currentCell.Append(character);
                }
            }

            cells.Add(currentCell.ToString());
            return cells.ToArray();
        }
    }
}