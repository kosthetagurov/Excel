using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Excel
{
    public class ExcelDataImporter
    {
        public List<ExcelRow> LoadData(string path)
        {
            var excelRows = new List<ExcelRow>();
            using var spreadSheetDocument = SpreadsheetDocument.Open(path, true);
            var workbookPart = spreadSheetDocument.WorkbookPart;
            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

            foreach (var sheet in sheets)
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var rows = sheetData.Descendants<Row>();

                var headerRow = rows.First();
                var columnNames = headerRow.Descendants<Cell>().Select(cell => GetCellValue(spreadSheetDocument, cell)).ToList();

                foreach (var row in rows.Skip(1))
                {
                    var excelRow = new ExcelRow();
                    int columnIndex = 0;

                    foreach (var cell in row.Descendants<Cell>())
                    {
                        int cellColumnIndex = GetColumnIndexFromName(GetColumnName(cell.CellReference)) - 1;

                        while (columnIndex < cellColumnIndex)
                        {
                            excelRow.Cells[columnNames[columnIndex]] = string.Empty;
                            columnIndex++;
                        }

                        excelRow.Cells[columnNames[columnIndex]] = GetCellValue(spreadSheetDocument, cell);
                        columnIndex++;
                    }

                    if (excelRow.IsRowEmpty() == false)
                    {
                        excelRows.Add(excelRow);
                    }
                }
            }

            return excelRows;
        }

        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            string value = cell.CellValue.InnerText;
            return cell.DataType != null && cell.DataType == CellValues.SharedString
                ? document.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText.Trim()
                : value.Trim();
        }

        private int GetColumnIndexFromName(string columnName)
        {
            int columnIndex = 0;
            foreach (char c in columnName)
            {
                columnIndex = (columnIndex * 26) + (c - 'A' + 1);
            }
            return columnIndex;
        }

        private string GetColumnName(string cellReference)
        {
            return Regex.Match(cellReference, "[A-Za-z]+").Value;
        }
    }
}
