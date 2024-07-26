using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ProcedureStat
{
    public static class SpreadsheetUtil
    {
        public static SheetData GetSheetData(SpreadsheetDocument spreadsheetDocument) 
            => GetWorksheetPart(spreadsheetDocument).Worksheet.GetFirstChild<SheetData>();

        public static WorksheetPart GetWorksheetPart(SpreadsheetDocument spreadsheetDocument)
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;

            var sheet = workbookPart.Workbook
                .Descendants<Sheet>()
                .FirstOrDefault();

            var worksheetPart = (WorksheetPart)workbookPart
                .GetPartById(sheet.Id);

            return worksheetPart;
        }

        public static string GetCellValue(SpreadsheetDocument spreadsheetDocument, Cell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            if (spreadsheetDocument == null || spreadsheetDocument.WorkbookPart == null ||
                spreadsheetDocument.WorkbookPart.SharedStringTablePart == null)
            {
                throw new ArgumentNullException();
            }

            var value = cell.InnerText;
            var sharedStringTable = spreadsheetDocument
                .WorkbookPart
                .SharedStringTablePart
                .SharedStringTable;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (int.TryParse(value, out int index) && index >= 0 && index < sharedStringTable.Count())
                {
                    value = sharedStringTable.ElementAt(index).InnerText;
                }
                else
                {
                    throw new ArgumentOutOfRangeException();
                }
            }

            return value;
        }

        public static Cell CreateCell(string columnName, uint columnIndex, string text)
            => new ()
            {
                CellReference = columnName + columnIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };

        public static string ColumnIndexToName(int columnIndex)
        {
            var currentIndex = columnIndex + 1;
            var columnName = string.Empty;

            while (currentIndex > 0)
            {
                int remainder = (currentIndex - 1) % Constant.AlphabetLength;
                columnName += (char)(Constant.FirstLetter + remainder);
                currentIndex = (currentIndex - remainder) / Constant.AlphabetLength;
            }
            return columnName;
        }
    }
}
