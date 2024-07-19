using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
            if (cell is null)
            {
                return string.Empty;
            }

            var value = cell.InnerText;
            var sharedStringTable = spreadsheetDocument
                .WorkbookPart
                .SharedStringTablePart
                .SharedStringTable;

            if (cell.DataType.Value == CellValues.SharedString) 
            {
                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
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
