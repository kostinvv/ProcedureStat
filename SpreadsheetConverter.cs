using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;

namespace ProcedureStat
{
    public class SpreadsheetConverter
    {
        private readonly string _spreadseetFilePath;

        public SpreadsheetConverter(string spreadseetFilePath) 
        {
            _spreadseetFilePath = spreadseetFilePath;
        }

        public DataTable ConvertToDataTable(int headerIndex = 1)
        {
            var dataTable = new DataTable();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(_spreadseetFilePath, isEditable: false))
            {
                SheetData sheetData = SpreadsheetUtil.GetSheetData(spreadsheetDocument);
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                PopulateDataTableFromRows(dataTable, rows, headerIndex, spreadsheetDocument);
            }

            return dataTable;
        }

        private static void PopulateDataTableFromRows(DataTable dt, IEnumerable<Row> rows, int headerIndex, SpreadsheetDocument document)
        {
            var rowIndex = 0;
            foreach (var row in rows)
            {
                rowIndex++;
                var cells = row.Descendants<Cell>();
                if (rowIndex == headerIndex)
                {
                    AddColumnsFromHeaderRow(dt, cells, document);
                }
                else
                {
                    AddDataRows(dt, cells, document);
                }
            }
        }

        private static void AddColumnsFromHeaderRow(DataTable dt, IEnumerable<Cell> cells, SpreadsheetDocument document)
        {
            foreach (var cell in cells)
            {
                string columnName = SpreadsheetUtil.GetCellValue(document, cell);
                dt.Columns.Add(columnName);
            }
        }

        private static void AddDataRows(DataTable dt, IEnumerable<Cell> cells, SpreadsheetDocument document)
        {
            dt.Rows.Add();
            var rowCount = dt.Rows.Count - 1;
            var columnIndex = 0;

            foreach (var cell in cells)
            {
                dt.Rows[rowCount][columnIndex] = SpreadsheetUtil.GetCellValue(document, cell);
                columnIndex++;
            }
        }
    }
}
