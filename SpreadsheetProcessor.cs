using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ProcedureStat
{
    public class SpreadsheetProcessor
    {
        private readonly string _inputFilePath;
        private readonly string _outputFilePath;

        public SpreadsheetProcessor(string inputFilePath, string outputFilePath)
        {
            _inputFilePath = inputFilePath;
            _outputFilePath = outputFilePath;
        }

        public void ProcessDocument(Dictionary<string, List<string>> details, List<string> columns, string objectNameColumn = Constant.ObjectNameColumn)
        {
            File.Copy(_inputFilePath, _outputFilePath, overwrite: true);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_outputFilePath, isEditable: true))
            {
                SheetData sheetData = SpreadsheetUtil.GetSheetData(spreadsheetDocument);

                Row headerRow = GetHeaderRow(sheetData);
                var objectNameColumnIndex = GetColumnIndex(headerRow, spreadsheetDocument, objectNameColumn);
                var headerStyleIndex = GetCellStyleIndex(headerRow);
                var dataCellStyleIndex = GetCellStyleIndex(sheetData.Elements<Row>().Skip(1).FirstOrDefault());

                AddColumnsToHeader(headerRow, columns, headerStyleIndex, spreadsheetDocument, 25);

                foreach (var row in sheetData.Elements<Row>().Skip(1))
                {
                    AddValuesToRow(row, objectNameColumnIndex, details, spreadsheetDocument);
                }

                SpreadsheetUtil.GetWorksheetPart(spreadsheetDocument).Worksheet.Save();
            }
        }

        private Row GetHeaderRow(SheetData sheetData)
            => sheetData.Elements<Row>().FirstOrDefault();

        private int GetColumnIndex(Row headerRow, SpreadsheetDocument spreadsheetDocument, string objectNameColumn)
        {
            int keyColumnIndex = -1;
            int columnIndex = 0;

            foreach (Cell cell in headerRow.Elements<Cell>())
            {
                var cellValue = SpreadsheetUtil.GetCellValue(spreadsheetDocument, cell);
                if (cellValue == objectNameColumn)
                {
                    keyColumnIndex = columnIndex;
                    break;
                }
                columnIndex++;
            }

            if (keyColumnIndex == -1)
            {
                throw new Exception($"Ключевой столбец '{objectNameColumn}' не найден.");
            }

            return columnIndex;
        }

        private static uint GetCellStyleIndex(Row row)
            => row.Elements<Cell>().FirstOrDefault().StyleIndex;

        private static void AddColumnsToHeader(Row headerRow, List<string> columns, uint styleIndex, SpreadsheetDocument doc, double width)
        {
            var columnIndex = headerRow.Elements<Cell>().Count();
            WorksheetPart worksheetPart = SpreadsheetUtil.GetWorksheetPart(doc);
            Columns cols = worksheetPart.Worksheet.Elements<Columns>().FirstOrDefault();

            if (cols == null)
            {
                cols = new Columns();
                worksheetPart.Worksheet.InsertAt(cols, 0);
            }

            foreach (var column in columns)
            {
                var columnName = SpreadsheetUtil.ColumnIndexToName(columnIndex);
                var rowIndex = headerRow.RowIndex.Value;

                Cell cell = SpreadsheetUtil.CreateCell(columnName, rowIndex, column);
                cell.StyleIndex = styleIndex;

                headerRow.AppendChild(cell);

                Column col = new Column()
                {
                    Min = (uint)columnIndex + 1,
                    Max = (uint)columnIndex + 1,
                    Width = width,
                    CustomWidth = true,
                };
                cols.Append(col);
                columnIndex++;
            }
            worksheetPart.Worksheet.Save();
        }

        private static void AddValuesToRow(Row row, int keyColumnIndex, Dictionary<string, List<string>> details, SpreadsheetDocument spreadsheetDocument)
        {
            var key = SpreadsheetUtil.GetCellValue(spreadsheetDocument,
                row.Elements<Cell>().ElementAt(keyColumnIndex));

            var stylesheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;

            uint redColorStyleIndex = CreateFillStyle(stylesheet, "fd7b7c");
            uint greenColorStyleIndex = CreateFillStyle(stylesheet, "57A639");
            uint yellowColorStyleIndex = CreateFillStyle(stylesheet, "fde910");

            if (details.ContainsKey(key))
            {
                List<string> values = details[key];
                int columnIndex = row.Elements<Cell>().Count();

                foreach (var value in values)
                {
                    var columnName = SpreadsheetUtil.ColumnIndexToName(columnIndex);
                    var rowIndex = row.RowIndex.Value;

                    Cell cell = SpreadsheetUtil.CreateCell(columnName, rowIndex, value);

                    if (value == "Есть ошибки")
                    {
                        cell.StyleIndex = redColorStyleIndex;
                    }

                    if (value == "Да")
                    {
                        cell.StyleIndex = greenColorStyleIndex;
                    }

                    if (value.StartsWith("(Обратить внимание)"))
                    {
                        cell.StyleIndex = yellowColorStyleIndex;
                    }

                    row.AppendChild(cell);
                    columnIndex++;
                }
            }
        }

        private static uint CreateFillStyle(Stylesheet stylesheet, string hexColor)
        {
            uint styleIndex = 0;
            if (stylesheet.Fills == null)
            {
                stylesheet.Fills = new Fills();
            }

            var fill = new Fill(
                new PatternFill(
                    new ForegroundColor() { Rgb = new HexBinaryValue() { Value = hexColor } }
                )
                { PatternType = PatternValues.Solid }
            );

            stylesheet.Fills.Append(fill);
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats();
                stylesheet.CellFormats.AppendChild(new CellFormat());
            }

            var cellFormat = new CellFormat() { FillId = stylesheet.Fills.Count - 1, ApplyFill = true };
            stylesheet.CellFormats.Append(cellFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            styleIndex = stylesheet.CellFormats.Count - 1;

            stylesheet.Save();
            return styleIndex;
        }
    }
}
