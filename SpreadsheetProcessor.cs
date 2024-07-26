using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ProcedureStat
{
    public class SpreadsheetProcessor
    {
        private readonly string _inputFilePath;
        private readonly string _outputFilePath;

        private const string RedColor = "FD7B7C";
        private const string GreenColor = "57A639";
        private const string YellowColor = "FDE910";

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

                var stylesheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;

                uint redFill = CreateFillStyle(stylesheet, hexColor: RedColor);
                uint greenFill = CreateFillStyle(stylesheet, hexColor: GreenColor);
                uint yellowFill = CreateFillStyle(stylesheet, hexColor: YellowColor);

                var styles = new List<uint>()
                {
                    redFill,
                    greenFill,
                    yellowFill,
                };

                foreach (var row in sheetData.Elements<Row>().Skip(1))
                {
                    AddValuesToRow(row, objectNameColumnIndex, details, spreadsheetDocument, styles);
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

        private static void AddValuesToRow(Row row, int keyColumnIndex, Dictionary<string, List<string>> details, SpreadsheetDocument spreadsheetDocument, List<uint> styles)
        {
            var key = SpreadsheetUtil.GetCellValue(spreadsheetDocument,
                row.Elements<Cell>().ElementAt(keyColumnIndex));

            var rowIndex = row.RowIndex.Value;
            int columnIndex = row.Elements<Cell>().Count();

            if (details.ContainsKey(key))
            {
                List<string> values = details[key];

                foreach (var value in values)
                {
                    var columnName = SpreadsheetUtil.ColumnIndexToName(columnIndex);
                    var cell = SpreadsheetUtil.CreateCell(columnName, rowIndex, text: value);

                    if (value.StartsWith("Есть ошибки") || value == "-")
                    {
                        cell.StyleIndex = styles[0];
                    }

                    if (value == "Да")
                    {
                        cell.StyleIndex = styles[1];
                    }

                    if (value.StartsWith("(Обратить внимание)"))
                    {
                        cell.StyleIndex = styles[2];
                    }

                    row.AppendChild(cell);
                    columnIndex++;
                }
            }
            else if (key.StartsWith(Constant.SchemeName))
            {
                var columnName = SpreadsheetUtil.ColumnIndexToName(columnIndex);
                var cell = SpreadsheetUtil.CreateCell(columnName, rowIndex, text: "Объект не найден");
                cell.StyleIndex = styles[0];

                row.AppendChild(cell);
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
