using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc
{
    public static class CofyXmlDocParser
    {
        public const int HEADER_ROW_INDEX = 1;
        private static Regex COLUMN_NAME_FILTER = new Regex("[a-zA-Z]");

        public class DataObject : Dictionary<string, string>
        {
            public DataObject subDataObject;
        }

        public class DataContainer : List<DataObject>
        {
        }

        private static bool IsSheetAvailable(Sheet sheet)
        {
            return sheet.State == null || sheet.State == SheetStateValues.Visible;
        }

        public static DataContainer ParseExcel(byte[] fileBytes)
        {
            using MemoryStream memoryStream = new MemoryStream(fileBytes);
            using var document = SpreadsheetDocument.Open(memoryStream, false);

            var workbookPart = document.WorkbookPart;
            if (workbookPart == null)
            {
                throw new InvalidOperationException("Excel does not have a workbook");
            }

            var sheets = workbookPart.Workbook.Descendants<Sheet>().Where(IsSheetAvailable);

            DataContainer rootDataContainer = new();
            foreach (var sheet in sheets)
            {
                if (sheet.Id == null || !sheet.Id.HasValue || sheet.Id.Value == null || !sheet.Id.HasValue)
                    continue;

                var sheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                ProcessSheet(sheetPart.Worksheet, workbookPart, rootDataContainer);
            }

            return rootDataContainer;
        }

        private static void ProcessSheet(Worksheet sheet, WorkbookPart workbookPart, in DataContainer rootDataContainer)
        {
            var rows = sheet.Descendants<Row>().Where(r => r.RowIndex is not null);

            var headers = new Dictionary<string, string>(); //<columnName, headerName>

            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sharedStringTable == null)
            {
                throw new InvalidOperationException("Detected excel has no shared string stringTable.");
            }

            foreach (var row in rows)
            {
                var rowIndex = row.RowIndex?.Value;
                if (rowIndex == null) continue;

                if (rowIndex == HEADER_ROW_INDEX)
                {
                     ProcessHeaderRow(row);
                }
                else
                {
                    DataObject rowData = new();
                    foreach (var element in row)
                    {
                        if (element is not Cell cell || string.IsNullOrEmpty(cell.CellReference))
                        {
                            throw new InvalidCastException(
                                $"Detected invalid or non cell element ({element.GetType()}) in row {rowIndex}");
                        }

                        var columnName = COLUMN_NAME_FILTER.Match(cell.CellReference).Value;
                        if (!headers.TryGetValue(columnName, out var key))
                        {
                            throw new KeyNotFoundException($"Header not found for column name {columnName}");
                        }

                        var value = GetCellValue(cell);
                        rowData[key] = value;
                    }

                    rootDataContainer.Add(rowData);
                }
            }

            void ProcessHeaderRow(Row row)
            {
                int columnIndex = 0;
                foreach (var element in row)
                {
                    if (element is not Cell cell || string.IsNullOrEmpty(cell.CellReference))
                    {
                        throw new InvalidCastException(
                            $"Detected invalid or non cell element ({element.GetType()}) in header row");
                    }

                    if (cell.CellReference is { Value: null }) continue;

                    var cellValue = GetCellValue(cell);
                    if (string.IsNullOrEmpty(cellValue)) continue; //empty column

                    var columnName = COLUMN_NAME_FILTER.Match(cell.CellReference).Value;
                    if (!headers.TryAdd(columnName, cellValue))
                    {
                        throw new ArgumentException($"duplicated header with column {cell.CellReference}");
                    }

                    columnIndex++;
                }
            }

            string GetCellValue(Cell cell)
            {
                if (cell.CellValue == null) return string.Empty;

                string value = cell.CellValue.InnerText;

                if (cell.DataType == null) return value;

                var dataType = cell.DataType.Value;

                if (dataType == CellValues.SharedString)
                    return sharedStringTable.ChildElements[int.Parse(value)].InnerText;

                if (dataType == CellValues.Boolean)
                    return value == "0" ? bool.FalseString : bool.TrueString;

                return value;
            }
        }
    }
}