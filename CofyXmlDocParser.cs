using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc
{
    public static class CofyXmlDocParser
    {
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

            Dictionary<string, string> headers = null; //<columnName, headerName>

            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sharedStringTable == null)
            {
                throw new InvalidOperationException("Detected excel has no shared string stringTable.");
            }
            
            var rowCellCount = 0;
            foreach (var row in rows)
            {
                var rowIndex = row.RowIndex?.Value;
                if (rowIndex == null) continue;

                if (headers == null)
                {
                    rowCellCount = row.Count();
                    headers = new Dictionary<string, string>(rowCellCount);
                    ProcessHeaderRow(row);
                }
                else
                {
                    DataObject rowData = new(rowCellCount);
                    foreach (var element in row)
                    {
                        if (element is not Cell cell || string.IsNullOrEmpty(cell.CellReference))
                        {
                            throw new InvalidCastException($"Detected invalid or non cell element ({element.GetType()}) in row {rowIndex}");
                        }

                        var columnName = GetColumnName(cell.CellReference);
                        if (!headers.TryGetValue(columnName, out var key)) continue;
                        
                        var value = GetCellValue(cell);

                        if (key.Contains("."))
                        {
                            if (!rowData.TryGetValue(key, out var obj))
                            {
                                obj = new List<string>();
                                rowData[key] = obj;
                            }
                            var list = obj as List<string>;
                            list?.Add(value);
                        }
                        else
                        {
                            rowData[key] = value;
                        }
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
                    if (string.IsNullOrEmpty(cellValue) || !char.IsLetter(cellValue[0])) continue; //empty column

                    var columnName = GetColumnName(cell.CellReference);
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

            string GetColumnName(string cellReference)
            {
                for (var i = 0; i < cellReference.Length; i++)
                {
                    if (char.IsLetter(cellReference, i))
                    {
                        return cellReference[..(i + 1)];
                    }
                }

                return string.Empty;
            }
        }
    }
}