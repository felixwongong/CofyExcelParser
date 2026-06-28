using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc;

public static class CofyXmlDocParser
{
    public static DataTable ParseExcel(byte[] fileBytes)
    {
        using var memoryStream = new MemoryStream(fileBytes);
        using var document = SpreadsheetDocument.Open(memoryStream, false);

        var workbookPart = document.WorkbookPart;
        if (workbookPart == null)
            throw new InvalidOperationException("Excel does not have a workbook");

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable == null)
            throw new InvalidOperationException("Detected excel has no shared string stringTable.");

        var rootDataTable = new DataTable();
        var sheets = workbookPart.Workbook.Descendants<Sheet>().Where(IsSheetAvailable);

        foreach (var sheet in sheets)
        {
            if (sheet.Id == null || !sheet.Id.HasValue || sheet.Id.Value == null)
                continue;

            var sheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            var rows = ExtractRows(sheetPart.Worksheet, sharedStringTable, sheet.Name);
            var table = new RawSheetTable(sheet.Name, rows);
            var data = SheetTableConverter.ToDataTable(table);
            foreach (var item in data)
                rootDataTable.Add(item);
        }

        return rootDataTable;
    }

    private static bool IsSheetAvailable(Sheet sheet)
    {
        return sheet.State == null || sheet.State == SheetStateValues.Visible;
    }

    private static IReadOnlyList<IReadOnlyList<string>> ExtractRows(
        Worksheet worksheet,
        SharedStringTable sharedStringTable,
        string sheetName)
    {
        var rows = worksheet.Descendants<Row>()
            .Where(r => r.RowIndex is not null && r.FirstChild != null && r.FirstChild.Any())
            .ToList();

        if (rows.Count == 0)
            return Array.Empty<IReadOnlyList<string>>();

        var maxColumnIndex = GetMaxColumnIndex(rows[0], sheetName, true);
        if (maxColumnIndex < 0)
            return Array.Empty<IReadOnlyList<string>>();

        var result = new List<IReadOnlyList<string>>(rows.Count);
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            var rowValues = new string[maxColumnIndex + 1];
            foreach (var element in row)
            {
                if (element is not Cell cell || string.IsNullOrEmpty(cell.CellReference))
                {
                    throw new InvalidCastException(
                        $"Detected invalid or non cell element ({element.GetType()}) in sheet '{sheetName}' row {rowIndex + 1}");
                }

                var columnIndex = GetColumnIndex(cell.CellReference);
                if (columnIndex < 0 || columnIndex > maxColumnIndex)
                    continue;

                rowValues[columnIndex] = GetCellValue(cell, sharedStringTable);
            }

            result.Add(rowValues);
        }

        return result;
    }

    private static int GetMaxColumnIndex(Row headerRow, string sheetName, bool isHeader)
    {
        var maxIndex = -1;
        foreach (var element in headerRow)
        {
            if (element is not Cell cell || string.IsNullOrEmpty(cell.CellReference))
            {
                throw new InvalidCastException(
                    $"Detected invalid or non cell element ({element.GetType()}) in sheet '{sheetName}' {(isHeader ? "header" : "data")} row");
            }

            var columnIndex = GetColumnIndex(cell.CellReference);
            if (columnIndex > maxIndex)
                maxIndex = columnIndex;
        }

        return maxIndex;
    }

    private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
    {
        if (cell.CellValue == null)
            return string.Empty;

        var value = cell.CellValue.InnerText;

        if (cell.DataType == null)
            return value;

        var dataType = cell.DataType.Value;

        if (dataType == CellValues.SharedString)
            return sharedStringTable.ChildElements[int.Parse(value)].InnerText;

        if (dataType == CellValues.Boolean)
            return value == "0" ? bool.FalseString : bool.TrueString;

        return value;
    }

    private static int GetColumnIndex(string cellReference)
    {
        var columnName = GetColumnName(cellReference);
        var index = 0;
        foreach (var character in columnName)
        {
            index = index * 26 + (character - 'A' + 1);
        }

        return index - 1;
    }

    private static string GetColumnName(string cellReference)
    {
        for (var i = 0; i < cellReference.Length; i++)
        {
            if (char.IsLetter(cellReference, i))
                return cellReference[..(i + 1)];
        }

        return string.Empty;
    }
}
