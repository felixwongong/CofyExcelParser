using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc;

public static class CofyXmlDocParser
{
    public const int HEADER_ROW_INDEX = 1;

    public class DataObject: Dictionary<string, string>
    {
        public DataObject? subDataObject;
    }

    public class DataContainer : List<DataObject>
    {
    }

    private static byte[] LoadExcelBytes(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            throw new ArgumentNullException(nameof(filePath), "filename is empty");
        }

        if (!File.Exists(filePath))
        {
            throw new ArgumentNullException(nameof(filePath), "file does not exist");
        }

        byte[] fileBytes;
        using var filestream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        fileBytes = new byte[filestream.Length];
        var byteLoaded = filestream.Read(fileBytes, 0, (int)filestream.Length);
        if (byteLoaded > filestream.Length)
        {
            throw new InvalidOperationException(
                $"Detect filestream read size ({filestream.Length} differ from byte load ({byteLoaded}))");
        }

        return fileBytes;
    }

    private static bool IsSheetAvailable(Sheet sheet)
    {
        return sheet.State == null || sheet.State == SheetStateValues.Visible;
    }
    
    public static DataContainer ParseExcel(string filePath)
    {
        byte[] fileBytes = LoadExcelBytes(filePath);

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
        
        var headers = new Dictionary<int, string>();    //<columnIndex, headerName>
        int maxHeaderColumnIndex = -1;

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable == null)
        {
            throw new InvalidOperationException("Detected excel has no shared string stringTable.");
        }

        foreach (var row in rows)
        {
            var rowIndex = row.RowIndex?.Value;
            if(rowIndex == null) continue;

            if (rowIndex == HEADER_ROW_INDEX)
            {
                maxHeaderColumnIndex = ProcessHeaderRow(row);
            }
            else
            {
                DataObject rowData = new();
                for (int i = 0; i < maxHeaderColumnIndex; i++)
                {
                    if(i >= row.Count()) break;
                    if (row.ElementAt(i) is not Cell cell)
                    {
                        throw new InvalidCastException(
                            $"Detected invalid or non cell element ({row.ElementAt(i).GetType()}) in row {rowIndex}");
                    }

                    if (!headers.TryGetValue(i, out var key))
                    {
                        throw new KeyNotFoundException($"Header not found in column index {i}");
                    }
                    var value = GetCellValue(cell);
                    rowData[key] = value;
                }
                rootDataContainer.Add(rowData);
            }
        }

        int ProcessHeaderRow(Row row)
        {
            int columnIndex = 0;
            foreach (var element in row)
            {
                if (element is not Cell cell)
                {
                    throw new InvalidCastException(
                        $"Detected invalid or non cell element ({element.GetType()}) in header row");
                }

                if (cell.CellReference is { Value: null }) continue;

                var cellValue = GetCellValue(cell); 
                if(string.IsNullOrEmpty(cellValue)) continue;   //empty column
                
                if (!headers.TryAdd(columnIndex, cellValue))
                {
                    throw new ArgumentException($"duplicated header with column index {columnIndex}");
                }

                columnIndex++;
            }

            return columnIndex;
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