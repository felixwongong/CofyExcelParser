using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc;

public class CofyXmlDocParser
{
    public const int HEADER_ROW_INDEX = 1;

    public class DataObject: Dictionary<string, string>
    {
        public DataObject subDataObject = new();
    }
    
    protected virtual byte[] LoadExcelBytes(string filePath)
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

    protected virtual bool IsSheetAvailable(Sheet sheet)
    {
        return sheet.State == null || sheet.State == SheetStateValues.Visible;
    }
    
    public void ReadExcelFile(string filePath)
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
        
        foreach (var sheet in sheets)
        {
            if (sheet.Id == null || !sheet.Id.HasValue || sheet.Id.Value == null || !sheet.Id.HasValue)
                continue;
            
            var sheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            ProcessSheet(sheetPart.Worksheet, workbookPart);
        }
    }

    protected virtual void ProcessSheet(Worksheet sheet, WorkbookPart workbookPart)
    {
        var rows = sheet.Descendants<Row>().Where(r => r.RowIndex is not null);
        
        var headers = new Dictionary<int, string>();    //<columnIndex, headerName>

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable == null)
        {
            throw new InvalidOperationException("Detected excel has no shared string stringTable.");
        }

        foreach (var row in rows)
        {
            int columnIndex = 0;
            
            ProcessRow();

            void ProcessRow()
            {
                foreach (var element in row)
                {
                    if (element is not Cell cell)
                    {
                        throw new InvalidCastException(
                            $"Detected invalid or non cell element ({element.GetType()}) in row");
                    }

                    if (cell.CellReference is { Value: null }) continue;

                    ProcessCell(cell);

                    columnIndex++;
                }
            }
            
            void ProcessCell(Cell cell)
            {
                if(row.RowIndex?.Value == null) return;

                if (row.RowIndex.Value == HEADER_ROW_INDEX)
                {
                    if (!headers.TryAdd(columnIndex, GetCellValue(cell)))
                    {
                        throw new ArgumentException($"duplicated header with column index {columnIndex}");
                    }
                }
                else
                {
                    
                }
            }

            string GetCellValue(Cell cell)
            {
                if (cell.CellValue == null) return string.Empty;
                
                string value = cell.CellValue.InnerText;

                if (cell.DataType != null
                    && cell.DataType.Value == CellValues.SharedString)
                {
                    return sharedStringTable.ChildElements[int.Parse(value)].InnerText;
                }
                else
                {
                    return value;
                }
            }
        }
    }
}