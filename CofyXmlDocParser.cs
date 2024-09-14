using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc;

public class CofyXmlDocParser
{
    public static void ReadExcelFile(string filePath)
    {
        byte[] loadFileBytes()
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

        byte[] fileBytes = loadFileBytes();

        using MemoryStream memoryStream = new MemoryStream(fileBytes);
        using var document = SpreadsheetDocument.Open(memoryStream, false);

        var workbookPart = document.WorkbookPart;
        if (workbookPart == null)
        {
            throw new InvalidOperationException("Excel does not have a workbook");
        }

        var sheets = workbookPart.Workbook.Descendants<Sheet>().Where(sheet => sheet.State == null || sheet.State == SheetStateValues.Visible );
        
        foreach (var sheet in sheets)
        {
            if (sheet.Id == null || !sheet.Id.HasValue || sheet.Id.Value == null || !sheet.Id.HasValue) continue;
            
            var sheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
            
        }
    }
}