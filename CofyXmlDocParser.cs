using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace CofyDev.Xml.Doc;

public class CofyXmlDocParser
{
    static void ReadExcelFile(string filename)
    {
        if (string.IsNullOrEmpty(filename))
        {
            throw new ArgumentNullException(nameof(filename), "filename is empty");
        }

        using var document = SpreadsheetDocument.Open(filename, false);
        
        var workbook = document.WorkbookPart;
        if (workbook == null)
        {
            throw new InvalidOperationException("Excel does not have a workbook");
        }

        var worksheets = workbook.WorksheetParts.ToList();
        if (!worksheets.Any())
        {
            throw new InvalidOperationException("Excel workbook does not have a worksheet");
        }
        
        foreach (var worksheet in worksheets)
        {
            var reader = OpenXmlReader.Create(worksheet);
        }
    }
}