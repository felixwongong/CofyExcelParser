using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CofyDev.Xml.Doc;

public class CofyXmlDocParser
{
    public static void ReadExcelFile(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
        {
            throw new ArgumentNullException(nameof(filePath), "filename is empty");
        }

        if (!File.Exists(filePath))
        {
            throw new ArgumentNullException(nameof(filePath), "file does not exist");
        }
        
        using var filestream = new FileStream(filePath, FileMode.Open);
        using var document = SpreadsheetDocument.Open(filestream, false);
        
        var workbook = document.WorkbookPart;
        if (workbook == null)
        {
            throw new InvalidOperationException("Excel does not have a workbook");
        }

        if (workbook.SharedStringTablePart == null)
        {
            throw new InvalidOperationException("Excel does not have text value");
        }

        var worksheets = workbook.WorksheetParts.ToList();
        if (!worksheets.Any())
        {
            throw new InvalidOperationException("Excel workbook does not have a worksheet");
        }

        var tableStringItems =
            workbook.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>();
        
        OpenXmlReader reader;
        List<string> cellValues = new();
        
        foreach (var worksheet in worksheets)
        {
            reader = OpenXmlReader.Create(worksheet);
            while (reader.Read())
            {
                if (reader.ElementType == typeof(CellValue))
                {
                    var cellStringTableKey = reader.GetText();
                    if(string.IsNullOrEmpty(cellStringTableKey)) continue;
                    Console.Write($"{cellStringTableKey}, ");
                    cellValues.Add(tableStringItems.ElementAt(Int32.Parse(cellStringTableKey)).Text.Text);
                }
            }
        }
        
        foreach (var cellValue in cellValues)
        {
            Console.Write($"{cellValue}, ");
        }
    }
}