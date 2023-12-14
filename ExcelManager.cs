using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;


public class ExcelManager
{
    private readonly string _filePath;
    public ExcelManager(string filePath)
    {
        _filePath = filePath;
    }

    public string ReadCell(string cellReference)
    {
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false);

        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();

        if (cell == null)
        {
            return null;
        }

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int sharedStringIndex = int.Parse(cell.InnerText);
            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringItem sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
            return sharedStringItem.Text.Text;
        }
        else
        {
            return cell.CellValue.Text;
        }
    }

    public void WriteCell(string cellReference, string data)
    {
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_filePath, true);

        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();
        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int sharedStringIndex = int.Parse(cell.InnerText);
            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
            SharedStringItem sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
            sharedStringItem.Text = new Text(data);
        }
        else
        {
            cell.CellValue = new CellValue(data);
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
        }
    }
}
