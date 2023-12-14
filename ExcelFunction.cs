using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

public static class ExcelFunctions
{
    [FunctionName("ReadExcelCellValue")]
    public static async Task<IActionResult> ReadExcelCellValueAsync(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = "excel/read-cell")] HttpRequest req,
        ILogger log)
    {
        try
        {
            var formCollection = await req.ReadFormAsync();
            var file = formCollection.Files["excelFile"];

            if (file == null || file.Length == 0)
            {
                return new BadRequestResult();
            }

            // Save the uploaded file to a temporary location
            var filePath = Path.GetTempFileName();
            using (var fileStream = File.Create(filePath))
            {
                await file.CopyToAsync(fileStream);
                fileStream.Flush();
            }

            // Read value from the specified cell in the uploaded Excel file
            var value = ReadCellValue(filePath, "H15");

            // Clean up: Delete the temporary file
            File.Delete(filePath);

            return new OkObjectResult(value);
        }
        catch (Exception ex)
        {
            log.LogError($"Error reading cell value: {ex.Message}");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError);
        }
    }

    private static string ReadCellValue(string filePath, string cellReference)
    {
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Last();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            Cell cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

            if (cell != null)
            {
                return GetCellValue(cell, workbookPart);
            }
            else
            {
                return "Cell not found";
            }
        }
    }

    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
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
}
