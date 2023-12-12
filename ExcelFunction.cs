using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;

namespace ExcelFileHelper
{
    public static class ExcelFunction
    {
        [FunctionName("ReadExcelCell")]
        public static IActionResult ReadExcelCell(
       [HttpTrigger(AuthorizationLevel.Function, "get", Route = "excel/read/{cellReference}")] HttpRequest req,
       string cellReference,
       ILogger log)
        {
            //Need to change the excel file path
            string filePath = "path/to/your/excel/file.xlsx";
            ExcelManager excelManager = new ExcelManager(filePath);
            string cellValue = excelManager.ReadCell(cellReference);
            return new OkObjectResult(cellValue);
        }

        [FunctionName("WriteExcelCell")]
        public static IActionResult WriteExcelCell(
       [HttpTrigger(AuthorizationLevel.Function, "post", Route = "excel/write/{cellReference}/{data}")] HttpRequest req,
       string cellReference,
       string data,
       ILogger log)
        {
            //Need to change the excel file path
            string filePath = "path/to/your/excel/file.xlsx";
            ExcelManager excelManager = new ExcelManager(filePath);
            excelManager.WriteCell(cellReference, data);
            return new OkResult();
        }
    }
}
