using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Excel;


namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReadExcelController : ControllerBase
    {
    
        private readonly ILogger<ReadExcelController> _logger;

        public ReadExcelController(ILogger<ReadExcelController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "ExcelRead")]
        public string Get()
        {
            string filePath = @"C:\Users\as835\Documents\csharp\monitor\ALC.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = (Worksheet)wb.Worksheets[0];

            return ws.Name;
            
        }
    }
}
