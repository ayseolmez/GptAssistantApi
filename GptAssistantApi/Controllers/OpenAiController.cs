using GptAssistant.Service.AiService;
using Microsoft.AspNetCore.Mvc;

namespace GptAssistant.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AiController : ControllerBase
    {
        private readonly IAiService _aiService;

        public AiController(IAiService aiService)
        {
            _aiService = aiService;
        }

        [HttpPost("processexcel")]
        public async Task<IActionResult> ProcessExcel()
        {
            var excelFilePath = "YOUR_EXCEL_PATH";
            if (string.IsNullOrEmpty(excelFilePath))
            {
                return BadRequest("Invalid request.");
            }

            await _aiService.ProcessExcelAndAskGpt(excelFilePath);
            return Ok("Excel file processed and answers written.");
        }
    }
}
