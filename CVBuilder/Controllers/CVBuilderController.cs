using CVBuilder.Dtos;
using CVBuilder.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace CVBuilder.Controllers
{
    [ApiController]
    [Route("[controller]")]
    [AllowAnonymous]
    public class CVBuilderController : ControllerBase
    {
        private readonly ICvBuilderService _cvBuilderService;

        public CVBuilderController(ICvBuilderService cvBuilderService)
        {
            _cvBuilderService = cvBuilderService;
        }

        [HttpPost("generate")]
        public async Task<IActionResult> GenerateCV(
            [FromBody] CVFieldValuesDto request)
        {
            var fileBytes = await _cvBuilderService.GenerateCVAsync(request);

            var fileName =
                $"CV_{request.FirstName}_{request.LastName}.docx";

            return File(
                fileBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                fileName);
        }
    }
}