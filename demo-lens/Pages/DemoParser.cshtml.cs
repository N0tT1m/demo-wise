using Azure.Core;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;

[IgnoreAntiforgeryToken]
[RequestFormLimits(MultipartBodyLengthLimit = 1073741824)]
[RequestSizeLimit(1073741824)]
public class DemoParserModel : PageModel
{
    private readonly ILogger<DemoParserModel> _logger;
    private readonly string _uploadsFolder;

    public DemoParserModel(ILogger<DemoParserModel> logger, IWebHostEnvironment environment)
    {
        _logger = logger;
        _uploadsFolder = Path.Combine(environment.WebRootPath, "uploads");
        Directory.CreateDirectory(_uploadsFolder);
    }

    public async Task<IActionResult> OnPostUpload()
    {
        try
        {
            _logger.LogInformation("Starting file upload");
            var file = Request.Form.Files.GetFile("demoFile");

            if (file == null)
            {
                _logger.LogWarning("No file received");
                return BadRequest("No file received");
            }

            if (!file.FileName.EndsWith(".dem"))
            {
                return BadRequest("Only .dem files are allowed");
            }

            var filePath = Path.Combine(_uploadsFolder, file.FileName);

            // Check if file exists
            if (System.IO.File.Exists(filePath))
            {
                return BadRequest($"A file with the name '{file.FileName}' already exists");
            }

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return new JsonResult(new { fileName = file.FileName, message = "Upload successful" });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error uploading file");
            return StatusCode(500, ex.Message);
        }
    }
}