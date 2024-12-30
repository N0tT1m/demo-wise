using Azure.Core;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;
using demo_lens.Services.DemoProcessing;
using Microsoft.AspNetCore.Mvc.Rendering;
using demo_lens.Data;
using ProcessResult = demo_lens.Data.ProcessResult;

[IgnoreAntiforgeryToken]
[RequestFormLimits(MultipartBodyLengthLimit = 1073741824)]
[RequestSizeLimit(1073741824)]
public class DemoParserModel : PageModel
{
    private readonly ILogger<DemoParserModel> _logger;
    private readonly string _uploadsFolder;
    private readonly DemoProcessor _demoProcessor;

    public DemoParserModel(
        ILogger<DemoParserModel> logger,
        IWebHostEnvironment environment,
        DemoProcessor demoProcessor)
    {
        _logger = logger;
        _demoProcessor = demoProcessor;
        _uploadsFolder = Path.Combine(environment.WebRootPath, "uploads");
        Directory.CreateDirectory(_uploadsFolder);
    }

    // Maps property remains the same...

    private string GetUniqueFilePath(string originalFilePath)
    {
        // If the file doesn't exist, return the original path
        if (!System.IO.File.Exists(originalFilePath))
        {
            return originalFilePath;
        }

        // Get the directory and file name without extension
        string directory = Path.GetDirectoryName(originalFilePath);
        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(originalFilePath);
        string extension = Path.GetExtension(originalFilePath);

        int counter = 1;
        string newFilePath;

        // Keep incrementing the counter until we find a filename that doesn't exist
        do
        {
            string newFileName = $"{fileNameWithoutExt}({counter}){extension}";
            newFilePath = Path.Combine(directory, newFileName);
            counter++;
        } while (System.IO.File.Exists(newFilePath));

        return newFilePath;
    }

    public IEnumerable<SelectListItem> Maps { get; } = new List<SelectListItem>
    {
        new SelectListItem { Value = "de_ancient", Text = "Ancient" },
        new SelectListItem { Value = "de_anubis", Text = "Anubis" },
        new SelectListItem { Value = "de_dust2", Text = "Dust 2" },
        new SelectListItem { Value = "de_inferno", Text = "Inferno" },
        new SelectListItem { Value = "de_mirage", Text = "Mirage" },
        new SelectListItem { Value = "de_nuke", Text = "Nuke" },
        new SelectListItem { Value = "de_overpass", Text = "Overpass" },
        new SelectListItem { Value = "de_train", Text = "Train" },
        new SelectListItem { Value = "de_vertigo", Text = "Vertigo" },
    };

    public async Task<IActionResult> OnPostUpload()
    {
        try
        {
            _logger.LogInformation("Starting file upload");
            var file = Request.Form.Files.GetFile("demoFile");
            var mapName = Request.Form["mapName"].ToString();

            if (file == null)
            {
                _logger.LogWarning("No file received");
                return BadRequest("No file received");
            }

            if (string.IsNullOrEmpty(mapName))
            {
                _logger.LogWarning("No map selected");
                return BadRequest("Please select a map");
            }

            if (!file.FileName.EndsWith(".dem"))
            {
                return BadRequest("Only .dem files are allowed");
            }

            // Create the initial file path
            var initialFilePath = Path.Combine(_uploadsFolder, file.FileName);

            // Get a unique file path that won't conflict with existing files
            var uniqueFilePath = GetUniqueFilePath(initialFilePath);

            // Save the file with the unique name
            using (var stream = new FileStream(uniqueFilePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            // Process the demo file with the selected map
            var result = await _demoProcessor.ProcessDemoFileAsync(uniqueFilePath, mapName);

            // Return detailed response with redirect URL
            return new JsonResult(new
            {
                success = result.Success,
                demoId = result.Id,
                fileName = Path.GetFileName(uniqueFilePath),
                message = result.Success
                    ? "Processing completed successfully"
                    : $"Processing failed: {result.Errors}",
                imagePath = result.ImagePath,
                output = result.Output,
                redirectUrl = "/ProcessedDemos" // Add this line
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing demo file");
            return StatusCode(500, ex.Message);
        }
    }
}