using System.Diagnostics;
using demo_lens.Data;
using DemoFile;
using DemoFile.Sdk;
using Microsoft.Data.SqlClient;
using Microsoft.VisualBasic;

namespace demo_lens.Services.DemoProcessing;

public class DemoProcessor
{
    private readonly ILogger<DemoProcessor> _logger;
    private readonly string _demoParserPath;
    private readonly string _mapImagesPath;
    private readonly ApplicationDbContext _dbContext;  // Add this
    private readonly string _dashboardBaseUrl;

    public DemoProcessor(
        ILogger<DemoProcessor> logger,
        IWebHostEnvironment environment,
        IConfiguration configuration,
        ApplicationDbContext dbContext)  // Add this parameter
    {
        _logger = logger;
        _demoParserPath = Path.Combine(environment.WebRootPath, "bin", "DemoParser.exe");
        _mapImagesPath = Path.Combine(environment.WebRootPath, "maps");
        _dbContext = dbContext;  // Initialize the context
        _dashboardBaseUrl = configuration["DashboardBaseUrl"];

        // Ensure directories exist
        Directory.CreateDirectory(Path.Combine(environment.WebRootPath, "bin"));
        Directory.CreateDirectory(_mapImagesPath);
    }

    public async Task<ProcessResult> ProcessDemoFileAsync(string demoFilePath, string mapName)
    {
        try
        {
            _logger.LogInformation($"Starting to process demo file: {Path.GetFileName(demoFilePath)}");

            // Create the ProcessResult object early
            var result = new ProcessResult
            {
                DemoFileName = Path.GetFileName(demoFilePath),
                MapName = mapName,
                ProcessedAt = DateTime.UtcNow
            };

            // Start the demo parser process
            var processStartInfo = new ProcessStartInfo
            {
                FileName = _demoParserPath,
                Arguments = $"\"{demoFilePath}\" \"{mapName}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = new Process { StartInfo = processStartInfo };
            var output = new List<string>();
            var errors = new List<string>();

            process.OutputDataReceived += (sender, e) =>
            {
                if (e.Data != null)
                {
                    output.Add(e.Data);
                    _logger.LogInformation($"Parser output: {e.Data}");
                }
            };

            process.ErrorDataReceived += (sender, e) =>
            {
                if (e.Data != null)
                {
                    errors.Add(e.Data);
                    _logger.LogError($"Parser error: {e.Data}");
                }
            };

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            await process.WaitForExitAsync();

            // Update the result object with process outcome
            result.Success = process.ExitCode == 0;
            result.ExitCode = process.ExitCode;
            result.Output = string.Join(Environment.NewLine, output);
            result.Errors = string.Join(Environment.NewLine, errors);

            // Set paths for generated content
            string imageFileName = Path.GetFileNameWithoutExtension(demoFilePath) + ".png";
            result.ImagePath = $"/maps/{imageFileName}";
            result.DashboardUrl = $"{_dashboardBaseUrl}/demo/{Path.GetFileNameWithoutExtension(demoFilePath)}";

            // Save to database and get the ID
            _dbContext.ProcessResults.Add(result);
            await _dbContext.SaveChangesAsync();

            _logger.LogInformation($"Demo processing completed with ID: {result.Id}");
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error processing demo file");
            throw;
        }
    }
}