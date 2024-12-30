namespace demo_lens.Services.FileCleanup;

public class FileCleanupService : IHostedService, IDisposable
{
    private readonly ILogger<FileCleanupService> _logger;
    private readonly string _uploadsPath;
    private readonly string _mapImagesPath;
    private Timer _timer;

    public FileCleanupService(
        ILogger<FileCleanupService> logger,
        IWebHostEnvironment environment)
    {
        _logger = logger;
        _uploadsPath = Path.Combine(environment.WebRootPath, "uploads");
        _mapImagesPath = Path.Combine(environment.WebRootPath, "maps");
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("File Cleanup Service is starting.");

        _timer = new Timer(DoCleanup, null, TimeSpan.Zero, 
            TimeSpan.FromHours(24)); // Run once per day

        return Task.CompletedTask;
    }

    private void DoCleanup(object state)
    {
        _logger.LogInformation("Running file cleanup");

        try
        {
            // Delete files older than 30 days
            var cutoff = DateTime.UtcNow.AddDays(-30);

            // Clean up demo files
            CleanDirectory(_uploadsPath, "*.dem", cutoff);

            // Clean up image files
            CleanDirectory(_mapImagesPath, "*.png", cutoff);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during file cleanup");
        }
    }

    private void CleanDirectory(string path, string pattern, DateTime cutoff)
    {
        if (!Directory.Exists(path)) return;

        foreach (var file in Directory.GetFiles(path, pattern))
        {
            var fileInfo = new FileInfo(file);
            if (fileInfo.CreationTimeUtc < cutoff)
            {
                try
                {
                    File.Delete(file);
                    _logger.LogInformation($"Deleted old file: {file}");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error deleting file: {file}");
                }
            }
        }
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("File Cleanup Service is stopping.");

        _timer?.Change(Timeout.Infinite, 0);

        return Task.CompletedTask;
    }

    public void Dispose()
    {
        _timer?.Dispose();
    }
}