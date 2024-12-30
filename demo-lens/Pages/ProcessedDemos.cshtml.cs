// ProcessedDemosModel.cs
using demo_lens.Data;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;

namespace demo_lens.Pages;

public class ProcessedDemosModel : PageModel
{
    private readonly ApplicationDbContext _dbContext;

    // Maps that have multiple levels
    private readonly HashSet<string> _multiLevelMaps = new()
    {
        "de_vertigo",
        "de_nuke",
        "de_train"
    };

    public ProcessedDemosModel(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }

    public List<ProcessResult> ProcessedDemos { get; set; }

    public async Task OnGetAsync()
    {
        ProcessedDemos = await _dbContext.ProcessResults
            .OrderByDescending(d => d.ProcessedAt)
            .ToListAsync();
    }

    // Helper to check if a map has multiple levels
    public bool HasMultipleLevels(string mapName) => _multiLevelMaps.Contains(mapName.ToLower());
}