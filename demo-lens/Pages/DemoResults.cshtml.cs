using demo_wise.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace demo_lens.Pages;

// Pages/DemoResults.cshtml.cs
public class DemoResultsModel : PageModel
{
    private readonly DemoDbContext _dbContext;
    private readonly ILogger<DemoResultsModel> _logger;

    public DemoResultsModel(DemoDbContext dbContext, ILogger<DemoResultsModel> logger)
    {
        _dbContext = dbContext;
        _logger = logger;
    }

    public ProcessResult DemoResult { get; set; }

    public async Task<IActionResult> OnGetAsync(int id)
    {
        DemoResult = await _dbContext.ProcessedDemos.FindAsync(id);
        
        if (DemoResult == null)
        {
            _logger.LogWarning($"Demo result with ID {id} not found");
            return NotFound();
        }

        return Page();
    }
}