using demo_lens.Services.DemoProcessing;
using Microsoft.EntityFrameworkCore;

namespace demo_lens.Data;

public class DemoDbContext : DbContext
{
    public DemoDbContext(DbContextOptions<DemoDbContext> options) : base(options)
    {
    }

    public DbSet<ProcessResult> ProcessedDemos { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<ProcessResult>()
            .Property(p => p.ProcessedAt)
            .HasDefaultValueSql("CURRENT_TIMESTAMP");
    }
}