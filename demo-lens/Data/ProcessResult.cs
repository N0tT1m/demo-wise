using System;
using System.ComponentModel.DataAnnotations;

namespace demo_lens.Data
{
    public class ProcessResult
    {
        // Primary key for database tracking
        [Key]
        public int Id { get; set; }
        
        // Demo processing status
        public bool Success { get; set; }
        public int ExitCode { get; set; }
        
        // Processing details and output
        public string Output { get; set; }
        public string Errors { get; set; }
        
        // Demo and map information
        public string MapName { get; set; }
        public string DemoFileName { get; set; }
        
        // Generated content paths
        public string ImagePath { get; set; }
        public string DashboardUrl { get; set; }
        
        // Tracking information
        public DateTime ProcessedAt { get; set; } = DateTime.UtcNow;
    }
}