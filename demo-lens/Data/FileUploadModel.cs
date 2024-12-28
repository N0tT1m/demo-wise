using System.ComponentModel.DataAnnotations;

public class FileUploadModel
{
    [Required]
    [Display(Name = "Demo file")]
    public IFormFile? FormFile { get; set; }
}