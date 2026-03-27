using System.ComponentModel.DataAnnotations;

namespace ASP.Models.ViewModels
{
    public class ProductImageViewModel
    {
        [Required(ErrorMessage = "Please select image")]
        public IFormFile ImageFile { get; set; }

        [StringLength(100, ErrorMessage = "Name too long")]
        public string? ImageName { get; set; }

        public int ProductId { get; set; }
    }
}
