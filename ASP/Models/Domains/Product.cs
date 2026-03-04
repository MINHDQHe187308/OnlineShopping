using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Domains
{
    public class Product
    {
        [Key]
        public int ProductId { get; set; }

        [Required]
        [StringLength(200)]
        public string ProductName { get; set; }

        [Required]
        public int CategoryId { get; set; }

        public string? Description { get; set; }

        public bool IsActive { get; set; } = true;

        [Range(0, int.MaxValue)]
        public int Quantity { get; set; } = 0;

        [StringLength(20)]
        public string? Size { get; set; }

        [StringLength(50)]
        public string? Color { get; set; }

        // Navigation
        [ForeignKey("CategoryId")]
        public Category Category { get; set; }

        public ICollection<ProductImage> ProductImages { get; set; }
        public ICollection<ProductVariant> ProductVariants { get; set; }
    }
}
