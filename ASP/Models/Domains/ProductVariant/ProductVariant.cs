using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Domains.ProductVariant
{
    public class ProductVariant
    {
        [Key]
        public int VariantId { get; set; }

        [Required]
        public int ProductId { get; set; }

        [Required]
        [StringLength(50)]
        public string SKU { get; set; }

        [Column(TypeName = "decimal(18,2)")]
        [Range(0, double.MaxValue)]
        public decimal Price { get; set; }

        public bool IsActive { get; set; } = true;

        [Range(0, int.MaxValue)]
        public int QuantityVariants { get; set; } = 0;

        [ForeignKey("ProductId")]
        public Product Product { get; set; }
    }
}
