using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Domains
{
    public class CartItem
    {
        [Key]
        public int CartItemId { get; set; }

        public int CartId { get; set; }

        public int VariantId { get; set; }

        [Range(1, int.MaxValue)]
        public int Quantity { get; set; }

        // Navigation
        [ForeignKey("CartId")]
        public Cart Cart { get; set; }

        [ForeignKey("VariantId")]
        public ProductVariant ProductVariant { get; set; }
    }
}