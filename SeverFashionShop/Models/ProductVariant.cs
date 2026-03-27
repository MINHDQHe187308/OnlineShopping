using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class ProductVariant
    {
        public ProductVariant()
        {
            CartItems = new HashSet<CartItem>();
            OrderDetails = new HashSet<OrderDetail>();
        }

        public int VariantId { get; set; }
        public int ProductId { get; set; }
        public string Sku { get; set; } = null!;
        public decimal Price { get; set; }
        public bool IsActive { get; set; }
        public int QuantityVariants { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
        public string? Color { get; set; }
        public string? Size { get; set; }

        public virtual Product Product { get; set; } = null!;
        public virtual ICollection<CartItem> CartItems { get; set; }
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }
    }
}
