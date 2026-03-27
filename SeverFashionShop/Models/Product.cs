using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class Product
    {
        public Product()
        {
            ProductImages = new HashSet<ProductImage>();
            ProductVariants = new HashSet<ProductVariant>();
        }

        public int ProductId { get; set; }
        public string ProductName { get; set; } = null!;
        public int CategoryId { get; set; }
        public string? Description { get; set; }
        public bool IsActive { get; set; }
        public int Quantity { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }

        public virtual Category Category { get; set; } = null!;
        public virtual ICollection<ProductImage> ProductImages { get; set; }
        public virtual ICollection<ProductVariant> ProductVariants { get; set; }
    }
}
