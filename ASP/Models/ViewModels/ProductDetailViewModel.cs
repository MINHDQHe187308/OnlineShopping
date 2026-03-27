using ASP.Models.Domains;
using System.Collections.Generic;

namespace ASP.Models.ViewModels
{
    public class ProductDetailViewModel
    {
        public Product Product { get; set; } = null!;

        public string MainImageUrl { get; set; } = "/images/no-image.jpg";

        public List<ProductImage> Images { get; set; } = new();

        public ProductVariant? DefaultVariant { get; set; }

        public decimal CurrentPrice { get; set; }

        public string CurrentColor { get; set; } = "Chưa có màu";
    }
}