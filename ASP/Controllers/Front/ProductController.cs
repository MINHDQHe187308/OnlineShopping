using Microsoft.AspNetCore.Mvc;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore; 

namespace ASP.Controllers.Front
{
    public class ProductController : Controller
    {
        private readonly ProductRepositoryInterface _productRepository;

        public ProductController(ProductRepositoryInterface productRepository)
        {
            _productRepository = productRepository;
        }

        public IActionResult Index(int? category = null)
        {
            var products = _productRepository.GetAllProducts();
            if (category.HasValue)
            {
                products = products.Where(p => p.CategoryId == category.Value).ToList();
            }

            ViewBag.SelectedCategory = category;
            return View("~/Views/Front/Products/Index.cshtml", products);
        }

        public async Task<IActionResult> Details(int id)
        {
            var product = await _productRepository.GetProductByIdAsync(id);
            if (product == null)
            {
                return NotFound();
            }

            var viewModel = new ProductDetailViewModel
            {
                Product = product,
                Images = product.ProductImages?.ToList() ?? new List<ProductImage>(),
                DefaultVariant = product.ProductVariants?.OrderBy(v => v.VariantId).FirstOrDefault(),
            };

            viewModel.MainImageUrl = product.ProductImages
                ?.FirstOrDefault(x => x.IsMain)?.ImageUrl
                ?? product.ProductImages?.FirstOrDefault()?.ImageUrl
                ?? "/images/no-image.jpg";

            viewModel.CurrentPrice = viewModel.DefaultVariant?.Price ?? 0;

           

            return View("~/Views/Front/Products/ProductDetail.cshtml",viewModel);
        }
    }
}