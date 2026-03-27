using ASP.Hubs;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.EntityFrameworkCore;

namespace ASP.Controllers.Admin
{
    [Route("admin/productimage")]
    public class ProductImageController : Controller
    {
        private readonly ProductImageRepositoryInterface _repo;
        private readonly IHubContext<AdminHub> _hub;

        public ProductImageController(ProductImageRepositoryInterface repo, IHubContext<AdminHub> hub)
        {
            _repo = repo;
            _hub = hub;
        }

        [HttpGet("")]
        public IActionResult Index(int productId, string? filter, int? categoryIdSearch, int page = 1)
        {
            var images = _repo.GetImagesByProductId(productId);

            ViewBag.ProductId = productId;
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            var product = _repo.GetProduct(productId);
            ViewBag.ProductName = product?.ProductName;

            return View("~/Views/Admin/ProductImage/Index.cshtml", images);
        }

        [HttpGet("create")]
        public IActionResult Create(int productId, string? filter, int? categoryIdSearch, int page = 1)
        {
            ViewBag.ProductId = productId;
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            return View("~/Views/Admin/ProductImage/Create.cshtml");
        }

        [HttpPost("create")]
        public async Task<IActionResult> Create(int productId, string? imageName, IFormFile imageFile, string? filter, int? categoryIdSearch, int page = 1)
        {
            if (imageFile == null || imageFile.Length == 0)
                return Json(new { success = false, message = "Please select image" });

            var extension = Path.GetExtension(imageFile.FileName);

            if (string.IsNullOrWhiteSpace(imageName))
            {
                imageName = Guid.NewGuid().ToString();
            }

            var fileName = imageName + extension;

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/images/products");

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            var filePath = Path.Combine(folderPath, fileName);

            int count = 1;

            while (System.IO.File.Exists(filePath))
            {
                fileName = imageName + "-" + count + extension;
                filePath = Path.Combine(folderPath, fileName);
                count++;
            }

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await imageFile.CopyToAsync(stream);
            }

            bool hasMain = _repo.HasMainImage(productId);

            var img = new ProductImage
            {
                ProductId = productId,
                ImageUrl = "/images/products/" + fileName,
                IsMain = !hasMain
            };

            await _repo.AddImageAsync(img);
            await _hub.Clients.All.SendAsync("ProductImageUpdated", productId);
            await _hub.Clients.All.SendAsync("ProductUpdated", productId);

            return RedirectToAction("Index", new
            {
                productId = productId,
                filter = filter,
                categoryIdSearch = categoryIdSearch,
                page = page
            });
        }


        //[HttpPost("delete/{id}")]
        //public async Task<IActionResult> DeleteAsync(int id)
        //{
        //    var img = _repo.GetImageById(id);

        //    if (img == null)
        //        return Json(new { success = false, message = "Image not found" });

        //    int productId = img.ProductId;

        //    if (img.IsMain)
        //    {
        //        var anotherImage = _repo.GetImagesByProductId(productId)
        //            .FirstOrDefault(x => x.ProductImageId != img.ProductImageId);

        //        if (anotherImage != null)
        //        {
        //            anotherImage.IsMain = true;
        //        }
        //    }

        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", img.ImageUrl.TrimStart('/'));

        //    if (System.IO.File.Exists(filePath))
        //    {
        //        System.IO.File.Delete(filePath);
        //    }

        //    _repo.DeleteImage(img);
        //    await _hub.Clients.All.SendAsync("ProductImageUpdated", productId);
        //    await _hub.Clients.All.SendAsync("ProductUpdated", productId);

        //    bool hasImage = _repo.HasAnyImage(productId);

        //    if (!hasImage)
        //    {
        //        var product = _repo.GetProduct(productId);

        //        if (product != null)
        //        {
        //            product.IsActive = false;
        //            _repo.UpdateProduct(product);
        //        }
        //    }

        //    return Json(new { success = true });
        //}



        [HttpPost("delete/{id}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var img = _repo.GetImageById(id);

            if (img == null)
                return Json(new { success = false, message = "Image not found" });

            int productId = img.ProductId;

            var product = _repo.GetProduct(productId);

            if (product == null)
                return Json(new { success = false, message = "Product not found" });

            var images = _repo.GetImagesByProductId(productId);

            if (product.ProductVariants.Any() && images.Count == 1)
            {
                return Json(new { success = false, message = "Cannot delete the last image of a product that has variants." });
            }

            if (img.IsMain)
            {
                var anotherImage = images.FirstOrDefault(x => x.ProductImageId != img.ProductImageId);
                if (anotherImage != null)
                {
                    anotherImage.IsMain = true;
                }
            }

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", img.ImageUrl.TrimStart('/'));
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }

            _repo.DeleteImage(img);

            await _hub.Clients.All.SendAsync("ProductImageUpdated", productId);
            await _hub.Clients.All.SendAsync("ProductUpdated", productId);


            if (!_repo.HasAnyImage(productId))
            {
                product.IsActive = false;
                _repo.UpdateProduct(product);
            }

            return Json(new { success = true });
        }


        [HttpPost("setmain/{id}")]
        public async Task<IActionResult> SetMainAsync(int id)
        {
            var img = _repo.GetImageById(id);

            if (img == null)
                return Json(new { success = false, message = "Image not found" });

            _repo.SetMainImage(img);
            await _hub.Clients.All.SendAsync("ProductImageUpdated", img.ProductId);
            await _hub.Clients.All.SendAsync("ProductUpdated", img.ProductId);

            return Json(new { success = true });
        }
    }
}
