using ASP.Models.ASPModel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ASP.Controllers.Admin
{
    [Route("admin/productimage")]
    public class ProductImageController : Controller
    {
        private readonly ASPDbContext _context;
        public ProductImageController(ASPDbContext context) 
        {
            _context = context;
        }

        [HttpGet("")]
        public IActionResult Index(int productId, string? filter, int? categoryIdSearch, int page = 1)
        {
            var images = _context.ProductImages
                .Where(x => x.ProductId == productId)
                .ToList();

            ViewBag.ProductId = productId;
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            var product = _context.Products.Find(productId);
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

            var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/images/productImage");

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

            bool hasMain = _context.ProductImages
                .Any(x => x.ProductId == productId && x.IsMain);

            var img = new ASP.Models.Domains.ProductImage
            {
                ProductId = productId,
                ImageUrl = "/images/productImage/" + fileName,
                IsMain = !hasMain
            };

            _context.ProductImages.Add(img);
            await _context.SaveChangesAsync();

            return RedirectToAction("Index", new
            {
                productId = productId,
                filter = filter,
                categoryIdSearch = categoryIdSearch,
                page = page
            });
        }


        [HttpPost("delete/{id}")]
        public IActionResult Delete(int id)
        {
            var img = _context.ProductImages.Find(id);

            if (img == null)
                return Json(new { success = false, message = "Image not found" });
            int productId = img.ProductId;

            if (img.IsMain)
            {
                var anotherImage = _context.ProductImages
                    .Where(x => x.ProductId == productId && x.ProductImageId != img.ProductImageId)
                    .FirstOrDefault();

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

            _context.ProductImages.Remove(img);
            _context.SaveChanges();

            bool hasImage = _context.ProductImages.Any(x => x.ProductId == productId);
            if (!hasImage)
            {
                var product = _context.Products.FirstOrDefault(x => x.ProductId == productId);
                if (product != null)
                {
                    product.IsActive = false;
                    _context.SaveChanges();
                }
            }


            return Json(new { success = true });
        }

        [HttpPost("setmain/{id}")]
        public IActionResult SetMain(int id)
        {
            var img = _context.ProductImages.Find(id);

            if (img == null)
                return Json(new { success = false, message = "Image not found" });

            var images = _context.ProductImages
                .Where(x => x.ProductId == img.ProductId)
                .ToList();

            foreach (var item in images)
                item.IsMain = false;

            img.IsMain = true;

            _context.SaveChanges();

            return Json(new { success = true });
        }
    }
}
