using ASP.Models.ASPModel;
using ASP.Models.Domains;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ReflectionIT.Mvc.Paging;

namespace ASP.Controllers.Admin
{
    [Route("admin/product")]
    public class ProductController : Controller
    {
        private readonly ASPDbContext _context;
        public ProductController(ASPDbContext context)
        {
            _context = context;
        }

        [Route("")]
        public IActionResult Index(string? filter, int? categoryIdSearch, int page = 1)
        {
            var query = _context.Products
                .Include(x => x.Category)
                .Include(x => x.ProductImages)
                .AsQueryable();

            if (!string.IsNullOrEmpty(filter))
            {
                query = query.Where(x => x.ProductName.Contains(filter));
            }

            if (categoryIdSearch != null && categoryIdSearch > 0)
            {
                query = query.Where(x => x.CategoryId == categoryIdSearch);
            }

            int pageSize = 10;

            int totalItems = query.Count();
            int totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

            if (page > totalPages && totalPages > 0)
            {
                page = totalPages;
            }

            var model = PagingList.Create(query, pageSize, page);

            model.RouteValue = new RouteValueDictionary
            {
                { "filter", filter },
                { "categoryIdSearch", categoryIdSearch }
            };

            ViewBag.Categories = _context.Categories.ToList();
            ViewBag.categoryIdSearch = categoryIdSearch;

            return View("~/Views/Admin/Product/Index.cshtml", model);
        }


        [HttpGet]
        [Route("create")]
        public IActionResult Create(string? filter, int? categoryIdSearch, int page = 1)
        {
            ViewBag.Categories = _context.Categories.ToList();
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;


            return View("~/Views/Admin/Product/Create.cshtml");
        }

        [HttpPost]
        [Route("create")]
        public IActionResult Create(Product product, string? filter, int? categoryIdSearch, int page = 1)
        {
            // bỏ validate navigation
            ModelState.Remove("Category");
            ModelState.Remove("ProductImages");
            ModelState.Remove("ProductVariants");

            if (ModelState.IsValid)
            {
                _context.Products.Add(product);
                _context.SaveChanges();

                return RedirectToAction("Index", new
                {
                    filter = filter,
                    categoryIdSearch = categoryIdSearch,
                    page = page
                });
            }

            ViewBag.Categories = _context.Categories.ToList();
            ViewBag.Filter = filter;
            ViewBag.Page = page;
            ViewBag.categoryIdSearch = categoryIdSearch;
            return View("~/Views/Admin/Product/Create.cshtml", product);
        }

        [HttpGet]
        [Route("edit/{id}")]
        public IActionResult Edit(int id, string? filter, int? categoryIdSearch, int page = 1)
        {
            var product = _context.Products
                .FirstOrDefault(x => x.ProductId == id);

            if (product == null)
            {
                return NotFound();
            }

            ViewBag.Categories = _context.Categories.ToList();
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            return View("~/Views/Admin/Product/Edit.cshtml", product);
        }

        [HttpPost]
        [Route("edit/{id}")]
        public IActionResult Edit(int id, Product model, string? filter, int? categoryIdSearch, int page = 1)
        {
            //System.Diagnostics.Debug.WriteLine("POST RUNNING");

            ModelState.Remove("Category");
            ModelState.Remove("ProductImages");
            ModelState.Remove("ProductVariants");

            var product = _context.Products.FirstOrDefault(x => x.ProductId == id);

            if (product == null)
                return NotFound();

            bool hasImage = _context.ProductImages.Any(x => x.ProductId == id);

            if (model.IsActive && !hasImage)
            {
                ModelState.AddModelError("IsActive", "Sản phẩm phải có ít nhất 1 ảnh trước khi kích hoạt.");
            }

            if (!ModelState.IsValid)
            {
                ViewBag.Categories = _context.Categories.ToList();
                ViewBag.Filter = filter;
                ViewBag.categoryIdSearch = categoryIdSearch;
                ViewBag.Page = page;
                return View("~/Views/Admin/Product/Edit.cshtml", model);
            }

            product.ProductName = model.ProductName;
            product.CategoryId = model.CategoryId;
            product.Description = model.Description;
            product.Quantity = model.Quantity;
            product.IsActive = model.IsActive;

            _context.SaveChanges();

            return RedirectToAction("Index", new
            {
                filter = filter,
                categoryIdSearch = categoryIdSearch,
                page = page
            });
        }


        [HttpPost]
        [Route("delete/{id}")]
        public IActionResult Delete(int id)
        {
            var product = _context.Products
                .Include(x => x.ProductImages)
                .Include(x => x.ProductVariants)
                .FirstOrDefault(x => x.ProductId == id);

            if (product == null)
            {
                return Json(new { success = false, message = "Product not found" });
            }

            // kiểm tra nếu còn ảnh
            if (product.ProductImages != null && product.ProductImages.Any())
            {
                return Json(new
                {
                    success = false,
                    message = "Không thể xóa vì sản phẩm vẫn còn ảnh"
                });
            }

            // kiểm tra nếu còn variant
            if (product.ProductVariants != null && product.ProductVariants.Any())
            {
                return Json(new
                {
                    success = false,
                    message = "Không thể xóa vì sản phẩm vẫn còn biến thể"
                });
            }

            try
            {
                _context.Products.Remove(product);
                _context.SaveChanges();

                return Json(new { success = true });
            }
            catch
            {
                return Json(new
                {
                    success = false,
                    message = "Delete failed"
                });
            }
        }
    }
}
