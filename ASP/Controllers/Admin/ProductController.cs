using ASP.Hubs;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.EntityFrameworkCore;
using ReflectionIT.Mvc.Paging;

namespace ASP.Controllers.Admin
{
    [Route("admin/product")]
    public class ProductController : Controller
    {
        private readonly ProductRepositoryInterface _repo;
        private readonly IHubContext<AdminHub> _hub;

        public ProductController(ProductRepositoryInterface repo, IHubContext<AdminHub> hub)
        {
            _repo = repo;
            _hub = hub;
        }

        [Route("")]
        public IActionResult Index(string? filter, int? categoryIdSearch, int page = 1)
        {
            var query = _repo.GetProducts(filter, categoryIdSearch);

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

            ViewBag.Categories = _repo.GetCategories();
            ViewBag.categoryIdSearch = categoryIdSearch;

            return View("~/Views/Admin/Product/Index.cshtml", model);
        }


        [HttpGet]
        [Route("create")]
        public IActionResult Create(string? filter, int? categoryIdSearch, int page = 1)
        {
            ViewBag.Categories = _repo.GetCategories();
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            return View("~/Views/Admin/Product/Create.cshtml");
        }

        [HttpGet]
        [Route("get-categories")]
        public IActionResult GetCategories()
        {
            var categories = _repo.GetCategories()
                .Select(c => new
                {
                    c.CategoryId,
                    c.CategoryName
                });

            return Json(categories);
        }


        [HttpPost]
        [Route("create")]
        public async Task<IActionResult> Create(Product product, string? filter, int? categoryIdSearch, int page = 1)
        {
            ModelState.Remove("Category");
            ModelState.Remove("ProductImages");
            ModelState.Remove("ProductVariants");
            ModelState.Remove("Description");

            if (ModelState.IsValid)
            {
                _repo.Add(product);

                await _hub.Clients.All.SendAsync("ProductUpdated", product.ProductId);
                TempData["SuccessMessage"] = "Thêm sản phẩm thành công";

                return RedirectToAction("Index", new
                {
                    filter = filter,
                    categoryIdSearch = categoryIdSearch,
                    page = page
                });
            }

            ViewBag.Categories = _repo.GetCategories();
            ViewBag.Filter = filter;
            ViewBag.Page = page;
            ViewBag.categoryIdSearch = categoryIdSearch;

            return View("~/Views/Admin/Product/Create.cshtml", product);
        }

        [HttpGet]
        [Route("edit/{id}")]
        public IActionResult Edit(int id, string? filter, int? categoryIdSearch, int page = 1)
        {
            var product = _repo.GetById(id);

            if (product == null)
            {
                return NotFound();
            }

            ViewBag.Categories = _repo.GetCategories();
            ViewBag.Filter = filter;
            ViewBag.categoryIdSearch = categoryIdSearch;
            ViewBag.Page = page;

            return View("~/Views/Admin/Product/Edit.cshtml", product);
        }

        [HttpPost]
        [Route("edit/{id}")]
        public async Task<IActionResult> Edit(int id, Product model, string? filter, int? categoryIdSearch, int page = 1)
        {
            ModelState.Remove("Category");
            ModelState.Remove("ProductImages");
            ModelState.Remove("ProductVariants");
            ModelState.Remove("Description");

            var product = _repo.GetById(id);

            if (product == null)
                return NotFound();

            bool hasImage = _repo.HasImage(id);

            if (model.IsActive && !hasImage)
            {
                ModelState.AddModelError("IsActive", "Sản phẩm phải có ít nhất 1 ảnh trước khi kích hoạt.");
            }

            if (!ModelState.IsValid)
            {
                ViewBag.Categories = _repo.GetCategories();
                ViewBag.Filter = filter;
                ViewBag.categoryIdSearch = categoryIdSearch;
                ViewBag.Page = page;
                return View("~/Views/Admin/Product/Edit.cshtml", model);
            }

            product.ProductName = model.ProductName;
            product.CategoryId = model.CategoryId;
            product.Description = model.Description;
            //product.Quantity = model.Quantity;
            product.IsActive = model.IsActive;

            _repo.Update(product);
            await _hub.Clients.All.SendAsync("ProductUpdated", product.ProductId);
            TempData["SuccessMessage"] = "Cập nhật sản phẩm thành công";

            return RedirectToAction("Index", new
            {
                filter = filter,
                categoryIdSearch = categoryIdSearch,
                page = page
            });
        }

        [HttpPost]
        [Route("delete/{id}")]
        public async Task<IActionResult> DeleteAsync(int id)
        {
            var product = _repo.GetById(id);

            if (product == null)
            {
                return Json(new { success = false, message = "Product not found" });
            }

            if (product.ProductImages != null && product.ProductImages.Any())
            {
                return Json(new
                {
                    success = false,
                    message = "Không thể xóa vì sản phẩm vẫn còn ảnh"
                });
            }

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
                _repo.Delete(product);
                await _hub.Clients.All.SendAsync("ProductUpdated", product.ProductId);

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
