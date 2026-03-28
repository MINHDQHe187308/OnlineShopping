using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using ASP.Policies;
using ASP.Models.Admin.Accounts;
using ASP.Models.Domains;
using ASP.Models.ASPModel;
using ASP.BaseCommon;

namespace ASP.Controllers.Admin
{
    [Authorize]
    public class ProductVariantController : Controller
    {
        private readonly IAuthorizationService _authService;
        private readonly ProductVariantRepositoryInterface _variantRepo;
        private readonly ProductRepositoryInterface _productRepo;

        public ProductVariantController(
            IAuthorizationService authService,
            ProductVariantRepositoryInterface variantRepo,
            ProductRepositoryInterface productRepo)
        {
            _authService = authService;
            _variantRepo = variantRepo;
            _productRepo = productRepo;
        }

        [HttpGet]
        [Route("admin/ProductVariant", Name = "admin.productvariants")]
        [Route("admin/ProductVariant/Index")]
        public async Task<IActionResult> Index(
            string? searchString = "",
            int psize = 10,
            int page = 1,
            string sort = "-VariantId")
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsView);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var list = await _variantRepo.GetAllByFilterAsync(searchString, psize, page, sort);
            return View("../Admin/ProductVariants/Index", list);
        }

        [HttpGet]
        [Route("admin/ProductVariant/Create", Name = "admin.productvariants.create")]
        public async Task<IActionResult> Create()
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var products = _productRepo.GetAllProducts1().OrderBy(p => p.ProductName);
            ViewBag.ProductsList = new SelectList(products, "ProductId", "ProductName");

            return View("../Admin/ProductVariants/Add", new ProductVariant { IsActive = true });
        }

        [HttpGet]
        [Route("admin/ProductVariant/GetProducts")]
        public IActionResult GetProducts()
        {
            var products = _productRepo.GetAllProducts1()
                .OrderBy(p => p.ProductName)
                .Select(p => new { p.ProductId, p.ProductName })
                .ToList();
            return Json(products);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/ProductVariant/Create", Name = "admin.productvariants.store")]
        public async Task<IActionResult> Store(ProductVariant variant)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            ModelState.Remove("Product");
            if (ModelState.IsValid)
            {
                await _variantRepo.CreateVariantAsync(variant);
                _productRepo.RecalculateQuantity(variant.ProductId);

                TempData["mess-type"] = "success";
                TempData["mess-detail"] = BaseController.BaseMessage("create_success");
                return RedirectToAction(nameof(Index));
            }

            var products = _productRepo.GetAllProducts1().OrderBy(p => p.ProductName);
            ViewBag.ProductsList = new SelectList(products, "ProductId", "ProductName", variant.ProductId);
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            return View("../Admin/ProductVariants/Add", variant);
        }

        [HttpGet]
        [Route("admin/ProductVariant/Edit/{id}", Name = "admin.productvariants.edit")]
        public async Task<IActionResult> Edit(int id)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var variant = await _variantRepo.GetVariantByIdAsync(id);
            if (variant == null) return NotFound();

            var products = _productRepo.GetAllProducts1().OrderBy(p => p.ProductName);
            ViewBag.ProductsList = new SelectList(products, "ProductId", "ProductName", variant.ProductId);

            return View("../Admin/ProductVariants/Edit", variant);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/ProductVariant/Edit/{id}", Name = "admin.productvariants.update")]
        public async Task<IActionResult> Update(int id, ProductVariant variant)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            if (id != variant.VariantId) return BadRequest();

            ModelState.Remove("Product");
            if (ModelState.IsValid)
            {
                var oldVariant = await _variantRepo.GetVariantByIdAsync(id);
                var oldProductId = oldVariant?.ProductId ?? variant.ProductId;

                var (success, message) = await _variantRepo.UpdateVariantAsync(variant);

                if (!success)
                {
                    var productList = _productRepo.GetAllProducts1().OrderBy(p => p.ProductName);
                    ViewBag.ProductsList = new SelectList(productList, "ProductId", "ProductName", variant.ProductId);
                    TempData["mess-type"] = "error";
                    TempData["mess-detail"] = message ?? BaseController.BaseMessage("update_fails");
                    return View("../Admin/ProductVariants/Edit", variant);
                }

                if (oldProductId != variant.ProductId)
                {
                    _productRepo.RecalculateQuantity(oldProductId);
                }
                _productRepo.RecalculateQuantity(variant.ProductId);

                TempData["mess-type"] = "success";
                TempData["mess-detail"] = BaseController.BaseMessage("update_success");
                return RedirectToAction(nameof(Index));
            }

            var products = _productRepo.GetAllProducts1().OrderBy(p => p.ProductName);
            ViewBag.ProductsList = new SelectList(products, "ProductId", "ProductName", variant.ProductId);
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            return View("../Admin/ProductVariants/Edit", variant);
        }

        [HttpGet]
        [Route("admin/ProductVariant/Delete/{id}", Name = "admin.productvariants.delete")]
        public async Task<IActionResult> Delete(int id)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsDelete);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var variant = await _variantRepo.GetVariantByIdAsync(id);
            var productId = variant?.ProductId ?? 0;

            await _variantRepo.DeleteVariantAsync(id);

            if (productId > 0)
            {
                _productRepo.RecalculateQuantity(productId);
            }

            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            return RedirectToAction(nameof(Index));
        }
    }
}
