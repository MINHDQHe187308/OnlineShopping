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

        public ProductVariantController(IAuthorizationService authService, ProductVariantRepositoryInterface variantRepo, ProductRepositoryInterface productRepo)
        {
            _authService = authService;
            _variantRepo = variantRepo;
            _productRepo = productRepo;
        }

        [HttpGet]
        [Route("admin/ProductVariant", Name = "admin.productvariants")]
        [Route("admin/ProductVariant/Index")]
        public async Task<IActionResult> Index()
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsView);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var variants = await _variantRepo.GetAllVariantsAsync();
            return View("../Admin/ProductVariants/Index", variants);
        }

        [HttpGet]
        [Route("admin/ProductVariant/Create", Name = "admin.productvariants.create")]
        public async Task<IActionResult> Create()
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPProductVariantsCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var products = _productRepo.GetAllProducts();
            ViewBag.ProductsList = new SelectList(products, "ProductId", "ProductName");

            return View("../Admin/ProductVariants/Add", new ProductVariant { IsActive = true });
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
                TempData["mess-type"] = "success";
                TempData["mess-detail"] = BaseController.BaseMessage("create_success");
                return RedirectToAction(nameof(Index));
            }

            var products = _productRepo.GetAllProducts();
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

            var products = _productRepo.GetAllProducts();
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
                await _variantRepo.UpdateVariantAsync(variant);
                TempData["mess-type"] = "success";
                TempData["mess-detail"] = BaseController.BaseMessage("update_success");
                return RedirectToAction(nameof(Index));
            }

            var products = _productRepo.GetAllProducts();
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

            await _variantRepo.DeleteVariantAsync(id);
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            return RedirectToAction(nameof(Index));
        }
    }
}
