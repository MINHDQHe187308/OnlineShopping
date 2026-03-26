using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ASP.Policies;
using ASP.Models.Admin.Accounts;
using ASP.Models.Domains;
using ASP.Models.ASPModel;
using ASP.BaseCommon;

namespace ASP.Controllers.Admin
{
    [Authorize]
    public class CategoryController : Controller
    {
        private readonly IAuthorizationService _authService;
        private readonly CategoryRepositoryInterface _categoryRepo;

        public CategoryController(IAuthorizationService authService, CategoryRepositoryInterface categoryRepo)
        {
            _authService = authService;
            _categoryRepo = categoryRepo;
        }

        [HttpGet]
        [Route("admin/Category", Name = "admin.categories")]
        [Route("admin/Category/Index")]
        public async Task<IActionResult> Index(
            string? searchString = "",
            int psize = 10,
            int page = 1,
            string sort = "-CategoryId")
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesView);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var list = await _categoryRepo.GetAllByFilterAsync(searchString, psize, page, sort);
            return View("../Admin/Categories/Index", list);
        }

        [HttpGet]
        [Route("admin/Category/Create", Name = "admin.categories.create")]
        public async Task<IActionResult> Create()
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            return View("../Admin/Categories/Add", new Category());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/Category/Create", Name = "admin.categories.store")]
        public async Task<IActionResult> Store(Category category)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            ModelState.Remove(nameof(category.Products));
            if (ModelState.IsValid)
            {
                var result = await _categoryRepo.CreateCategoryAsync(category);
                if (result)
                {
                    TempData["mess-type"] = "success";
                    TempData["mess-detail"] = BaseController.BaseMessage("create_success");
                    return RedirectToAction(nameof(Index));
                }
                ModelState.AddModelError("", "Could not save to database.");
            }

            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            return View("../Admin/Categories/Add", category);
        }

        [HttpGet]
        [Route("admin/Category/Edit/{id}", Name = "admin.categories.edit")]
        public async Task<IActionResult> Edit(int id)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var category = await _categoryRepo.GetCategoryByIdAsync(id);
            if (category == null) return NotFound();

            return View("../Admin/Categories/Edit", category);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/Category/Edit/{id}", Name = "admin.categories.update")]
        public async Task<IActionResult> Update(int id, Category category)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();

            if (id != category.CategoryId) return BadRequest();

            ModelState.Remove(nameof(category.Products));
            if (ModelState.IsValid)
            {
                var result = await _categoryRepo.UpdateCategoryAsync(category);
                if (result)
                {
                    TempData["mess-type"] = "success";
                    TempData["mess-detail"] = BaseController.BaseMessage("update_success");
                    return RedirectToAction(nameof(Index));
                }
                ModelState.AddModelError("", "Could not update in database.");
            }

            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            return View("../Admin/Categories/Edit", category);
        }

        [HttpGet]
        [Route("admin/Category/Delete/{id}", Name = "admin.categories.delete")]
        public async Task<IActionResult> Delete(int id)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPCategoriesDelete);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var success = await _categoryRepo.DeleteCategoryAsync(id);
            if (success)
            {
                TempData["mess-type"] = "success";
                TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            }
            else
            {
                TempData["mess-type"] = "error";
                TempData["mess-detail"] = "Failed to delete from database. It is possible that this category is referenced elsewhere.";
            }
            
            return RedirectToAction(nameof(Index));
        }
    }
}
