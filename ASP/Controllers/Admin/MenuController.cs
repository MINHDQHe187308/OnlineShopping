using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ASP.BaseCommon;
using ASP.Models.Admin;
using ASP.Models.Admin.Menus;
using ASP.Policies;
using ASP.Models;
using ASP.Models.ASPModel;
using AdminMenu = ASP.Models.Admin.Menus.Menu;
namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.menus")]
    public class MenuController : Controller, MenuListenerInterface
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        public MenuRepositoryInterface menu;
        protected string photosPath;
        public BaseController _baseController;
        public MenuController(IAuthorizationService authService, ASPDbContext context, MenuRepositoryInterface menu, BaseController baseController)
        {
            _authService = authService;
            _context = context;
            this.menu = menu;
            _baseController = baseController;
        }
        public async Task<IActionResult> Index(string filter = "", int pagesize = 10, int page = 1, string sort = "-ID")
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPMenusView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            var list = this.menu.GetAllByLimit(filter, pagesize, page, sort).Result;
            return View("../Admin/Menus/Index", list);
        }
        //
        [HttpGet]
        [Route("Create", Name = "admin.menus.create")]
        public async Task<ActionResult> Create(AdminMenu menu)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPMenusCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            ViewBag.MenuStatics = _baseController.GetMenuStatics();
            return View("../Admin/Menus/Add", menu);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Create", Name = "admin.menus.store")]
        public async Task<IActionResult> Store(AdminMenu menu)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPMenusCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            if (!ModelState.IsValid)
            {
                ViewBag.MenuStatics = _baseController.GetMenuStatics();
                return await this.CreateMenuFails(menu);
            }
            return await this.menu.CreateMenu(this, menu);
        }
        //
        [Route("Edit/{id?}", Name = "admin.menus.show")]
        public async Task<IActionResult> Show(int id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPMenusUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            ViewBag.MenuStatics = _baseController.GetMenuStatics();
            var menu = this.menu.GetMenuById(id);
            if(menu == null) return new ForbidResult();
            return View("../Admin/Menus/Edit", menu );
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Edit", Name = "admin.menus.edit")]
        public async Task<IActionResult> Edit(int id, AdminMenu menu)
        {
            ModelState.Remove("Language");
            if (!ModelState.IsValid)
            {
                ViewBag.MenuStatics = _baseController.GetMenuStatics();
                return await this.UpdateMenuFails(menu);
            }
            return await this.menu.UpdateMenuById(id, this, menu);
        }
        [HttpGet]
        [Route("Delete/{id?}", Name = "admin.menus.delete")]
        public async Task<IActionResult> Delete(int id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPMenusDelete);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await this.menu.RemoveMenuById(id, this);
        }
        //
        [HttpPost]
        public async Task<IActionResult> CreateMenuSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("create_success");
            return await Task.Run<ActionResult>(() => { return RedirectToAction(nameof(Index)); });
        }
        [HttpPost]
        public async Task<IActionResult> CreateMenuFails(AdminMenu menu)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            return await Task.Run<ActionResult>(() => { return View("../Admin/Menus/Add", menu); });
        }
        [HttpPost]
        public async Task<IActionResult> UpdateMenuSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("update_success");
            return await Task.Run<ActionResult>(() => { return RedirectToAction(nameof(Index)); });
        }
        [HttpPost]
        public async Task<IActionResult> UpdateMenuFails(AdminMenu menu)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            return await Task.Run<ActionResult>(() => { return View("../Admin/Menus/Edit", menu); });
        }
        [HttpPost]
        public async Task<IActionResult> DeleteMenuFails()
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_fails");
            return await Task.Run<ActionResult>(() => { return RedirectToAction(nameof(Index)); });
        }
        [HttpPost]
        public async Task<IActionResult> DeleteMenuSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            return await Task.Run<ActionResult>(() => { return RedirectToAction(nameof(Index)); });
        }
        [HttpPost]
        public async Task<IActionResult> PageNotFound()
        {
            TempData["mess-type"] = "warning";
            TempData["mess-detail"] = BaseController.BaseMessage("row_fails");
            return await Task.Run<ActionResult>(() => { return RedirectToAction(nameof(Index)); });
        }
    }
}
