using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using ASP.Policies;
using ASP.BaseCommon;
using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    [Authorize]
    public class AccountController : Controller, AccountListenerInterface
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        public AccountRepositoryInterface user;
        public BaseController baseController;
        protected string photosPath;
        public AccountController(IAuthorizationService authService, ASPDbContext context, AccountRepositoryInterface user, BaseController baseController)
        {
            _authService = authService;
            _context = context;
            this.user = user;
            this.baseController = baseController;
        }
        [HttpGet]
        [Route("admin/Account", Name = "admin.accounts")]
        [Route("admin/Account/Index")]
        public async Task<IActionResult> Index(string filter, int? fLevelManage, string fRoleName, int? fStatus = 0, int psize = 10, int page = 1, string sort = "-CreatedDate")
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            var list = user.GetAllByLimit(filter, fLevelManage, fRoleName, fStatus, psize, page, sort).Result;
            //get roles 
            ViewBag.sltRoles = user.GetRoleByDropdown("0");
            return View("../Admin/Accounts/Index", list);
        }
        [HttpGet]
        [Route("admin/Account/Create", Name = "admin.accounts.create")]
        public async Task<ActionResult> Create(ApplicationUser user)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            //get roles 
            ViewBag.sltRoles = this.user.GetRoleByDropdown("0");
            //
            return View("../Admin/Accounts/Add", user);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/Account/Create", Name = "admin.accounts.store")]
        public async Task<IActionResult> Store(ApplicationUser user, IFormCollection request, IFormFile Avatar)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            #region validation
            if (string.IsNullOrEmpty(user.PassWord))
            {
                ModelState.AddModelError("Password", "The Password is required.");
            }
            ModelState.Remove("Avatar");
            //
            var chkUserExits = this.user.GetAccountById(null, user.UserName);
            if (chkUserExits != null)
            {
                ModelState.AddModelError("UserName", "Tài khoản đã tồn tại trên hệ thống.");
            }
            if (!ModelState.IsValid)
            {
                //get roles 
                ViewBag.sltRoles = this.user.GetRoleByDropdown("0");
                return CreateAccountFails(user);
            }
            #endregion
            return await this.user.CreateAccount(this, user, Avatar);
        }
        [Route("admin/Account/Edit/{id?}", Name = "admin.accounts.show")]
        public async Task<IActionResult> Show(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            //get roles 
            ViewBag.sltRoles = user.GetRoleByDropdown(id);
            //
            var sltLevelManage = Enum.GetValues(typeof(EnumLevelManage)).Cast<EnumLevelManage>()
            .Select(se => new SelectListItem
            {
                Text = se.ToString(),
                Value = ((int)se).ToString()
            }).FirstOrDefault(w => w.Value == 15.ToString());
            return View("../Admin/Accounts/Edit", user.GetAccountById(id, null));
        }
        [Route("admin/Account/ResetPassword/{id?}", Name = "admin.accounts.resetpw")]
        public async Task<IActionResult> ResetPassword(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion

            return View("../Admin/Accounts/ResetPassword", user.GetAccountById(id, null));
        }
        //
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/Account/Edit", Name = "admin.accounts.edit")]
        public async Task<IActionResult> Edit(string id, ApplicationUser user, IFormFile Avatar)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            #region validation
            ModelState.Remove("Avatar");
            //if (!ModelState.IsValid)
            //{
            //    //get roles 
            //    ViewBag.sltRoles = this.user.GetRoleByDropdown(id);
            //    return UpdateAccountFails(user);
            //}
            #endregion
            return await this.user.UpdateAccountById(id, this, user, Avatar);
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("admin/Account/ResetPassword", Name = "admin.accounts.resetpassword")]
        public async Task<IActionResult> ResetPassword(string id, ApplicationUser user)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion

            return await this.user.ResetPassword(id, this, user);
        }
        [HttpGet]
        [Route("admin/Account/Banned/{id?}", Name = "admin.accounts.banned")]
        public async Task<IActionResult> Banned(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersBanned);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await user.BannedAccountById(id, this);
        }
        [HttpGet]
        [Route("admin/Account/Delete/{id?}", Name = "admin.accounts.delete")]
        public async Task<IActionResult> Delete(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPUsersDelete);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await user.RemoveAccountById(id, this);
        }
        //
        [HttpPost]
        public IActionResult CreateAccountSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("create_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult CreateAccountFails(ApplicationUser user)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            return View("../Admin/Accounts/Add", user);
        }
        [HttpPost]
        public IActionResult UpdateAccountSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("update_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult UpdateAccountFails(ApplicationUser user)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            return View("../Admin/Accounts/Edit", user);
        }
        [HttpPost]
        public IActionResult BannedAccountSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("banned_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult BannedAccountFails()
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("banned_fails");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult DeleteAccountSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult DeleteAccountFails()
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_fails");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult PageNotFound()
        {
            TempData["mess-type"] = "warning";
            TempData["mess-detail"] = BaseController.BaseMessage("row_fails");
            return RedirectToAction(nameof(Index));
        }
        //
    }
}
