using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using ASP.BaseCommon;
using ASP.ConfigCommon;
using ASP.Models.Admin.Roles;
using ASP.Policies;
using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.roles")]
    public class RoleController : Controller, RoleListenerInterface
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        //
        public RoleRepositoryInterface _role;
        protected string photosPath;
        public RoleController(IAuthorizationService authService, ASPDbContext context, UserManager<ApplicationUser> userMgr, RoleRepositoryInterface role)
        {
            _authService = authService;
            _context = context;
            _userManager = userMgr;
            _role = role;
        }
        // GET: RoleController
        [HttpGet]
        public async Task<ActionResult> Index(string filter, int pagesize = 10, int page = 1, string sort = "-UpdatedDate")
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            var list = await _role.GetAllByLimit(filter, pagesize, page, sort);
            return View("../Admin/Roles/Index", list);
        }
        //
        [Route("create", Name = "admin.roles.create")]
        public async Task<ActionResult> Create()
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            ViewBag.Permissions = Permissions.GetPermissions();
            return View("../Admin/Roles/Add");
        }
        //
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Create", Name = "admin.roles.store")]
        public async Task<IActionResult> Store(Role role, IFormCollection request)
        {
            if (role.Status == (int)EnumStatusUser.InActive && role.DefaultRole)
            {
                // ko co nhom nao Mac Dinh => bao loi
                ModelState.AddModelError("Name", "Lỗi! Không tồn tại bản ghi Mặc định & trạng thái InActive.");
            }
            if (!ModelState.IsValid)
            {
                return CreateRoleFails(role);
            }
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesCreate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await _role.CreateRole(this, role, request);
        }
        //
        [Route("Edit/{id?}", Name = "admin.roles.show")]
        public async Task<IActionResult> Show(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            ViewBag.Permissions = Permissions.GetPermissions();
            return View("../Admin/Roles/Edit", _role.GetRoleById(id));
        }
        //
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Edit/{id?}", Name = "admin.roles.edit")]
        public async Task<IActionResult> Edit(string id, Role role, IFormCollection request)
        {
            if (role.Status == (int)EnumStatusUser.InActive && role.DefaultRole)
            {
                // ko co nhom nao Mac Dinh => bao loi
                ModelState.AddModelError("Name", "Lỗi! Không tồn tại bản ghi Mặc định & trạng thái InActive.");
            }
            if (!ModelState.IsValid)
            {
                return UpdateRoleFails(role);
            }
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            string userID = _userManager.GetUserId(User);
            return await _role.UpdateRoleById(id, userID, this, role, request);
        }
        //
        [HttpGet]
        [Route("Banned/{id?}", Name = "admin.roles.banned")]
        public async Task<IActionResult> Banned(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesBanned);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await _role.BannedRoleById(id, this);
        }
        //
        [HttpGet]
        [Route("Delete/{id?}", Name = "admin.roles.delete")]
        public async Task<IActionResult> Delete(string id)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPRolesDelete);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            return await _role.RemoveRoleById(id, this);
        }
        //
        [HttpPost]
        public IActionResult CreateRoleSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("create_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult CreateRoleFails(Role role)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("create_fails");
            //
            ViewBag.Permissions = Permissions.GetPermissions();
            return View("../Admin/Roles/Add", role);
        }
        [HttpPost]
        public IActionResult UpdateRoleSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("update_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult UpdateRoleFails(Role role)
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            //
            ViewBag.Permissions = Permissions.GetPermissions();
            return View("../Admin/Roles/Edit", role);
        }
        [HttpPost]
        public IActionResult BannedRoleSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("banned_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult BannedRoleFails()
        {
            TempData["mess-type"] = "error";
            TempData["mess-detail"] = BaseController.BaseMessage("banned_fails");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult DeleteRoleSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("delete_success");
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult DeleteRoleFails(string strEx = null)
        {
            TempData["mess-type"] = "error";
            if (string.IsNullOrEmpty(strEx))
            {
                TempData["mess-detail"] = BaseController.BaseMessage("delete_fails");
            }
            else
            {
                TempData["mess-detail"] = strEx;
            }
            return RedirectToAction(nameof(Index));
        }
        [HttpPost]
        public IActionResult PageNotFound()
        {
            TempData["mess-type"] = "warning";
            TempData["mess-detail"] = BaseController.BaseMessage("row_fails");
            return RedirectToAction(nameof(Index));
        }
    }
}
