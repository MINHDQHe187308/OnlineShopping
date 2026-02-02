using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ASP.BaseCommon;
using ASP.Models.Admin.Logs;
using ASP.Policies;
using Microsoft.AspNetCore.Mvc.Rendering;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.logs")]
    public class LogController : Controller
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        public LogRepositoryInterface log;
        public LogController(IAuthorizationService authService, ASPDbContext context, LogRepositoryInterface log)
        {
            _authService = authService;
            _context = context;
            this.log = log;
        }
        public async Task<IActionResult> Index(string filter = "", int pagesize = 10, int page = 1, string sort = "-ID")
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPLogsView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            ViewBag.vblogs = EnumTypeLog.GetLogTypeDescription().Select(s => new SelectListItem { Value = s.Key, Text = s.Value }).ToList();
            var list = log.GetAllByLimit(filter, pagesize, page, sort).Result;
            return View("../Admin/Logs/Index", list);
        }
    }
}
