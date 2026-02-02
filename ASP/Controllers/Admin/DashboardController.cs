using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using ASP.BaseCommon;
using ASP.Policies;
using ASP.Models.Admin.Accounts;
using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.dashboard")]
    public class DashboardController : Controller
    {
        private readonly ASPDbContext _context;
        private readonly IAuthorizationService _authService;
        private UserManager<ApplicationUser> _userManager;
        private BaseController _baseController;

        public DashboardController(ASPDbContext context, BaseController baseController, IAuthorizationService authService, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _baseController = baseController;
            _authService = authService;
            _userManager = userMgr;
        }
        public async Task<IActionResult> Index()
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPDashboardView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            //var sCommon = new ServiceReferenceCommon.Service1Client();
            //var resultEmail = sCommon.GetEmailFromCodeAsync("970638").Result;
            ////
            //var serviceEz = new ServiceEz.ServiceManPowerEzClient();
            //bool chkAccAD = serviceEz.CheckAccountAdAsync("dmvn970638", "123qwe!!!!!").Result;

            //var servicePPS = new ServiceReferencePPS.Service1Client();
            //var test = servicePPS.GetListProductAsync();
            //var test4 = servicePPS.GetListProductAsync().Result;
            //List<string> obj = new List<string>
            //{
            //    "asadsadsdsad",
            //    "ffdfdfdfdfdd"
            //};
            //var test2 = servicePPS.GetListProductFromListPartAsync(test4.ToArray());

            return View("../Admin/Dashboard/Index");
        }
    }
}
