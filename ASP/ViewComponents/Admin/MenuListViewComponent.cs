using ASP.ConfigCommon;
using ASP.Models.Admin;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;

namespace ASP.ViewComponents.Front
{
    public class MenuListViewComponent : ViewComponent
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        public MenuListViewComponent(ASPDbContext context, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _userManager = userMgr;
        }
        public async Task<IViewComponentResult> InvokeAsync(bool isDone)
        {
            var items = await GetItemsAsync(isDone);
            return View("Default", items);
        }
        private async Task<List<PermissionDetail>> GetItemsAsync(bool isDone)
        {
            var userID = _userManager.GetUserId(Request.HttpContext.User);
            var arrClaims = _context.UserClaims.Where(w => w.UserId == userID).ToList();
            var allPermissions = Permissions.GetPermissions().Where(p => p.Proute != "");
            var getPermissions = new List<PermissionDetail>();
            foreach (var item in allPermissions)
            {
                var chkClaim = arrClaims.Any(f => f.ClaimType == item.Pcan && f.ClaimValue == "1");
                if (chkClaim)
                {
                    getPermissions.Add(item);
                }
            }
            return getPermissions;
        }
    }
}
