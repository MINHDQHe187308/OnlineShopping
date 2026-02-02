using ASP.ConfigCommon;
using ASP.Models.Admin;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Menus;
using ASP.Models.ASPModel;

namespace ASP.ViewComponents.Front
{
    public class HeaderViewComponent : ViewComponent
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        public HeaderViewComponent(ASPDbContext context, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _userManager = userMgr;
        }
        public async Task<IViewComponentResult> InvokeAsync(bool isDone)
        {
            // check login user
            ApplicationUser userCurrent = await _userManager.GetUserAsync(HttpContext.User);
            ViewBag.vbUsername = (userCurrent != null) ? userCurrent.UserName : null;
            // get logo
            var find = _context.ThemeOptions.FirstOrDefault(w => w.Name == "_optLogo");
            ViewBag.strLogo = (find != null) ? find.Value : "";
            
            // get header menu
            var findMenu = _context.Menus.FirstOrDefault(w => w.Id == 1);

            List<MenuDetail> objMenu = new List<MenuDetail>();
            if (findMenu != null)
            {
                objMenu = JsonConvert.DeserializeObject<List<MenuDetail>>(findMenu.Content).ToList();
            }
            return View(objMenu);
        }
    }
}
