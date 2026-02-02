using ASP.Models.Admin;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASP.Models.Admin.Menus;
using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;

namespace ASP.ViewComponents.Front
{
    public class HeaderStyleViewComponent : ViewComponent
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        public HeaderStyleViewComponent(ASPDbContext context, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _userManager = userMgr;
        }
        public async Task<IViewComponentResult> InvokeAsync(bool isDone)
        {
            // get favicon
            var findFav = _context.ThemeOptions.FirstOrDefault(w => w.Name == "_optFavicon");
            ViewBag.strFav = (findFav != null) ? findFav.Value : "";
            return View("Default");
        }
    }
}
