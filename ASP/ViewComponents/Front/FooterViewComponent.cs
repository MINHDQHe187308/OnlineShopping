using ASP.ConfigCommon;
using ASP.Models.Admin;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using ASP.Models.Admin.Menus;
using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;

namespace ASP.ViewComponents.Front
{
    public class FooterViewComponent : ViewComponent
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        public FooterViewComponent(ASPDbContext context, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _userManager = userMgr;
        }
        public async Task<IViewComponentResult> InvokeAsync(bool isDone)
        {
            // _optFooterContent
            var findFooterContent = _context.ThemeOptions.FirstOrDefault(w => w.Name == "_optFooterContent");
            ViewBag.strFooterContent = (findFooterContent != null) ? findFooterContent.Value : "";
            return View("Default");
        }

    }
}
