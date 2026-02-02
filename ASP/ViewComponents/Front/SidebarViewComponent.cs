using ASP.Models.Admin;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Menus;
using ASP.Models.ASPModel;

namespace ASP.ViewComponents.Front
{
    public class SidebarViewComponent : ViewComponent
    {
        private readonly ASPDbContext _context;
        private UserManager<ApplicationUser> _userManager;
        public SidebarViewComponent(ASPDbContext context, UserManager<ApplicationUser> userMgr)
        {
            _context = context;
            _userManager = userMgr;
        }
        public async Task<IViewComponentResult> InvokeAsync(bool isDone)
        {
            // get header menu
            var findMenu = _context.Menus.FirstOrDefault(w => w.Id == 5);
            List<MenuDetail> objMenu = new List<MenuDetail>();
            if (findMenu != null)
            {
                objMenu = JsonConvert.DeserializeObject<List<MenuDetail>>(findMenu.Content).ToList();
            }
            return View("Default", objMenu);
        }
    }
}
