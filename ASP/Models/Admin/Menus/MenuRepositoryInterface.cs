using Microsoft.AspNetCore.Mvc;
using ReflectionIT.Mvc.Paging;
using ASP.Models.Admin.Menus;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Menus
{
    public interface MenuRepositoryInterface
    {
        public Task<IActionResult> CreateMenu(MenuListenerInterface listener, Menu request);
        //
        public Task<IActionResult> UpdateMenuById(int id, MenuListenerInterface listener, Menu request);
        //
        public Task<IActionResult> RemoveMenuById(int id, MenuListenerInterface listener);
        //
        public Menu GetMenuById(int id);
        //
        public Task<PagingList<Menu>> GetAllByLimit(string filter, int numberOfPageToShow, int limit, string sort);
        //
    }
}
