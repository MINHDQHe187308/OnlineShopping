using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Menus
{
    public interface MenuListenerInterface
    {
        public Task<IActionResult> CreateMenuSuccess();
        public Task<IActionResult> CreateMenuFails(Menu menu);
        public Task<IActionResult> UpdateMenuSuccess();
        public Task<IActionResult> UpdateMenuFails(Menu menu);
        public Task<IActionResult> DeleteMenuSuccess();
        public Task<IActionResult> DeleteMenuFails();
        public Task<IActionResult> PageNotFound();
    }
}
