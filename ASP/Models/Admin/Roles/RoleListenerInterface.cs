using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Roles
{
    public interface RoleListenerInterface
    {
        public IActionResult CreateRoleSuccess();
        public IActionResult CreateRoleFails(Role role);
        public IActionResult UpdateRoleSuccess();
        public IActionResult UpdateRoleFails(Role role);
        public IActionResult BannedRoleSuccess();
        public IActionResult BannedRoleFails();
        public IActionResult DeleteRoleSuccess();
        public IActionResult DeleteRoleFails(string strEx = null);
        public IActionResult PageNotFound();
    }
}
