using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ReflectionIT.Mvc.Paging;
using ASP.Models.Admin.Roles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Roles
{
    public interface RoleRepositoryInterface
    {
        public Task<IActionResult> CreateRole(RoleListenerInterface listener, Role model, IFormCollection request);
        //
        public Task<IActionResult> UpdateRoleById(string id,string userID, RoleListenerInterface listener, Role model, IFormCollection request);
        //
        public Task<IActionResult> BannedRoleById(string id, RoleListenerInterface listener);
        //
        public Task<IActionResult> RemoveRoleById(string id, RoleListenerInterface listener);
        //
        public Role GetRoleById(string id);
        //
        public Task<PagingList<Role>> GetAllByLimit(string filter, int numberOfPageToShow, int limit, string sort);

    }
}
