using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Accounts
{
    public interface AccountRepositoryInterface
    {
        public Task<IActionResult> CreateAccount(AccountListenerInterface listener, ApplicationUser request, IFormFile formFile);
        //
        public Task<IActionResult> UpdateAccountById(string id, AccountListenerInterface listener, ApplicationUser request, IFormFile formFile);
        //
        public Task<IActionResult> BannedAccountById(string id, AccountListenerInterface listener);
        //
        public Task<IActionResult> RemoveAccountById(string id, AccountListenerInterface listener);
        //
        public Task<IActionResult> ResetPassword(string id, AccountListenerInterface listener, ApplicationUser request);
        //
        public ApplicationUser GetAccountById(string id, string uname);
        //
        public Task<PagingList<ApplicationUser>> GetAllByLimit(string filter, int? fLevelManage, string fRoleName, int? fStatus, int numberOfPageToShow, int limit, string sort);
        //
        public List<SelectListItem> GetRoleByDropdown(string userID);
        // lay danh sach mail tech & administrator
        public Task<List<ApplicationUser>> GetListTechAndAdministrator();
    }
}
