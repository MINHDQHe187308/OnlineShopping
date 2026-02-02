using Microsoft.AspNetCore.Cryptography.KeyDerivation;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Routing;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Transactions;
using System.Reflection;
using Microsoft.Extensions.Logging;
using ASP.Models.Admin.Logs;
using ASP.Models.Admin.Roles;
using ASP.BaseCommon;
using ASP.ConfigCommon;
using ASP.Models.ASPModel;

namespace ASP.Models.Admin.Accounts
{
    public class AccountRepository : AccountRepositoryInterface
    {
        private readonly ILogger<AccountRepository> _logger;
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;
        private UserManager<ApplicationUser> userManager;

        public AccountRepository(ILogger<AccountRepository> logger, ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log, UserManager<ApplicationUser> usrMgr)
        {
            _logger = logger;
            _context = context;
            this.env = env;
            this.log = log;
            photosPath = this.env.WebRootPath + "/assets/users/";
            userManager = usrMgr;
        }
        //
        public async Task<IActionResult> CreateAccount(AccountListenerInterface listener, ApplicationUser request, IFormFile formFile)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    #region image 
                    if (!Directory.Exists(photosPath))
                    {
                        Directory.CreateDirectory(photosPath);
                    }
                    string image_name = null;
                    if (formFile != null)
                    {
                        image_name = "avatar_" + DateTime.Now.Ticks + formFile.FileName;
                        var filePath = Path.Combine(photosPath, image_name);
                        var extension = Path.GetExtension(filePath);
                        using (var stream = File.Create(filePath))
                        {
                            formFile.CopyTo(stream);
                        }
                    }
                    request.Avatar = image_name;
                    request.CreatedDate = DateTime.Now;
                    request.UpdatedDate = DateTime.Now;
                    #endregion
                    IdentityResult result = await userManager.CreateAsync(request, request.PassWord);
                    #region add user_roles
                    _context.UserRoles.Add(new IdentityUserRole<string>
                    {
                        UserId = request.Id,
                        RoleId = request.RoleId
                    });
                    #endregion

                    #region update UserClaims: remove & add new  
                    // lay role id
                    var findRole = _context.Roles.Find(request.RoleId);
                    if (findRole != null)
                    {
                        var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(findRole.Content).ToList();
                        foreach (var claim in objContent)
                        {
                            var findClaim = _context.UserClaims.FirstOrDefault(f => f.UserId == request.Id && f.ClaimType == claim.Pkey);
                            if (findClaim != null)
                            {
                                findClaim.ClaimValue = claim.Pvalue;
                            }
                            else
                            {
                                _context.UserClaims.Add(new IdentityUserClaim<string>
                                {
                                    UserId = request.Id,
                                    ClaimType = claim.Pkey,
                                    ClaimValue = claim.Pvalue
                                });
                            }
                        }
                        await _context.SaveChangesAsync();
                        //_context.SaveChanges();
                        //
                    }
                    #endregion

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Thêm mới " + EnumTypeLog.APP_LOG_USER + " ID:" + request.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", null, request.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", null, request.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", null, request.PassWord != null ? request.PassWord : "");
                    logContent += EnumTypeLog.SetLogLine("Email", null, request.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", null, request.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Vai trò", null, request.RoleId.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", null, request.PhoneNumber != null ? request.PhoneNumber.ToString() : "");
                    logContent += EnumTypeLog.SetLogLine("Avatar", null, image_name != null ? image_name : "");
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", null, request.Status.ToString());
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.CreateAccountSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.CreateAccountFails(request);
                }
            }
        }
        //
        public async Task<IActionResult> UpdateAccountById(string id, AccountListenerInterface listener, ApplicationUser request, IFormFile formFile)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    ApplicationUser find = await userManager.FindByIdAsync(id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Sửa " + EnumTypeLog.APP_LOG_USER + " ID: " + find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", find.UserName, request.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", find.FullName, request.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", find.PassWord, request.PassWord);
                    logContent += EnumTypeLog.SetLogLine("Email", find.Email, request.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", find.LevelManage.ToString(), request.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", find.PhoneNumber, request.PhoneNumber);
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), request.Status.ToString());
                    //
                    #endregion
                    find.UserName = string.IsNullOrEmpty(request.UserName) ? null: request.UserName.Trim();
                    find.FullName = !string.IsNullOrEmpty(request.FullName) ? request.FullName.Trim() : null;
                    //find.PassWord = request.PassWord;
                    //if (!string.IsNullOrEmpty(request.PassWord))
                    //{
                    //    var password = new PasswordHasher<ApplicationUser>();
                    //    var hashed = password.HashPassword(find, request.PassWord);
                    //    find.PasswordHash = hashed;
                    //}
                    //
                    find.Email = request.Email;
                    find.LevelManage = request.LevelManage;
                    find.PhoneNumber = request.PhoneNumber;
                    find.Status = request.Status;
                    find.UpdatedDate = DateTime.Now;
                    #region image 
                    // full path to file in temp location
                    if (!Directory.Exists(photosPath))
                    {
                        Directory.CreateDirectory(photosPath);
                    }
                    string image_name = null;
                    if (formFile != null)
                    {
                        // delete old file
                        if (find.Avatar != null)
                        {
                            if (File.Exists(Path.Combine(photosPath, find.Avatar)))
                            {
                                File.Delete(Path.Combine(photosPath, find.Avatar));
                            }
                        }
                        // add new file
                        image_name = "avatar_" + DateTime.Now.Ticks + formFile.FileName;
                        var filePath = Path.Combine(photosPath, image_name);
                        var extension = Path.GetExtension(filePath);
                        //
                        using (var stream = File.Create(filePath))
                        {
                            formFile.CopyTo(stream);
                        }
                        logContent += EnumTypeLog.SetLogLine("Avatar", find.Avatar, image_name);
                        find.Avatar = image_name;
                    }
                    else
                    {
                        if (request.rmavatar == "remove")
                        {
                            // delete old file
                            if (find.Avatar != null)
                            {
                                if (File.Exists(Path.Combine(photosPath, find.Avatar)))
                                {
                                    File.Delete(Path.Combine(photosPath, find.Avatar));
                                }
                            }
                            // remove avatar
                            logContent += EnumTypeLog.SetLogLine("Avatar", find.Avatar, image_name);
                            find.Avatar = null;
                        }
                    }
                    #endregion

                    #region update user_roles
                    var findUserRoles = _context.UserRoles.FirstOrDefault(w => w.UserId == find.Id);
                    if (findUserRoles != null)
                    {
                        if (findUserRoles.UserId == find.Id && findUserRoles.RoleId != request.RoleId)
                        {
                            logContent += EnumTypeLog.SetLogLine("Vai trò", findUserRoles.RoleId.ToString(), request.RoleId.ToString());
                            _context.Remove(findUserRoles);
                            //
                            _context.UserRoles.Add(new IdentityUserRole<string>
                            {
                                UserId = request.Id,
                                RoleId = request.RoleId
                            });
                        }
                    }
                    else
                    {
                        _context.UserRoles.Add(new IdentityUserRole<string>
                        {
                            UserId = request.Id,
                            RoleId = request.RoleId
                        });
                    }
                    #endregion

                    #region update UserClaims: remove & add new  
                    // lay role id
                    var findRole = _context.Roles.FirstOrDefault(f => f.Id == request.RoleId);
                    if (findRole != null)
                    {
                        var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(findRole.Content).ToList();
                        foreach (var claim in objContent)
                        {
                            var findClaim = _context.UserClaims.FirstOrDefault(f => f.UserId == request.Id && f.ClaimType == claim.Pkey);
                            if (findClaim != null)
                            {
                                findClaim.ClaimValue = claim.Pvalue;
                            }
                            else
                            {
                                _context.UserClaims.Add(new IdentityUserClaim<string>
                                {
                                    UserId = request.Id,
                                    ClaimType = claim.Pkey,
                                    ClaimValue = claim.Pvalue
                                });
                            }
                        }
                        //_context.SaveChanges();
                        await _context.SaveChangesAsync();
                        //
                    }
                    #endregion

                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    var result = await userManager.UpdateAsync(find);
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.UpdateAccountSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.UpdateAccountFails(request);
                }
            }

        }
        //
        public async Task<IActionResult> BannedAccountById(string id, AccountListenerInterface listener)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    var find = _context.Users.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Khóa " + EnumTypeLog.APP_LOG_USER + " Thành viên");
                    logContent += EnumTypeLog.SetLogLine("ID", null, find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", null, find.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", null, find.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", null, find.PassWord);
                    logContent += EnumTypeLog.SetLogLine("Email", null, find.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", null, find.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", null, find.PhoneNumber);
                    logContent += EnumTypeLog.SetLogLine("Avatar", null, find.Avatar);
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), EnumStatusUser.InActive.ToString());

                    #region get user_roles
                    var findAccountRoles = _context.UserRoles.FirstOrDefault(w => w.UserId == find.Id);
                    if (findAccountRoles != null)
                    {
                        logContent += EnumTypeLog.SetLogLine("Vai trò", null, findAccountRoles.RoleId.ToString());
                    }
                    #endregion
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    #endregion
                    // update Status: inactive
                    find.Status = (int)EnumStatusUser.InActive;
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.BannedAccountSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.BannedAccountFails();
                }
            }

        }
        //
        public async Task<IActionResult> RemoveAccountById(string id, AccountListenerInterface listener)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    var find = _context.Users.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Xóa " + EnumTypeLog.APP_LOG_USER + " Thành viên");
                    logContent += EnumTypeLog.SetLogLine("ID", null, find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", null, find.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", null, find.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", null, find.PassWord);
                    logContent += EnumTypeLog.SetLogLine("Email", null, find.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", null, find.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", null, find.PhoneNumber);
                    logContent += EnumTypeLog.SetLogLine("Avatar", null, find.Avatar);
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), "Xóa");

                    #region get user_roles
                    var findAccountRoles = _context.UserRoles.FirstOrDefault(w => w.UserId == find.Id);
                    if (findAccountRoles != null)
                    {
                        logContent += EnumTypeLog.SetLogLine("Role", null, findAccountRoles.RoleId.ToString());
                    }
                    #endregion
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    #endregion
                    //
                    var userRoles = _context.UserRoles.Where(w => w.UserId == find.Id);
                    _context.UserRoles.RemoveRange(userRoles);
                    var userClaims = _context.UserClaims.Where(w => w.UserId == find.Id);
                    _context.UserClaims.RemoveRange(userClaims);
                    _context.Users.Remove(find);
                    //
                    //_context.SaveChanges();

                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.DeleteAccountSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.DeleteAccountFails();
                }
            }

        }
        //
        public async Task<IActionResult> ResetPassword(string id, AccountListenerInterface listener, ApplicationUser request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled)) {
                try
                {
                    var find = _context.Users.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Sửa " + EnumTypeLog.APP_LOG_USER + " ID: " + find.Id);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", find.PassWord, request.PassWord);
                    #endregion

                    find.PassWord = request.PassWord;
                    if (!string.IsNullOrEmpty(request.PassWord))
                    {
                        var password = new PasswordHasher<ApplicationUser>();
                        var hashed = password.HashPassword(find, request.PassWord);
                        find.PasswordHash = hashed;
                    }
                    find.UpdatedDate = DateTime.Now;

                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    var result = await userManager.UpdateAsync(find);

                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.UpdateAccountSuccess();
                }
                catch (DbException ex) {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.DeleteAccountFails();
                }
            }
        }
        //
        public ApplicationUser GetAccountById(string id, string uname)
        {
            if (uname != null)
            {
                var find = _context.Users.AsNoTracking().FirstOrDefault(f => f.UserName == uname);
                if (find != null)
                {
                    find.PassWord = null;
                    find.PasswordHash = null;
                }
                return find;
            }
            else
            {
                var find = _context.Users.AsNoTracking().FirstOrDefault(f => f.Id == id);
                find.PassWord = null;
                find.PasswordHash = null;
                return find;
            }
        }
        //
        public async Task<PagingList<ApplicationUser>> GetAllByLimit(string filter = "", int? fLevelManage = null, string fRoleName = "", int? fStatus = 0, int numberOfPageToShow = 15, int page = 0, string sort = null)
        {
            //
            var qry = (from u in _context.Users
                       join urole in _context.UserRoles on u.Id equals urole.UserId
                       join role in _context.Roles on urole.RoleId equals role.Id
                       select new ApplicationUser()
                       {
                           Id = u.Id,
                           UserName = u.UserName,
                           FullName = u.FullName,
                           Email = u.Email,
                           Avatar = u.Avatar,
                           LevelManage = u.LevelManage,
                           RoleName = role.Name,
                           RoleId = role.Id,
                           Status = u.Status,
                           CreatedDate = u.CreatedDate,
                           UpdatedDate = u.UpdatedDate
                       }).OrderByDescending(f => f.Id).AsQueryable();
            //
            if (!string.IsNullOrWhiteSpace(filter))
            {
                qry = qry.Where(p => p.UserName.Contains(filter));
            }
            if (fLevelManage != null)
            {
                qry = qry.Where(p => p.LevelManage == fLevelManage);
            }
            if (!string.IsNullOrWhiteSpace(fRoleName))
            {
                qry = qry.Where(p => p.RoleId.Contains(fRoleName));
            }
            if (fStatus > 0)
            {
                qry = qry.Where(p => p.Status == fStatus);
            }
            var objs = await PagingList.CreateAsync(qry, numberOfPageToShow, page, sort, "ID");
            objs.RouteValue = new RouteValueDictionary { { "filter", filter }, { "fLevelManage", fLevelManage }, { "fRoleName", fRoleName }, { "fStatus", fStatus }, { "psize", numberOfPageToShow } };
            return objs;
        }
        //
        public List<SelectListItem> GetRoleByDropdown(string userID = "0")
        {
            var getAccountRoles = _context.UserRoles.FirstOrDefault(f => f.UserId == userID);
            var roleID = getAccountRoles != null ? getAccountRoles.RoleId.ToString() : "0";
            var slt = _context.Roles.Where(w => w.Status == (int)EnumStatusUser.Active).Select(s => new SelectListItem()
            {
                Text = s.Name,
                Value = s.Id.ToString(),
                Selected = s.Id == roleID ? true : false
            }).ToList();
            return slt;
        }

        public async Task<List<ApplicationUser>> GetListTechAndAdministrator()
        {
            return await _context.Users.Where(w => w.LevelManage == (int)EnumLevelManage.Editor || w.LevelManage == (int)EnumLevelManage.Administrator).Select(u => new ApplicationUser()
            {
                Id = u.Id,
                UserName = u.UserName,
                FullName = u.FullName,
                Email = u.Email,
                Avatar = u.Avatar,
                LevelManage = u.LevelManage,
                Status = u.Status,
                CreatedDate = u.CreatedDate,
                UpdatedDate = u.UpdatedDate
            }).ToListAsync();
        }
    }
}
