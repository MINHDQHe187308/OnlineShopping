using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Data.Common;
using System.Transactions;
using ASP.BaseCommon;
using ASP.ConfigCommon;
using ASP.Models.Admin.Logs;
using System.Reflection;
using ASP.Models.Admin.Accounts;
using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;

namespace ASP.Models.Admin.Auths
{
    public class AuthRepository : AuthRepositoryInterface
    {
        private readonly ILogger<AccountRepository> _logger;
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;
        private UserManager<ApplicationUser> userManager;

        public AuthRepository(ILogger<AccountRepository> logger, ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log, UserManager<ApplicationUser> usrMgr)
        {
            _logger = logger;
            //_context = context;
            _context = context;
            this.env = env;
            this.log = log;
            this.photosPath = this.env.WebRootPath + "/assets/users";
            userManager = usrMgr;
        }
        //
        public async Task<IActionResult> CreateAccount(AuthListenerInterface listener, Register request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    #region image 
                    ApplicationUser objUser = new ApplicationUser();
                    objUser.UserName = request.UserName;
                    objUser.FullName = request.FullName;
                    objUser.Email = request.Email;
                    objUser.PassWord = request.PassWord;
                    objUser.Avatar = null;
                    //
                    objUser.LevelManage = (int)EnumLevelManage.User;
                    objUser.Status = (int)EnumStatusUser.Pending;
                    //
                    objUser.CreatedDate = DateTime.Now;
                    objUser.UpdatedDate = DateTime.Now;

                    #endregion
                    IdentityResult result = await userManager.CreateAsync(objUser, request.PassWord);
                    #region add user_roles
                    var getRoleDefault = _context.Roles.FirstOrDefault(f => f.DefaultRole == true);
                    if (getRoleDefault != null)
                    {
                        //
                        objUser.RoleId = getRoleDefault.Id;
                        //
                        var addObj = new IdentityUserRole<string>();
                        addObj.UserId = objUser.Id;
                        addObj.RoleId = getRoleDefault.Id;
                        _context.Add(addObj);
                    }
                    else
                    {
                        return listener.RegisterAccountFails(request);
                    }
                    #endregion
                    #region update UserClaims: remove & add new  
                    // lay role id
                    var findRole = _context.Roles.FirstOrDefault(f => f.Id == getRoleDefault.Id);
                    if (findRole != null)
                    {
                        var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(findRole.Content).ToList();
                        foreach (var claim in objContent)
                        {
                            var findClaim = _context.UserClaims.FirstOrDefault(f => f.UserId == objUser.Id && f.ClaimType == claim.Pkey);
                            if (findClaim != null)
                            {
                                findClaim.ClaimValue = claim.Pvalue;
                            }
                            else
                            {
                                _context.UserClaims.Add(new IdentityUserClaim<string>
                                {
                                    UserId = objUser.Id,
                                    ClaimType = claim.Pkey,
                                    ClaimValue = claim.Pvalue
                                });
                            }
                        }
                        _context.SaveChanges();
                        //
                    }
                    #endregion

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Thêm mới " + EnumTypeLog.APP_LOG_USER + " ID:" + objUser.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", null, objUser.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", null, objUser.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", null, objUser.PassWord);
                    logContent += EnumTypeLog.SetLogLine("Email", null, objUser.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", null, objUser.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Vai trò", null, objUser.RoleId.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", null, objUser.PhoneNumber);
                    logContent += EnumTypeLog.SetLogLine("Avatar", null, "");
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", null, objUser.Status.ToString());
                    //
                    if (logContent != null)
                    {
                        this.log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.RegisterAccountSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}: {1}", MethodBase.GetCurrentMethod().DeclaringType, ex.Message);
                    return listener.RegisterAccountFails(request);
                }
            }
        }
        //
        public bool CheckAccount(Register request)
        {
            return _context.Users.Any(f => f.UserName == request.UserName);
        }
        public async Task<bool> CheckAccountAD(Register request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    #region image 
                    ApplicationUser objUser = new ApplicationUser();
                    objUser.UserName = request.UserName;
                    objUser.FullName = request.FullName;
                    objUser.Email = request.Email;
                    objUser.PassWord = request.PassWord;
                    objUser.Avatar = null;
                    //
                    objUser.LevelManage = (int)EnumLevelManage.User;
                    objUser.Status = (int)EnumStatusUser.Active;
                    //
                    objUser.CreatedDate = DateTime.Now;
                    objUser.UpdatedDate = DateTime.Now;
                    #endregion
                    IdentityResult result = await userManager.CreateAsync(objUser, request.PassWord);
                    #region add user_roles
                    var getRoleDefault = _context.Roles.FirstOrDefault(f => f.DefaultRole == true);
                    if (getRoleDefault != null)
                    {
                        objUser.RoleId = getRoleDefault.Id;
                        var addObj = new IdentityUserRole<string>();
                        addObj.UserId = objUser.Id;
                        addObj.RoleId = getRoleDefault.Id;
                        _context.Add(addObj);
                    }
                    else
                    {
                        return false;
                    }
                    #endregion
                    #region update UserClaims: remove & add new  
                    // lay role id
                    var findRole = _context.Roles.FirstOrDefault(f => f.Id == getRoleDefault.Id);
                    if (findRole != null)
                    {
                        var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(findRole.Content).ToList();
                        foreach (var claim in objContent)
                        {
                            var findClaim = _context.UserClaims.FirstOrDefault(f => f.UserId == objUser.Id && f.ClaimType == claim.Pkey);
                            if (findClaim != null)
                            {
                                findClaim.ClaimValue = claim.Pvalue;
                            }
                            else
                            {
                                _context.UserClaims.Add(new IdentityUserClaim<string>
                                {
                                    UserId = objUser.Id,
                                    ClaimType = claim.Pkey,
                                    ClaimValue = claim.Pvalue
                                });
                            }
                        }
                        _context.SaveChanges();
                        //
                    }
                    #endregion

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Thêm mới " + EnumTypeLog.APP_LOG_USER + " ID:" + objUser.Id);
                    logContent += EnumTypeLog.SetLogLine("Tài khoản", null, objUser.UserName);
                    logContent += EnumTypeLog.SetLogLine("Họ tên", null, objUser.FullName);
                    logContent += EnumTypeLog.SetLogLine("Mật khẩu", null, objUser.PassWord);
                    logContent += EnumTypeLog.SetLogLine("Email", null, objUser.Email);
                    logContent += EnumTypeLog.SetLogLine("Loại tài khoản", null, objUser.LevelManage.ToString());
                    logContent += EnumTypeLog.SetLogLine("Vai trò", null, objUser.RoleId.ToString());
                    logContent += EnumTypeLog.SetLogLine("Điện thoại", null, objUser.PhoneNumber);
                    logContent += EnumTypeLog.SetLogLine("Avatar", null, "");
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", null, objUser.Status.ToString());
                    //
                    if (logContent != null)
                    {
                        this.log.CreateLog(EnumTypeLog.APP_LOG_USER, logContent);
                    }
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return true;
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}: {1}", MethodBase.GetCurrentMethod().DeclaringType, ex.Message);
                    return false;
                }
            }
        }
    }
}
