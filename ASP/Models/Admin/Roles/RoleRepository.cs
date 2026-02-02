using ASP.BaseCommon;
using ASP.ConfigCommon;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ReflectionIT.Mvc.Paging;
using System.Data.Common;
using System.Transactions;
using ASP.Models.Admin.Logs;
using ASP.Models.Admin.Accounts;
using System.Reflection;
using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;

namespace ASP.Models.Admin.Roles
{
    public class RoleRepository : RoleRepositoryInterface
    {
        private readonly ILogger<RoleRepository> _logger;
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;

        public RoleRepository(ILogger<RoleRepository> logger, ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log)
        {
            _logger = logger;
            _context = context;
            this.env = env;
            this.log = log;
            photosPath = this.env.WebRootPath + "/assets/roles";
        }
        //
        public async Task<IActionResult> CreateRole(RoleListenerInterface listener, Role model, IFormCollection request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    List<ActionDetail> permissions = new List<ActionDetail>();
                    List<string> strContent = new List<string>();
                    foreach (var item in request)
                    {
                        if (item.Key.Contains(Permissions.APP_NAME))
                        {
                            permissions.Add(new ActionDetail() { Pkey = item.Key, Pvalue = item.Value });
                            // handle log
                            string setVal = "Deny";
                            if (item.Value == "1")
                            {
                                setVal = "Allow";
                            }
                            strContent.Add(item.Key.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":<strong>" + setVal + "</strong>");
                        }
                    }
                    #region check default role
                    if (model.DefaultRole)
                    {
                        var getDefaultRoles = _context.Roles.Where(w => w.DefaultRole == true && w.Id != model.Id).ToList();
                        foreach (var item in getDefaultRoles)
                        {
                            item.DefaultRole = false;
                        }
                    }
                    #endregion
                    //
                    model.Content = JsonConvert.SerializeObject(permissions);
                    //
                    model.CreatedDate = DateTime.Now;
                    model.UpdatedDate = DateTime.Now;
                    _context.Add(model);
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Thêm mới " + EnumTypeLog.APP_LOG_ROLE + " ID:" + model.Id);
                    logContent += EnumTypeLog.SetLogLine("Tiêu đề", null, model.Name);
                    logContent += EnumTypeLog.SetLogLine("Mặc định", null, model.DefaultRole.ToString());
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", null, model.Status.ToString());
                    logContent += EnumTypeLog.SetLogLine("Quyền hạn", null, string.Join(", ", strContent));
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_ROLE, logContent);
                    }
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.CreateRoleSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.CreateRoleFails(model);
                }
            }
        }
        //
        public async Task<IActionResult> UpdateRoleById(string id, string userID, RoleListenerInterface listener, Role model, IFormCollection request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    var find = _context.Roles.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    List<ActionDetail> permissions = new List<ActionDetail>();
                    foreach (var item in request)
                    {
                        if (item.Key.Contains(Permissions.APP_NAME))
                        {
                            permissions.Add(new ActionDetail() { Pkey = item.Key, Pvalue = item.Value });
                            //
                        }
                    }
                    #region UserRoles
                    var arrUser = _context.UserRoles.Where(w => w.RoleId == find.Id).Select(s => s.UserId).ToList();
                    #endregion
                    //
                    #region check old => new permission
                    //
                    var objOldContent = JsonConvert.DeserializeObject<List<ActionDetail>>(find.Content).ToList();
                    List<string> strContent = new List<string>();

                    foreach (var item in permissions)
                    {
                        #region update permission 
                        foreach (var itemUser in arrUser)
                        {
                            var obj = _context.UserClaims.FirstOrDefault(f => f.UserId == itemUser && f.ClaimType == $"{item.Pkey}");
                            if (obj != null)
                            {
                                obj.ClaimValue = item.Pvalue;
                            }
                            else
                            {
                                _context.UserClaims.Add(new IdentityUserClaim<string>
                                {
                                    UserId = itemUser,
                                    ClaimType = item.Pkey,
                                    ClaimValue = item.Pvalue
                                });
                            }
                        }// endforeach;
                        //_context.SaveChanges();
                        await _context.SaveChangesAsync();
                        #endregion
                        //
                        #region add log
                        var findOld = objOldContent.FirstOrDefault(f => f.Pkey == item.Pkey);
                        if (findOld != null)
                        {
                            var chk = findOld.Pvalue != item.Pvalue ? item.Pvalue == "1" ? item.Pkey.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":Deny <strong>=></strong> Allow" : item.Pkey.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":Allow <strong>=></strong> Deny" : null;
                            if (chk != null)
                            {
                                strContent.Add(chk);
                            }
                        }
                        else
                        {
                            // handle log
                            string setVal = "Deny";
                            if (item.Pvalue == "1")
                            {
                                setVal = "Allow";
                            }
                            strContent.Add(item.Pkey.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":<strong>" + setVal + "</strong>");
                        }
                        #endregion
                    }
                    #endregion

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Sửa " + EnumTypeLog.APP_LOG_ROLE + " ID: " + find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tiêu đề", find.Name, model.Name);
                    logContent += EnumTypeLog.SetLogLine("Mặc định", find.DefaultRole.ToString(), model.DefaultRole.ToString());
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), model.Status.ToString());
                    logContent += EnumTypeLog.SetLogLine("Quyền hạn", null, string.Join(", ", strContent));
                    //
                    #endregion

                    #region check default role
                    if (model.DefaultRole)
                    {
                        var getDefaultRoles = _context.Roles.Where(w => w.DefaultRole == true && w.Id != find.Id).ToList();
                        foreach (var item in getDefaultRoles)
                        {
                            item.DefaultRole = false;
                        }
                    }
                    else
                    {
                        // neu ko co nhom nao = default
                        #region check default role
                        var getDefaultRoles = _context.Roles.Where(w => w.DefaultRole == true && w.Id != find.Id).ToList();
                        if (getDefaultRoles.Count() == 0)
                        {
                            model.DefaultRole = true;
                        }
                        else
                        {
                            int totalDefaultRole = getDefaultRoles.Count();
                            if (totalDefaultRole > 1)
                            {
                                int countR = 0;
                                foreach (var item in getDefaultRoles)
                                {
                                    countR++;
                                    if (countR < totalDefaultRole)
                                    {
                                        item.DefaultRole = false;
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion

                    find.Name = string.IsNullOrEmpty(model.Name) ? model.Name : model.Name.Trim();
                    find.Content = JsonConvert.SerializeObject(permissions);
                    find.DefaultRole = model.DefaultRole;
                    find.Status = model.Status;
                    ///
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_ROLE, logContent);
                    }
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.UpdateRoleSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.UpdateRoleFails(model);
                }
            }
        }
        //
        public async Task<IActionResult> BannedRoleById(string id, RoleListenerInterface listener)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {

                    var find = _context.Roles.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Khóa " + EnumTypeLog.APP_LOG_ROLE + " ID: " + find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tiêu đề", null, find.Name);
                    logContent += EnumTypeLog.SetLogLine("Mặc định", null, find.DefaultRole.ToString());
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), EnumStatusUser.InActive.ToString());
                    //
                    var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(find.Content).ToList();
                    List<string> strContent = new List<string>();
                    foreach (var item in objContent)
                    {
                        if (item.Pkey.Contains(Permissions.APP_NAME))
                        {
                            // handle log
                            string setVal = "Deny";
                            if (item.Pvalue == "1")
                            {
                                setVal = "Allow";
                            }
                            strContent.Add(item.Pkey.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":<strong>" + setVal + "</strong>");
                        }
                    }
                    logContent += EnumTypeLog.SetLogLine("Quyền hạn", null, string.Join(", ", strContent));
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_ROLE, logContent);
                    }
                    #endregion

                    #region check default role
                    if (find.DefaultRole)
                    {
                        var getDefaultRoles = _context.Roles.FirstOrDefault(w => w.DefaultRole == false);
                        if (getDefaultRoles != null)
                        {
                            getDefaultRoles.DefaultRole = true;
                        }
                        else
                        {
                            // ko co nhom nao Mac Dinh => bao loi
                            scope.Dispose();
                            return listener.DeleteRoleFails("Lỗi! Danh sách Vai trò phải tồn tại bản ghi dạng Mặc định.");
                        }
                    }
                    #endregion
                    //
                    find.Status = (int)EnumStatusUser.InActive;
                    find.DefaultRole = false;
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.BannedRoleSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.BannedRoleFails();
                }
            }
        }
        //
        public async Task<IActionResult> RemoveRoleById(string id, RoleListenerInterface listener)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {

                    var find = _context.Roles.FirstOrDefault(f => f.Id == id);
                    if (find == null)
                    {
                        return listener.PageNotFound();
                    }
                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Xóa " + EnumTypeLog.APP_LOG_ROLE + " ID: " + find.Id);
                    logContent += EnumTypeLog.SetLogLine("Tiêu đề", null, find.Name);
                    logContent += EnumTypeLog.SetLogLine("Mặc định", null, find.DefaultRole.ToString());
                    logContent += EnumTypeLog.SetLogLine("Trạng thái", find.Status.ToString(), "Remove");
                    //
                    var objContent = JsonConvert.DeserializeObject<List<ActionDetail>>(find.Content).ToList();
                    List<string> strContent = new List<string>();
                    foreach (var item in objContent)
                    {
                        if (item.Pkey.Contains(Permissions.APP_NAME))
                        {
                            // handle log
                            string setVal = "Deny";
                            if (item.Pvalue == "1")
                            {
                                setVal = "Allow";
                            }
                            strContent.Add(item.Pkey.Replace(".", " ").Replace("permissions[", "").Replace("]", "") + ":<strong>" + setVal + "</strong>");
                        }
                    }
                    logContent += EnumTypeLog.SetLogLine("Quyền hạn", null, string.Join(", ", strContent));
                    #endregion
                    #region log user
                    var qryUser = (from u in _context.Users
                                   join urole in _context.UserRoles on u.Id equals urole.UserId
                                   join role in _context.Roles on urole.RoleId equals role.Id
                                   select new ApplicationUser()
                                   {
                                       Id = u.Id,
                                       UserName = u.UserName,
                                       FullName = u.FullName,
                                       Email = u.Email,
                                       LevelManage = u.LevelManage,
                                       RoleName = role.Name,
                                       RoleId = role.Id,
                                       Status = u.Status
                                   }).Where(w => w.RoleId == id).OrderByDescending(f => f.Id).ToList();
                    string strLogUser = "";
                    foreach (var item in qryUser)
                    {
                        strLogUser += "<div>";
                        strLogUser += "<strong>Tài khoản:</strong>" + item.UserName;
                        strLogUser += ", <strong>Họ tên:</strong>" + item.FullName;
                        strLogUser += ", <strong>Cấp bậc:</strong>" + item.LevelManage.ToString();
                        strLogUser += ", <strong>Vai trò:</strong>" + item.RoleName;
                        strLogUser += "</div>";
                        // delete user
                        var userClaims = _context.UserClaims.Where(w => w.UserId == item.Id);
                        _context.UserClaims.RemoveRange(userClaims);
                        var users = _context.Users.Where(w => w.Id == item.Id);
                        _context.Users.RemoveRange(users);
                    }
                    logContent += EnumTypeLog.SetLogLine("Thành viên", null, string.Join(", ", strLogUser));
                    //
                    if (logContent != null)
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_ROLE, logContent);
                    }
                    #endregion
                    #region check default role
                    if (find.DefaultRole)
                    {
                        var getDefaultRoles = _context.Roles.FirstOrDefault(w => w.DefaultRole == false);
                        if (getDefaultRoles != null)
                        {
                            getDefaultRoles.DefaultRole = true;
                        }
                        else
                        {
                            // ko co nhom nao Mac Dinh => bao loi
                            scope.Dispose();
                            return listener.DeleteRoleFails("Lỗi! Danh sách Vai trò phải tồn tại bản ghi dạng Mặc định.");
                        }
                    }
                    #endregion
                    #region delete
                    // delete user role
                    var userRoles = _context.UserRoles.Where(w => w.RoleId == find.Id);
                    _context.UserRoles.RemoveRange(userRoles);
                    //delete role
                    _context.Roles.Remove(find);
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return listener.DeleteRoleSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.DeleteRoleFails(ex.Message);
                }
            }
        }
        //
        public Role GetRoleById(string id)
        {
            var find = _context.Roles.Select(s => new Role()
            {
                Id = s.Id,
                Name = s.Name,
                Content = s.Content,
                DefaultRole = s.DefaultRole,
                Status = s.Status
            }).FirstOrDefault(f => f.Id == id);
            return find;
        }
        //
        public async Task<PagingList<Role>> GetAllByLimit(string filter = "", int numberOfPageToShow = 15, int page = 0, string sort = null)
        {
            var allUserRoles = _context.UserRoles;
            var qry = _context.Roles.Select(s => new Role() //Where(w => w.Status != (int)EnumStatus.Delete).
            {
                Id = s.Id,
                Name = s.Name,
                Content = s.Content,
                DefaultRole = s.DefaultRole,
                Status = s.Status,
                UpdatedDate = s.UpdatedDate,
                CountUser = allUserRoles.Count(ur => ur.RoleId == s.Id),
                //CountUser = s.UserRoles.Count(),
            }).OrderByDescending(f => f.Id).AsQueryable().AsNoTracking();
            //
            if (!string.IsNullOrWhiteSpace(filter))
            {
                qry = qry.Where(p => p.Name.Contains(filter));
            }
            var objs = await PagingList.CreateAsync(qry, numberOfPageToShow, page, sort, "Id");
            objs.RouteValue = new RouteValueDictionary { { "filter", filter }, { "psize", numberOfPageToShow } };
            return objs;
        }
    }
}
