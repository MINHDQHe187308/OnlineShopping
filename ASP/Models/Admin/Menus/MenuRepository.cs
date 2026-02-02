using ASP.BaseCommon;
using Microsoft.AspNetCore.Mvc;
using ReflectionIT.Mvc.Paging;
using System.Data.Common;
using System.Transactions;
using ASP.Models.Admin.Logs;
using System.Reflection;
using ASP.Models.ASPModel;
using ASP.Models;
using Microsoft.AspNetCore.Routing;

// alias để phân biệt
using AdminMenu = ASP.Models.Admin.Menus.Menu;
using DbMenu = ASP.Models.Menu;

namespace ASP.Models.Admin.Menus
{
    public class MenuRepository : MenuRepositoryInterface
    {
        private readonly ILogger<MenuRepository> _logger;
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;
        private object userManager;

        public MenuRepository(ILogger<MenuRepository> logger, ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log)
        {
            _logger = logger;
            _context = context;
            this.env = env;
            this.log = log;
            photosPath = this.env.WebRootPath + "/assets/menus";
        }

        public async Task<IActionResult> CreateMenu(MenuListenerInterface listener, AdminMenu request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    request.Name = !string.IsNullOrEmpty(request.Name) ? request.Name.Trim() : null;
                    request.Description = !string.IsNullOrEmpty(request.Description) ? request.Description.Trim() : null;
                    request.Content = request.out_menu;

                    // map sang DbMenu để lưu
                    var dbMenu = new DbMenu
                    {
                        Name = request.Name,
                        Description = request.Description,
                        Content = request.Content,
                        Language = request.Language,
                        CreatedDate = DateTime.Now,
                        UpdatedDate = DateTime.Now
                    };

                    _context.Add(dbMenu);
                    await _context.SaveChangesAsync();

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Create " + EnumTypeLog.APP_LOG_MENU + " Menu ID:" + dbMenu.Id);
                    logContent += EnumTypeLog.SetLogLine("Name", null, request.Name);
                    logContent += EnumTypeLog.SetLogLine("Description", null, request.Description);
                    logContent += EnumTypeLog.SetLogLine("Language", null, request.Language);
                    logContent += EnumTypeLog.SetLogLine("Content", null, request.out_menu);
                    if (!string.IsNullOrEmpty(logContent))
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_MENU, logContent);
                    }
                    #endregion

                    scope.Complete();
                    return await listener.CreateMenuSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}",
                        MethodBase.GetCurrentMethod().DeclaringType,
                        MethodBase.GetCurrentMethod().Name,
                        ex.Message);
                    return await listener.CreateMenuFails(request);
                }
            }
        }

        public AdminMenu GetMenuById(int id)
        {
            var dbMenu = _context.Menus.FirstOrDefault(f => f.Id == id);
            if (dbMenu == null) return null;

            return new AdminMenu
            {
                ID = dbMenu.Id,
                Name = dbMenu.Name,
                Description = dbMenu.Description,
                Language = dbMenu.Language,
                Content = dbMenu.Content,
                CreatedDate = dbMenu.CreatedDate,
                UpdatedDate = dbMenu.UpdatedDate
            };
        }

        public async Task<IActionResult> UpdateMenuById(int id, MenuListenerInterface listener, AdminMenu request)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    var dbMenu = await _context.Menus.FindAsync(id);
                    if (dbMenu == null)
                    {
                        return await listener.PageNotFound();
                    }

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Update " + EnumTypeLog.APP_LOG_MENU + " Menu ID: " + dbMenu.Id);
                    logContent += EnumTypeLog.SetLogLine("Name", dbMenu.Name, request.Name);
                    logContent += EnumTypeLog.SetLogLine("Description", dbMenu.Description, request.Description);
                    logContent += EnumTypeLog.SetLogLine("Language", dbMenu.Language, request.Language);
                    logContent += EnumTypeLog.SetLogLine("Content", dbMenu.Content, request.out_menu);
                    #endregion

                    dbMenu.Name = !string.IsNullOrEmpty(request.Name) ? request.Name.Trim() : null;
                    dbMenu.Description = !string.IsNullOrEmpty(request.Description) ? request.Description.Trim() : null;
                    dbMenu.Content = request.out_menu;
                    dbMenu.UpdatedDate = DateTime.Now;

                    if (!string.IsNullOrEmpty(logContent))
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_MENU, logContent);
                    }

                    _context.Menus.Update(dbMenu);
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return await listener.UpdateMenuSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType,
                        MethodBase.GetCurrentMethod().Name, ex.Message);
                    return await listener.DeleteMenuFails();
                }
            }
        }

        public async Task<IActionResult> RemoveMenuById(int id, MenuListenerInterface listener)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    var dbMenu = _context.Menus.FirstOrDefault(f => f.Id == id);
                    if (dbMenu == null)
                    {
                        return await listener.PageNotFound();
                    }

                    #region log
                    var logContent = EnumTypeLog.SetLogTitle("Delete " + EnumTypeLog.APP_LOG_MENU + " Menu");
                    logContent += EnumTypeLog.SetLogLine("ID", null, dbMenu.Id.ToString());
                    logContent += EnumTypeLog.SetLogLine("Name", null, dbMenu.Name);
                    logContent += EnumTypeLog.SetLogLine("Description", null, dbMenu.Description);
                    logContent += EnumTypeLog.SetLogLine("Language", null, dbMenu.Language);
                    logContent += EnumTypeLog.SetLogLine("Content", null, dbMenu.Content);
                    if (!string.IsNullOrEmpty(logContent))
                    {
                        log.CreateLog(EnumTypeLog.APP_LOG_MENU, logContent);
                    }
                    #endregion

                    _context.Remove(dbMenu);
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    return await listener.DeleteMenuSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType,
                        MethodBase.GetCurrentMethod().Name, ex.Message);
                    return await listener.DeleteMenuFails();
                }
            }
        }

        public async Task<PagingList<AdminMenu>> GetAllByLimit(string filter = "", int numberOfPageToShow = 10, int page = 0, string sort = null)
        {
            var qry = (from m in _context.Menus
                       select new AdminMenu
                       {
                           ID = m.Id,
                           Name = m.Name,
                           Description = m.Description,
                           Language = m.Language,
                           Content = m.Content,
                           CreatedDate = m.CreatedDate,
                           UpdatedDate = m.UpdatedDate
                       }).OrderByDescending(f => f.ID).AsQueryable();

            if (!string.IsNullOrWhiteSpace(filter))
            {
                qry = qry.Where(p => p.Name.Contains(filter));
            }

            var objs = await PagingList.CreateAsync(qry, numberOfPageToShow, page, sort, "ID");
            objs.RouteValue = new RouteValueDictionary { { "filter", filter }, { "psize", numberOfPageToShow } };
            return objs;
        }
    }
}
