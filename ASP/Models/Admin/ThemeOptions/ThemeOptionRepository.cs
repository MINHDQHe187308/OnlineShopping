using ASP.BaseCommon;
using ASP.ConfigCommon;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using Newtonsoft.Json;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Transactions;
using ASP.Models.Admin.Logs;
using Microsoft.Extensions.Logging;
using System.Reflection;
using Microsoft.AspNetCore.SignalR;
using ASP.Hubs;
using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;

namespace ASP.Models.Admin.ThemeOptions
{
    public class ThemeOptionRepository : ThemeOptionRepositoryInterface
    {
        private readonly ILogger<ThemeOptionRepository> _logger;
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;
        private object userManager;
        private readonly IHubContext<NotificationHub> _hubContext;
        public ThemeOptionRepository(ILogger<ThemeOptionRepository> logger, ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log, IHubContext<NotificationHub> hubContext)
        {
            _logger = logger;
            _context = context;
            this.env = env;
            this.log = log;
            photosPath = this.env.WebRootPath + "/assets/themeoptions";
            _hubContext = hubContext;
        }

        public async Task<IActionResult> UpdateThemeOptionById(ThemeOptionListenerInterface listener, IFormCollection request, List<IFormFile> formFiles)
        {

            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                try
                {
                    #region log
                    string logContent = "";
                    //
                    #endregion

                    #region image file
                    foreach (var itemFile in request.Files)
                    {
                        // full path to file in temp location
                        if (!Directory.Exists(photosPath))
                        {
                            Directory.CreateDirectory(photosPath);
                        }
                        if (itemFile != null)
                        {

                            if (itemFile.Name == "_optLogo")
                            {
                                #region logo
                                var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optLogo");
                                string image_name = null;
                                if (itemFile != null)
                                {
                                    // delete old file
                                    if (find.Value != null)
                                    {
                                        if (File.Exists(Path.Combine(photosPath, find.Value)))
                                        {
                                            File.Delete(Path.Combine(photosPath, find.Value));
                                        }
                                    }
                                    // add new file
                                    image_name = "themeoptions-" + DateTime.Now.Ticks + itemFile.FileName;
                                    var filePath = Path.Combine(photosPath, image_name);
                                    var extension = Path.GetExtension(filePath);
                                    //
                                    using (var stream = File.Create(filePath))
                                    {
                                        itemFile.CopyTo(stream);
                                    }
                                    logContent += EnumTypeLog.SetLogLine("Logo", find.Value, image_name);
                                    find.Value = image_name;
                                }
                                else
                                {
                                    if (request["rmavatar"].First() == "remove")
                                    {
                                        // delete old file
                                        if (find.Value != null)
                                        {
                                            if (File.Exists(Path.Combine(photosPath, find.Value)))
                                            {
                                                File.Delete(Path.Combine(photosPath, find.Value));
                                            }
                                        }
                                        // remove avatar
                                        logContent += EnumTypeLog.SetLogLine("Logo", find.Value, null);
                                        find.Value = null;
                                    }
                                }
                                #endregion

                            }// endif;
                            if (itemFile.Name == "_optFavicon")
                            {
                                #region logo Favicon
                                var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optFavicon");
                                string image_name = null;
                                if (itemFile != null)
                                {
                                    // delete old file
                                    if (find.Value != null)
                                    {
                                        if (File.Exists(Path.Combine(photosPath, find.Value)))
                                        {
                                            File.Delete(Path.Combine(photosPath, find.Value));
                                        }
                                    }
                                    // add new file
                                    image_name = "themeoptions-" + DateTime.Now.Ticks + itemFile.FileName;
                                    var filePath = Path.Combine(photosPath, image_name);
                                    var extension = Path.GetExtension(filePath);
                                    //
                                    using (var stream = File.Create(filePath))
                                    {
                                        itemFile.CopyTo(stream);
                                    }
                                    logContent += EnumTypeLog.SetLogLine("_optFavicon", find.Value, image_name);
                                    find.Value = image_name;
                                }
                                else
                                {
                                    if (request["rmavatar2"].First() == "remove")
                                    {
                                        // delete old file
                                        if (find.Value != null)
                                        {
                                            if (File.Exists(Path.Combine(photosPath, find.Value)))
                                            {
                                                File.Delete(Path.Combine(photosPath, find.Value));
                                            }
                                        }
                                        // remove avatar
                                        logContent += EnumTypeLog.SetLogLine("_optFavicon", find.Value, null);
                                        find.Value = null;
                                    }
                                }
                                #endregion

                            }// endif;
                        }
                    }
                    #endregion
                    //
                    #region other info

                    foreach (var item in request.Keys)
                    {
                        if (item == "_optLogo" && request["_optLogo"].Any() && request["rmavatar"].First() == "remove")
                        {
                            // image logo == null: delete file
                            // delete old file
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optLogo");
                            if (find.Value != null)
                            {
                                if (File.Exists(Path.Combine(photosPath, find.Value)))
                                {
                                    File.Delete(Path.Combine(photosPath, find.Value));
                                }
                            }
                            // remove avatar
                            logContent += EnumTypeLog.SetLogLine("Logo", find.Value, null);
                            find.Value = null;
                        }
                        else if (item == "_optFavicon" && request["_optFavicon"].Any() && request["rmavatar2"].First() == "remove")
                        {
                            // image logo == null: delete file
                            // delete old file
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optFavicon");
                            if (find.Value != null)
                            {
                                if (File.Exists(Path.Combine(photosPath, find.Value)))
                                {
                                    File.Delete(Path.Combine(photosPath, find.Value));
                                }
                            }
                            // remove avatar
                            logContent += EnumTypeLog.SetLogLine("_optFavicon", find.Value, null);
                            find.Value = null;
                        }
                        else if (item == "_optTrialText")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optTrialText");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Tiêu đề Trial", find.Value, request["_optTrialText"].First());
                                find.Value = request["_optTrialText"].First();
                            }
                        }
                        else if (item == "_optTimeSlide")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optTimeSlide");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Thời gian chuyển slide(Millisecond)", find.Value, request["_optTimeSlide"].First());
                                find.Value = request["_optTimeSlide"].First();
                            }
                        }
                        else if (item == "_optFooterContent")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optFooterContent");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Footer", find.Value, request["_optFooterContent"].First());
                                find.Value = request["_optFooterContent"].First();
                            }
                        }
                        else if (item == "_optguide")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optguide");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Hướng dẫn", find.Value, request["_optguide"].First());
                                find.Value = request["_optguide"].First();
                            }
                        }
                        else if (item == "_optShowAllInventory")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optShowAllInventory");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Hiển thị thông tin Sản lượng", find.Value, request["_optShowAllInventory"].First());
                                find.Value = request["_optShowAllInventory"].First();
                            }
                        }
                        else if (item == "_optWorkerService")
                        {
                            if (!string.IsNullOrEmpty(request["_optWorkerService"].ToString()))
                            {
                                var chk = JsonConvert.DeserializeObject<List<WorkerServiceModel>>(request["_optWorkerService"].First());
                                if (chk != null)
                                {
                                    if (chk.Any(f => f.time_start == "" || f.time_start == null || f.time_repeat == ""))
                                    {
                                        // error
                                        scope.Dispose();
                                        return listener.UpdateThemeOptionFails("Thời gian bắt đầu, Thời gian lặp(số phút) không để trống.");
                                    }
                                }
                                //
                            }
                            //
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optWorkerService");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Worker Service", find.Value, request["_optWorkerService"].First());
                                find.Value = request["_optWorkerService"].First();
                                find.Language = "en-US"; // default
                            }
                        }
                        else if (item == "_optTotalWorkingDay")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optTotalWorkingDay");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Tham số Worker Service _optTotalWorkingDay", find.Value, request["_optTotalWorkingDay"].First());
                                find.Value = request["_optTotalWorkingDay"].First();
                                find.Language = "en-US"; // default
                            }
                        }
                        else if (item == "_optTotalRequire30")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optTotalRequire30");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Tham số Worker Service _optTotalRequire30", find.Value, request["_optTotalRequire30"].First());
                                find.Value = request["_optTotalRequire30"].First();
                                find.Language = "en-US"; // default
                            }
                        }
                        else if (item == "_optTotalRequire120")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optTotalRequire120");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Tham số Worker Service _optTotalRequire120", find.Value, request["_optTotalRequire120"].First());
                                find.Value = request["_optTotalRequire120"].First();
                                find.Language = "en-US"; // default
                            }
                        }
                        else if (item == "_optNumKeepLog")
                        {
                            var find = _context.ThemeOptions.FirstOrDefault(f => f.Name == "_optNumKeepLog");
                            if (find != null)
                            {
                                logContent += EnumTypeLog.SetLogLine("Tham số Worker Service _optNumKeepLog", find.Value, request["_optNumKeepLog"].First());
                                find.Value = request["_optNumKeepLog"].First();
                                find.Language = "en-US"; // default
                            }
                        }
                        //
                    }
                    #endregion

                    #region logs
                    if (!string.IsNullOrEmpty(logContent))
                    {
                        logContent = EnumTypeLog.SetLogTitle("Update " + EnumTypeLog.APP_LOG_THEMEOPTION + " ThemeOptions:") + logContent;
                        log.CreateLog(EnumTypeLog.APP_LOG_THEMEOPTION, logContent);
                    }
                    #endregion
                    //_context.SaveChanges();
                    await _context.SaveChangesAsync();
                    scope.Complete();
                    await _hubContext.Clients.All.SendAsync("Notify", $"Inventory");
                    return listener.UpdateThemeOptionSuccess();
                }
                catch (DbException ex)
                {
                    scope.Dispose();
                    _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                    return listener.UpdateThemeOptionFails("");
                }
            }
        }

        public List<ThemeOption> GetAll(string typeData)
        {
            var objs = _context.ThemeOptions.Where(w => w.TypeData == typeData).OrderBy(o => o.CreatedDate).ToList();
            return objs;
        }

        public ThemeOption GetThemeOption(string strName, string typeData)
        {
            var find = _context.ThemeOptions.FirstOrDefault(w => w.Name == strName && w.TypeData == typeData);
            return find;
        }
    }
}
