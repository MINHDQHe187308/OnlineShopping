using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using ASP.BaseCommon;
using ASP.Policies;
using ASP.SeedData;
using System.ServiceProcess;
using ASP.Models.Admin.ThemeOptions;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.themoptions")]
    public class ThemeOptionController : Controller, ThemeOptionListenerInterface
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        public ThemeOptionRepositoryInterface _themeOption;
        protected string photosPath;
        public ThemeOptionController(IAuthorizationService authService, ASPDbContext context, ThemeOptionRepositoryInterface themeOption)
        {
            _authService = authService;
            _context = context;
            _themeOption = themeOption;
        }

        // GET: ThemeOptionController
        [HttpGet]
        public async Task<IActionResult> Index()
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPThemoptionsView);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            //home tab
            var list = _themeOption.GetAll("home_tab");
            //worker services tab
            var workerServices = _themeOption.GetAll("worker_services_tab");
            ViewBag.workerServices = workerServices;
            //append grid home tab
            var jsonStr = "[]";
            var slides = workerServices.FirstOrDefault(f => f.Name == "_optWorkerService" && f.Value != null);
            if (slides != null)
            {
                jsonStr = slides.Value;
                var checkServive = Newtonsoft.Json.JsonConvert.DeserializeObject<List<WorkerServiceModel>>(jsonStr);
                foreach (var item in checkServive)
                {
                    string serviceName = item.name ?? "xxx";//.NET Joke Service
                    string strStatus = "";
                    if (DoesServiceExist(serviceName))
                    {
                        ServiceController serviceController = new ServiceController(serviceName);
                        if (serviceController.Status.Equals(ServiceControllerStatus.Running) || serviceController.Status.Equals(ServiceControllerStatus.StartPending))
                        {
                            strStatus = "running";
                        }
                        else
                        {
                            strStatus = "stopped";
                        }

                    }
                    item.status = $"{strStatus}";
                }
                jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(checkServive);
            }
            ViewBag.jsonStr = jsonStr;
            //inventory tab
            var inventory = _themeOption.GetAll("inventory_tab");
            ViewBag.inventory = inventory;
            //
            return View("../Admin/ThemeOptions/Index", list);
        }
        //
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Route("Edit", Name = "admin.themoptions.edit")]
        public async Task<IActionResult> Edit(IFormCollection request, List<IFormFile> formFiles)
        {
            #region check access
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPThemoptionsUpdate);
            if (!hasAccess.Succeeded) return new ForbidResult();
            #endregion
            #region validate
            //var b = "";
            //if (string.IsNullOrEmpty(request["_optWorkerService"].ToString()))
            //{
            //}
            //var test = Newtonsoft.Json.JsonConvert.DeserializeObject<List<WorkerServiceModel>>(slides.Value);
            #endregion
            return await _themeOption.UpdateThemeOptionById(this, request, formFiles);
        }
        //
        [HttpGet]
        [Route("RestartWindowsService", Name = "admin.themoptions.restart")]
        public JsonResult RestartWindowsService(string codeValue, string codeType, string nameValue)
        {

            try
            {
                string serviceName = nameValue;//".NET Joke Service";
                if (DoesServiceExist(serviceName))
                {
                    ServiceController serviceController = new ServiceController(serviceName);
                    if (codeType == "stop")
                    {
                        //stop
                        if (serviceController.Status.Equals(ServiceControllerStatus.Running) || serviceController.Status.Equals(ServiceControllerStatus.StartPending))
                        {
                            serviceController.Stop();
                        }
                        serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                    }
                    if (codeType == "start")
                    {
                        //
                        if (serviceController.Status.Equals(ServiceControllerStatus.Stopped) || serviceController.Status.Equals(ServiceControllerStatus.StopPending))
                        {
                            serviceController.Start();

                        }
                        serviceController.WaitForStatus(ServiceControllerStatus.Running);
                    }
                    //
                    return Json(new { result = true, message = "" });
                }
                else
                {
                    return Json(new { result = false, message = $"Lỗi! Service '{serviceName}' không tồn tại." });
                }
            }
            catch (Exception ex)
            {
                return Json(new { result = false, message = $"{ex.Message}" });
            }
        }
        public static bool DoesServiceExist(string serviceName)
        {
            return ServiceController.GetServices().Any(serviceController => serviceController.ServiceName.Equals(serviceName));
        }
        //
        public IActionResult UpdateThemeOptionSuccess()
        {
            TempData["mess-type"] = "success";
            TempData["mess-detail"] = BaseController.BaseMessage("update_success");
            return RedirectToAction(nameof(Index));
        }
        //
        public IActionResult UpdateThemeOptionFails(string message = "")
        {
            TempData["mess-type"] = "error";
            if (string.IsNullOrEmpty(message))
            {
                TempData["mess-detail"] = BaseController.BaseMessage("update_fails");
            }
            else
            {
                TempData["mess-detail"] = $"{message}";
            }
            return RedirectToAction(nameof(Index));
        }
        //
        public IActionResult PageNotFound()
        {
            TempData["mess-type"] = "warning";
            TempData["mess-detail"] = BaseController.BaseMessage("row_fails");
            return RedirectToAction(nameof(Index));
        }
    }
}
