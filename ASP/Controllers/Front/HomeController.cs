using ASP.BaseCommon;
using ASP.Hubs;
using ASP.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using ASP.Models.Admin.ThemeOptions;
using System.Diagnostics;
using ASP.Models.Admin.Roles;
using Spire.Xls;
using System.Reflection;
using Microsoft.EntityFrameworkCore;
using MimeKit;
using System.Drawing;
using System.Data;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Front
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public ThemeOptionRepositoryInterface _themeOption;
        private readonly IHubContext<NotificationHub> _hubContext;
        private readonly ASPDbContext _context;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IAuthorizationService _authService;
        public BaseController _baseController;

        public HomeController(ASPDbContext context, ILogger<HomeController> logger, ThemeOptionRepositoryInterface themeOption, IWebHostEnvironment webHostEnvironment, IAuthorizationService authService, IHubContext<NotificationHub> hubContext, BaseController baseController)
        {
            _context = context;
            _logger = logger;
            _hubContext = hubContext;
            _themeOption = themeOption;
            _webHostEnvironment = webHostEnvironment;
            _authService = authService;
            _baseController = baseController;
        }

        public IActionResult Index()
        {
            return View("../Front/Home/Index");
        }        

        [Route("notification", Name = "front.home.Notify")]
        public async Task<IActionResult> Notify(string message)
        {
            var cLog = _context.Logs.Count();
            await _hubContext.Clients.All.SendAsync("Notify", $"{cLog}");
            return View("../Front/Home/Notification");
        }

        public IActionResult Privacy()
        {
            return View("../Front/Home/Privacy");
        }

        public async Task<IActionResult> Guide()
        {
            var obj = _themeOption.GetThemeOption("_optguide", "home_tab");
            ViewBag.vbGuide = "";
            if (obj != null)
            {
                ViewBag.vbGuide = obj.Value;
            }
            return View("../Front/Home/Guide");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}