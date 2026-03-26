using System;
using System.Linq;
using System.Threading.Tasks;
using ASP.BaseCommon;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Policies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace ASP.Controllers.Admin
{
    [Authorize]
    [Route("admin/[controller]", Name = "admin.orders")]
    public class OrderController : Controller
    {
        private readonly IAuthorizationService _authService;
        private readonly ASPDbContext _context;
        private readonly OrderRepositoryInterface _orderRepository;

        public OrderController(
            IAuthorizationService authService,
            ASPDbContext context,
            OrderRepositoryInterface orderRepository)
        {
            _authService = authService;
            _context = context;
            _orderRepository = orderRepository;
        }

        [HttpGet]
        public async Task<IActionResult> Index(
            string? status = "",
            DateTime? fromDate = null,
            DateTime? toDate = null,
            int psize = 10,
            int page = 1,
            string sort = "-OrderId")
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPOrdersView);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var statusList = _context.Orders
                .Select(o => o.Status)
                .Distinct()
                .OrderBy(s => s)
                .ToList();

            ViewBag.StatusList = statusList
                .Select(s => new SelectListItem
                {
                    Value = s,
                    Text = s,
                    Selected = !string.IsNullOrEmpty(status) && s == status
                })
                .ToList();

            var list = await _orderRepository.GetAllByFilterAsync(status, fromDate, toDate, psize, page, sort);
            return View("../Admin/Orders/Index", list);
        }

        [HttpGet("details/{id}", Name = "admin.orders.details")]
        public async Task<IActionResult> Details(int id)
        {
            var hasAccess = await _authService.AuthorizeAsync(User, new DocumentAuth(), PolicyOperations.ASPOrdersView);
            if (!hasAccess.Succeeded) return new ForbidResult();

            var order = await _orderRepository.GetByIdAsync(id);
            if (order == null)
            {
                return NotFound();
            }

            return View("../Admin/Orders/Details", order);
        }
    }
}

