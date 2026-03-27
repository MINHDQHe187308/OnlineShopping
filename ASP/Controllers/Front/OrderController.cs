using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using ASP.Models.ASPModel;
using ASP.Models.Domains;

namespace ASP.Controllers.Front
{
    public class OrderController : Controller
    {
        private readonly ASPDbContext _context;

        public OrderController(ASPDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> History()
        {
            string userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Challenge();

            var orders = await _context.Orders
                .Where(o => o.UserId == userId)
                .Include(o => o.OrderDetails)
                    .ThenInclude(oi => oi.ProductVariant)
                        .ThenInclude(pv => pv.Product)
                            .ThenInclude(p => p.ProductImages)
                .OrderByDescending(o => o.OrderDate)
                .ToListAsync();

            return View("~/Views/Front/Orders/OrderHistory.cshtml", orders);
        }
        [HttpPost]
        public async Task<IActionResult> UpdateStatus(int orderId, string status)
        {
            var order = await _context.Orders.FindAsync(orderId);

            if (order == null)
                return Json(new { success = false, message = "Không tìm thấy đơn hàng." });

            if (order.Status == "Success" || order.Status == "Canceled")
            {
                return Json(new { success = false, message = "Đơn hàng đã kết thúc, không thể thay đổi trạng thái!" });
            }

            if (order.Status == "Pending")
            {
                if (status == "Success" || status == "Canceled")
                {
                    order.Status = status;

                    _context.Orders.Update(order);
                    await _context.SaveChangesAsync();

                    return Json(new { success = true, message = "Cập nhật trạng thái thành công." });
                }
            }

            return Json(new { success = false, message = "Thao tác không hợp lệ." });
        }
    }
}