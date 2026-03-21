using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace ASP.Controllers.Front
{
    [Authorize]
    public class CheckoutController : Controller
    {
        private readonly ASPDbContext _context;

        public CheckoutController(ASPDbContext context)
        {
            _context = context;
        }

        // GET: Checkout page
        public async Task<IActionResult> Index()
        {

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            var cart = await _context.Carts
    .Include(c => c.CartItems)
        .ThenInclude(ci => ci.ProductVariant)
            .ThenInclude(pv => pv.Product)
                .ThenInclude(p => p.ProductImages) // <--- Lấy danh sách ảnh ở đây
    .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
                return RedirectToAction("Index", "Cart");

            var vm = new CheckoutViewModel
            {
                CartItems = cart.CartItems.ToList(),
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * x.ProductVariant.Price)
            };

            return View("~/Views/Front/Checkout/Index.cshtml", vm);
        }

        // POST: Handle checkout
        [HttpPost]
        public async Task<IActionResult> PlaceOrder(CheckoutViewModel model)
        {
            var userId = User.Identity.Name;

            var cart = await _context.Carts
                .Include(c => c.CartItems)
                .ThenInclude(ci => ci.ProductVariant)
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
                return RedirectToAction("Index");

            // 1. Create Order
            var order = new Order
            {
                UserId = userId,
                OrderDate = DateTime.Now,
                Status = "Pending",
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * x.ProductVariant.Price)
            };

            _context.Orders.Add(order);
            await _context.SaveChangesAsync();

            // 2. Create OrderDetails
            foreach (var item in cart.CartItems)
            {
                var orderDetail = new OrderDetail
                {
                    OrderId = order.OrderId,
                    VariantId = item.VariantId,
                    Quantity = item.Quantity,
                    UnitPrice = item.ProductVariant.Price
                };

                _context.OrderDetails.Add(orderDetail);
            }


            _context.CartItems.RemoveRange(cart.CartItems);

            await _context.SaveChangesAsync();

            return RedirectToAction("Success");
        }

        public IActionResult Success()
        {
            return View();
        }
    }
}
