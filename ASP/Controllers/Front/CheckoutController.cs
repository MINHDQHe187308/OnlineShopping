using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using Microsoft.AspNetCore.SignalR;

namespace ASP.Controllers.Front
{
    [Authorize]
    public class CheckoutController : Controller
    {
        private readonly ASPDbContext _context;
        private readonly UserManager<ApplicationUser> _userManager;
        public CheckoutController(ASPDbContext context, UserManager<ApplicationUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        // GET: Checkout page
        public async Task<IActionResult> Index()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            var cart = await _context.Carts           
                .Include(c => c.CartItems)
                .ThenInclude(ci => ci.ProductVariant)
                .ThenInclude(pv => pv.Product)
                .ThenInclude(p => p.ProductImages)
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
                return RedirectToAction("Index", "Cart");

            var shippingAddress = await _context.ShippingAddresses
                .Where(s => s.UserId == userId)
                .ToListAsync();

            var defaultAddress = shippingAddress.FirstOrDefault(s => s.IsDefault)
                                 ?? shippingAddress.FirstOrDefault();

            var user = await _userManager.GetUserAsync(User);

            var vm = new CheckoutViewModel
            {
                CartItems = cart.CartItems.ToList(),
                Addresses = shippingAddress,
                Address = defaultAddress,
                user = user,
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * (x.ProductVariant?.Price ?? 0))
            };

            return View("~/Views/Front/Checkout/Index.cshtml", vm);
        }

        [HttpPost]
        public async Task<IActionResult> PlaceOrder(CheckoutViewModel model)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            var cart = await _context.Carts
                  .Include(c => c.User)
                .Include(c => c.CartItems)
                .ThenInclude(ci => ci.ProductVariant)
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
            {
                throw new Exception("Cart is null");
            }


            string formAddress = Request.Form["address"];
            string formCity = Request.Form["city"];
            string formWard = Request.Form["ward"];
            var shippingAddress = _context.ShippingAddresses
                .FirstOrDefault(a => a.UserId == userId && a.AddressLine == formAddress && a.City == formCity);
            var order = new Order
            {
                UserId = userId,
                OrderDate = DateTime.Now,
                CreatedBy = cart.User.FullName,
                ShippingAddress = shippingAddress,
                Status = "Pending",
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * x.ProductVariant.Price)
            };

            _context.Orders.Add(order);
            await _context.SaveChangesAsync();

            foreach (var item in cart.CartItems)
            {
                var orderDetail = new OrderDetail
                {
                    OrderId = order.OrderId,
                    VariantId = item.VariantId,
                    Quantity = item.Quantity,
                    UnitPrice = item.ProductVariant.Price
                };

                Product product = await _context.Products.FirstOrDefaultAsync(p => p.ProductId == item.ProductVariant.ProductId);
                product.Quantity -= item.Quantity;
                _context.OrderDetails.Add(orderDetail);
            }


            _context.CartItems.RemoveRange(cart.CartItems);

            await _context.SaveChangesAsync();

            return RedirectToAction("Success");
        }

        public IActionResult Success()
        {
            return View("~/Views/Front/Checkout/Success.cshtml");
        }
    }
}
