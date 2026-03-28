using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

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

        
        public async Task<IActionResult> Index()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            // 🔹 Lấy cart
            var cart = await _context.Carts
                .Include(c => c.User)
                .Include(c => c.CartItems)
                    .ThenInclude(ci => ci.ProductVariant)
                        .ThenInclude(pv => pv.Product)
                            .ThenInclude(p => p.ProductImages)
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
            {
                return RedirectToAction("Index", "Cart");
            }

         
            var addresses = await _context.ShippingAddresses
                .Where(s => s.UserId == userId)
                .Include(a => a.User) 
                .ToListAsync();

        
            if (!addresses.Any())
            {
                TempData["Error"] = "Vui lòng thêm địa chỉ giao hàng trước!";
                return RedirectToAction("Index", "AddAddress");
            }

            var defaultAddress = addresses.FirstOrDefault(a => a.IsDefault)
                                 ?? addresses.FirstOrDefault();

            var user = await _userManager.GetUserAsync(User);

            var vm = new CheckoutViewModel
            {
                CartItems = cart.CartItems?.ToList() ?? new List<CartItem>(),
                Addresses = addresses,
                Address = defaultAddress,
                user = user,
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * (x.ProductVariant?.Price ?? 0))
            };

            return View("~/Views/Front/Checkout/Index.cshtml", vm);
        }

     
        [HttpPost]
        public async Task<IActionResult> PlaceOrder()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

           
            var cart = await _context.Carts
                .Include(c => c.User)
                .Include(c => c.CartItems)
                    .ThenInclude(ci => ci.ProductVariant)
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (cart == null || !cart.CartItems.Any())
            {
                TempData["Error"] = "Giỏ hàng trống!";
                return RedirectToAction("Index", "Cart");
            }

        
            string formAddress = Request.Form["address"];
            string formCity = Request.Form["city"];
            string formDistrict = Request.Form["district"];
            string formWard = Request.Form["ward"];

            var shippingAddress = await _context.ShippingAddresses
                .Include(a => a.User)
                .FirstOrDefaultAsync(a =>
                    a.UserId == userId &&
                    a.AddressLine == formAddress &&
                    a.City == formCity &&
                    a.District == formDistrict &&
                    a.Ward == formWard
                );

           
            if (shippingAddress == null)
            {
                TempData["Error"] = "Không tìm thấy địa chỉ giao hàng!";
                return RedirectToAction("Index");
            }

          
            var order = new Order
            {
                UserId = userId,
                OrderDate = DateTime.Now,
                CreatedBy = cart.User?.FullName ?? "Unknown",
                ShippingAddress = shippingAddress,
                Status = "Pending",
                TotalAmount = cart.CartItems.Sum(x => x.Quantity * (x.ProductVariant?.Price ?? 0))
            };

            _context.Orders.Add(order);
            await _context.SaveChangesAsync();

            
            foreach (var item in cart.CartItems)
            {
                if (item.ProductVariant == null)
                    continue; 

                var orderDetail = new OrderDetail
                {
                    OrderId = order.OrderId,
                    VariantId = item.VariantId,
                    Quantity = item.Quantity,
                    UnitPrice = item.ProductVariant.Price
                };

                var product = await _context.Products
                    .FirstOrDefaultAsync(p => p.ProductId == item.ProductVariant.ProductId);

                var variant = await _context.ProductVariants
                    .FirstOrDefaultAsync(pv => pv.VariantId == item.VariantId);

                if (variant != null)
                {
                    variant.QuantityVariants -= item.Quantity;
                }

                if (product != null)
                {
                    product.Quantity -= item.Quantity;
                }

                _context.OrderDetails.Add(orderDetail);
            }

            
            _context.CartItems.RemoveRange(cart.CartItems);

            await _context.SaveChangesAsync();

            TempData["SuccessMessage"] = "Đặt hàng thành công!";
            return RedirectToAction("Success");
        }

      
        public IActionResult Success()
        {
            return View("~/Views/Front/Checkout/Success.cshtml");
        }
    }
}