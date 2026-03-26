using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SkiaSharp;
using System.Security.Claims;

namespace ASP.Controllers.Front
{
    public class ShippingController : Controller
    {
        private readonly ASPDbContext _context;

        public ShippingController(ASPDbContext context)
        {
            _context = context;
        }
        public async Task<IActionResult> Index(int id)
        {
            string userId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            if (string.IsNullOrEmpty(userId)) return Challenge();

            var viewModel = new CheckoutViewModel
            {
                Addresses = new List<ShippingAddress>(),
                CartItems = new List<CartItem>()
            };

            viewModel.Addresses = await _context.ShippingAddresses
                .Where(x => x.UserId == userId)
                .ToListAsync() ?? new List<ShippingAddress>();

            viewModel.user = await _context.Users.FirstOrDefaultAsync(u => u.Id == userId);

            viewModel.CartItems = await _context.CartItems
                .Include(c => c.ProductVariant)
                .Where(c => c.Cart.UserId == userId)
                .ToListAsync() ?? new List<CartItem>();

            viewModel.TotalAmount = viewModel.CartItems
                .Where(x => x.ProductVariant != null) 
                .Sum(x => x.Quantity * (x.ProductVariant?.Price ?? 0));

            viewModel.Address = viewModel.Addresses.FirstOrDefault(a => a.AddressId ==id)
                                ?? viewModel.Addresses.FirstOrDefault();

            return View("~/Views/Front/Checkout/UpdateAddress.cshtml", viewModel);
        }
        
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Update(CheckoutViewModel model)
        {
            

            var existing = await _context.ShippingAddresses
                .FirstOrDefaultAsync(x => x.AddressId == model.Address.AddressId);

            if (existing == null) return NotFound();

            if (model.Address.IsDefault)
            {
                var otherAddresses = await _context.ShippingAddresses
                    .Where(a => a.UserId == existing.UserId && a.AddressId != existing.AddressId)
                    .ToListAsync();
                foreach (var addr in otherAddresses) addr.IsDefault = false;
            }

            existing.FullName = model.Address.FullName;
            existing.Phone = model.Address.Phone;
            existing.City = model.Address.City;
            existing.District = model.Address.District;
            existing.Ward = model.Address.Ward;
            existing.AddressLine = model.Address.AddressLine;
            existing.IsDefault = model.Address.IsDefault;

            _context.ShippingAddresses.Update(existing);
            await _context.SaveChangesAsync();
            TempData["SuccessMessage"] = "Cập nhật địa chỉ thành công!";
            return RedirectToAction("Index", "Checkout", new { id = existing.AddressId });
        }
    }
}
