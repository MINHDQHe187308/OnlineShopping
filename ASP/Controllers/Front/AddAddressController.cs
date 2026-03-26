using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ASP.Models.ViewModels;
using ASP.Models.Domains;
using System.Security.Claims;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Front
{
    public class AddAddressController : Controller
    {
        private readonly ASPDbContext _context;

        public AddAddressController(ASPDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        public async Task<IActionResult> Index()
        {
            string userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Challenge();

            var user = await _context.Users.FirstOrDefaultAsync(u => u.Id == userId);

            var model = new CheckoutViewModel
            {
                Address = new ShippingAddress
                {
                    UserId = userId,
                    FullName = user?.FullName,
                    Phone = user?.PhoneNumber
                }
            };

            return View("~/Views/Front/Checkout/AddAddress.cshtml", model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Add(CheckoutViewModel model)
        {
            

            var newAddress = model.Address;

            if (newAddress.IsDefault)
            {
                var otherAddresses = await _context.ShippingAddresses
                    .Where(a => a.UserId == newAddress.UserId)
                    .ToListAsync();

                foreach (var addr in otherAddresses)
                {
                    addr.IsDefault = false;
                }
            }

            _context.ShippingAddresses.Add(newAddress);
            await _context.SaveChangesAsync();

            TempData["SuccessMessage"] = "Thêm địa chỉ mới thành công!";

            return RedirectToAction("Index", "Checkout");
        }
    }
}