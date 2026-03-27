using ASP.Hubs;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace ASP.Controllers.Front
{
    public class AddAddressController : Controller
    {
        private readonly ASPDbContext _context;
        private readonly IHubContext<AddressHub> _hubContext;

        public AddAddressController(ASPDbContext context, IHubContext<AddressHub> hubContext)
        {
            _context = context;
            _hubContext = hubContext;
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
            var addressToSync = new
            {
                addressId = newAddress.AddressId,
                fullName = newAddress.FullName,
                phone = newAddress.Phone,
                addressLine = newAddress.AddressLine,
                city = newAddress.City,
                district = newAddress.District,
                ward = newAddress.Ward,
                isDefault = newAddress.IsDefault
            };

            await _hubContext.Clients.All.SendAsync("AddressMessage", addressToSync);

            return RedirectToAction("Index", "Checkout");
        }
    }
}