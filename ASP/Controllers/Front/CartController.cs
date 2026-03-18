using Microsoft.AspNetCore.Mvc;
using ASP.Models.Domains;
using System.Threading.Tasks;
using System.Security.Claims;

namespace ASP.Controllers.Front
{
    public class CartController : Controller
    {
        private readonly CartRepositoyInterface _cartRepo;
        private readonly CartItemRepositoryInterface _cartItemRepo;

        public CartController(
            CartRepositoyInterface cartRepo,
            CartItemRepositoryInterface cartItemRepo)
        {
            _cartRepo = cartRepo;
            _cartItemRepo = cartItemRepo;
        }

        private string GetUserId()
        {
            return User.FindFirstValue(ClaimTypes.NameIdentifier);
        }

      
        public async Task<IActionResult> Index()
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return RedirectToAction("Login", "Account");

            var cart = await _cartRepo.GetCartWithItemsAsync(userId);

            if (cart == null)
                return Content("Cart is empty");

            ViewBag.CartItemCount = cart.CartItems?.Sum(ci => ci.Quantity) ?? 0;

            return View("~/Views/Front/Carts/Index.cshtml", cart);
        }

        [HttpPost]
        public async Task<IActionResult> AddToCart([FromBody] AddToCartModel model)
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Unauthorized();

            var cart = await _cartRepo.GetOrCreateCartAsync(userId);

            await _cartItemRepo.AddToCartAsync(
                cart.CartId,
                model.VariantId,
                model.Quantity
            );

            await _cartRepo.SaveChangesAsync();

            return Ok(new
            {
                success = true,
                message = "Đã thêm vào giỏ hàng"
            });
        }

       
        [HttpGet]
        public async Task<IActionResult> GetCartCount()
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Json(0);

            var cart = await _cartRepo.GetCartWithItemsAsync(userId);

            int count = cart?.CartItems?.Sum(i => i.Quantity) ?? 0;

            return Json(count);
        }

       
        [HttpPost]
        public async Task<IActionResult> RemoveItem([FromBody] RemoveItemModel model)
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Unauthorized();

            var cart = await _cartRepo.GetCartWithItemsAsync(userId);

            if (cart == null || !cart.CartItems.Any())
                return BadRequest("Giỏ hàng trống");

            var item = cart.CartItems
                .FirstOrDefault(x => x.CartItemId == model.CartItemId);

            if (item == null)
                return BadRequest("Không tìm thấy sản phẩm");

            _cartItemRepo.Delete(item);

            await _cartRepo.SaveChangesAsync();

            return Ok(new
            {
                success = true,
                message = "Đã xoá sản phẩm"
            });
        }
    }


    public class AddToCartModel
    {
        public int VariantId { get; set; }

        public int Quantity { get; set; } = 1;
    }

    public class RemoveItemModel
    {
        public int CartItemId { get; set; }
    }
}