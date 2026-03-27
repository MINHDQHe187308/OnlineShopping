using ASP.Models.Domains;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using System.Threading.Tasks;

namespace ASP.Controllers.Front
{
    public class CartController : Controller
    {
        private readonly CartRepositoyInterface _cartRepo;
        private readonly CartItemRepositoryInterface _cartItemRepo;
        private readonly ASP.Models.ASPModel.ASPDbContext _context;

        public CartController(
            CartRepositoyInterface cartRepo,
            CartItemRepositoryInterface cartItemRepo,
            ASP.Models.ASPModel.ASPDbContext context)
        {
            _cartRepo = cartRepo;
            _cartItemRepo = cartItemRepo;
            _context = context;
        }

        private string GetUserId()
        {
            return User.FindFirstValue(ClaimTypes.NameIdentifier);
        }

      
        public async Task<IActionResult> Index()
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Redirect("/Login");

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
                return Unauthorized(new { message = "Vui lòng đăng nhập" });

            if (model.Quantity < 1)
                return BadRequest(new { message = "Số lượng phải lớn hơn 0" });

            var variant = await _context.ProductVariants
                .Include(v => v.Product)
                .FirstOrDefaultAsync(v => v.VariantId == model.VariantId);

            if (variant == null)
                return BadRequest(new { message = "Không tìm thấy biến thể sản phẩm" });

            if (!variant.IsActive)
                return BadRequest(new { message = "Biến thể này hiện không còn hoạt động" });

            if (variant.Product == null || !variant.Product.IsActive)
                return BadRequest(new { message = "Sản phẩm này hiện không còn hoạt động" });

            if (variant.QuantityVariants <= 0)
                return BadRequest(new { message = "Biến thể này đã hết hàng" });

            var cart = await _cartRepo.GetOrCreateCartAsync(userId);

            var existingItem = await _context.CartItems
                .FirstOrDefaultAsync(ci => ci.CartId == cart.CartId && ci.VariantId == model.VariantId);

            var currentQtyInCart = existingItem?.Quantity ?? 0;

            if (currentQtyInCart + model.Quantity > variant.QuantityVariants)
            {
                return BadRequest(new
                {
                    message = $"Số lượng vượt quá tồn kho. Chỉ còn {variant.QuantityVariants} sản phẩm."
                });
            }

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

        [HttpPost]
        public async Task<IActionResult> UpdateItem([FromBody] UpdateCartItemModel model)
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { message = "Vui lòng đăng nhập" });

            if (model.Quantity < 1)
                return BadRequest(new { message = "Số lượng phải lớn hơn 0" });

            var cart = await _cartRepo.GetCartWithItemsAsync(userId);

            if (cart == null || !cart.CartItems.Any())
                return BadRequest(new { message = "Giỏ hàng trống" });

            var item = cart.CartItems.FirstOrDefault(x => x.CartItemId == model.CartItemId);

            if (item == null)
                return BadRequest(new { message = "Không tìm thấy sản phẩm trong giỏ" });

            await _cartItemRepo.UpdateQuantityAsync(model.CartItemId, model.Quantity);
            await _cartRepo.SaveChangesAsync();

            return Ok(new
            {
                success = true,
                message = "Cập nhật số lượng thành công"
            });
        }

        [HttpPost]
        public async Task<IActionResult> ClearCart()
        {
            var userId = GetUserId();

            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { message = "Vui lòng đăng nhập" });

            var cart = await _cartRepo.GetCartWithItemsAsync(userId);

            if (cart == null || !cart.CartItems.Any())
                return BadRequest(new { message = "Giỏ hàng trống" });

            await _cartItemRepo.ClearCartAsync(cart.CartId);
            await _cartRepo.SaveChangesAsync();

            return Ok(new
            {
                success = true,
                message = "Đã xóa toàn bộ giỏ hàng"
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
    public class UpdateCartItemModel
    {
        public int CartItemId { get; set; }
        public int Quantity { get; set; }
    }
}