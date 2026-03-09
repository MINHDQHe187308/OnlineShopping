using Microsoft.AspNetCore.Mvc;
using ASP.Models.Domains;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Razor.Language.Intermediate;

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

        public async Task<IActionResult> Index()
        {
            int customerId = 2;
            var cart = await _cartRepo.GetCartWithItemsAsync(customerId);

            if (cart == null)
            {
                return RedirectToAction("Index", "Product");
            }

            ViewBag.CartItemCount = cart.CartItems?.Sum(ci => ci.Quantity) ?? 0;
            return View("~/Views/Front/Carts/Index.cshtml", cart);
        }
        //public async Task<IActionResult> AddToCart(int variantId)
        //{
        //    int customerId = 2;
        //    var cart = await _cartRepo.GetOrCreateCartAsync(customerId);
        //    await _cartItemRepo.AddToCartAsync(cart.CartId, variantId, 1);
        //    await _cartRepo.SaveChangesAsync();
        //    return RedirectToAction("Index", "Product");
        //}
        [HttpPost]
        public async Task<IActionResult> AddToCart([FromBody] AddToCartModel model)
        {
            int customerId = 2; 

            var cart = await _cartRepo.GetOrCreateCartAsync(customerId);
            await _cartItemRepo.AddToCartAsync(cart.CartId, model.VariantId, model.Quantity);
            await _cartRepo.SaveChangesAsync();

         
            return Ok(new { success = true, message = "Đã thêm vào giỏ hàng" });
        }

        public class AddToCartModel
        {
            public int VariantId { get; set; }
            public int Quantity { get; set; } = 1;
        }

       
        [HttpGet]
        public async Task<IActionResult> GetCartCount()
        {
            int customerId = 2;
            var cart = await _cartRepo.GetCartWithItemsAsync(customerId);
            int count = cart?.CartItems?.Sum(i => i.Quantity) ?? 0;
            return Json(count);
        }
        [HttpPost]
        public async Task<IActionResult> RemoveItem([FromBody] RemoveItemModel model)
        {
            int customerId = 2;

            var cart = await _cartRepo.GetCartWithItemsAsync(customerId);
            if (cart == null || !cart.CartItems.Any())
            {
                return BadRequest("Giỏ hàng không tồn tại hoặc trống");
            }

            var itemToRemove = cart.CartItems.FirstOrDefault(ci => ci.CartItemId == model.CartItemId);
            if (itemToRemove == null)
            {
                return BadRequest("Không tìm thấy sản phẩm trong giỏ hàng");
            }

            _cartItemRepo.Delete(itemToRemove);
            await _cartRepo.SaveChangesAsync();

            return Ok(new { success = true, message = "Đã xóa sản phẩm khỏi giỏ hàng" });
        }
    }
}
public class RemoveItemModel
{
    public int CartItemId { get; set; }
}