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
        public async Task<IActionResult> AddToCart(int variantId)
        {
            int customerId = 2;
            var cart = await _cartRepo.GetOrCreateCartAsync(customerId);        
            await _cartItemRepo.AddToCartAsync(cart.CartId, variantId, 1);      
            await _cartRepo.SaveChangesAsync();
            return RedirectToAction("Index", "Product");
        }
    }
}