using ASP.Models.ASPModel;
using Microsoft.EntityFrameworkCore;
using ASP.Models.Domains;
namespace ASP.Models.Domains
{
    public class CartItemRepository : CartItemRepositoryInterface
    {
        private readonly ASPDbContext _context;

        public CartItemRepository(ASPDbContext context)
        {
            _context = context;
        }

        public async Task AddToCartAsync(int cartId, int variantId, int quantity = 1)
        {
            var existingItem = await _context.CartItems
                .FirstOrDefaultAsync(ci => ci.CartId == cartId && ci.VariantId == variantId);

            if (existingItem != null)
            {
                existingItem.Quantity += quantity;
            }
            else
            {
                var item = new CartItem
                {
                    CartId = cartId,
                    VariantId = variantId,
                    Quantity = quantity
                };

                _context.CartItems.Add(item);
            }
        }
        public void Delete(CartItem item)
        {
            _context.CartItems.Remove(item);
        }
        public async Task UpdateQuantityAsync(int cartItemId, int quantity)
        {
            var item = await _context.CartItems.FirstOrDefaultAsync(x => x.CartItemId == cartItemId);
            if (item != null)
            {
                item.Quantity = quantity;
            }
        }

        public async Task ClearCartAsync(int cartId)
        {
            var items = await _context.CartItems
                .Where(x => x.CartId == cartId)
                .ToListAsync();

            _context.CartItems.RemoveRange(items);
        }
    }


}