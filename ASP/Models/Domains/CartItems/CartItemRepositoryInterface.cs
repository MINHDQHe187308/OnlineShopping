namespace ASP.Models.Domains
{
    public interface CartItemRepositoryInterface
    {
        Task AddToCartAsync(int cartId, int variantId, int quantity = 1);
        Task UpdateQuantityAsync(int cartItemId, int quantity);
        Task ClearCartAsync(int cartId);
        void Delete(CartItem item);
    }
}
