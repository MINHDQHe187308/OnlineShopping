namespace ASP.Models.Domains
{
    public interface CartItemRepositoryInterface
    {
        Task AddToCartAsync(int cartId, int variantId, int quantity = 1);
        void Delete(CartItem item);
    }
}
