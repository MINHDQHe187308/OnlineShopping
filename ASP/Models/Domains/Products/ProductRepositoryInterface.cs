namespace ASP.Models.Domains
{
    public interface ProductRepositoryInterface
    {
        IEnumerable<Product> GetAllProducts();
        Task<Product?> GetProductByIdAsync(int id);
    }
}