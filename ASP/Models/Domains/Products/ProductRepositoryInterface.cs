namespace ASP.Models.Domains
{
    public interface ProductRepositoryInterface
    {
        IEnumerable<Product> GetAllProducts();
        Task<Product?> GetProductByIdAsync(int id);
        Task ImportProductsAsync(List<Product> products);
        IEnumerable<Product> GetAllProducts1();
    }
}