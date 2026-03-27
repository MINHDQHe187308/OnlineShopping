namespace ASP.Models.Domains
{
    public interface ProductRepositoryInterface
    {
        IEnumerable<Product> GetAllProducts();
        Task<Product?> GetProductByIdAsync(int id);
        Task ImportProductsAsync(List<Product> products);
        IEnumerable<Product> GetAllProducts1();
        IQueryable<Product> QueryProducts();
        Task<List<Product>> GetRelatedProductsAsync(int productId, int categoryId, int take = 4);
    }
}