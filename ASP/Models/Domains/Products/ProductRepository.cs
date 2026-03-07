using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;

namespace ASP.Models.Domains
{
    public class ProductRepository : ProductRepositoryInterface
    {
        private readonly ASPDbContext _context;

        public ProductRepository(ASPDbContext context)
        {
            _context = context;
        }

        public IEnumerable<Product> GetAllProducts()
        {
            return _context.Products
                .Include(p => p.Category)
                .Include(p => p.ProductImages)
                .Include(p => p.ProductVariants)
                .OrderByDescending(p => p.ProductId)
                .ToList();
        }

        public IEnumerable<Product> GetBestSellingProducts(int take = 8)
        {
            return _context.Products
                .Include(p => p.Category)
                .Include(p => p.ProductImages)
                .Include(p => p.ProductVariants)
                .Where(p => p.IsActive)
                .Take(take)
                .ToList();
        }

        public IEnumerable<Product> GetNewArrivals(int take = 4)
        {
            return _context.Products
                .Include(p => p.Category)
                .Include(p => p.ProductImages)
                .Include(p => p.ProductVariants)
                .Where(p => p.IsActive)
                .OrderByDescending(p => p.ProductId)
                .Take(take)
                .ToList();
        }
    
    public async Task<Product?> GetProductByIdAsync(int id)
        {
            return await _context.Products
                .Include(p => p.Category)
                .Include(p => p.ProductImages)
                .Include(p => p.ProductVariants)
                .FirstOrDefaultAsync(p => p.ProductId == id);
        }

    }
}