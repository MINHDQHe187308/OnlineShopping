using ASP.Models.ASPModel;
using Microsoft.EntityFrameworkCore;

namespace ASP.Models.Domains
{
    
        public class ProductImageRepository : ProductImageRepositoryInterface
        {
        private readonly ASPDbContext _context;
        public ProductImageRepository(ASPDbContext context) => _context = context;

        //lấy list image theo product
        public List<ProductImage> GetImagesByProductId(int productId)
        {
            return _context.ProductImages
                .Where(x => x.ProductId == productId)
                .ToList();
        }

        // lấy image theo id
        public ProductImage? GetImageById(int id)
        {
            return _context.ProductImages.FirstOrDefault(x => x.ProductImageId == id);
        }

        // check main image
        public bool HasMainImage(int productId)
        {
            return _context.ProductImages
                .Any(x => x.ProductId == productId && x.IsMain);
        }

        // check product có image
        public bool HasAnyImage(int productId)
        {
            return _context.ProductImages
                .Any(x => x.ProductId == productId);
        }

        // thêm image
        public async Task AddImageAsync(ProductImage image)
        {
            await _context.ProductImages.AddAsync(image);
            await _context.SaveChangesAsync();
        }

        // xóa image
        public void DeleteImage(ProductImage image)
        {
            _context.ProductImages.Remove(image);
            _context.SaveChanges();
        }

        // set main image
        public void SetMainImage(ProductImage img)
        {
            var images = _context.ProductImages
                .Where(x => x.ProductId == img.ProductId)
                .ToList();

            foreach (var item in images)
                item.IsMain = false;

            img.IsMain = true;

            _context.SaveChanges();
        }

        // lấy product
        public Product? GetProduct(int productId)
        {
            return _context.Products.Include(p => p.ProductVariants).FirstOrDefault(x => x.ProductId == productId);
        }

        // update product
        public void UpdateProduct(Product product)
        {
            _context.Products.Update(product);
            _context.SaveChanges();
        }
    }
    }

