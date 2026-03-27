using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;

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
            try
            {
                return GetProductsFromServer().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lấy dữ liệu từ Server: {ex.Message}");
              
                return new List<Product>();
            }
        }

       
        private async Task<IEnumerable<Product>> GetProductsFromServer()
        {
            string serverIp = "127.0.0.1";   
            int serverPort = 5000;

            using TcpClient client = new TcpClient();
            await client.ConnectAsync(serverIp, serverPort);

            using NetworkStream stream = client.GetStream();

            
            string command = "GET_PRODUCTS";
            byte[] requestBytes = Encoding.UTF8.GetBytes(command);
            await stream.WriteAsync(requestBytes, 0, requestBytes.Length);

            
            byte[] buffer = new byte[65536]; 
            int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);

            if (bytesRead == 0)
                return new List<Product>();

            string responseJson = Encoding.UTF8.GetString(buffer, 0, bytesRead).Trim();

       
            var response = JsonSerializer.Deserialize<JsonResponse>(responseJson, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (response?.success != true || response.products == null)
                return new List<Product>();

           
            var products = new List<Product>();

            foreach (var item in response.products)
            {
                var product = new Product
                {
                    ProductId = item.ProductId,
                    ProductName = item.ProductName,
                    Quantity = item.Quantity,
                   
                };

                
                if (!string.IsNullOrEmpty(item.Category) && item.Category != "N/A")
                {
                    product.Category = new Category { CategoryName = item.Category };
                }

                products.Add(product);
            }

            return products;
        }

        public IEnumerable<Product> GetAllProducts1()
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

        public async Task ImportProductsAsync(List<Product> products)
        {
            await _context.Products.AddRangeAsync(products);
            await _context.SaveChangesAsync();
        }
    }

    
    public class JsonResponse
    {
        public bool success { get; set; }
        public string command { get; set; } = string.Empty;
        public int count { get; set; }
        public List<ProductDto> products { get; set; } = new();
        public DateTime timestamp { get; set; }
    }

    public class ProductDto
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public string MainImage { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }
}