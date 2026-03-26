using ASP.Models.ASPModel;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Routing;
using ReflectionIT.Mvc.Paging;
using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.SignalR;
using ASP.Hubs;

namespace ASP.Models.Domains
{
    public class ProductVariantRepository : ProductVariantRepositoryInterface
    {
        private readonly ASPDbContext _context;
        private readonly IHubContext<AdminHub> _hubContext;

        public ProductVariantRepository(ASPDbContext context, IHubContext<AdminHub> hubContext)
        {
            _context = context;
            _hubContext = hubContext;
        }

        public async Task<PagingList<ProductVariant>> GetAllByFilterAsync(string? searchString, int pageSize, int page, string? sort)
        {
            var query = _context.ProductVariants
                .Include(pv => pv.Product)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(searchString))
            {
                searchString = searchString.Trim().ToLower();
                query = query.Where(pv =>
                    (pv.Product != null && pv.Product.ProductName.ToLower().Contains(searchString)) ||
                    (pv.SKU != null && pv.SKU.ToLower().Contains(searchString)) ||
                    (pv.Size != null && pv.Size.ToLower().Contains(searchString)) ||
                    (pv.Color != null && pv.Color.ToLower().Contains(searchString))
                );
            }

            query = query.OrderByDescending(pv => pv.VariantId);

            var list = await PagingList.CreateAsync(query, pageSize, page, sort, "VariantId");

            list.RouteValue = new RouteValueDictionary
            {
                { "searchString", searchString },
                { "psize", pageSize }
            };

            return list;
        }

        public async Task<ProductVariant?> GetVariantByIdAsync(int id)
        {
            return await _context.ProductVariants
                .Include(pv => pv.Product)
                .FirstOrDefaultAsync(pv => pv.VariantId == id);
        }

        public async Task<bool> CreateVariantAsync(ProductVariant variant)
        {
            try
            {
                _context.ProductVariants.Add(variant);
                await _context.SaveChangesAsync();
                
                await _hubContext.Clients.All.SendAsync("VariantCreated", variant);
                
                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> UpdateVariantAsync(ProductVariant variant)
        {
            try
            {
                _context.ProductVariants.Update(variant);
                await _context.SaveChangesAsync();
                
                await _hubContext.Clients.All.SendAsync("VariantUpdated", variant);
                
                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> DeleteVariantAsync(int id)
        {
            try
            {
                var variant = await _context.ProductVariants.FindAsync(id);
                if (variant == null) return false;

                int productId = variant.ProductId;

                bool isReferenced = await _context.OrderDetails.AnyAsync(od => od.VariantId == id) ||
                                    await _context.CartItems.AnyAsync(ci => ci.VariantId == id);

                if (isReferenced)
                {
                    variant.IsActive = false;
                    _context.ProductVariants.Update(variant);
                    await _context.SaveChangesAsync();
                    
                    await _hubContext.Clients.All.SendAsync("VariantUpdated", variant);
                }
                else
                {
                    _context.ProductVariants.Remove(variant);
                    await _context.SaveChangesAsync();
                    
                    await _hubContext.Clients.All.SendAsync("VariantDeleted", id, productId);
                }
                
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
