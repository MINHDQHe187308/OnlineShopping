using ASP.Models.ASPModel;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
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

        public async Task<IEnumerable<ProductVariant>> GetAllVariantsAsync()
        {
            return await _context.ProductVariants
                .Include(pv => pv.Product)
                .OrderByDescending(pv => pv.VariantId)
                .ToListAsync();
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
                
                // Real-time notification
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
                
                // Real-time notification
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

                // Cache ProductId to send in event
                int productId = variant.ProductId;

                _context.ProductVariants.Remove(variant);
                await _context.SaveChangesAsync();
                
                // Real-time notification
                await _hubContext.Clients.All.SendAsync("VariantDeleted", id, productId);
                
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
