using ASP.Models.ASPModel;
using ASP.Models.ViewModels.OrderDetails;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public class OrderDetailRepository : OrderDetailRepositoryInterface
    {
        private readonly ASPDbContext _context;

        public OrderDetailRepository(ASPDbContext context)
        {
            _context = context;
        }
        public async Task<List<OrderDetailItemViewModel>> GetByOrderIdAsync(int orderId)
        {
            return await _context.OrderDetails
                .Where(od => od.OrderId == orderId)
                .Include(od => od.ProductVariant)
                    .ThenInclude(pv => pv.Product)
                        .ThenInclude(p => p.ProductImages)
                .Select(od => new OrderDetailItemViewModel
                {
                    OrderDetailId = od.OrderDetailId,
                    OrderId = od.OrderId,
                    VariantId = od.VariantId,
                    ProductName = od.ProductVariant.Product.ProductName,
                    VariantName = $"{od.ProductVariant.Size} - {od.ProductVariant.Color}",
                    ProductImage = od.ProductVariant.Product.ProductImages.FirstOrDefault().ImageUrl,
                    Quantity = od.Quantity,
                    UnitPrice = od.UnitPrice,
                    TotalPrice = od.TotalPrice
                })
                .ToListAsync();
        }
    }
}
