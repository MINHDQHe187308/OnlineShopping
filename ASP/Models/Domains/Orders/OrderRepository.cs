using ASP.Models.ASPModel;
using ASP.Models.ViewModels.Orders;
using ASP.Models.ViewModels.OrderDetails;
using Microsoft.AspNetCore.Routing;
using Microsoft.EntityFrameworkCore;
using ReflectionIT.Mvc.Paging;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public class OrderRepository : OrderRepositoryInterface
    {
        private readonly ASPDbContext _context;
        private readonly OrderDetailRepositoryInterface _orderDetailRepository;

        public OrderRepository(ASPDbContext context, OrderDetailRepositoryInterface orderDetailRepository)
        {
            _context = context;
            _orderDetailRepository = orderDetailRepository;
        }

        public async Task<PagingList<OrderListItem>> GetAllByFilterAsync(
            string? status,
            DateTime? fromDate,
            DateTime? toDate,
            int pageSize,
            int page,
            string? sort)
        {
            var query = _context.Orders
                .Include(o => o.User)
                .Select(o => new OrderListItem
                {
                    OrderId = o.OrderId,
                    UserName = o.User != null ? o.User.UserName : null,
                    OrderDate = o.OrderDate,
                    TotalAmount = o.TotalAmount,
                    Status = o.Status
                })
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(status))
            {
                query = query.Where(o => o.Status == status);
            }

            if (fromDate.HasValue)
            {
                var from = fromDate.Value.Date;
                query = query.Where(o => o.OrderDate >= from);
            }

            if (toDate.HasValue)
            {
                var to = toDate.Value.Date.AddDays(1).AddTicks(-1);
                query = query.Where(o => o.OrderDate <= to);
            }

            query = query.OrderByDescending(o => o.OrderId);

            var list = await PagingList.CreateAsync(query, pageSize, page, sort, "OrderId");

            list.RouteValue = new RouteValueDictionary
            {
                { "status", status },
                { "fromDate", fromDate?.ToString("yyyy-MM-dd") },
                { "toDate", toDate?.ToString("yyyy-MM-dd") },
                { "psize", pageSize }
            };

            return list;
        }

        public async Task<OrderDetailViewModel?> GetByIdAsync(int orderId)
        {
            var order = await _context.Orders
                .Include(o => o.User)
                .Include(o => o.ShippingAddress)
                .FirstOrDefaultAsync(o => o.OrderId == orderId);

            if (order == null) return null;

            var orderDetails = await _orderDetailRepository.GetByOrderIdAsync(orderId);

            var viewModel = new OrderDetailViewModel
            {
                OrderId = order.OrderId,
                UserName = order.User?.UserName ?? string.Empty,
                UserEmail = order.User?.Email ?? string.Empty,
                UserPhone = order.User?.PhoneNumber ?? string.Empty,
                OrderDate = order.OrderDate,
                TotalAmount = order.TotalAmount,
                Status = order.Status,
                CreatedBy = order.CreatedBy,
                ShippingAddress = order.ShippingAddress?.AddressLine,
                ShippingCity = order.ShippingAddress?.City,
                ShippingCountry = "Việt Nam", // Default country as not in model
                ShippingPostalCode = order.ShippingAddress?.Ward, // Using Ward as postal code alternative
                OrderDetails = orderDetails
            };

            return viewModel;
        }
    }
}
