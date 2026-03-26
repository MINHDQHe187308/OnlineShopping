using ReflectionIT.Mvc.Paging;
using System;
using System.Threading.Tasks;
using ASP.Models.ViewModels.Orders;
using ASP.Models.ViewModels.OrderDetails;

namespace ASP.Models.Domains
{
    public interface OrderRepositoryInterface
    {
        Task<PagingList<OrderListItem>> GetAllByFilterAsync(
            string? status,
            DateTime? fromDate,
            DateTime? toDate,
            int pageSize,
            int page,
            string? sort);
            
        Task<OrderDetailViewModel?> GetByIdAsync(int orderId);
    }
}
