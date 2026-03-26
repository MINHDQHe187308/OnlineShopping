using ASP.Models.ViewModels.OrderDetails;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public interface OrderDetailRepositoryInterface
    {
        Task<List<OrderDetailItemViewModel>> GetByOrderIdAsync(int orderId);
    }
}
