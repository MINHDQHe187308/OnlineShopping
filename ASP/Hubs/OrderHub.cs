using Microsoft.AspNetCore.SignalR;

namespace ASP.Hubs
{
    public class OrderHub : Hub
    {
        public async Task UpdateOrderStatus(string orderUid, int newStatus)
        {
            await Clients.All.SendAsync("OrderStatusUpdated", orderUid, newStatus);
        }
    }
}