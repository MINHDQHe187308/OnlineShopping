using ASP.Models.Domains;
using Microsoft.AspNetCore.SignalR;

namespace ASP.Hubs
{
    public class AddressHub : Hub
    {
        public async Task SendMessage(ShippingAddress addressData)
        {
            await Clients.All.SendAsync("AddressMessage", addressData);
        }
        public async Task AddMessage(ShippingAddress addressData)
        {
            await Clients.All.SendAsync("AddAddressMessage", addressData);
        }
    }
}
