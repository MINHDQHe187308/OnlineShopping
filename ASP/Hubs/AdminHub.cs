using Microsoft.AspNetCore.SignalR;
using System.Threading.Tasks;

namespace ASP.Hubs
{
    public class AdminHub : Hub
    {
        // Hub used for broadcasting Admin-specific events like product variant updates.
        // Methods can be empty if we only push from Server to Client using IHubContext.
    }
}
