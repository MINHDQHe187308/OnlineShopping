using Microsoft.AspNetCore.SignalR;
using ASP.Models.ASPModel;

namespace ASP.Hubs
{
    public class PrivacyHub:Hub
    {
        private readonly ASPDbContext _context;
        public PrivacyHub(ASPDbContext context)
        {
            _context = context;
        }
        public async Task SendMessage(string message)
        {
            var a = message;
            var cLog = _context.Logs.Count();
            await Clients.All.SendAsync("ReceiveMessage", message, cLog);
        }
    }
}
