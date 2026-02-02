using Microsoft.AspNetCore.SignalR;

namespace ASP.Hubs
{
    public class ChatHub:Hub
    {
        //public async Task SendMessage(string user, string message)
        //{
        //    //await Clients.All.SendAsync("ReceiveMessagexx", user, message);
        //    await Clients.All.SendAsync("Notify", $"Home page loaded at: {DateTime.Now}");
        //}
        public async Task SendMessage(string user, string message)
        {
            await Clients.All.SendAsync("ReceiveMessage", user, message);
        }
    }
}
