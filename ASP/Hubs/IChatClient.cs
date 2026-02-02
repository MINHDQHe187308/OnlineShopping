namespace ASP.Hubs
{
    public interface IChatClient
    {
        Task ReceiveMessage(string user, string message);
    }
}
