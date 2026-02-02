using ODCS.BaseCommon;
using ODCS.Hubs;
using ODCS.Models.Admin.Logs;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Caching.Memory;
using ODCS.Models;

namespace ODCS.WorkerService
{
    public class NotifyWorker : BackgroundService
    {
        private readonly ILogger<NotifyWorker> _logger;
        private readonly IHubContext<NotificationHub> _hubContext;
        protected readonly StockControlContext _context;
        private readonly StockControlContext _dataService;
        public LogRepositoryInterface _log;
        public NotifyWorker(ILogger<NotifyWorker> logger, IHubContext<NotificationHub> hubContext, IServiceScopeFactory factory)
        {
            _logger = logger;
            _hubContext = hubContext;
            _context = factory.CreateScope().ServiceProvider.GetRequiredService<StockControlContext>();
            _log = factory.CreateScope().ServiceProvider.GetRequiredService<LogRepositoryInterface>();
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            //APP_LOG_MENU
            var countBefore = _log.CountTotalAll(EnumTypeLog.APP_LOG_INVENTORIES_WORKER_SERVICES.ToString());
            while (!stoppingToken.IsCancellationRequested)
            {
                var countCurrent = _log.CountTotalAll(EnumTypeLog.APP_LOG_INVENTORIES_WORKER_SERVICES.ToString());
                if (countCurrent != countBefore)
                {
                    // SignalR
                    _logger.LogInformation("SignalR - Worker running at: {Time}", DateTime.Now);
                    //SignalR; no changes found
                    await _hubContext.Clients.All.SendAsync("Notify", $"Inventory");
                    countBefore = countCurrent;
                }
                else
                {
                    //SignalR; no changes found
                    await _hubContext.Clients.All.SendAsync("Notify", $"no changes found");
                }
                await Task.Delay(2000, stoppingToken);
            }
        }
    }
}
