using ASP.DTO.DensoDTO;
using ASP.Hubs;
using ASP.Models.ASPModel;
using ASP.Models.Front;
using ASP.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Controllers.Front
{
    [ApiController]
    [Route("api/[controller]")]
    public class DelayHistoryController : Controller
    {
        private readonly DelayHistoryRepositoryInterface _delayHistoryRepository;
        private readonly OrderRepositoryInterface _orderRepository;
        private readonly EmailServiceInterface _emailService;
        private readonly ILogger<DelayHistoryController> _logger;
        private readonly IConfiguration _configuration;
       
        public DelayHistoryController(
            DelayHistoryRepositoryInterface delayHistoryRepository,
            OrderRepositoryInterface orderRepository,
            EmailServiceInterface emailService,
            ILogger<DelayHistoryController> logger,
            IConfiguration configuration)
        {
            _delayHistoryRepository = delayHistoryRepository;
            _orderRepository = orderRepository;
            _emailService = emailService;
            _logger = logger;
            _configuration = configuration;
        }

        [HttpGet("{orderId}")]
        public async Task<IActionResult> GetByOrderId(Guid orderId)
        {
            try
            {
                var histories = await _delayHistoryRepository.GetByOrderIdAsync(orderId);
                return Ok(histories);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving delay history for OrderId: {OrderId}", orderId);
                return BadRequest(new { success = false, message = "Error retrieving delay history: " + ex.Message });
            }
        }

        [HttpPost("SaveDelay")]
        public async Task<IActionResult> SaveDelay([FromBody] DelaySaveDto delayDto)
        {
            try
            {
                if (!Guid.TryParse(delayDto.OrderId, out Guid orderId) || string.IsNullOrEmpty(delayDto.Reason))
                {
                    _logger.LogWarning("Invalid data in SaveDelay: OrderId={OrderId}, Reason={Reason}", delayDto.OrderId, delayDto.Reason);
                    return BadRequest(new { success = false, message = "Invalid data" });
                }

                var startTime = DateTime.Parse(delayDto.StartTime);
                var changeTime = DateTime.Parse(delayDto.ChangeTime);

                var delayHistory = new DelayHistory
                {
                    OId = orderId,
                    DelayType = delayDto.DelayType,
                    Reason = delayDto.Reason,
                    StartTime = startTime,
                    ChangeTime = changeTime,
                    DelayTime = delayDto.DelayTime,
                    IsAdvance = delayDto.IsAdvance,
                    CreatedDate = DateTime.Now
                };

                await _delayHistoryRepository.CreateAsync(delayHistory);
                // Pass along IsAdvance so repository can set correct status (Advance vs Delay)
                await _orderRepository.UpdateOrderStatusToDelay(orderId, startTime, delayDto.DelayTime, delayDto.IsAdvance);

                // Gửi email sau save thành công
                await SendDelayNotificationEmail(orderId, delayDto);

                _logger.LogInformation("Delay saved successfully for OrderId: {OrderId}, DelayTime: {DelayTime}h", orderId, delayDto.DelayTime);
                return Ok(new { success = true, message = "Delay saved, status updated, and email sent successfully!" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "SaveDelay ERROR for OrderId: {OrderId}", delayDto?.OrderId ?? "Unknown");
                return BadRequest(new { success = false, message = ex.Message });
            }
        }

        // FIXED: SendDelayNotificationEmail - Thêm format Delay StartTime + logs emojis consistent
        private async Task SendDelayNotificationEmail(Guid orderId, DelaySaveDto delayDto)
        {
            try
            {
                var order = await _orderRepository.GetOrderById(orderId);
                if (order == null)
                {
                    _logger.LogWarning("Order not found - Skip email for {OrderId}", orderId);
                    return;
                }

                string customerCode = order.CustomerCode ?? "";
                _logger.LogInformation("CustomerCode from DB: '{CustomerCode}' for OrderId: {OrderId}", customerCode, orderId);

                // Lấy recipients từ config
                var recipients = _configuration.GetSection("DelayNotification:Recipients").Get<List<string>>() ?? new List<string> { "nguyen.duc.tuan.a7d@ap.denso.com" };
                var ccList = _configuration.GetSection("DelayNotification:CcRecipients").Get<List<string>>() ?? new List<string> { "minhduongqh0311@gmail.com" };

                // TEST OVERRIDE: Uncomment để test Gmail To (force external)
                // recipients = new List<string> { "anhqh0311@gmail.com" };

                var subject = $"Delay Notification: Order Customer - {customerCode} Sent By Dương Minh Hẹ Hẹ :3";


                string formattedDelayStart = delayDto.StartTime;  // Fallback raw
                if (DateTime.TryParse(delayDto.StartTime, out DateTime delayStartParsed))
                {
                    formattedDelayStart = delayStartParsed.ToString("MM/dd/yyyy HH:mm");
                    _logger.LogDebug("DelayStartTime parsed and formatted: {Formatted}", formattedDelayStart);
                }
                else
                {
                    _logger.LogWarning("Failed to parse DelayStartTime: {StartTime}, using raw", delayDto.StartTime);
                }

                // FIXED: Duration plural (1 hour vs hours)
                string durationText = $"{delayDto.DelayTime} {(delayDto.DelayTime == 1 ? "hour" : "hours")}";

                var body = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; background-color: #f4f4f4; padding: 20px; }}
        .container {{ max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 30px; border-radius: 10px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }}
        .header {{ text-align: center; background-color: #ff6b6b; color: white; padding: 20px; border-radius: 10px 10px 0 0; }}
        .header h1 {{ margin: 0; font-size: 24px; }}
        .content {{ padding: 20px 0; }}
        .info-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
        .info-table th {{ background-color: #f8f9fa; padding: 12px; text-align: left; font-weight: bold; border-bottom: 2px solid #dee2e6; width: 30%; }}
        .info-table td {{ padding: 12px; border-bottom: 1px solid #dee2e6; }}
        .icon {{ font-size: 18px; margin-right: 8px; }}
        .footer {{ text-align: center; padding: 20px 0; font-size: 12px; color: #6c757d; border-top: 1px solid #dee2e6; }}
        .alert-icon {{ font-size: 48px; }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='header'>
            <div class='alert-icon'>⏰</div>
            <h1>Delay Alert for Order</h1>
        </div>
        <div class='content'>
            <table class='info-table'>
                <tr>
                    <th><span class='icon'>🏢</span>Customer:</th>
                    <td>{customerCode}</td>
                </tr>
                <tr>
                    <th><span class='icon'>🕒</span>Delay Start:</th>
                    <td>{formattedDelayStart}</td>
                </tr>
                <tr>
                    <th><span class='icon'>⏱️</span>Duration:</th>
                    <td>{durationText}</td>
                </tr>
                <tr>
                    <th><span class='icon'>📝</span>Reason:</th>
                    <td>{delayDto.Reason}</td>
                </tr>
                <tr>
                    <th><span class='icon'>📅</span>Ship Date:</th>
                    <td>{order.ShipDate:yyyy-MM-dd}</td>  
                </tr>
            </table>
        </div>
        <div class='footer'>
            <hr>
            <p>Denso Warehouse System - Automated Notification</p>
        </div>
    </div>
</body>
</html>";

                _logger.LogInformation("TEST: Email Body Preview: {Preview}... for OrderId: {OrderId}", body.Substring(0, 200), orderId);

                // Gọi service với lists (match old) - single string cho compat với interface
                await _emailService.SendEmailAsync(string.Join(",", recipients), string.Join(",", ccList), subject, body, true);  // true=HTML

                _logger.LogInformation("Email SENT: To={To}, CC={Cc}, Order={OrderId}", string.Join(",", recipients), string.Join(",", ccList), order.UId);
            }
            catch (Exception emailEx)
            {
                _logger.LogError(emailEx, "EMAIL FAILED (Delay still saved): {Message} for OrderId: {OrderId}", emailEx.Message, orderId);

            }
        }
    }
}