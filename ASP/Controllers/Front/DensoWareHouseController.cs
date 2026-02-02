using ASP.DTO.DensoDTO;
using ASP.Models.ASPModel;
using ASP.Models.Front;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
namespace ASP.Controllers.Front
{
    public class DensoWareHouseController : Controller
    {
        private readonly OrderRepositoryInterface _orderRepository;
        private readonly CustomerRepositoryInterface _customerRepository;
        private readonly OrderDetailRepositoryInterface _orderDetailRepository;
        private readonly DelayHistoryRepositoryInterface _delayHistoryRepository;
        public DensoWareHouseController(
            OrderRepositoryInterface orderRepository,
            CustomerRepositoryInterface customerRepository,
            OrderDetailRepositoryInterface orderDetailRepository,
            DelayHistoryRepositoryInterface delayHistoryRepository)
        {
            _orderRepository = orderRepository;
            _customerRepository = customerRepository;
            _orderDetailRepository = orderDetailRepository;
            _delayHistoryRepository = delayHistoryRepository;
        }
        public async Task<IActionResult> Calendar()
        {
            var today = DateTime.Today;
            var orders = await _orderRepository.GetOrdersByDate(today);
            if (orders.Any())
            {
                Console.WriteLine($"First Order Details: UId={orders.First().UId}, ProgressStatus={orders.First().OrderStatus}, ApiStatus={orders.First().ApiOrderStatus}, StartTime={orders.First().StartTime}");
            }
            var allCustomers = await _customerRepository.GetAllCustomers();
            var customerCodesWithOrders = orders.Select(o => o.CustomerCode).Distinct().ToHashSet();
            var customers = allCustomers.Where(c => customerCodesWithOrders.Contains(c.CustomerCode)).ToList();

            var ordersForViewList = new List<object>();
            foreach (var o in orders)
            {
                // If order currently in Delay, check whether delay has already ended; if so, re-evaluate status
                if (o.OrderStatus == 4 && o.DelayStartTime.HasValue && (o.DelayTime ?? 0) > 0)
                {
                    var delayEnd = o.DelayStartTime.Value.AddHours(o.DelayTime ?? 0);
                    if (delayEnd < DateTime.Now)
                    {
                        // Recalculate order status (this will persist changes and notify clients)
                        await _orderRepository.UpdateOrderStatusIfNeeded(o.UId);
                    }
                }

                // Build cumulative counts
                int collectCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Collected &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;
                int prepareCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Exported &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;
                int loadCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Delivered &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;

                // Always gather delay intervals from DelayHistory so the UI can render past/future delays
                var historyItems = (await _delayHistoryRepository.GetByOrderIdAsync(o.UId))
                    .Select(h => new
                    {
                        Start = h.StartTime.ToString("o"),
                        End = h.StartTime.AddHours(h.DelayTime).ToString("o"),
                        DelayTime = h.DelayTime,
                        Reason = h.Reason,
                        DelayType = h.DelayType,
                        IsAdvance = h.IsAdvance
                    }).ToArray();

                var orderObj = new
                {
                    UId = o.UId.ToString(),
                    Resource = o.CustomerCode ?? "Unknown",
                    ShipDate = o.ShipDate.ToString("yyyy-MM-dd"),
                    StartTime = o.StartTime.ToString("o"),
                    EndTime = o.EndTime.ToString("o"),
                    AcStartTime = o.AcStartTime?.ToString("o") ?? "",
                    AcEndTime = o.AcEndTime?.ToString("o") ?? "",
                    Status = o.OrderStatus,
                    ApiStatus = o.ApiOrderStatus,
                    TotalPallet = o.TotalPallet,
                    CollectPallet = $"{collectCount} / {o.TotalPallet}",
                    ThreePointScan = $"{prepareCount} / {o.TotalPallet}",
                    LoadCont = $"{loadCount} / {o.TotalPallet}",
                    TransCd = o.TransCd ?? "N/A",
                    TransMethod = o.TransMethod.ToString(),
                    ContSize = o.ContSize.ToString(),
                    TotalColumn = o.TotalColumn,
                    DelayIntervals = historyItems,
                    DelayStartTime = o.DelayStartTime?.ToString("o"),
                    DelayTime = o.DelayTime ?? 0
                };

                ordersForViewList.Add(orderObj);
            }

            var ordersForView = ordersForViewList.ToArray();

            var customersForView = customers.Select(c => new
            {
                CustomerCode = c.CustomerCode,
                CustomerName = c.CustomerName
            }).ToArray();
            var modelForView = new
            {
                Orders = ordersForView,
                Customers = customersForView
            };
            return View("~/Views/Front/DensoWareHouse/Calendar.cshtml", modelForView);
        }
        [HttpGet]
        public async Task<JsonResult> GetOrderDetails(string orderId)
        {
            try
            {
                if (string.IsNullOrEmpty(orderId) || !Guid.TryParse(orderId, out Guid parsedOrderId))
                {
                    return Json(new { success = false, message = "Invalid orderId format" });
                }
                var orderDetails = await _orderDetailRepository.GetOrderDetailsByOrderId(parsedOrderId);
                var order = await _orderRepository.GetOrderById(parsedOrderId);
                foreach (var od in orderDetails)
                {
                    var slCount = od.ShoppingLists?.Count ?? 0;
                    var tpcTotal = od.ShoppingLists?.Count(sl => sl.ThreePointCheck != null) ?? 0;
                }
                var detailsForView = orderDetails.Select(od => {
                    var progress = od.GetProgress();
                    return new
                    {
                        UId = progress.UId,
                        partNo = progress.PartNo,
                        quantity = progress.Quantity,
                        totalPallet = progress.TotalPallet,
                        palletSize = od.PalletSize,
                        warehouse = progress.Warehouse,
                        contNo = progress.ContNo,
                        bookContStatus = progress.BookContStatus,
                        collectPercent = progress.CollectPercent,
                        preparePercent = progress.PreparePercent,
                        loadingPercent = progress.LoadingPercent,
                        currentStage = progress.CurrentStage,
                        status = progress.Status // Progress status từ local
                    };
                }).ToList();
                object? orderSummary = null;
                if (order != null && order.OrderStatus == 4)
                {
                    DateTime? delayStart = order.DelayStartTime;
                    double delayTime = order.DelayTime ?? 0;
                    if (delayStart.HasValue)
                    {
                        DateTime baseStart = order.AcStartTime.HasValue ? order.AcStartTime.Value : order.StartTime;
                        DateTime baseEnd = order.AcEndTime.HasValue ? order.AcEndTime.Value : order.EndTime;
                        DateTime newEnd = baseEnd.AddHours(delayTime);
                        string newTimeRange = $"{baseStart:HH:mm:ss} - {newEnd:HH:mm:ss}";
                        orderSummary = new
                        {
                            newTimeRange = newTimeRange,
                            delayTime = delayTime,
                            apiStatus = order.ApiOrderStatus // Thêm ApiStatus nếu cần trace delay với API
                        };
                    }
                }
                return Json(new { success = true, data = detailsForView, orderSummary = orderSummary });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }
        [HttpGet]
        public async Task<JsonResult> GetCalendarData()
        {
            // Accept optional date query parameter (yyyy-MM-dd or ISO); fallback to today
            var dateStr = Request.Query["date"].FirstOrDefault();
            DateTime targetDate;
            if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out var parsed))
            {
                targetDate = parsed.Date;
            }
            else
            {
                targetDate = DateTime.Today;
            }
            var orders = await _orderRepository.GetOrdersByDate(targetDate);
            var allCustomers = await _customerRepository.GetAllCustomers();
            var customerCodesWithOrders = orders.Select(o => o.CustomerCode).Distinct().ToHashSet();
            var customers = allCustomers.Where(c => customerCodesWithOrders.Contains(c.CustomerCode)).ToList();
            // Build ordersForView with DelayIntervals included (DelayHistory) so frontend can render delay/advance intervals
            var ordersForViewList = new List<object>();
            foreach (var o in orders)
            {
                // Cumulative counts: Đã đạt mốc này trở lên, loại trừ Canceled
                int collectCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Collected &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;
                int prepareCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Exported &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;
                int loadCount = o.OrderDetails?.Sum(od => od.ShoppingLists?.Count(sl =>
                    sl.PLStatus >= (short)CollectionStatusEnumDTO.Delivered &&
                    sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled) ?? 0) ?? 0;

                // Gather delay intervals (DelayHistory) for this order so frontend can draw past/future delays and advances
                var historyItems = (await _delayHistoryRepository.GetByOrderIdAsync(o.UId))
                    .Select(h => new
                    {
                        Start = h.StartTime.ToString("o"),
                        End = h.StartTime.AddHours(h.DelayTime).ToString("o"),
                        DelayTime = h.DelayTime,
                        Reason = h.Reason,
                        DelayType = h.DelayType,
                        IsAdvance = h.IsAdvance
                    }).ToArray();

                string? delayStartTime = null;
                double delayTime = 0;
                if (o.OrderStatus == 4)
                {
                    delayStartTime = o.DelayStartTime?.ToString("o");
                    delayTime = o.DelayTime ?? 0;
                }

                ordersForViewList.Add(new
                {
                    UId = o.UId.ToString(),
                    Resource = o.CustomerCode ?? "Unknown",
                    ShipDate = o.ShipDate.ToString("yyyy-MM-dd"),
                    StartTime = o.StartTime.ToString("o"),
                    EndTime = o.EndTime.ToString("o"),
                    AcStartTime = o.AcStartTime?.ToString("o") ?? "",
                    AcEndTime = o.AcEndTime?.ToString("o") ?? "",
                    Status = o.OrderStatus,
                    ApiStatus = o.ApiOrderStatus,
                    TotalPallet = o.TotalPallet,
                    CollectPallet = $"{collectCount} / {o.TotalPallet}",
                    ThreePointScan = $"{prepareCount} / {o.TotalPallet}",
                    LoadCont = $"{loadCount} / {o.TotalPallet}",
                    TransCd = o.TransCd ?? "N/A",
                    TransMethod = o.TransMethod.ToString(),
                    ContSize = o.ContSize.ToString(),
                    TotalColumn = o.TotalColumn,
                    DelayStartTime = delayStartTime,
                    DelayTime = delayTime,
                    DelayIntervals = historyItems
                });
            }
            var ordersForView = ordersForViewList.ToArray();
            var customersForView = customers.Select(c => new
            {
                CustomerCode = c.CustomerCode,
                CustomerName = c.CustomerName
            }).ToArray();

            // Build advanceEvents: prefer explicit Order.AdvanceStartTime/AdvanceEndTime when present,
            // otherwise include DelayHistory items where IsAdvance == true
            var advanceEvents = new List<object>();
            foreach (var o in orders)
            {
                // If Order table contains explicit advance markers, use them
                if (o.IsAdvance == true && o.AdvanceStartTime.HasValue)
                {
                    var advStart = o.AdvanceStartTime.Value.ToString("o");
                    var advEnd = (o.AdvanceEndTime.HasValue) ? o.AdvanceEndTime.Value.ToString("o") : o.AdvanceStartTime.Value.ToString("o");
                    advanceEvents.Add(new
                    {
                        id = "advance-" + o.UId.ToString(),
                        UId = o.UId.ToString(),
                        Resource = o.CustomerCode,
                        Start = advStart,
                        End = advEnd,
                        Title = "Advance",
                        IsAdvance = true
                    });
                }
                else
                {
                    // fallback: read from DelayHistory entries for this order (isAdvance entries)
                    var hist = (await _delayHistoryRepository.GetByOrderIdAsync(o.UId))
                        .Where(h => h.IsAdvance)
                        .Select(h => new
                        {
                            id = "advance-" + o.UId.ToString() + "-" + h.UId.ToString(),
                            UId = o.UId.ToString(),
                            Resource = o.CustomerCode,
                            Start = h.StartTime.ToString("o"),
                            End = h.StartTime.AddHours(h.DelayTime).ToString("o"),
                            Title = "Advance",
                            IsAdvance = true
                        });
                    advanceEvents.AddRange(hist);
                }
            }

            return Json(new { orders = ordersForView, customers = customersForView, advanceEvents = advanceEvents.ToArray(), requestedDate = targetDate.ToString("yyyy-MM-dd") });
        }
        private string MapOrderStatusToString(short orderStatus)
        {
            return orderStatus switch
            {
                0 => "Planned",
                1 => "Pending",
                2 => "Completed",
                3 => "Shipped",
                4 => "Delay",
                _ => "Planned"
            };
        }
    }
}