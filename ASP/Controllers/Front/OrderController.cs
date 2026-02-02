using Microsoft.AspNetCore.Mvc;
using ASP.Models.Front;
using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using ASP.DTO.DensoDTO;
namespace ASP.Controllers.Front
{
    public class OrderController : Controller
    {
        private readonly OrderRepositoryInterface _orderRepository;
        private readonly OrderDetailRepositoryInterface _orderDetailRepository;
        public OrderController(OrderRepositoryInterface orderRepository, OrderDetailRepositoryInterface orderDetailRepository)
        {
            _orderRepository = orderRepository;
            _orderDetailRepository = orderDetailRepository;
        }
        public async Task<IActionResult> OrderList(int page = 1, int pageSize = 50)
        {
            var orders = await _orderRepository.GetPagedOrdersAsync(page, pageSize);
            var totalCount = await _orderRepository.GetTotalOrdersCountAsync();
            var totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
            ViewBag.Page = page;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalPages = totalPages;
            ViewBag.TotalCount = totalCount;
            return View("~/Views/Front/DensoWareHouse/OrderList.cshtml", orders);
        }
        [HttpGet]
        public async Task<IActionResult> ExportExcel(Guid orderId)
        {
            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("Your Name");
                var order = await _orderRepository.GetOrderById(orderId);
                if (order == null)
                {
                    return NotFound("Order not found.");
                }
                var orderDetails = order.OrderDetails ?? await _orderDetailRepository.GetOrderDetailsByOrderId(orderId);
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Picking List");
                int currentRow = 1;
                int headerCols = 8;
                worksheet.Cells[currentRow, 1, currentRow, headerCols].Merge = true;
                worksheet.Cells[currentRow, 1].Value = $"ORDER INFORMATION";
                worksheet.Cells[currentRow, 1].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 1].Style.Font.Size = 14;
                worksheet.Cells[currentRow, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(68, 114, 196)); // Blue header
                worksheet.Cells[currentRow, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells[currentRow, 1, currentRow, headerCols].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                currentRow += 2;
                var orderInfo = new object[] {
                    "Order UId:", order.UId.ToString(),
                    "Customer:", order.CustomerCode,
                    "Ship Date:", order.ShipDate.ToString("yyyy-MM-dd"),
                    "Total Pallets:", order.TotalPallet.ToString()
                };
                worksheet.Cells[currentRow, 1, currentRow, orderInfo.Length].LoadFromArrays(new[] { orderInfo });
                worksheet.Cells[currentRow, 1, currentRow, orderInfo.Length].Style.Font.Bold = true;
                worksheet.Cells[currentRow, 1, currentRow, orderInfo.Length].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[currentRow, 1, currentRow, orderInfo.Length].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(236, 240, 241)); // Light gray
                currentRow += 2;
                var dataHeaders = new string[] { "Part No", "Pallet Size", "Quantity", "Total Pallets", "Warehouse", "Pallet No", "PL Status", "Collected Date" };
                var headerRange = worksheet.Cells[currentRow, 1, currentRow, dataHeaders.Length];
                headerRange.LoadFromArrays(new[] { dataHeaders });
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(52, 152, 219)); // Lighter blue
                headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
                headerRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                currentRow++;
                foreach (var od in orderDetails.OrderBy(od => od.PartNo))
                {
                    var shoppingLists = od.ShoppingLists?.OrderBy(sl => sl.PalletNo).ToList() ?? new List<ShoppingList>();
                    if (!shoppingLists.Any())
                    {
                        var emptyRow = new object[] {
                            od.PartNo, od.PalletSize, od.Quantity, od.TotalPallet, od.Warehouse,
                            "No Pallets", "N/A", "N/A"
                        };
                        worksheet.Cells[currentRow, 1, currentRow, dataHeaders.Length].LoadFromArrays(new[] { emptyRow });
                        worksheet.Cells[currentRow, 1, currentRow, dataHeaders.Length].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        currentRow++;
                    }
                    else
                    {
                        foreach (var sl in shoppingLists)
                        {
                            string plStatusText = sl.PLStatus switch
                            {
                                (short)CollectionStatusEnumDTO.None => "None",
                                (short)CollectionStatusEnumDTO.Collected => "Đã Thu Thập",
                                (short)CollectionStatusEnumDTO.Exported => "Đã Check 3 điểm",
                                (short)CollectionStatusEnumDTO.Delivered => "Đã Load lên Cont",
                                (short)CollectionStatusEnumDTO.Canceled => "Bị Huỷ",
                                _ => $"Unknown ({sl.PLStatus})"
                            };
                            var dataRow = new object[] {
                                od.PartNo,
                                od.PalletSize,
                                od.Quantity,
                                od.TotalPallet,
                                od.Warehouse,
                                sl.PalletNo,
                                plStatusText,
                                sl.CollectedDate?.ToString("yyyy-MM-dd HH:mm") ?? "N/A"
                            };
                            worksheet.Cells[currentRow, 1, currentRow, dataHeaders.Length].LoadFromArrays(new[] { dataRow });
                            if (sl.PLStatus >= (short)CollectionStatusEnumDTO.Collected && sl.PLStatus != (short)CollectionStatusEnumDTO.Canceled)
                            {
                                var statusRange = worksheet.Cells[currentRow, 7];
                                statusRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                statusRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 238, 144));
                                var dateRange = worksheet.Cells[currentRow, 8];
                                dateRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                dateRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 238, 144));
                            }
                            else if (sl.PLStatus == (short)CollectionStatusEnumDTO.Canceled)
                            {
                                var statusRange = worksheet.Cells[currentRow, 7];
                                statusRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                statusRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 182, 193));
                            }
                            worksheet.Cells[currentRow, 1, currentRow, dataHeaders.Length].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                            currentRow++;
                        }
                    }
                }
                worksheet.Cells.AutoFitColumns(15);
                if (currentRow > 3)
                {
                    worksheet.Cells[3, 1, currentRow - 1, dataHeaders.Length].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                }
                var fileName = $"PickingList_Order_{order.UId}_{order.ShipDate:yyyyMMdd}.xlsx";
                var fileBytes = package.GetAsByteArray();
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error generating Excel: {ex.Message}");
            }
        }
    }
}