using ASP.Models.Front;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
namespace ASP.Controllers.Front
{
    public class CustomerController : Controller
    {
        private readonly CustomerRepositoryInterface _customerRepository;
        private readonly LeadtimeMasterRepositoryInterface _leadtimeRepository;
        private readonly ShippingScheduleRepositoryInterface _shippingScheduleRepository;
        public CustomerController(
            CustomerRepositoryInterface customerRepository,
            LeadtimeMasterRepositoryInterface leadtimeRepository,
            ShippingScheduleRepositoryInterface shippingScheduleRepository)
        {
            _customerRepository = customerRepository;
            _leadtimeRepository = leadtimeRepository;
            _shippingScheduleRepository = shippingScheduleRepository;
        }
        private bool IsRowEmpty(ExcelWorksheet sheet, int row)
        {
            for (int col = 1; col <= 8; col++)
            {
                var value = sheet.Cells[row, col].Value;
                if (value != null && !string.IsNullOrWhiteSpace(value.ToString().Trim()))
                    return false; // Có ít nhất 1 ô có dữ liệu → Không trống
            }
            return true; // Toàn bộ trống
        }
        // Helper để xử lý merge cells
        private object GetMergedValue(ExcelWorksheet sheet, int row, int col)
        {
            var cell = sheet.Cells[row, col];
            var mergedAddress = sheet.MergedCells[row, col];
            if (!string.IsNullOrEmpty(mergedAddress))
            {
                var mergedRange = new ExcelAddress(mergedAddress);
                var topLeftCell = sheet.Cells[mergedRange.Start.Row, mergedRange.Start.Column];
                return topLeftCell.Value;
            }
            return cell.Value;
        }
        // Danh sách Customers
        public async Task<IActionResult> CustomerList()
        {
            var customers = await _customerRepository.GetAllCustomers();
            if (customers == null || customers.Count == 0)
            {
                TempData["ErrorMessage"] = "No customers found.";
                return View("~/Views/Front/Home/CustomerList.cshtml", new List<Customer>());
            }
            return View("~/Views/Front/Home/CustomerList.cshtml", customers);
        }
        // Tạo mới Customer (Ajax gọi đến)
        [HttpPost]
        public async Task<JsonResult> AddSupplier([FromBody] Customer request)
        {
            var success = await _customerRepository.CreateCustomer(request);
            if (success)
                return Json(new { success = true });
            return Json(new { success = false, message = "Cannot create supplier" });
        }
        // Cập nhật Customer
        [HttpPost]
        public async Task<JsonResult> UpdateSupplier([FromBody] Customer request)
        {
            if (request == null || string.IsNullOrEmpty(request.CustomerCode))
            {
                return Json(new { success = false, message = "Mã khách hàng không hợp lệ" });
            }
            var result = await _customerRepository.UpdateCustomerByCode(request.CustomerCode, request);
            if (result.Success)
                return Json(new { success = true, message = result.Message });
            return Json(new { success = false, message = result.Message });
        }

        // Xóa Customer
        [HttpPost]
        public async Task<JsonResult> DeleteSupplier(string code)
        {
            var success = await _customerRepository.RemoveCustomerByCode(code);
            if (success)
            {
                return Json(new { success = true, message = "Xóa khách hàng thành công!" });
            }
            return Json(new { success = false, message = "Không tìm thấy khách hàng để xóa!" });
        }

        // Xóa tất cả Customer
        [HttpPost]
        public async Task<JsonResult> DeleteAllCustomers()
        {
            var success = await _customerRepository.DeleteAllCustomers();
            if (success)
            {
                return Json(new { success = true, message = "Đã xóa tất cả khách hàng thành công!" });
            }
            return Json(new { success = false, message = "Có lỗi xảy ra khi xóa tất cả khách hàng!" });
        }

        // Lấy danh sách LeadtimeMaster theo CustomerCode
        [HttpGet]
        public async Task<JsonResult> GetLeadtimesByCustomer(string customerCode)
        {
            if (string.IsNullOrEmpty(customerCode))
            {
                return Json(new { success = false, message = "CustomerCode không hợp lệ" });
            }
            var leadtimes = await _leadtimeRepository.GetAllLeadtimesByCustomer(customerCode);
            return Json(new { success = true, data = leadtimes });
        }

        // Thêm LeadtimeMaster
        [HttpPost]
        public async Task<JsonResult> AddLeadtime([FromBody] LeadtimeMaster request)
        {
            if (request == null || string.IsNullOrEmpty(request.CustomerCode) || string.IsNullOrEmpty(request.TransCd))
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ" });
            }
            var success = await _leadtimeRepository.CreateLeadtime(request);
            if (success)
                return Json(new { success = true });
            return Json(new { success = false, message = "Cannot create leadtime" });
        }

        // Cập nhật LeadtimeMaster
        [HttpPost]
        public async Task<JsonResult> UpdateLeadtime([FromBody] LeadtimeMaster request)
        {
            if (request == null || string.IsNullOrEmpty(request.CustomerCode) || string.IsNullOrEmpty(request.TransCd))
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ" });
            }
            var result = await _leadtimeRepository.UpdateLeadtimeByKey(request.CustomerCode, request.TransCd, request);
            if (result.Success)
                return Json(new { success = true, message = result.Message });
            return Json(new { success = false, message = result.Message });
        }

        // Xóa LeadtimeMaster
        [HttpPost]
        public async Task<JsonResult> DeleteLeadtime(string customerCode, string transCd)
        {
            var success = await _leadtimeRepository.RemoveLeadtimeByKey(customerCode, transCd);
            if (success)
            {
                return Json(new { success = true, message = "Xóa leadtime thành công!" });
            }
            return Json(new { success = false, message = "Không tìm thấy leadtime để xóa!" });
        }

        [HttpPost]
        public async Task<JsonResult> DeleteAllLeadtimesByCustomer()
        {
            try
            {
                // Đọc body như string (an toàn cho JSON)
                using var reader = new StreamReader(Request.Body);
                var jsonBody = await reader.ReadToEndAsync();
                if (string.IsNullOrEmpty(jsonBody))
                {
                    return Json(new { success = false, message = "Request body trống hoặc không hợp lệ." });
                }

                // Parse JSON: { "customerCode": "ABC123" }
                using var doc = JsonDocument.Parse(jsonBody);
                if (!doc.RootElement.TryGetProperty("customerCode", out var prop) || prop.ValueKind != JsonValueKind.String)
                {
                    return Json(new { success = false, message = "Không tìm thấy 'customerCode' trong request." });
                }

                var customerCode = prop.GetString()?.Trim();
                if (string.IsNullOrEmpty(customerCode))
                {
                    return Json(new { success = false, message = "CustomerCode không hợp lệ (trống hoặc null)." });
                }

                // Gọi repository
                var success = await _leadtimeRepository.DeleteAllLeadtimes(customerCode);
                if (success)
                {
                    return Json(new { success = true, message = $"Đã xóa tất cả leadtime của customer '{customerCode}' thành công!" });
                }
                return Json(new { success = false, message = "Có lỗi xảy ra khi xóa leadtime!" });
            }
            catch (JsonException ex)
            {
                return Json(new { success = false, message = $"Lỗi parse JSON: {ex.Message}" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }
        // Lấy danh sách ShippingSchedule theo CustomerCode
        [HttpGet]
        public async Task<JsonResult> GetShippingSchedulesByCustomer(string customerCode)
        {
            if (string.IsNullOrEmpty(customerCode))
            {
                return Json(new { success = false, message = "CustomerCode không hợp lệ" });
            }
            var schedules = await _shippingScheduleRepository.GetAllShippingSchedulesByCustomer(customerCode);
            return Json(new { success = true, data = schedules });
        }

        // Thêm ShippingSchedule
        [HttpPost]
        public async Task<JsonResult> AddShippingSchedule([FromBody] ShippingSchedule request)
        {
            if (request == null || string.IsNullOrEmpty(request.CustomerCode) || string.IsNullOrEmpty(request.TransCd))
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ" });
            }
            var success = await _shippingScheduleRepository.CreateShippingSchedule(request);
            if (success)
                return Json(new { success = true });
            return Json(new { success = false, message = "Cannot create shipping schedule" });
        }

        // Cập nhật ShippingSchedule
        [HttpPost]
        public async Task<JsonResult> UpdateShippingSchedule([FromBody] ShippingSchedule request)
        {
            if (request == null || string.IsNullOrEmpty(request.CustomerCode) || string.IsNullOrEmpty(request.TransCd))
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ" });
            }
            var result = await _shippingScheduleRepository.UpdateShippingScheduleByKey(request.CustomerCode, request.TransCd, request.Weekday, request);
            if (result.Success)
                return Json(new { success = true, message = result.Message });
            return Json(new { success = false, message = result.Message });
        }

        // Xóa ShippingSchedule
        [HttpPost]
        public async Task<JsonResult> DeleteShippingSchedule(string customerCode, string transCd, int weekday)
        {
            var success = await _shippingScheduleRepository.RemoveShippingScheduleByKey(customerCode, transCd, (DayOfWeek)weekday);
            if (success)
            {
                return Json(new { success = true, message = "Xóa lịch vận chuyển thành công!" });
            }
            return Json(new { success = false, message = "Không tìm thấy lịch vận chuyển để xóa!" });
        }
        [HttpPost]
        public async Task<JsonResult> DeleteAllShippingSchedulesByCustomer()
        {
            try
            {
                // Đọc body như string (an toàn cho JSON)
                using var reader = new StreamReader(Request.Body);
                var jsonBody = await reader.ReadToEndAsync();
                if (string.IsNullOrEmpty(jsonBody))
                {
                    return Json(new { success = false, message = "Request body trống hoặc không hợp lệ." });
                }

                // Parse JSON: { "customerCode": "ABC123" }
                using var doc = JsonDocument.Parse(jsonBody);
                if (!doc.RootElement.TryGetProperty("customerCode", out var prop) || prop.ValueKind != JsonValueKind.String)
                {
                    return Json(new { success = false, message = "Không tìm thấy 'customerCode' trong request." });
                }

                var customerCode = prop.GetString()?.Trim();
                if (string.IsNullOrEmpty(customerCode))
                {
                    return Json(new { success = false, message = "CustomerCode không hợp lệ (trống hoặc null)." });
                }

                // Gọi repository
                var success = await _shippingScheduleRepository.DeleteAllShippingSchedules(customerCode);
                if (success)
                {
                    return Json(new { success = true, message = $"Đã xóa tất cả shipping schedule của customer '{customerCode}' thành công!" });
                }
                return Json(new { success = false, message = "Có lỗi xảy ra khi xóa shipping schedule!" });
            }
            catch (JsonException ex)
            {
                return Json(new { success = false, message = $"Lỗi parse JSON: {ex.Message}" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile excelFile)
        {
            ExcelPackage.License.SetNonCommercialPersonal("Your Name");
            if (excelFile == null || excelFile.Length == 0)
            {
                return Json(new { success = false, message = "No file uploaded." });
            }
            var ext = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm")
            {
                return Json(new { success = false, message = "Only .xlsx or .xlsm files are supported." });
            }
            var results = new ImportResult
            {
                CustomersAdded = 0,
                LeadtimesAdded = 0,
                ShippingSchedulesAdded = 0,
                Errors = new List<string>()
            };
            try
            {
                using var stream = new MemoryStream();
                await excelFile.CopyToAsync(stream);
                using var package = new ExcelPackage(stream);
                var sheet = package.Workbook.Worksheets["ShippingSchedule"];
                if (sheet != null)
                {
                    await ImportFromSingleSheet(sheet, results);
                }
                else
                {
                    results.Errors.Add("Sheet 'ShippingSchedule' not found. Please use a single sheet named 'ShippingSchedule'.");
                }
            }
            catch (Exception ex)
            {
                results.Errors.Add($"Error processing file: {ex.Message}");
            }
            return Json(new
            {
                success = !results.Errors.Any(),
                message = results.Errors.Any() ? string.Join("; ", results.Errors) : $"Import successful! Processed {results.CustomersAdded} customers, {results.LeadtimesAdded} leadtimes, {results.ShippingSchedulesAdded} shipping schedules."
            });
        }
        private async Task ImportFromSingleSheet(ExcelWorksheet sheet, ImportResult results)
        {
            var rowCount = sheet.Dimension?.Rows ?? 0;
            while (rowCount >= 4 && IsRowEmpty(sheet, rowCount))
            {
                rowCount--;
            }
            if (rowCount < 4)
            {
                results.Errors.Add("No data rows found (only header or empty sheet).");
                return;
            }
            List<ImportRow> validRows = new List<ImportRow>();
            var customerCodes = new HashSet<string>();
            var ltKeys = new HashSet<string>();
            var ssKeys = new HashSet<string>();
            // Pass 1: Parse và validate rows, collect valid data và keys
            for (int row = 4; row <= rowCount; row++)
            {
                if (IsRowEmpty(sheet, row))
                {
                    continue;
                }
                try
                {
                    var customerCodeRaw = GetMergedValue(sheet, row, 1);
                    var customerCode = customerCodeRaw?.ToString()?.Trim();
                    var customerNameRaw = GetMergedValue(sheet, row, 2);
                    var customerName = customerNameRaw?.ToString()?.Trim();
                    var transCdRaw = GetMergedValue(sheet, row, 3);
                    var transCd = transCdRaw?.ToString()?.Trim();
                    var collectTimeRaw = GetMergedValue(sheet, row, 4);
                    decimal? collectTime = null;
                    if (collectTimeRaw != null)
                    {
                        if (decimal.TryParse(collectTimeRaw.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var ct))
                            collectTime = ct;
                    }
                    var prepareTimeRaw = GetMergedValue(sheet, row, 5);
                    decimal? prepareTime = null;
                    if (prepareTimeRaw != null)
                    {
                        if (decimal.TryParse(prepareTimeRaw.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var pt))
                            prepareTime = pt;
                    }
                    var loadingTimeRaw = GetMergedValue(sheet, row, 6);
                    decimal? loadingTime = null;
                    if (loadingTimeRaw != null)
                    {
                        if (decimal.TryParse(loadingTimeRaw.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var lt))
                            loadingTime = lt;
                    }
                    // **FIX PARSE WEEKDAY: Hỗ trợ số 2-7 trực tiếp, tên ngày, DateTime → Giữ nguyên 2-7, sau convert sang DayOfWeek**
                    var weekdayCell = sheet.Cells[row, 7].Value;
                    var weekdayStr = (weekdayCell?.ToString() ?? "").Trim();
                    var weekdayText = sheet.Cells[row, 7].Text?.Trim() ?? ""; // Ưu tiên Text (e.g., "2" hoặc "Thứ 2")
                    int weekdayInt = -1; // Sẽ là 2-7 từ Excel
                    var missingFields = new List<string>();
                    if (string.IsNullOrEmpty(customerCode)) missingFields.Add("Customer Code");
                    if (string.IsNullOrEmpty(customerName)) missingFields.Add("Customer Name");
                    if (string.IsNullOrEmpty(transCd)) missingFields.Add("TransCode");
                    if (!collectTime.HasValue) missingFields.Add("CollectTimePerPallet");
                    if (!prepareTime.HasValue) missingFields.Add("PrepareTimePerPallet");
                    if (!loadingTime.HasValue) missingFields.Add("LoadingTimePerPallet");
                    if (weekdayCell == null) missingFields.Add("Day of the week");
                    if (!string.IsNullOrEmpty(weekdayText))
                    {
                        // Hỗ trợ tên ngày tiếng Việt/Anh (dễ dùng cho user)
                        var dayNamesVi = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                        {
                            { "thứ 2", 2 }, { "thu 2", 2 }, { "monday", 2 }, { "mon", 2 },
                            { "thứ 3", 3 }, { "thu 3", 3 }, { "tuesday", 3 }, { "tue", 3 },
                            { "thứ 4", 4 }, { "thu 4", 4 }, { "wednesday", 4 }, { "wed", 4 },
                            { "thứ 5", 5 }, { "thu 5", 5 }, { "thursday", 5 }, { "thu", 5 },
                            { "thứ 6", 6 }, { "thu 6", 6 }, { "friday", 6 }, { "fri", 6 },
                            { "thứ 7", 7 }, { "thu 7", 7 }, { "saturday", 7 }, { "sat", 7 },
                            { "chủ nhật", 1 }, { "cn", 1 }, { "sunday", 1 }, { "sun", 1 } // Optional: 1=Sunday
                        };
                        if (dayNamesVi.TryGetValue(weekdayText, out var dayFromName))
                        {
                            weekdayInt = dayFromName;
                        }
                        else if (int.TryParse(weekdayText, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedInt))
                        {
                            weekdayInt = parsedInt;
                        }
                        else if (DateTime.TryParse(weekdayText, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDt))
                        {
                            // Map DayOfWeek (0=Sun→1, 1=Mon→2, etc.) để khớp thứ N
                            weekdayInt = (int)parsedDt.DayOfWeek + 1;
                        }
                    }
                    // Fallback từ Value nếu Text fail
                    if (weekdayInt < 0)
                    {
                        if (weekdayCell is DateTime weekdayDt)
                        {
                            weekdayInt = (int)weekdayDt.DayOfWeek + 1;
                        }
                        else if (weekdayCell is double weekdayDbl)
                        {
                            try
                            {
                                var dtFromSerial = DateTime.FromOADate(weekdayDbl);
                                weekdayInt = (int)dtFromSerial.DayOfWeek + 1;
                            }
                            catch { }
                        }
                        else if (int.TryParse(weekdayStr, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsedFallback))
                        {
                            weekdayInt = parsedFallback;
                        }
                        else if (Enum.TryParse<DayOfWeek>(weekdayStr, true, out var parsedDay))
                        {
                            weekdayInt = (int)parsedDay + 1; // DayOfWeek Mon=1 → thứ 2=2
                        }
                    }
                    // **VALIDATE: 1-7 (1=Sun optional, 2-7=Mon-Sat)**
                    if (weekdayInt < 1 || weekdayInt > 7)
                    {
                        results.Errors.Add($"Row {row}: Invalid weekday '{weekdayCell}' (Text='{weekdayText}'). Expected 2-7 for Thứ 2-7 (or 1 for CN), or day name.");
                        missingFields.Add("Day of the week");
                    }
                    if (missingFields.Any())
                    {
                        results.Errors.Add($"Row {row}: Missing required fields: {string.Join(", ", missingFields)}.");
                        continue;
                    }
                    // FIXED: Xử lý CutOffTime
                    var cutOffCell = sheet.Cells[row, 8].Value;
                    var cutOffTimeStr = cutOffCell?.ToString()?.Trim() ?? ""; // FIXED: Dùng ToString() để handle Double/String/DateTime
                    var cutOffText = sheet.Cells[row, 8].Text?.Trim() ?? ""; // Ưu tiên Text cho formatted value
                    bool isCutOffEmpty = cutOffCell == null; // FIXED: Chỉ empty nếu null (Double 0.5 != null)
                    TimeOnly cutOffTimeOnly = TimeOnly.Parse("00:00");
                    bool parsedCutOff = false;
                    if (!isCutOffEmpty)
                    {
                        // PRIORITY: Parse cell.Text trước (e.g., "12:00" trực tiếp)
                        if (!string.IsNullOrEmpty(cutOffText))
                        {
                            var normalizedText = cutOffText.Replace(',', ':').Replace(' ', ':').Replace(".", ":").Trim().ToUpper();
                            if (TimeOnly.TryParse(cutOffText, CultureInfo.InvariantCulture, DateTimeStyles.None, out cutOffTimeOnly) ||
                                TimeOnly.TryParse(cutOffText, new CultureInfo("vi-VN"), DateTimeStyles.None, out cutOffTimeOnly))
                            {
                                parsedCutOff = true;
                            }
                            else
                            {
                                var formats = new[] { "H:MM", "HH:MM", "H:MM:SS", "HH:MM:SS", "H:MM TT", "HH:MM TT", "H:M", "HH:M", "H:M:S", "HH:M:S" };
                                foreach (var fmt in formats)
                                {
                                    if (TimeOnly.TryParseExact(normalizedText, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out cutOffTimeOnly) ||
                                        TimeOnly.TryParseExact(normalizedText, fmt, new CultureInfo("vi-VN"), DateTimeStyles.None, out cutOffTimeOnly))
                                    {
                                        parsedCutOff = true;
                                        break;
                                    }
                                }
                                if (!parsedCutOff && DateTime.TryParse(normalizedText, out var dtTextFallback))
                                {
                                    cutOffTimeOnly = TimeOnly.FromDateTime(dtTextFallback);
                                    parsedCutOff = true;
                                }
                            }
                        }
                        // Branch 1: DateTime (common for formatted time)
                        if (!parsedCutOff && cutOffCell is DateTime dt)
                        {
                            cutOffTimeOnly = TimeOnly.FromDateTime(dt);
                            parsedCutOff = true;
                        }
                        // Branch 2: Double (fraction) - Robust với OADate
                        if (!parsedCutOff && cutOffCell is double dbl)
                        {
                            if (Math.Abs(dbl) > 0.000001) // Giảm threshold cho precision
                            {
                                try
                                {
                                    // Sử dụng FromOADate thay TimeSpan để tránh precision loss
                                    var dtFromOADate = DateTime.FromOADate(dbl);
                                    cutOffTimeOnly = TimeOnly.FromDateTime(dtFromOADate);
                                    parsedCutOff = true;
                                }
                                catch (Exception exOADate)
                                {
                                    // Fallback TimeSpan
                                    var ts = TimeSpan.FromDays(dbl);
                                    cutOffTimeOnly = TimeOnly.FromTimeSpan(ts);
                                    parsedCutOff = true;
                                }
                            }
                        }
                        // Branch 3: String from Value
                        if (!parsedCutOff && !string.IsNullOrEmpty(cutOffTimeStr))
                        {
                            var normalizedStr = cutOffTimeStr.Replace(',', ':').Replace(' ', ':').Replace(".", ":").Trim().ToUpper(); // Normalize mạnh hơn, handle . cho decimal
                            var formats = new[] { "H:MM", "HH:MM", "H:MM:SS", "HH:MM:SS", "H:MM TT", "HH:MM TT", "H:M", "HH:M", "H:M:S", "HH:M:S" }; // Upper cho AM/PM
                            foreach (var fmt in formats)
                            {
                                if (TimeOnly.TryParseExact(normalizedStr, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out cutOffTimeOnly) ||
                                    TimeOnly.TryParseExact(normalizedStr, fmt, new CultureInfo("vi-VN"), DateTimeStyles.None, out cutOffTimeOnly))
                                {
                                    parsedCutOff = true;
                                    break;
                                }
                            }
                            if (!parsedCutOff)
                            {
                                if (TimeOnly.TryParse(normalizedStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out cutOffTimeOnly) ||
                                    TimeOnly.TryParse(normalizedStr, new CultureInfo("vi-VN"), DateTimeStyles.None, out cutOffTimeOnly))
                                {
                                    parsedCutOff = true;
                                }
                            }
                            if (!parsedCutOff && DateTime.TryParse(normalizedStr, out var dtFallback))
                            {
                                cutOffTimeOnly = TimeOnly.FromDateTime(dtFallback);
                                parsedCutOff = true;
                            }
                        }
                        if (!parsedCutOff)
                        {
                            results.Errors.Add($"Row {row}: Invalid cutOffTime '{cutOffCell}' (Text='{cutOffText}'). Set to default 00:00.");
                        }
                    }
                    var importRow = new ImportRow(customerCode, customerName, transCd, collectTime.Value, prepareTime.Value, loadingTime.Value, weekdayInt, cutOffTimeOnly);
                    validRows.Add(importRow);
                    customerCodes.Add(customerCode);
                    ltKeys.Add($"{customerCode}-{transCd}");
                    // **NEW: Thu thập ssKeys với DB weekday để tránh duplicate key**
                    int weekdayDBForKey = (weekdayInt == 1 ? 0 : weekdayInt - 1);
                    ssKeys.Add($"{customerCode}-{transCd}-{weekdayDBForKey}");
                }
                catch (Exception ex)
                {
                    results.Errors.Add($"Row {row}: {ex.Message}");
                }
            }
            if (validRows.Count == 0)
            {
                if (!results.Errors.Any())
                {
                    results.Errors.Add("No valid data rows found.");
                }
                return;
            }
            // Pass 2: Fetch existing data using collected keys
            var existingCustomers = new Dictionary<string, Customer>();
            foreach (var code in customerCodes)
            {
                var c = await _customerRepository.GetCustomerByCode(code);
                if (c != null)
                {
                    existingCustomers[code] = c;
                }
            }
            var existingLeadtimes = new Dictionary<string, LeadtimeMaster>();
            foreach (var custCode in customerCodes)
            {
                var lts = await _leadtimeRepository.GetAllLeadtimesByCustomer(custCode) ?? new List<LeadtimeMaster>();
                foreach (var lt in lts)
                {
                    var key = $"{lt.CustomerCode}-{lt.TransCd}";
                    if (ltKeys.Contains(key))
                    {
                        existingLeadtimes[key] = lt;
                    }
                }
            }
            var existingSchedules = new Dictionary<string, ShippingSchedule>();
            foreach (var custCode in customerCodes)
            {
                var sss = await _shippingScheduleRepository.GetAllShippingSchedulesByCustomer(custCode) ?? new List<ShippingSchedule>();
                foreach (var ss in sss)
                {
                    var key = $"{ss.CustomerCode}-{ss.TransCd}-{(int)ss.Weekday}";
                    if (ssKeys.Contains(key))
                    {
                        existingSchedules[key] = ss;
                    }
                }
            }
            // Pass 3: Process updates/creates - Check changes before update/insert
            var processedCustomers = new HashSet<string>();
            var processedLeadtimes = new HashSet<string>();
            var processedSsKeys = new HashSet<string>(); // **NEW: Tránh duplicate SS key (fix overwrite lệch thứ tự)**
            foreach (var rowData in validRows)
            {
                // Process Customer (once per code) - Check if name changed
                if (!processedCustomers.Contains(rowData.CustomerCode))
                {
                    var existingC = existingCustomers.GetValueOrDefault(rowData.CustomerCode);
                    bool isNewCustomer = existingC == null;
                    bool customerChanged = false;
                    Customer customerToSave = existingC ?? new Customer();
                    if (!isNewCustomer)
                    {
                        if (existingC.CustomerName != rowData.CustomerName)
                        {
                            existingC.CustomerName = rowData.CustomerName;
                            existingC.UpdateBy = "ExcelImport";
                            customerChanged = true;
                        }
                    }
                    else
                    {
                        customerToSave.CustomerCode = rowData.CustomerCode;
                        customerToSave.CustomerName = rowData.CustomerName;
                        customerToSave.Descriptions = "";
                        customerToSave.CreateBy = "ExcelImport";
                        customerChanged = true; // New always "changed"
                    }
                    if (customerChanged)
                    {
                        bool success;
                        if (isNewCustomer)
                        {
                            success = await _customerRepository.CreateCustomer(customerToSave);
                        }
                        else
                        {
                            var res = await _customerRepository.UpdateCustomerByCode(rowData.CustomerCode, existingC);
                            success = res.Success;
                        }
                        if (success)
                        {
                            results.CustomersAdded++;
                            if (isNewCustomer)
                            {
                                existingCustomers[rowData.CustomerCode] = customerToSave;
                            }
                        }
                        else
                        {
                            results.Errors.Add($"Failed to {(isNewCustomer ? "create" : "update")} customer {rowData.CustomerCode}.");
                        }
                    }
                    processedCustomers.Add(rowData.CustomerCode);
                }
                // Process Leadtime (once per cust-trans) - Check if times changed
                var ltKey = $"{rowData.CustomerCode}-{rowData.TransCd}";
                if (!processedLeadtimes.Contains(ltKey))
                {
                    var existingLt = existingLeadtimes.GetValueOrDefault(ltKey);
                    bool isNewLt = existingLt == null;
                    bool ltChanged = false;
                    LeadtimeMaster ltToSave = existingLt ?? new LeadtimeMaster();
                    if (!isNewLt)
                    {
                        if (existingLt.CollectTimePerPallet != (double)rowData.CollectTimePerPallet ||
                            existingLt.PrepareTimePerPallet != (double)rowData.PrepareTimePerPallet ||
                            existingLt.LoadingTimePerColumn != (double)rowData.LoadingTimePerColumn)
                        {
                            existingLt.CollectTimePerPallet = (double)rowData.CollectTimePerPallet;
                            existingLt.PrepareTimePerPallet = (double)rowData.PrepareTimePerPallet;
                            existingLt.LoadingTimePerColumn = (double)rowData.LoadingTimePerColumn;
                            existingLt.UpdateBy = "ExcelImport";
                            ltChanged = true;
                        }
                    }
                    else
                    {
                        ltToSave.CustomerCode = rowData.CustomerCode;
                        ltToSave.TransCd = rowData.TransCd;
                        ltToSave.CollectTimePerPallet = (double)rowData.CollectTimePerPallet;
                        ltToSave.PrepareTimePerPallet = (double)rowData.PrepareTimePerPallet;
                        ltToSave.LoadingTimePerColumn = (double)rowData.LoadingTimePerColumn;
                        ltToSave.CreateBy = "ExcelImport";
                        ltChanged = true; // New always "changed"
                    }
                    if (ltChanged)
                    {
                        bool success;
                        if (isNewLt)
                        {
                            success = await _leadtimeRepository.CreateLeadtime(ltToSave);
                        }
                        else
                        {
                            var res = await _leadtimeRepository.UpdateLeadtimeByKey(rowData.CustomerCode, rowData.TransCd, existingLt);
                            success = res.Success;
                        }
                        if (success)
                        {
                            results.LeadtimesAdded++;
                            if (isNewLt)
                            {
                                existingLeadtimes[ltKey] = ltToSave;
                            }
                        }
                        else
                        {
                            results.Errors.Add($"Failed to {(isNewLt ? "create" : "update")} leadtime {ltKey}.");
                        }
                    }
                    processedLeadtimes.Add(ltKey);
                }
                // **FIX PROCESS SHIPPING SCHEDULE**
                int excelWeekday = rowData.Weekday; // 2-7 từ Excel
                int dbWeekday = (excelWeekday == 1 ? 0 : excelWeekday - 1); // Convert sang 0-6 cho DayOfWeek
                var ssKey = $"{rowData.CustomerCode}-{rowData.TransCd}-{dbWeekday}"; // Key dùng DB value
                if (processedSsKeys.Contains(ssKey))
                {
                    continue; // Skip để giữ CutOffTime của row đầu tiên
                }
                var existingSs = existingSchedules.GetValueOrDefault(ssKey);
                bool isNewSs = existingSs == null;
                bool ssChanged = false;
                ShippingSchedule ssToSave = existingSs ?? new ShippingSchedule();
                if (!isNewSs)
                {
                    // Check changed: So sánh TimeOnly (không phụ thuộc Excel/DB weekday)
                    if (!existingSs.CutOffTime.Equals(rowData.CutOffTime) && !(rowData.CutOffTime.Hour == 0 && rowData.CutOffTime.Minute == 0))
                    {
                        existingSs.CutOffTime = rowData.CutOffTime;
                        existingSs.UpdatedBy = "ExcelImport";
                        ssChanged = true;
                    }
                }
                else
                {
                    ssToSave.CustomerCode = rowData.CustomerCode;
                    ssToSave.TransCd = rowData.TransCd;
                    ssToSave.Weekday = (DayOfWeek)dbWeekday; // **CAST VÀO ENUM DayOfWeek**
                    ssToSave.CutOffTime = rowData.CutOffTime;
                    ssToSave.Description = "";
                    ssToSave.CreatedBy = "ExcelImport";
                    ssChanged = true;
                }
                if (ssChanged)
                {
                    bool success;
                    if (isNewSs)
                    {
                        success = await _shippingScheduleRepository.CreateShippingSchedule(ssToSave);
                    }
                    else
                    {
                        var res = await _shippingScheduleRepository.UpdateShippingScheduleByKey(
                            rowData.CustomerCode, rowData.TransCd, (DayOfWeek)dbWeekday, existingSs); // Truyền DayOfWeek đúng
                        success = res.Success;
                    }
                    if (success)
                    {
                        results.ShippingSchedulesAdded++;
                        if (isNewSs)
                        {
                            existingSchedules[ssKey] = ssToSave;
                        }
                    }
                    else
                    {
                        results.Errors.Add($"Failed to {(isNewSs ? "create" : "update")} shipping schedule {ssKey} (Excel Weekday={excelWeekday}).");
                    }
                }
                processedSsKeys.Add(ssKey); // Mark as processed
            }
        }
        private record ImportRow(string CustomerCode, string CustomerName, string TransCd, decimal CollectTimePerPallet, decimal PrepareTimePerPallet, decimal LoadingTimePerColumn, int Weekday, TimeOnly CutOffTime);
        [HttpGet]
        public async Task<IActionResult> DownloadCustomerTemplate(string customerCodes)
        {
            if (string.IsNullOrEmpty(customerCodes))
            {
                customerCodes = "";
            }
            var codeList = customerCodes.Split(',')
                                        .Where(c => !string.IsNullOrEmpty(c.Trim()))
                                        .Select(c => c.Trim())
                                        .Distinct()
                                        .ToList();
            List<Customer> customers;
            if (!codeList.Any())
            {
                customers = new List<Customer>
                {
                    new Customer { CustomerCode = "SAMPLE", CustomerName = "Sample Customer" }
                };
            }
            else
            {
                customers = new List<Customer>();
                foreach (var code in codeList)
                {
                    var customer = await _customerRepository.GetCustomerByCode(code);
                    if (customer != null)
                    {
                        customers.Add(customer);
                    }
                }
            }
            if (!customers.Any())
            {
                return NotFound("No valid customers found.");
            }
            ExcelPackage.License.SetNonCommercialPersonal("Your Name");
            var stream = new MemoryStream();
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("ShippingSchedule");
                worksheet.View.RightToLeft = false;
                worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                string titleText;
                if (codeList.Any())
                {
                    if (customers.Count == 1)
                    {
                        titleText = $"CUSTOMER SHIPPING SCHEDULE - {customers[0].CustomerName} ({customers[0].CustomerCode})";
                    }
                    else
                    {
                        titleText = "MULTIPLE CUSTOMER SHIPPING SCHEDULES";
                    }
                }
                else
                {
                    titleText = "SHIPPING SCHEDULE TEMPLATE";
                }
                worksheet.Cells[1, 1].Value = titleText;
                using (var titleRange = worksheet.Cells[1, 1, 1, 8])
                {
                    titleRange.Merge = true;
                    titleRange.Style.Font.Bold = true;
                    titleRange.Style.Font.Size = 16;
                    titleRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    titleRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    titleRange.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                }
                int headerRow = 3;
                string[] headers = { "Customer Code", "Customer Name", "TransCode", "CollectTimePerPallet (minute)",
                  "PrepareTimePerPallet (minute)", "LoadingTimePerPallet (minute)", "Day of the week", "CutOffTime" };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = worksheet.Cells[headerRow, i + 1];
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                int currentRow = headerRow + 1;
                foreach (var customer in customers)
                {
                    List<LeadtimeMaster> leadtimes;
                    List<ShippingSchedule> schedules;
                    if (customer.CustomerCode == "SAMPLE")
                    {
                        leadtimes = new List<LeadtimeMaster>();
                        schedules = new List<ShippingSchedule>();
                    }
                    else
                    {
                        leadtimes = await _leadtimeRepository.GetAllLeadtimesByCustomer(customer.CustomerCode);
                        schedules = await _shippingScheduleRepository.GetAllShippingSchedulesByCustomer(customer.CustomerCode);
                    }
                    var leadtimeDict = leadtimes.GroupBy(l => l.TransCd)
                                                .ToDictionary(g => g.Key, g => g.First());
                    var scheduleDict = schedules.GroupBy(s => new { s.TransCd, s.Weekday })
                                                .ToDictionary(g => (g.Key.TransCd, (int)g.Key.Weekday), g => g.First().CutOffTime);
                    var transCodes = leadtimeDict.Keys.Union(scheduleDict.Keys.Select(k => k.Item1)).OrderBy(t => t).ToList();
                    if (!transCodes.Any())
                    {
                        transCodes.Add("TRANS_DEFAULT");
                    }
                    int customerStartRow = currentRow;
                    foreach (var transCode in transCodes)
                    {
                        var lt = leadtimeDict.TryGetValue(transCode, out var l) ? l : new LeadtimeMaster
                        {
                            CollectTimePerPallet = 0.0,
                            PrepareTimePerPallet = 0.0,
                            LoadingTimePerColumn = 0.0
                        };
                        int transStartRow = currentRow;
                        for (int day = 1; day <= 6; day++) // Monday=1 to Saturday=6
                        {
                            int weekday = day + 1; // 2=Monday to 7=Saturday
                            double? timeFraction = null;
                            if (scheduleDict.TryGetValue((transCode, day), out var c))
                            {
                                // Robust fraction: Sử dụng TimeSpan để tránh float error
                                var ts = new TimeSpan(c.Hour, c.Minute, 0);
                                timeFraction = ts.TotalDays;
                            }
                            int row = currentRow;
                            int col = 1;
                            bool isFirstCustomerRow = (row == customerStartRow);
                            bool isFirstTransRow = (day == 1);
                            if (isFirstCustomerRow)
                            {
                                worksheet.Cells[row, col++].Value = customer.CustomerCode;
                                worksheet.Cells[row, col++].Value = customer.CustomerName;
                            }
                            else
                            {
                                col += 2;
                            }
                            if (isFirstTransRow)
                            {
                                worksheet.Cells[row, col++].Value = transCode;
                                worksheet.Cells[row, col++].Value = lt.CollectTimePerPallet;
                                worksheet.Cells[row, col++].Value = lt.PrepareTimePerPallet;
                                worksheet.Cells[row, col++].Value = lt.LoadingTimePerColumn;
                            }
                            else
                            {
                                col += 4;
                            }
                            worksheet.Cells[row, col++].Value = weekday;
                            int timeCol = col;
                            if (timeFraction.HasValue)
                            {
                                worksheet.Cells[row, timeCol].Value = timeFraction.Value;
                            }
                            else
                            {
                                // Default 12:00 for new
                                worksheet.Cells[row, timeCol].Value = 12.0 / 24.0;
                            }
                            using (var range = worksheet.Cells[row, 1, row, headers.Length])
                            {
                                range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                            currentRow++;
                        }
                        int transEndRow = currentRow - 1;
                        if (transStartRow < transEndRow)
                        {
                            worksheet.Cells[transStartRow, 3, transEndRow, 3].Merge = true;
                            worksheet.Cells[transStartRow, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[transStartRow, 4, transEndRow, 4].Merge = true;
                            worksheet.Cells[transStartRow, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[transStartRow, 5, transEndRow, 5].Merge = true;
                            worksheet.Cells[transStartRow, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[transStartRow, 6, transEndRow, 6].Merge = true;
                            worksheet.Cells[transStartRow, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }
                    int customerEndRow = currentRow - 1;
                    if (customerStartRow < customerEndRow)
                    {
                        worksheet.Cells[customerStartRow, 1, customerEndRow, 1].Merge = true;
                        worksheet.Cells[customerStartRow, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[customerStartRow, 2, customerEndRow, 2].Merge = true;
                        worksheet.Cells[customerStartRow, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                }
                worksheet.Column(1).Width = 15;
                worksheet.Column(2).Width = 30;
                worksheet.Column(3).Width = 20;
                worksheet.Column(4).Width = 20;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 20;
                worksheet.Column(7).Width = 15;
                worksheet.Column(8).Width = 15;
                int dataStartRow = headerRow + 1;
                int dataEndRow = currentRow - 1;
                if (dataStartRow <= dataEndRow)
                {
                    using (var numRange = worksheet.Cells[dataStartRow, 4, dataEndRow, 6])
                    {
                        numRange.Style.Numberformat.Format = "0.00";
                    }
                    using (var timeRange = worksheet.Cells[dataStartRow, 8, dataEndRow, 8])
                    {
                        timeRange.Style.Numberformat.Format = "hh:mm";
                    }
                    using (var wdRange = worksheet.Cells[dataStartRow, 7, dataEndRow, 7])
                    {
                        wdRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                // Embed VBA Macro (giữ nguyên)
                package.Workbook.CreateVBAProject();
                var vbaCode = new StringBuilder();
                vbaCode.AppendLine("Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)");
                vbaCode.AppendLine(" If Target.Column = 3 And Target.Row > 3 Then");
                vbaCode.AppendLine(" Cancel = True");
                vbaCode.AppendLine(" Call AddNewTransCode(Target.Row)");
                vbaCode.AppendLine(" ElseIf (Target.Column = 1 Or Target.Column = 2) And Target.Row > 3 Then");
                vbaCode.AppendLine(" Cancel = True");
                vbaCode.AppendLine(" Call AddNewCustomer(Target.Row)");
                vbaCode.AppendLine(" End If");
                vbaCode.AppendLine("End Sub");
                vbaCode.AppendLine("Sub AddNewTransCode(startRow As Long)");
                vbaCode.AppendLine(" Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(\"ShippingSchedule\")");
                vbaCode.AppendLine(" Dim transArea As Range: Set transArea = ws.Cells(startRow, 3).MergeArea");
                vbaCode.AppendLine(" Dim transStartRow As Long, transEndRow As Long: transStartRow = transArea.Row: transEndRow = transStartRow + transArea.Rows.Count - 1");
                vbaCode.AppendLine(" Dim customerStartRow As Long: customerStartRow = ws.Cells(startRow, 1).MergeArea.Row");
                vbaCode.AppendLine(" Dim insertPos As Long: insertPos = transEndRow + 1");
                vbaCode.AppendLine(" ws.Rows(insertPos & \":\" & insertPos + 5).Insert Shift:=xlDown");
                vbaCode.AppendLine(" Dim newStartRow As Long: newStartRow = insertPos");
                vbaCode.AppendLine(" Dim i As Long");
                vbaCode.AppendLine(" For i = 0 To 5");
                vbaCode.AppendLine(" Dim currentRow As Long: currentRow = newStartRow + i");
                vbaCode.AppendLine(" If i = 0 Then");
                vbaCode.AppendLine(" ws.Cells(currentRow, 3).Value = \"TRANS_NEW\"");
                vbaCode.AppendLine(" ws.Cells(currentRow, 4).Value = 1.5");
                vbaCode.AppendLine(" ws.Cells(currentRow, 5).Value = 2.0");
                vbaCode.AppendLine(" ws.Cells(currentRow, 6).Value = 0.5");
                vbaCode.AppendLine(" End If");
                vbaCode.AppendLine(" ws.Cells(currentRow, 7).Value = 2 + i");
                vbaCode.AppendLine(" With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8))");
                vbaCode.AppendLine(" .Borders.LineStyle = xlContinuous");
                vbaCode.AppendLine(" .Borders.Weight = xlThin");
                vbaCode.AppendLine(" .VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" End With");
                vbaCode.AppendLine(" Next i");
                vbaCode.AppendLine(" Dim transEndRowNew As Long: transEndRowNew = newStartRow + 5");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 3), ws.Cells(transEndRowNew, 3)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 4), ws.Cells(transEndRowNew, 4)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 5), ws.Cells(transEndRowNew, 5)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 6), ws.Cells(transEndRowNew, 6)).Merge");
                vbaCode.AppendLine(" ws.Cells(newStartRow, 3).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" Dim customerEndRow As Long: customerEndRow = transEndRowNew");
                vbaCode.AppendLine(" ws.Range(ws.Cells(customerStartRow, 1), ws.Cells(customerEndRow, 1)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(customerStartRow, 2), ws.Cells(customerEndRow, 2)).Merge");
                vbaCode.AppendLine(" ws.Cells(customerStartRow, 1).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" ws.Cells(customerStartRow, 2).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 4), ws.Cells(transEndRowNew, 6)).NumberFormat = \"0.00\"");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 8), ws.Cells(transEndRowNew, 8)).NumberFormat = \"hh:mm\"");
                vbaCode.AppendLine(" ws.Columns(\"A:H\").AutoFit");
                vbaCode.AppendLine(" MsgBox \"Duong Minh Da Phu Phep Ra Row Transcode moi:))\"");
                vbaCode.AppendLine("End Sub");
                vbaCode.AppendLine("Sub AddNewCustomer(startRow As Long)");
                vbaCode.AppendLine(" Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(\"ShippingSchedule\")");
                vbaCode.AppendLine(" Dim customerArea As Range: Set customerArea = ws.Cells(startRow, 1).MergeArea");
                vbaCode.AppendLine(" Dim customerStartRow As Long, customerEndRow As Long: customerStartRow = customerArea.Row: customerEndRow = customerStartRow + customerArea.Rows.Count - 1");
                vbaCode.AppendLine(" Dim insertPos As Long: insertPos = customerEndRow + 1");
                vbaCode.AppendLine(" ws.Rows(insertPos & \":\" & insertPos + 5).Insert Shift:=xlDown");
                vbaCode.AppendLine(" Dim newStartRow As Long: newStartRow = insertPos");
                vbaCode.AppendLine(" Dim i As Long");
                vbaCode.AppendLine(" For i = 0 To 5");
                vbaCode.AppendLine(" Dim currentRow As Long: currentRow = newStartRow + i");
                vbaCode.AppendLine(" If i = 0 Then");
                vbaCode.AppendLine(" ws.Cells(currentRow, 1).Value = \"CUST_NEW\"");
                vbaCode.AppendLine(" ws.Cells(currentRow, 2).Value = \"New Customer Name\"");
                vbaCode.AppendLine(" ws.Cells(currentRow, 3).Value = \"TRANS_DEFAULT\"");
                vbaCode.AppendLine(" ws.Cells(currentRow, 4).Value = 1.5");
                vbaCode.AppendLine(" ws.Cells(currentRow, 5).Value = 2.0");
                vbaCode.AppendLine(" ws.Cells(currentRow, 6).Value = 0.5");
                vbaCode.AppendLine(" End If");
                vbaCode.AppendLine(" ws.Cells(currentRow, 7).Value = 2 + i");
                vbaCode.AppendLine(" With ws.Range(ws.Cells(currentRow, 1), ws.Cells(currentRow, 8))");
                vbaCode.AppendLine(" .Borders.LineStyle = xlContinuous");
                vbaCode.AppendLine(" .Borders.Weight = xlThin");
                vbaCode.AppendLine(" .VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" End With");
                vbaCode.AppendLine(" Next i");
                vbaCode.AppendLine(" Dim newEndRow As Long: newEndRow = newStartRow + 5");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 1), ws.Cells(newEndRow, 1)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 2), ws.Cells(newEndRow, 2)).Merge");
                vbaCode.AppendLine(" ws.Cells(newStartRow, 1).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" ws.Cells(newStartRow, 2).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 3), ws.Cells(newEndRow, 3)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 4), ws.Cells(newEndRow, 4)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 5), ws.Cells(newEndRow, 5)).Merge");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 6), ws.Cells(newEndRow, 6)).Merge");
                vbaCode.AppendLine(" ws.Cells(newStartRow, 3).VerticalAlignment = xlCenter");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 4), ws.Cells(newEndRow, 6)).NumberFormat = \"0.00\"");
                vbaCode.AppendLine(" ws.Range(ws.Cells(newStartRow, 8), ws.Cells(newEndRow, 8)).NumberFormat = \"hh:mm\"");
                vbaCode.AppendLine(" ws.Columns(\"A:H\").AutoFit");
                vbaCode.AppendLine(" MsgBox \"Duong Minh Da Phu Phep Ra Row Customer moi :))\"");
                vbaCode.AppendLine("End Sub");
                worksheet.CodeModule.Code = vbaCode.ToString();
                package.Save();
            }
            stream.Position = 0;
            string fileName;
            if (codeList.Any())
            {
                if (customers.Count == 1)
                {
                    fileName = $"ShippingSchedule_{customers[0].CustomerCode}_{DateTime.Now:yyyyMMdd}.xlsm";
                }
                else
                {
                    fileName = $"ShippingSchedules_Multiple_{DateTime.Now:yyyyMMdd}.xlsm";
                }
            }
            else
            {
                fileName = $"ShippingSchedule_Template.xlsm";
            }
            return File(stream.ToArray(), "application/vnd.ms-excel.sheet.macroEnabled.12", fileName);
        }
        [HttpGet]
        public IActionResult DownloadTemplate()
        {
            return RedirectToAction("DownloadCustomerTemplate");
        }
    }
    public class ImportResult
    {
        public int CustomersAdded { get; set; }
        public int LeadtimesAdded { get; set; }
        public int ShippingSchedulesAdded { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }
}