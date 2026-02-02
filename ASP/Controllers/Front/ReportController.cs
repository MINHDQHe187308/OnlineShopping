using StockControl.BaseCommon;
using StockControl.Controllers.Admin;
using StockControl.Models.Front.Reports;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MimeKit;
using Newtonsoft.Json;
using Spire.Xls;
using StockControl.BaseCommon;
using StockControl.Models;
using System.Reflection;

namespace StockControl.Controllers.Front
{
    public class ReportController : Controller, ReportListenerInterface
    {
        private readonly ILogger<ReportController> _logger;
        private readonly IAuthorizationService _authService;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly StockControlContext _context;
        public BaseController _baseControllerI;
        protected string photosPath;
        private readonly IWebHostEnvironment _env;
        public ReportRepositoryInterface _report;
        public ReportController(ILogger<ReportController> logger, IAuthorizationService authService, IWebHostEnvironment webHostEnvironment, StockControlContext context, BaseController baseControllerI, ReportRepositoryInterface report, IWebHostEnvironment env)
        {
            _logger = logger;
            _authService = authService;
            _webHostEnvironment = webHostEnvironment;
            _context = context;
            _baseControllerI = baseControllerI;
            _report = report;
            _env = env;
            photosPath = _env.WebRootPath + "/templates/export_excel";
        }
        /**
         * linh kien
         * ***/
        public IActionResult Factory(int fYear = 0)
        {
            ViewBag.sltYears = _baseControllerI.GetYears(2022);
            ViewBag.sltMonths = _baseControllerI.GetMonths(true);
            #region day du lieu vao truc X
            List<string> xAxisCategories = new List<string>();
            xAxisCategories.Add($"FY {(fYear - 1)}");
            foreach (var item in _baseControllerI.GetMonths(true))
            {
                xAxisCategories.Add(item.Text);
            }
            ViewBag.xAxisCategories = JsonConvert.SerializeObject(xAxisCategories);
            #endregion

            #region day du lieu vao bieu do
            /**
             * table
             * ***
             * chart
             * **/
            var result = _report.GetStockFactory(fYear);
            var chart = _report.GetChartFactory(fYear);

            #region tinh avg min, max, target cua nam hien tai - 1
            var resultPast = _report.GetStockFactory(fYear - 1);
            double avgMinPast = 0.0, avgMaxPast = 0.0, avgTargetPast = 0.0;
            foreach (var item in resultPast.ListValues)
            {
                if (item.Key == "Min")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgMinPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
                else if (item.Key == "Max")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgMaxPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
                else if (item.Key == "Target")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgTargetPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
            }
            #endregion

            foreach (var item in result.ListValues)
            {
                var seriItem = new SeriesModel<double>();
                //
                seriItem.name = item.Key;
                seriItem.type = "line";
                seriItem.data = item.Value;
                //
                chart.Series.Add(seriItem);
            }
            #region day du lieu nam hien tai - 1, vao bieu do
            var chartPast = _report.GetChartFactory(fYear - 1);
            var dicDataPast = new Dictionary<string, double>();
            foreach (var item in chartPast.Series)
            {
                // ktra tung loai hang
                if (item.type != "line")
                {
                    var avg = (item.data.Count() > 0) ? item.data.Sum() / 12 : 0;
                    dicDataPast.Add(item.name, Math.Round(avg, 2, MidpointRounding.ToEven));
                }

            }
            #endregion

            foreach (var item in chart.Series)
            {
                // ktra tung loai hang: add nam hien tai - 1
                if (dicDataPast.ContainsKey(item.name) == true)
                {
                    // co key
                    item.data.Insert(0, dicDataPast[item.name]);
                }
                else
                {
                    if (item.name == "Min")
                    {
                        item.data.Insert(0, avgMinPast);
                    }
                    else if (item.name == "Max")
                    {
                        item.data.Insert(0, avgMaxPast);
                    }
                    else if (item.name == "Target")
                    {
                        item.data.Insert(0, avgTargetPast);
                    }
                    else
                    {
                        item.data.Insert(0, 0);
                    }
                }


            }
            ViewBag.series = JsonConvert.SerializeObject(chart);
            #endregion
            #region day du lieu vao table
            var old_values = result.ListValues;
            var new_values = new Dictionary<string, List<double>>();
            foreach (var item in chart.Series)
            {
                if (item.type != "line")
                {
                    new_values.Add(item.name, item.data);
                }
            }
            result.ListValues = new_values;
            old_values.ToList().ForEach(x => result.ListValues.Add(x.Key, x.Value));
            #endregion
            #region judge
            result.Evaluations = chart.Evaluations;
            #endregion
            return View("../Front/Reports/Factory", result);
        }
        public IActionResult GetFactoryExplain(int year = 0, int month = 0)
        {
            try
            {
                int nId = 0;
                string explain = "";
                var find = _statistical.GetStatisticalFactoryById(null, year, month, EnumTypeStatistical.Factory.ToString());
                if (find != null)
                {
                    nId = find.ID;
                    explain = (string.IsNullOrEmpty(find.RptExplain)) ? "No content was found." : find.RptExplain;
                }
                else
                {
                    nId = 0;
                    explain = "No content was found.";
                }
                return Json(new { Result = "success", Data = explain, id = nId });
            }
            catch (Exception ex)
            {
                return Json(new { Result = "error", Data = ex.Message, id = 0 });
            }
        }
        /**
         * tung loai hang
         * ***/
        public IActionResult Category(int fYear = 0, string fCat = "")
        {
            ViewBag.sltYears = _baseControllerI.GetYears(2022);
            ViewBag.sltMonths = _baseControllerI.GetMonths(true);
            var sltTypeProduct = _categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Category, true);
            ViewBag.sltTypeProduct = sltTypeProduct;
            ViewBag.CatName = fCat;
            #region day du lieu vao truc X
            List<string> xAxisCategories = new List<string>();
            xAxisCategories.Add($"FY {(fYear - 1)}");
            foreach (var item in _baseControllerI.GetMonths(true))
            {
                xAxisCategories.Add(item.Text);
            }
            ViewBag.xAxisCategories = JsonConvert.SerializeObject(xAxisCategories);
            #endregion

            #region day du lieu vao bieu do
            /**
             * table
             * ***
             * chart
             * **/
            var result = _report.GetStockCategory(fYear, fCat);
            var chart = _report.GetChartCategory(fYear, fCat);

            #region tinh avg min, max, target cua nam hien tai - 1
            var resultPast = _report.GetStockCategory(fYear - 1, fCat);
            double avgMinPast = 0.0, avgMaxPast = 0.0, avgTargetPast = 0.0;
            foreach (var item in resultPast.ListValues)
            {
                if (item.Key == "Min")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgMinPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
                else if (item.Key == "Max")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgMaxPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
                else if (item.Key == "Target")
                {
                    var avgN = item.Value.Sum() / item.Value.Count();
                    avgTargetPast = (Double.IsNaN(avgN)) ? 0 : Math.Round(avgN, 2, MidpointRounding.ToEven);
                }
            }
            #endregion

            foreach (var item in result.ListValues)
            {
                var seriItem = new SeriesModel<double>();
                //
                seriItem.name = item.Key;
                seriItem.type = "line";
                seriItem.data = item.Value;
                //
                chart.Series.Add(seriItem);
            }
            #region day du lieu nam hien tai - 1, vao bieu do
            var chartPast = _report.GetChartCategory(fYear - 1, fCat);
            var dicDataPast = new Dictionary<string, double>();
            foreach (var item in chartPast.Series)
            {
                // ktra tung loai hang
                if (item.type != "line")
                {
                    var avg = (item.data.Count() > 0) ? item.data.Sum() / 12 : 0;
                    dicDataPast.Add(item.name, Math.Round(avg, 2, MidpointRounding.ToEven));
                }

            }
            #endregion

            foreach (var item in chart.Series)
            {
                // ktra tung loai hang: add nam hien tai - 1
                if (dicDataPast.ContainsKey(item.name) == true)
                {
                    // co key
                    item.data.Insert(0, dicDataPast[item.name]);
                }
                else
                {
                    if (item.name == "Min")
                    {
                        item.data.Insert(0, avgMinPast);
                    }
                    else if (item.name == "Max")
                    {
                        item.data.Insert(0, avgMaxPast);
                    }
                    else if (item.name == "Target")
                    {
                        item.data.Insert(0, avgTargetPast);
                    }
                    else
                    {
                        item.data.Insert(0, 0);
                    }
                }


            }
            ViewBag.series = JsonConvert.SerializeObject(chart);
            #endregion
            #region day du lieu vao table
            var old_values = result.ListValues;
            var new_values = new Dictionary<string, List<double>>();
            foreach (var item in chart.Series)
            {
                if (item.type != "line")
                {
                    new_values.Add(item.name, item.data);
                }
            }
            result.ListValues = new_values;
            old_values.ToList().ForEach(x => result.ListValues.Add(x.Key, x.Value));
            #endregion
            #region judge
            result.Evaluations = chart.Evaluations;
            #endregion
            return View("../Front/Reports/Category", result);
        }
        public IActionResult GetCategoryExplain(int year = 0, int month = 0, string cat = "")
        {
            try
            {
                int nId = 0;
                string explain = "";
                var find = _statistical.GetStatisticalCatById(null, year, month, EnumTypeStatistical.Category.ToString(), cat);
                if (find != null)
                {
                    nId = find.ID;
                    explain = (string.IsNullOrEmpty(find.RptExplain)) ? "No content was found." : find.RptExplain;
                }
                else
                {
                    nId = 0;
                    explain = "No content was found.";
                }
                return Json(new { Result = "success", Data = explain, id = nId });
            }
            catch (Exception ex)
            {
                return Json(new { Result = "error", Data = ex.Message, id = 0 });
            }
        }
        /**
         * tung nha cung cap
         * ***/
        public IActionResult Supplier(int fYear = 0, string fSupplier = "")
        {
            ViewBag.sltYears = _baseControllerI.GetYears(2022);
            ViewBag.sltTypeProduct = _product.GetSupplierCode();
            ViewBag.CatName = fSupplier;
            #region day du lieu vao truc X
            List<string> xAxisCategories = new List<string>();
            foreach (var item in _baseControllerI.GetMonths(true))
            {
                xAxisCategories.Add(item.Text);
            }
            xAxisCategories.Add($"FY {fYear}");
            ViewBag.xAxisCategories = JsonConvert.SerializeObject(xAxisCategories);
            #endregion

            #region day du lieu vao bieu do
            /**
             * chart
             * **/
            var chart = _report.GetChartSupplier(fYear, fSupplier);
            #region day du lieu AVG cua nam vao bieu do
            foreach (var item in chart.Series)
            {
                // ktra tung loai hang
                if (item.type != "line")
                {
                    var avg = (item.data.Count() > 0) ? item.data.Where(w => w > 0).Sum() / item.data.Where(w => w > 0).Count() : 0;
                    item.data.Add(Math.Round(avg, 2, MidpointRounding.ToEven));
                }

            }
            #endregion
            ViewBag.series = JsonConvert.SerializeObject(chart);
            #endregion

            return View("../Front/Reports/Supplier");
        }
        //
        [HttpGet]
        //[ValidateAntiForgeryToken]
        //[Route("ExportData", Name = "front.inventoryinfos.exportdata")]
        public async Task<JsonResult> ExportData(int fYear = 0, string fSupplier = "")
        {
            try
            {
                #region validate
                if (fYear == 0)
                {
                    return Json(new { result = false, fileName = "", rootsPath = this.photosPath, message = "Năm không để trống." });
                }
                #endregion

                var sltYears = _baseControllerI.GetYears(2022);
                var sltSupplierCode = _product.GetSupplierCode();
                #region day du lieu vao truc X
                List<string> xAxisCategories = new List<string>();
                foreach (var item in _baseControllerI.GetMonths(true))
                {
                    xAxisCategories.Add(item.Text);
                }
                xAxisCategories.Add($"FY {fYear}");
                ViewBag.xAxisCategories = JsonConvert.SerializeObject(xAxisCategories);
                #endregion

                #region day du lieu vao bieu do
                /**
                 * chart
                 * **/
                var listChart = new List<HighChartModel>();
                if (!string.IsNullOrEmpty(fSupplier))
                {
                    // lay 1 nha cung cap
                    sltSupplierCode = new List<SelectListItem>(){
                        new SelectListItem() { Value = fSupplier, Text = fSupplier }
                    };
                }
                foreach (var itemSupplier in sltSupplierCode)
                {
                    var chart = _report.GetChartSupplier(fYear, itemSupplier.Text);
                    #region day du lieu AVG cua nam vao bieu do
                    foreach (var item in chart.Series)
                    {
                        // ktra tung loai hang
                        if (item.type != "line")
                        {
                            var avg = (item.data.Count() > 0) ? item.data.Where(w => w > 0).Sum() / item.data.Where(w => w > 0).Count() : 0;
                            if (Double.IsInfinity(avg) || Double.IsNaN(avg))
                            {
                                // Put your logic here.
                                avg = 0.0;
                            }
                            item.data.Add(Math.Round(avg, 2, MidpointRounding.ToEven));
                        }

                    }
                    #endregion
                    listChart.Add(chart);
                    #endregion
                }

                #region Excel Export

                var fileName = "Excel_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                var filePathNew = Path.Combine(this.photosPath, fileName);
                var fileSource = Path.Combine(this.photosPath, "export_front_inventory_info.xlsx");
                //
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(fileSource);
                Worksheet sheet = workbook.Worksheets[1];

                #region export data into row
                // add data into excel file 
                int cRow = 1, countM = 1, countNCC = 1;
                var dicMonth = _baseControllerI.GetMonthExportExcel();
                sheet.Range["A1"].Text = "NCC/Tháng";
                foreach (var item in xAxisCategories)
                {
                    sheet.Range[$"{dicMonth[countM]}" + cRow].Text = item;
                    countM++;
                }
                // NCC
                foreach (var item in listChart)
                {
                    foreach (var itemSeri in item.Series)
                    {
                        cRow++;
                        sheet.Range[$"A" + cRow].Text = itemSeri.name;
                        int countData = 1;
                        foreach (var itemData in itemSeri.data)
                        {
                            //sheet.Range[$"{dicMonth[countData]}" + cRow].Text = itemData.ToString();
                            sheet.Range[$"{dicMonth[countData]}" + cRow].NumberValue = itemData;
                            countData++;
                        }
                    }
                } 
                #endregion
                sheet.Activate();
                workbook.SaveToFile(filePathNew, ExcelVersion.Version2016);
                #endregion
                return Json(new { result = true, fileName = fileName, rootsPath = this.photosPath, message = "" });
            }
            catch (Exception ex)
            {
                _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                return Json(new { result = false, fileName = "", rootsPath = this.photosPath, message = ex.Message });
            }

        }
        [HttpGet]
        [DeleteFile]
        //[Route("DownloadEx", Name = "front.inventoryinfos.downloadex")]
        public IActionResult DownloadEx(string file, string rootsPath)
        {
            try
            {
                if (!string.IsNullOrEmpty(file) && !string.IsNullOrEmpty(rootsPath))
                {
                    string fullPath = Path.Combine(photosPath, file);
                    if (System.IO.File.Exists(fullPath))
                    {
                        return PhysicalFile(fullPath, MimeTypes.GetMimeType(fullPath), Path.GetFileName(fullPath));
                    }
                }
                return new ForbidResult();
            }
            catch (Exception ex)
            {
                _logger.LogError("{0}/{1}: {2}", MethodBase.GetCurrentMethod().DeclaringType, MethodBase.GetCurrentMethod().Name, ex.Message);
                return new ForbidResult();
            }
       
        }
    }
}
