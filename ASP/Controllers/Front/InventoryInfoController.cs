using StockControl.BaseCommon;
using StockControl.Models.Front.InventoriesInfo;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using MimeKit;
using Spire.Xls;
using StockControl.BaseCommon;
using StockControl.Models;
using StockControl.Models.Admin.Categories;
using StockControl.Models.Front.InventoriesInfo;
using System.Reflection;

namespace StockControl.Controllers.Front
{
    public class InventoryInfoController : Controller, InventoryInfoListenerInterface
    {
        private readonly ILogger<InventoryInfoController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IAuthorizationService _authService;
        private readonly StockControlContext _context;
        public InventoryInfoRepositoryInterface _inventory;
        protected string photosPath;
        private readonly IWebHostEnvironment _env;
        public BaseController _baseController;
        public CategoriesRepositoryInterface categories;
        public InventoryInfoController(ILogger<InventoryInfoController> logger, IWebHostEnvironment webHostEnvironment, IAuthorizationService authService, StockControlContext context, IWebHostEnvironment env, InventoryInfoRepositoryInterface inventory, BaseController baseController, CategoriesRepositoryInterface categories)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            _authService = authService;
            _context = context;
            _inventory = inventory;
            _baseController = baseController;
            _env = env;
            photosPath = _env.WebRootPath + "/templates/export_excel";
            this.categories = categories;
        }
        public async Task<IActionResult> Index(DateTime? fProductDate, string fPartNo, string fTypeProduct, string fSupplierCode, DateTime? fShortageDateTotal1A, string fLineName, int fStatus = 0, int pagesize = 10, int page = 1, string sort = "-ProductDate")
        {
            //get taxonomy 
            var sltWareHouse = this.categories.GetTaxonomyByDropdown((int)EnumTaxonomy.WareHouse, true);
            ViewBag.sltWareHouse = sltWareHouse;
            //loai linh kien
            var sltTypeProduct = this.categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Category, true);
            sltTypeProduct.Insert(0, new SelectListItem { Value = "null", Text = "Null" });
            ViewBag.sltTypeProduct = sltTypeProduct;
            //get taxonomy 
            //Dictionary<string, string> dicInit = new Dictionary<string, string>();
            ViewData["DataPic"] = this.categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Pic, false).ToDictionary(x => x.Value, x => x.Text);
            ViewData["DataReason"] = this.categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Reason, false).ToDictionary(x => x.Value, x => x.Text);
            //get taxonomy 
            var sltReason = this.categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Reason, false);
            ViewBag.sltReason = sltReason;
            //
            var list = _inventory.GetAllByLimit(fProductDate, fPartNo, fTypeProduct, fSupplierCode, fShortageDateTotal1A, fLineName, fStatus, pagesize, page, sort).Result;
            return View("../Front/InventoryInfo/Index", list);
        }
        //
        [HttpGet]
        //[ValidateAntiForgeryToken]
        //[Route("ExportData", Name = "front.inventoryinfos.exportdata")]
        public async Task<JsonResult> ExportData(DateTime? fProductDate, string fPartNo = null, string fTypeProduct = null, string fSupplierCode = null, DateTime? fShortageDateTotal1A = null, string fLineName = null, int fStatus = 0)
        {
            try
            {
                #region validate
                if (fProductDate == null)
                {
                    return Json(new { result = false, fileName = "", rootsPath = this.photosPath, message = "Ngày không để trống." });
                }
                #endregion
                var fileName = "Excel_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                var filePathNew = Path.Combine(this.photosPath, fileName);
                var fileSource = Path.Combine(this.photosPath, "export_front_inventory_info.xlsx");
                //
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(fileSource);
                Worksheet sheet = workbook.Worksheets[0];

                #region export data into row
                var list = await _inventory.GetAll(fProductDate, fPartNo, fTypeProduct, fSupplierCode, fShortageDateTotal1A, fLineName, fStatus);
                // add data into excel file 
                int cRow = 2, count = 1;
                if (list != null)
                {
                    foreach (var item in list)
                    {
                        sheet.Range["A" + cRow].Text = count.ToString();
                        sheet.Range["B" + cRow].Text = item.ProductDate.ToString("dd/MM/yyyy");
                        sheet.Range["C" + cRow].Text = item.PartNo ?? "";
                        sheet.Range["D" + cRow].Text = item.PartName;
                        sheet.Range["E" + cRow].Text = item.TypeProduct ?? "";
                        sheet.Range["F" + cRow].Text = item.SupplierCode ?? "";
                        //
                        //sheet.Range["G" + cRow].Text = item.StockQtyTotal1A.ToString() ?? "";
                        //sheet.Range["H" + cRow].Text = item.StockDOHTotal1A.ToString() ?? "";
                        sheet.Range["G" + cRow].NumberValue = item.StockQtyTotal1A ?? 0;
                        sheet.Range["H" + cRow].NumberValue = item.StockDOHTotal1A ?? 0;
                        //
                        if (item.StockQtyTotal1A.Value != null && item.Pallet.Value != null)
                        {
                            var palletQty = (double.IsInfinity(item.StockQtyTotal1A.Value / item.Pallet.Value) == true || double.IsNaN(item.StockQtyTotal1A.Value / item.Pallet.Value) == true) ? 0 : Math.Round(item.StockQtyTotal1A.Value / item.Pallet.Value, 2);
                            //sheet.Range["I" + cRow].Text = palletQty.ToString() ?? "";
                            sheet.Range["I" + cRow].NumberValue = palletQty;
                        }
                        //
                        sheet.Range["J" + cRow].Text = (item.NextReceiving != null) ? item.NextReceiving.Value.ToString("dd/MM/yyyy") : "";
                        sheet.Range["K" + cRow].Text = (item.ShortageDateTotal1A != null) ? item.ShortageDateTotal1A.Value.ToString("dd/MM/yyyy") : "";
                        //sheet.Range["L" + cRow].Text = item.ShortageQtyTotal1A.ToString() ?? "";
                        sheet.Range["L" + cRow].NumberValue = item.ShortageQtyTotal1A ?? 0;
                        //
                        string strStatus = "";
                        if (item.Status == (int)EnumStatusInventory.InActive)
                        {
                            strStatus = "Low";
                        }
                        else if (item.Status == (int)EnumStatusInventory.Pending)
                        {
                            strStatus = "High";
                        }
                        else if (item.Status == (int)EnumStatusInventory.Active)
                        {
                            strStatus = "Normal";
                        }
                        sheet.Range["M" + cRow].Text = strStatus;
                        //
                        var dataPic = categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Pic, false).ToDictionary(x => x.Value, x => x.Text);
                        if (item.Pic != null && dataPic != null)
                        {
                            if (dataPic.ContainsKey(item.Pic.Value.ToString()))
                            {
                                sheet.Range["N" + cRow].Text = dataPic[item.Pic.Value.ToString()];
                            }
                        }
                        //
                        var dataReason = categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Reason, false).ToDictionary(x => x.Value, x => x.Text);
                        if (item.Reason != null && dataReason != null)
                        {
                            if (dataReason.ContainsKey(item.Reason.Value.ToString()))
                            {
                                sheet.Range["O" + cRow].Text = dataReason[item.Reason.Value.ToString()];
                            }
                        }
                        //
                        sheet.Range["P" + cRow].Text = item.Remark ?? "";
                        sheet.Range["Q" + cRow].Text = item.LineName ?? "";
                        cRow++;
                        count++;
                    }
                }
                #endregion
                workbook.SaveToFile(filePathNew, ExcelVersion.Version2016);
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
