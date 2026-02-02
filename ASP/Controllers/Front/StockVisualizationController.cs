using StockControl.BaseCommon;
using StockControl.Hubs;
using StockControl.Models.Front.InventoriesInfo;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;
using StockControl.BaseCommon;
using StockControl.Models;
using StockControl.Models.Admin.Categories;
using StockControl.Models.Front.InventoriesInfo;
using StockControl.Models.Front.StockVisualizations;
using System.Diagnostics;

namespace StockControl.Controllers.Front
{
    public class StockVisualizationController : Controller
    {
        // GET: StockVisualizationController
        private readonly ILogger<InventoryInfoController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IAuthorizationService _authService;
        private readonly StockControlContext _context;
        public InventoryInfoRepositoryInterface _inventory;
        private readonly IWebHostEnvironment _env;
        public BaseController _baseController;
        public CategoriesRepositoryInterface _categories;
        public StockVisualizationController(ILogger<InventoryInfoController> logger, IWebHostEnvironment webHostEnvironment, IAuthorizationService authService, StockControlContext context, IWebHostEnvironment env, InventoryInfoRepositoryInterface inventory, BaseController baseController, CategoriesRepositoryInterface categories)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            _authService = authService;
            _context = context;
            _inventory = inventory;
            _baseController = baseController;
            _env = env;
            _categories = categories;
        }
        public async Task<IActionResult> Index()
        {
            //
            var obj = new StockVisualization();
            // get kind of parts
            var listCats = _categories.GetAllCategoriesByTax((int)EnumTaxonomy.Category).Select(s => new Categories()
            {
                ID = s.ID,
                Title = s.Title,
                CatsMeta = s.CatsMeta
            }).ToList();
            ViewBag.TotalCats = listCats.Count();
            //
            var listCatTop10 = new Dictionary<string, List<InventoryInfo>>();
            if (listCats.Count() > 0)
            {
                obj.KindOfPart = listCats.Select(s => s.Title).ToList();
                // kho A
                var qryInventoryA = (from p in _context.Products
                                     join m in _context.Inventories on p.ID equals m.ProductRefId
                                     where m.House == "A" &&
                                     m.ProductDate == DateTime.Now.Date &&
                                     m.StockQtyTotal1A > 0 &&
                                     m.NextReceiving != null &&
                                     m.ShortageDate != null &&
                                     m.ShortageQty != null
                                     select new InventoryInfo()
                                     {
                                         ProductDate = m.ProductDate,
                                         PartNo = m.PartNo,
                                         PartName = p.PartName,
                                         TypeProduct = p.TypeProduct,
                                         SupplierCode = p.SupplierCode,
                                         House = m.House,
                                         StockQty = m.StockQty,
                                         StockBegin = m.StockBegin,
                                         InOut = m.InOut,
                                         StockDOH = m.StockDOH,
                                         ShortageDate = m.ShortageDate,
                                         ShortageDateStr = (m.ShortageDate != null) ? m.ShortageDate.Value.ToString("MM/dd/yyyy") : "",
                                         NextReceiving = m.NextReceiving,
                                         ShortageQty = m.ShortageQty,
                                         ProductPlanQty = m.ProductPlanQty,
                                         Mins = m.Mins,
                                         Maxs = m.Maxs,
                                         Pallet = m.Pallet,
                                         PalletPolicy = m.PalletPolicy,
                                         Status = m.Status,
                                         Pic = m.Pic,
                                         Reason = m.Reason,
                                         Fixing = m.Fixing,
                                         Remark = m.Remark,
                                         LineName = m.LineName,
                                         LineNameTop = BaseController.HandleSubString(m.LineName ?? "", p.LineName ?? "", 25),
                                         ProductRefId = m.ProductRefId,
                                         //
                                         StockQtyTotal1A = m.StockQtyTotal1A,
                                         StockDOHTotal1A = m.StockDOHTotal1A,
                                         ShortageDateTotal1A = m.ShortageDateTotal1A,
                                         ShortageQtyTotal1A = m.ShortageQtyTotal1A,
                                         //
                                         CreatedDate = m.CreatedDate,
                                         UpdatedDate = m.UpdatedDate
                                     }).OrderByDescending(f => f.ProductDate).ThenByDescending(f => f.House).ToList();
                // kho 1
                var qryInventory1 = (from p in _context.Products
                                     join m in _context.Inventories on p.ID equals m.ProductRefId
                                     where m.House == "1" &&
                                     m.ProductDate == DateTime.Now.Date &&
                                     m.StockQtyTotal1A > 0 &&
                                     m.NextReceiving != null &&
                                     m.ShortageDate != null &&
                                     m.ShortageQty != null
                                     select new InventoryInfo()
                                     {
                                         ProductDate = m.ProductDate,
                                         PartNo = m.PartNo,
                                         PartName = p.PartName,
                                         TypeProduct = p.TypeProduct,
                                         SupplierCode = p.SupplierCode,
                                         House = m.House,
                                         StockQty = m.StockQty,
                                         StockBegin = m.StockBegin,
                                         InOut = m.InOut,
                                         StockDOH = m.StockDOH,
                                         ShortageDate = m.ShortageDate,
                                         ShortageDateStr = (m.ShortageDate != null) ? m.ShortageDate.Value.ToString("MM/dd/yyyy") : "",
                                         NextReceiving = m.NextReceiving,
                                         ShortageQty = m.ShortageQty,
                                         ProductPlanQty = m.ProductPlanQty,
                                         Mins = m.Mins,
                                         Maxs = m.Maxs,
                                         Pallet = m.Pallet,
                                         PalletPolicy = m.PalletPolicy,
                                         Status = m.Status,
                                         Pic = m.Pic,
                                         Reason = m.Reason,
                                         Fixing = m.Fixing,
                                         Remark = m.Remark,
                                         LineName = m.LineName,
                                         LineNameTop = BaseController.HandleSubString(m.LineName ?? "", p.LineName ?? "", 25),
                                         ProductRefId = m.ProductRefId,
                                         //
                                         StockQtyTotal1A = m.StockQtyTotal1A,
                                         StockDOHTotal1A = m.StockDOHTotal1A,
                                         ShortageDateTotal1A = m.ShortageDateTotal1A,
                                         ShortageQtyTotal1A = m.ShortageQtyTotal1A,
                                         //
                                         CreatedDate = m.CreatedDate,
                                         UpdatedDate = m.UpdatedDate
                                     }).OrderByDescending(f => f.ProductDate).ThenByDescending(f => f.House).ToList();
                // kho R
                var qryInventoryR = (from p in _context.Products
                                     join m in _context.Inventories on p.ID equals m.ProductRefId
                                     where m.House == "R" &&
                                     m.ProductDate == DateTime.Now.Date &&
                                     m.StockQtyTotal1A > 0 &&
                                     m.NextReceiving != null &&
                                     m.ShortageDate != null &&
                                     m.ShortageQty != null
                                     select new InventoryInfo()
                                     {
                                         ProductDate = m.ProductDate,
                                         PartNo = m.PartNo,
                                         PartName = p.PartName,
                                         TypeProduct = p.TypeProduct,
                                         SupplierCode = p.SupplierCode,
                                         House = m.House,
                                         StockQty = m.StockQty,
                                         StockBegin = m.StockBegin,
                                         InOut = m.InOut,
                                         StockDOH = m.StockDOH,
                                         ShortageDate = m.ShortageDate,
                                         ShortageDateStr = (m.ShortageDate != null) ? m.ShortageDate.Value.ToString("MM/dd/yyyy") : "",
                                         NextReceiving = m.NextReceiving,
                                         ShortageQty = m.ShortageQty,
                                         ProductPlanQty = m.ProductPlanQty,
                                         Mins = m.Mins,
                                         Maxs = m.Maxs,
                                         Pallet = m.Pallet,
                                         PalletPolicy = m.PalletPolicy,
                                         Status = m.Status,
                                         Pic = m.Pic,
                                         Reason = m.Reason,
                                         Fixing = m.Fixing,
                                         Remark = m.Remark,
                                         LineName = m.LineName,
                                         LineNameTop = BaseController.HandleSubString(m.LineName ?? "", p.LineName ?? "", 25),
                                         ProductRefId = m.ProductRefId,
                                         //
                                         StockQtyTotal1A = m.StockQtyTotal1A,
                                         StockDOHTotal1A = m.StockDOHTotal1A,
                                         ShortageDateTotal1A = m.ShortageDateTotal1A,
                                         ShortageQtyTotal1A = m.ShortageQtyTotal1A,
                                         //
                                         CreatedDate = m.CreatedDate,
                                         UpdatedDate = m.UpdatedDate
                                     }).OrderByDescending(f => f.ProductDate).ThenByDescending(f => f.House).ToList();
                //
                var qryInventory1Top = (from p in _context.Products
                                        join m in _context.Inventories on p.ID equals m.ProductRefId
                                        where
                                        m.House == "1" &&
                                        m.ProductDate == DateTime.Now.Date &&
                                        m.Status == (int)EnumStatusInventory.InActive
                                        select new InventoryInfo()
                                        {
                                            ProductDate = m.ProductDate,
                                            PartNo = m.PartNo,
                                            PartName = p.PartName,
                                            TypeProduct = p.TypeProduct,
                                            SupplierCode = p.SupplierCode,
                                            House = m.House,
                                            StockQty = m.StockQty,
                                            StockBegin = m.StockBegin,
                                            InOut = m.InOut,
                                            StockDOH = m.StockDOH,
                                            ShortageDate = m.ShortageDate,
                                            ShortageDateStr = (m.ShortageDate != null) ? m.ShortageDate.Value.ToString("MM/dd/yyyy") : "",
                                            NextReceiving = m.NextReceiving,
                                            ShortageQty = m.ShortageQty,
                                            ProductPlanQty = m.ProductPlanQty,
                                            Mins = m.Mins,
                                            Maxs = m.Maxs,
                                            Pallet = m.Pallet,
                                            PalletPolicy = m.PalletPolicy,
                                            Status = m.Status,
                                            Pic = m.Pic,
                                            Reason = m.Reason,
                                            Fixing = m.Fixing,
                                            Remark = m.Remark,
                                            LineName = m.LineName,
                                            LineNameTop = BaseController.HandleSubString(m.LineName ?? "", p.LineName ?? "", 25),
                                            ProductRefId = m.ProductRefId,
                                            //
                                            StockQtyTotal1A = m.StockQtyTotal1A,
                                            StockDOHTotal1A = m.StockDOHTotal1A,
                                            ShortageDateTotal1A = m.ShortageDateTotal1A,
                                            ShortageQtyTotal1A = m.ShortageQtyTotal1A,
                                            //
                                            CreatedDate = m.CreatedDate,
                                            UpdatedDate = m.UpdatedDate
                                        }).OrderByDescending(f => f.ProductDate).ThenByDescending(f => f.House).ToList();
                var listVisualizeStocks2 = new List<StockDetail>();
                //
                var listVisualizeDohs = new List<double>();
                var listVisualizeDOHPOLICY = new List<string>();
                var listVisualizePIC = new List<string>();
                //
                var listVisualizePolicy1 = new List<double>();
                var listVisualizeActual1 = new List<double>();
                var listVisualizePolicyA = new List<double>();
                var listVisualizeActualA = new List<double>();
                var listVisualizeActualR = new List<double>();
                //
                var listVisualizeLine = new List<string>();
                foreach (var item in listCats)
                {
                    string catStock = "";
                    int itemStock = (int)EnumStatusInventory.Active;//10;
                    double itemDOH = 0.0;
                    string itemDOHPOLICY = "";
                    string itemPIC = "";
                    //
                    double itemPolicy1 = 0.0;
                    double itemActual1 = 0.0;
                    double itemPolicyA = 0.0;
                    double itemActualA = 0.0;
                    double itemActualR = 0.0;
                    //
                    string itemLine = "";
                    var objTypeProduct = qryInventory1.Where(w => w.TypeProduct == item.Title).ToList();//
                    var objTypeProductA = qryInventoryA.Where(w => w.TypeProduct == item.Title).ToList();//
                    var objTypeProductR = qryInventoryR.Where(w => w.TypeProduct == item.Title).ToList();//
                    if (objTypeProduct.Count() > 0)
                    {
                        // DOH
                        var totalStockQty = objTypeProduct.Sum(s => s.StockQtyTotal1A);
                        var require30 = objTypeProduct.Sum(f => f.ProductPlanQty);
                        itemDOH = Math.Round(Convert.ToDouble(totalStockQty) / Convert.ToDouble(require30) * 24, 1);
                        if (double.IsNaN(itemDOH) || double.IsInfinity(itemDOH))
                        {
                            itemDOH = 0;
                        }
                        // Stock
                        catStock = item.Title;
                        // check Min Category & DOH: category(min trong category) > min DOH tổng kho
                        double catMin = 0.0, catMax = 0.0;
                        if (item.CatsMeta.Count() > 0)
                        {
                            // min
                            var findMinCat = item.CatsMeta.FirstOrDefault(f => f.MetaKey == "Min" && f.CatId == item.ID);
                            if (findMinCat != null)
                            {
                                catMin = Convert.ToDouble(findMinCat.MetaValue);
                            }
                            //
                            var findMaxCat = item.CatsMeta.FirstOrDefault(f => f.MetaKey == "Max" && f.CatId == item.ID);
                            if (findMaxCat != null)
                            {
                                catMax = Convert.ToDouble(findMaxCat.MetaValue);
                            }
                        }
                        // < 08.02.2023
                        //if (objTypeProduct.Any(f => f.Status == (int)EnumStatusInventory.InActive) || catMin > itemDOH)
                        //{
                        //    itemStock = (int)EnumStatusInventory.InActive; //- 5;
                        //}
                        //else if (catMax < itemDOH)
                        //{
                        //    //objTypeProduct.Any(f => f.Status == (int)EnumStatusInventory.Pending)
                        //    itemStock = (int)EnumStatusInventory.Pending;//5;
                        //}

                        // doi dieu kien, ap dung tu ngay: > 08.02.2023
                        if (catMin > itemDOH)
                        {
                            itemStock = (int)EnumStatusInventory.InActive; //- 5;
                        }
                        else if (catMax < itemDOH)
                        {
                            //objTypeProduct.Any(f => f.Status == (int)EnumStatusInventory.Pending)
                            itemStock = (int)EnumStatusInventory.Pending;//5;
                        }
                        // DOH POLICY
                        itemDOHPOLICY = $"{catMin} - {catMax}";
                        //PIC
                        var itemsPIC = objTypeProduct.Select(w => w.Pic).Distinct().ToList();
                        var listPIC = _context.Categories.Where(w => itemsPIC.Contains(w.ID)).Select(s => s.Title).ToList();
                        if (listPIC.Count() > 0)
                        {
                            string combinedStringPIC = string.Join(" ,", listPIC);
                            itemPIC = combinedStringPIC;
                        }
                        //Policy
                        itemPolicy1 = objTypeProduct.Sum(s => s.PalletPolicy ?? 0);
                        if (double.IsNaN(itemPolicy1) || double.IsInfinity(itemPolicy1))
                        {
                            itemPolicy1 = 0;
                        }
                        foreach (var item2 in objTypeProduct)
                        {
                            if (item2.StockQty > 0 && item2.Pallet > 0)
                            {
                                //itemActual1 += Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value);
                                itemActual1 += Math.Ceiling(Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value));
                            }
                        }
                        // so pallet: Slow moving cho kho R
                        foreach (var item2 in objTypeProductR)
                        {
                            if (item2.StockQty > 0 && item2.Pallet > 0)
                            {
                                //itemActualR += Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value);
                                itemActualR += Math.Ceiling(Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value));
                            }
                        }
                        //
                        itemPolicyA = objTypeProduct.Sum(s => s.PalletPolicy ?? 0);
                        if (double.IsNaN(itemPolicyA) || double.IsInfinity(itemPolicyA))
                        {
                            itemPolicyA = 0;
                        }
                        foreach (var item2 in objTypeProductA)
                        {
                            if (item2.StockQty > 0 && item2.Pallet > 0)
                            {
                                //itemActualA += Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value);
                                itemActualA += Math.Ceiling(Convert.ToDouble(item2.StockQty.Value / item2.Pallet.Value));
                            }
                        }
                        //Line
                        var itemsLine = objTypeProduct.Select(w => w.LineName).Distinct().ToList();
                        if (itemsLine.Count() > 0)
                        {
                            string combinedStringLine = string.Join(" ,", itemsLine);
                            itemLine = combinedStringLine;
                        }
                    }
                    //STOCK
                    bool chkRemark = false;
                    // ko dung tu ngay: 08.02.2023
                    //if (itemStock == (int)EnumStatusInventory.InActive)
                    //{
                    //    chkRemark = (objTypeProduct.Any(f => f.Remark == null && f.Status == (int)EnumStatusInventory.InActive) == true) ? false : true;
                    //}
                    listVisualizeStocks2.Add(new StockDetail()
                    {
                        Key = item.Title,
                        Value = itemStock,
                        CheckRemark = chkRemark
                    });

                    //DOH
                    listVisualizeDohs.Add(itemDOH);
                    //DOH POLICY
                    listVisualizeDOHPOLICY.Add(itemDOHPOLICY);
                    //PIC
                    listVisualizePIC.Add(itemPIC);
                    //Policy
                    listVisualizePolicy1.Add(Math.Round(itemPolicy1, 1));
                    listVisualizeActual1.Add(Math.Round(itemActual1, 1));
                    listVisualizePolicyA.Add(Math.Round(itemPolicyA, 1));
                    listVisualizeActualA.Add(Math.Round(itemActualA, 1));
                    listVisualizeActualR.Add(Math.Round(itemActualR, 1));
                    //Line
                    listVisualizeLine.Add(itemLine);
                    // top 10, lay theo tung loai hang
                    // kho 1 edit: 10.02.2023
                    var top10 = qryInventory1Top.Where(w => w.Status == (int)EnumStatusInventory.InActive && w.ShortageDate != null && w.ShortageDate >= DateTime.Now.Date && w.TypeProduct == item.Title).OrderByDescending(o => o.ShortageDate).Take(10).ToList();
                    if (top10.Count() > 0)
                    {
                        listCatTop10.Add(item.Title, top10);
                    }
                }
                //
                obj.Stocks = listVisualizeStocks2;
                obj.Dohs = listVisualizeDohs;
                obj.DohPolicy = listVisualizeDOHPOLICY;
                obj.Pics = listVisualizePIC;
                obj.Policy1 = listVisualizePolicy1;
                obj.Actual1 = listVisualizeActual1;
                obj.PolicyA = listVisualizePolicyA;
                obj.ActualA = listVisualizeActualA;
                obj.ActualR = listVisualizeActualR;
                obj.Lines = listVisualizeLine;
            }
            //
            //var list = _inventory.GetStockVisualization();
            ViewBag.Top10 = listCatTop10;
            return View("../Front/StockVisualizations/Index", obj);
        }
    }
}
