using StockControl.BaseCommon;
using StockControl.Models.Admin.Logs;
using StockControl.Models.Admin.Stocks;
using StockControl.Models.Admin.Categories;
using StockControl.Models.Admin.Products;

namespace StockControl.Models.Front.Reports
{
    public class ReportRepository : ReportRepositoryInterface
    {
        private readonly ILogger<ReportRepository> _logger;
        protected readonly StockControlContext _context;
        protected LogRepositoryInterface _log;
        private readonly IWebHostEnvironment _env;
        protected string _photosPath;
        private object userManager;
        public BaseController _baseControllerI;
        public CategoriesRepositoryInterface _categories;
        public ProductRepositoryInterface _product;
        public ReportRepository(ILogger<ReportRepository> logger, StockControlContext context, IWebHostEnvironment env, LogRepositoryInterface log, BaseController baseController, CategoriesRepositoryInterface categories, ProductRepositoryInterface product)
        {
            _logger = logger;
            _context = context;
            _env = env;
            _log = log;
            _photosPath = _env.WebRootPath + "/UploadExcel";
            _baseControllerI = baseController;
            _categories = categories;
            _product = product;
        }
        public HighChartModel? GetChartFactory(int year)
        {
            //service Common
            var serviceCommon = new ServiceReferenceCommon.Service1Client();
            // lay bang ti gia
            var listExchangeRate = _context.ExchangeRates.Where(m => (m.ExYear == year && m.ExMonth >= 5) || (m.ExYear == (year + 1) && m.ExMonth <= 4)).ToList();
            // list danh muc
            var sltTypeProduct = _categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Category, true);
            // list DOH: loai hang theo thang: 
            HighChartModel model = new HighChartModel();
            var series = new List<SeriesModel<double>>();
            // loai hang, type, data
            var listDOHTotalStock = new List<double>();
            var listDOHTotalRequire = new List<double>();
            //chi lay kho 1, A
            var allStockYear = (from p in _context.Products
                                join m in _context.Stocks on p.ID equals m.ProductRefId
                                where p.NotKpi != true && (m.House == "1" || m.House == "A" || m.House == "R") && ((m.StockYear == year && m.StockMonth >= 5) || (m.StockYear == (year + 1) && m.StockMonth <= 4))
                                select new Stock()
                                {
                                    StockYear = m.StockYear,
                                    StockMonth = m.StockMonth,
                                    PartNo = m.PartNo,
                                    House = m.House,
                                    Invqy = m.Invqy,
                                    Rqrqy = m.Rqrqy,
                                    Price = m.Price,
                                    Scost = m.Scost,
                                    TypeProduct = p.TypeProduct,
                                    Currency = p.Currency,
                                    ProductRefId = m.ProductRefId,
                                    NotKpi = m.NotKpi,
                                    CreatedDate = m.CreatedDate,
                                    UpdatedDate = m.UpdatedDate
                                }).OrderByDescending(f => f.StockYear).ThenByDescending(f => f.StockMonth).ToList();
            var listDataChartModel = new List<DataChartModel>();
            var dicMonthYear = new Dictionary<int, int>();
            if (allStockYear.Count() <= 0)
            {
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"Null";
                seriItem.type = "column";
                seriItem.data = new List<double>()
                {
                    0,0,0,0,0,0,0,0,0,0,0,0
                };
                //
                series.Add(seriItem);
                model.Series = series;
                return model;
            }
            foreach (var itemCat in sltTypeProduct)
            {
                foreach (var itemMonth in _baseControllerI.GetMonthReport2())
                {
                    var yearCurrent = (itemMonth.Value >= 5) ? year : (year + 1);
                    // ngay lam viec theo thang
                    int mHoliday = serviceCommon.GetHolidayFromYearFacAsync(yearCurrent).Result.Where(w => w.Year == yearCurrent && w.Day.Month == itemMonth.Value).Count();
                    int mDays = DateTime.DaysInMonth(yearCurrent, itemMonth.Value);
                    int totalWorkingDayInMonth = mDays - mHoliday;
                    //
                    var allStockCat = allStockYear.Where(w => w.StockYear == yearCurrent && w.StockMonth == itemMonth.Value && w.TypeProduct == itemCat.Value).ToList();
                    double totalStock = 0.0, totalRequire = 0.0;
                    foreach (var itemProduct in allStockCat)
                    {
                        // coding hear
                        /*
                         * tinh ton kho cuoi thang
                         * gia quy doi ra USD
                        */
                        double price = 0.0;
                        if (!string.IsNullOrEmpty(itemProduct.Currency))
                        {
                            var findExchangeRate = listExchangeRate.FirstOrDefault(f => f.ExYear == yearCurrent && f.ExMonth == itemMonth.Value && f.Currency == itemProduct.Currency);
                            if (findExchangeRate != null)
                            {
                                price = (itemProduct.Price ?? 0) * Convert.ToDouble(findExchangeRate.ExRateUSD);
                            }
                        }
                        else
                        {
                            // ko co ti gia => = 1
                            price = itemProduct.Price ?? 0;
                        }
                        // ton kho tung ma
                        double oneStockPartNo = 0;
                        if (itemProduct.Invqy != null)
                        {
                            oneStockPartNo = itemProduct.Invqy.Value * price;
                        }
                        totalStock += oneStockPartNo;
                        /*
                         * require thang tiep theo
                         * chi lay kho 1
                         */
                        if (itemProduct.House == "1")
                        {
                            double oneRequirePartNo = 0;
                            if (itemProduct.Rqrqy != null)
                            {
                                oneRequirePartNo = itemProduct.Rqrqy.Value * price;
                            }
                            totalRequire += oneRequirePartNo;
                        }
                    }
                    var dataChartDetail = new DataChartModel();
                    dataChartDetail.Year = yearCurrent;
                    dataChartDetail.Month = itemMonth.Value;
                    dataChartDetail.CatName = itemCat.Value;
                    dataChartDetail.Stock = totalStock;
                    dataChartDetail.Require = totalRequire;
                    dataChartDetail.TotalWorkingDays = totalWorkingDayInMonth;
                    listDataChartModel.Add(dataChartDetail);
                    // them du lieu theo thang, nam
                    if (dicMonthYear.ContainsKey(itemMonth.Value) == false)
                    {
                        dicMonthYear.Add(itemMonth.Value, yearCurrent);
                    }
                }// endmonth
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"{itemCat.Value}";
                seriItem.type = "column";
                //
                series.Add(seriItem);
            }//end cat
            var listEvals = new List<EvaluationModel>();
            foreach (var item in series)
            {
                var dataItem = new List<double>();
                foreach (var item2 in dicMonthYear)
                {
                    var totalRequire = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key).Sum(s => s.Require);
                    var totalStock = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.Stock);
                    var totalWDays = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.TotalWorkingDays);
                    double doh = (totalStock / totalRequire) * totalWDays;
                    doh = (Double.IsNaN(doh)) ? 0 : doh;
                    dataItem.Add(Math.Round(doh, 2, MidpointRounding.ToEven));
                    //
                    var obj = new EvaluationModel();
                    obj.Year = item2.Value;
                    obj.Month = item2.Key;
                    obj.CatName = item.name;
                    obj.TotalDoh = Math.Round(doh, 2, MidpointRounding.ToEven);
                    listEvals.Add(obj);
                }
                item.data = dataItem;
            }
            // Evaluations DOH theo hang
            var evalsDOH = new List<EvaluationModel>();
            foreach (var item2 in dicMonthYear)
            {
                DateTime rewriteDt = new DateTime(item2.Value, item2.Key, 1);
                /**
                 * tra ve thang mac dinh theo nam tai chinh
                 * Du lieu bieu do lay theo nam tai chinh: thang + 1
                 * **/
                rewriteDt = rewriteDt.AddMonths(-1);
                //
                var totalDoh = listEvals.Where(w => w.Year == item2.Value && w.Month == item2.Key).Sum(s => s.TotalDoh);
                var obj = new EvaluationModel();
                obj.Year = rewriteDt.Year;
                obj.Month = rewriteDt.Month;
                obj.TotalDoh = totalDoh;
                //min, max, target
                var objStatistical = _context.Statisticals.FirstOrDefault(w => w.TypeReport == EnumTypeStatistical.Factory.ToString() && w.RptYear == rewriteDt.Year && w.RptMonth == rewriteDt.Month);
                if (objStatistical != null)
                {
                    obj.Min = Math.Round(objStatistical.RptMin, 1, MidpointRounding.ToEven);
                    obj.Max = Math.Round(objStatistical.RptMax, 1, MidpointRounding.ToEven);
                    obj.Target = Math.Round(objStatistical.RptTarget, 1, MidpointRounding.ToEven);
                }
                else
                {
                    obj.Min = 0;
                    obj.Max = 0;
                    obj.Target = 0;
                }
                //
                if (obj.TotalDoh >= obj.Min && obj.TotalDoh <= obj.Max)
                {
                    obj.Judge = "O";
                }
                else
                {
                    obj.Judge = "X";
                }
                //
                evalsDOH.Add(obj);
            }
            //
            model.Series = series;
            model.Evaluations = evalsDOH;
            return model;
        }
        public ReportFactoryModel? GetStockFactory(int year)
        {
            var obj = new ReportFactoryModel();
            //nam - thang
            var listMonth = new Dictionary<string, Dictionary<int, int>>();
            var dicMonthDetailDefault = new Dictionary<int, int>();
            dicMonthDetailDefault.Add(0, 0);
            listMonth.Add("#", dicMonthDetailDefault);
            listMonth.Add($"FY {(year - 1)}", dicMonthDetailDefault);
            //Min, Max, Target
            var objStatistical = _context.Statisticals.Where(w => w.TypeReport == EnumTypeStatistical.Factory.ToString() && ((w.RptYear == year && w.RptMonth >= 4) || (w.RptYear == (year + 1) && w.RptMonth <= 3))).OrderBy(o => o.RptYear).ThenBy(t => t.RptMonth).ToList();
            var dic = new Dictionary<string, List<double>>();
            //
            var listMinMonth = new List<double>();
            var listMaxMonth = new List<double>();
            var listTargetMonth = new List<double>();

            foreach (var item in _baseControllerI.GetMonthReport())
            {
                var yearCurrent = (item.Value >= 4) ? year : (year + 1);
                // danh sach thang
                var dicMonthDetail = new Dictionary<int, int>();
                dicMonthDetail.Add(yearCurrent, item.Value);
                listMonth.Add(item.Key, dicMonthDetail);
                // danh sach Min, Max, Target
                var find = objStatistical.FirstOrDefault(f => f.RptYear == yearCurrent && f.RptMonth == item.Value && f.TypeReport == EnumTypeStatistical.Factory.ToString());
                if (find != null)
                {
                    listMinMonth.Add(Math.Round(find.RptMin, 1, MidpointRounding.ToEven));
                    listMaxMonth.Add(Math.Round(find.RptMax, 1, MidpointRounding.ToEven));
                    listTargetMonth.Add(Math.Round(find.RptTarget, 1, MidpointRounding.ToEven));
                }
                else
                {
                    listMinMonth.Add(0);
                    listMaxMonth.Add(0);
                    listTargetMonth.Add(0);
                }
            }
            //
            dic.Add("Min", listMinMonth);
            dic.Add("Max", listMaxMonth);
            dic.Add("Target", listTargetMonth);
            //
            obj.ListMonths = listMonth;
            obj.ListValues = dic;
            return obj;
        }
        public HighChartModel? GetChartCategory(int year, string cat)
        {
            //service Common
            var serviceCommon = new ServiceReferenceCommon.Service1Client();
            // lay bang ti gia
            var listExchangeRate = _context.ExchangeRates.Where(m => (m.ExYear == year && m.ExMonth >= 5) || (m.ExYear == (year + 1) && m.ExMonth <= 4)).ToList();
            // list danh muc
            var sltTypeProduct = _categories.GetTaxonomyByDropdown((int)EnumTaxonomy.Category, true).Where(w => w.Value == cat).ToList();
            // list DOH: loai hang theo thang: 
            HighChartModel model = new HighChartModel();
            var series = new List<SeriesModel<double>>();
            // loai hang, type, data
            var listDOHTotalStock = new List<double>();
            var listDOHTotalRequire = new List<double>();
            //chi lay kho 1, A
            var allStockYear = (from p in _context.Products
                                join m in _context.Stocks on p.ID equals m.ProductRefId
                                where (
                                (p.NotKpi == false || p.NotKpi == null) && (m.House == "1" || m.House == "A" || m.House == "R") && p.TypeProduct == cat && ((m.StockYear == year && m.StockMonth >= 5) || (m.StockYear == (year + 1) && m.StockMonth <= 4))
                                )
                                select new Stock()
                                {
                                    StockYear = m.StockYear,
                                    StockMonth = m.StockMonth,
                                    PartNo = m.PartNo,
                                    House = m.House,
                                    Invqy = m.Invqy,
                                    Rqrqy = m.Rqrqy,
                                    Price = m.Price,
                                    Scost = m.Scost,
                                    TypeProduct = p.TypeProduct,
                                    Currency = p.Currency,
                                    ProductRefId = m.ProductRefId,
                                    NotKpi = m.NotKpi,
                                    CreatedDate = m.CreatedDate,
                                    UpdatedDate = m.UpdatedDate
                                }).OrderByDescending(f => f.StockYear).ThenByDescending(f => f.StockMonth).ToList();
            var listDataChartModel = new List<DataChartModel>();
            var dicMonthYear = new Dictionary<int, int>();
            if (allStockYear.Count() <= 0)
            {
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"Null";
                seriItem.type = "column";
                seriItem.data = new List<double>()
                {
                    0,0,0,0,0,0,0,0,0,0,0,0
                };
                //
                series.Add(seriItem);
                model.Series = series;
                return model;
            }
            foreach (var itemCat in sltTypeProduct)
            {
                // GetMonthReport2: cấu hình từ tháng 5 => 4 năm tiếp theo, chỉ đổi tên hiển thị
                foreach (var itemMonth in _baseControllerI.GetMonthReport2())
                {
                    var yearCurrent = (itemMonth.Value >= 5) ? year : (year + 1);
                    // ngay lam viec theo thang => + 1, theo cong thuc
                    //var dtExtra1 = new DateTime(yearCurrent, itemMonth.Value, 1).AddMonths(1);
                    int mHoliday = serviceCommon.GetHolidayFromYearFacAsync(yearCurrent).Result.Where(w => w.Year == yearCurrent && w.Day.Month == itemMonth.Value).Count();
                    int mDays = DateTime.DaysInMonth(yearCurrent, itemMonth.Value);
                    int totalWorkingDayInMonth = mDays - mHoliday;
                    // ngay ton kho thi lay nhu bo loc
                    //if (itemMonth.Value == 1)
                    //{
                    //    int ax = 0;
                    //}
                    // lay theo thang
                    //var allStockCat = allStockYear.Where(w => w.StockYear == yearCurrent && w.StockMonth == itemMonth.Value && w.TypeProduct == itemCat.Value).ToList();
                    // lay (thang + 1)
                    var allStockCat = allStockYear.Where(w => w.StockYear == yearCurrent && w.StockMonth == itemMonth.Value && w.TypeProduct == itemCat.Value).ToList();

                    double totalStock = 0.0, totalRequire = 0.0;
                    foreach (var itemProduct in allStockCat)
                    {
                        // coding hear
                        /*
                         * tinh ton kho cuoi thang
                         * gia quy doi ra USD VN012099-0010
                        */
                        //if (itemProduct.PartNo == "VN012039-0030")
                        //{
                        //    var test = 0;
                        //}
                        double price = 0.0;
                        if (!string.IsNullOrEmpty(itemProduct.Currency))
                        {
                            var findExchangeRate = listExchangeRate.FirstOrDefault(f => f.ExYear == yearCurrent && f.ExMonth == itemMonth.Value && f.Currency == itemProduct.Currency);
                            if (findExchangeRate != null)
                            {
                                price = (itemProduct.Price ?? 0) * Convert.ToDouble(findExchangeRate.ExRateUSD);
                            }
                        }
                        else
                        {
                            // ko co ti gia => = 1
                            price = itemProduct.Price ?? 0;
                        }
                        double oneStockPartNo = 0;
                        if (itemProduct.Invqy != null)
                        {
                            oneStockPartNo = itemProduct.Invqy.Value * price;
                        }
                        totalStock += oneStockPartNo;
                        /*
                         * require thang tiep theo
                         * chi lay kho 1
                         */
                        if (itemProduct.House == "1")
                        {
                            double oneRequirePartNo = 0;
                            if (itemProduct.Rqrqy != null)
                            {
                                oneRequirePartNo = itemProduct.Rqrqy.Value * price;
                            }
                            totalRequire += oneRequirePartNo;
                        }
                    }
                    var dataChartDetail = new DataChartModel();
                    dataChartDetail.Year = yearCurrent;
                    dataChartDetail.Month = itemMonth.Value;
                    dataChartDetail.CatName = itemCat.Value;
                    dataChartDetail.Stock = totalStock;
                    dataChartDetail.Require = totalRequire;
                    dataChartDetail.TotalWorkingDays = totalWorkingDayInMonth;
                    listDataChartModel.Add(dataChartDetail);
                    // them du lieu theo thang, nam
                    if (dicMonthYear.ContainsKey(itemMonth.Value) == false)
                    {
                        dicMonthYear.Add(itemMonth.Value, yearCurrent);
                    }
                }// endmonth
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"{itemCat.Value}";
                seriItem.type = "column";
                //
                series.Add(seriItem);
            }//end cat
            var listEvals = new List<EvaluationModel>();
            foreach (var item in series)
            {
                var dataItem = new List<double>();
                foreach (var item2 in dicMonthYear)
                {
                    var totalRequire = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key).Sum(s => s.Require);
                    var totalStock = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.Stock);
                    var totalWDays = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.TotalWorkingDays);
                    // amount
                    double doh = (totalStock / totalRequire) * totalWDays;
                    doh = (Double.IsNaN(doh) || Double.IsInfinity(doh)) ? 0 : doh;
                    dataItem.Add(Math.Round(doh, 2, MidpointRounding.ToEven));
                    //
                    var obj = new EvaluationModel();
                    obj.Year = item2.Value;
                    obj.Month = item2.Key;
                    obj.CatName = item.name;
                    obj.TotalDoh = Math.Round(doh, 2, MidpointRounding.ToEven);
                    listEvals.Add(obj);
                }
                item.data = dataItem;
            }
            // Evaluations DOH theo hang
            var evalsDOH = new List<EvaluationModel>();
            foreach (var item2 in dicMonthYear)
            {
                DateTime rewriteDt = new DateTime(item2.Value, item2.Key, 1);
                /**
                 * tra ve thang mac dinh theo nam tai chinh
                 * Du lieu bieu do lay theo nam tai chinh: thang + 1
                 * **/
                rewriteDt = rewriteDt.AddMonths(-1);
                //
                var totalDoh = listEvals.Where(w => w.Year == item2.Value && w.Month == item2.Key).Sum(s => s.TotalDoh);
                var obj = new EvaluationModel();
                obj.Year = rewriteDt.Year;
                obj.Month = rewriteDt.Month;
                obj.TotalDoh = totalDoh;
                //min, max, target
                var objStatistical = _context.Statisticals.FirstOrDefault(w => w.TypeReport == EnumTypeStatistical.Category.ToString() && w.RptCat == cat && w.RptYear == rewriteDt.Year && w.RptMonth == rewriteDt.Month);
                if (objStatistical != null)
                {
                    obj.Min = Math.Round(objStatistical.RptMin, 1, MidpointRounding.ToEven);
                    obj.Max = Math.Round(objStatistical.RptMax, 1, MidpointRounding.ToEven);
                    obj.Target = Math.Round(objStatistical.RptTarget, 1, MidpointRounding.ToEven);
                }
                else
                {
                    obj.Min = 0;
                    obj.Max = 0;
                    obj.Target = 0;
                }
                //
                if (obj.TotalDoh >= obj.Min && obj.TotalDoh <= obj.Max)
                {
                    obj.Judge = "O";
                }
                else
                {
                    obj.Judge = "X";
                }
                //
                evalsDOH.Add(obj);
            }
            //
            model.Series = series;
            model.Evaluations = evalsDOH;
            return model;
        }
        public ReportFactoryModel? GetStockCategory(int year, string cat)
        {
            var obj = new ReportFactoryModel();
            //nam - thang
            var listMonth = new Dictionary<string, Dictionary<int, int>>();
            var dicMonthDetailDefault = new Dictionary<int, int>();
            dicMonthDetailDefault.Add(0, 0);
            listMonth.Add("#", dicMonthDetailDefault);
            listMonth.Add($"FY {(year - 1)}", dicMonthDetailDefault);
            //Min, Max, Target
            var objStatistical = _context.Statisticals.Where(w => w.TypeReport == EnumTypeStatistical.Category.ToString() && w.RptCat == cat && ((w.RptYear == year && w.RptMonth >= 4) || (w.RptYear == (year + 1) && w.RptMonth <= 3))).OrderBy(o => o.RptYear).ThenBy(t => t.RptMonth).ToList();
            var dic = new Dictionary<string, List<double>>();
            //
            var listMinMonth = new List<double>();
            var listMaxMonth = new List<double>();
            var listTargetMonth = new List<double>();

            foreach (var item in _baseControllerI.GetMonthReport())
            {
                var yearCurrent = (item.Value >= 4) ? year : (year + 1);
                // danh sach thang
                var dicMonthDetail = new Dictionary<int, int>();
                dicMonthDetail.Add(yearCurrent, item.Value);
                listMonth.Add(item.Key, dicMonthDetail);
                // danh sach Min, Max, Target
                var find = objStatistical.FirstOrDefault(f => f.RptYear == yearCurrent && f.RptMonth == item.Value && f.TypeReport == EnumTypeStatistical.Category.ToString());
                if (find != null)
                {
                    listMinMonth.Add(Math.Round(find.RptMin, 1, MidpointRounding.ToEven));
                    listMaxMonth.Add(Math.Round(find.RptMax, 1, MidpointRounding.ToEven));
                    listTargetMonth.Add(Math.Round(find.RptTarget, 1, MidpointRounding.ToEven));
                }
                else
                {
                    listMinMonth.Add(0);
                    listMaxMonth.Add(0);
                    listTargetMonth.Add(0);
                }
            }
            //
            dic.Add("Min", listMinMonth);
            dic.Add("Max", listMaxMonth);
            dic.Add("Target", listTargetMonth);
            //
            obj.ListMonths = listMonth;
            obj.ListValues = dic;
            return obj;
        }
        public HighChartModel? GetChartSupplier(int year, string supplierCode)
        {
            //service Common
            var serviceCommon = new ServiceReferenceCommon.Service1Client();
            // lay bang ti gia
            var listExchangeRate = _context.ExchangeRates.Where(m => (m.ExYear == year && m.ExMonth >= 5) || (m.ExYear == (year + 1) && m.ExMonth <= 4)).ToList();
            // list danh muc
            var sltTypeSupplier = _product.GetSupplierCode().Where(w => w.Value == supplierCode).ToList();
            // list DOH: loai hang theo thang: 
            HighChartModel model = new HighChartModel();
            var series = new List<SeriesModel<double>>();
            // loai hang, type, data
            var listDOHTotalStock = new List<double>();
            var listDOHTotalRequire = new List<double>();
            //chi lay kho 1, A, R
            // lay het: p.NotKpi != true &&
            var allStockYear = (from p in _context.Products
                                join m in _context.Stocks on p.ID equals m.ProductRefId
                                where (m.House == "1" || m.House == "A" || m.House == "R") && p.SupplierCode == supplierCode && ((m.StockYear == year && m.StockMonth >= 5) || (m.StockYear == (year + 1) && m.StockMonth <= 4))
                                select new Stock()
                                {
                                    StockYear = m.StockYear,
                                    StockMonth = m.StockMonth,
                                    PartNo = m.PartNo,
                                    House = m.House,
                                    Invqy = m.Invqy,
                                    Rqrqy = m.Rqrqy,
                                    Price = m.Price,
                                    Scost = m.Scost,
                                    TypeProduct = p.TypeProduct,
                                    SupplierCode = p.SupplierCode,
                                    Currency = p.Currency,
                                    ProductRefId = m.ProductRefId,
                                    NotKpi = m.NotKpi,
                                    CreatedDate = m.CreatedDate,
                                    UpdatedDate = m.UpdatedDate
                                }).OrderByDescending(f => f.StockYear).ThenByDescending(f => f.StockMonth).ToList();
            var listDataChartModel = new List<DataChartModel>();
            var dicMonthYear = new Dictionary<int, int>();
            if (allStockYear.Count() <= 0)
            {
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"Null";
                seriItem.type = "column";
                seriItem.data = new List<double>()
                {
                    0,0,0,0,0,0,0,0,0,0,0,0
                };
                //
                series.Add(seriItem);
                model.Series = series;
                return model;
            }
            foreach (var itemCat in sltTypeSupplier)
            {
                foreach (var itemMonth in _baseControllerI.GetMonthReport2())
                {
                    var yearCurrent = (itemMonth.Value >= 5) ? year : (year + 1);
                    // ngay lam viec theo thang
                    int mHoliday = serviceCommon.GetHolidayFromYearFacAsync(yearCurrent).Result.Where(w => w.Year == yearCurrent && w.Day.Month == itemMonth.Value).Count();
                    int mDays = DateTime.DaysInMonth(yearCurrent, itemMonth.Value);
                    int totalWorkingDayInMonth = mDays - mHoliday;
                    //
                    var allStockCat = allStockYear.Where(w => w.StockYear == yearCurrent && w.StockMonth == itemMonth.Value && w.SupplierCode == itemCat.Value).ToList();
                    double totalStock = 0.0, totalRequire = 0.0;
                    foreach (var itemProduct in allStockCat)
                    {
                        // coding hear
                        /*
                         * tinh ton kho cuoi thang
                         * gia quy doi ra USD
                        */
                        double price = 0.0;
                        if (!string.IsNullOrEmpty(itemProduct.Currency))
                        {
                            var findExchangeRate = listExchangeRate.FirstOrDefault(f => f.ExYear == yearCurrent && f.ExMonth == itemMonth.Value && f.Currency == itemProduct.Currency);
                            if (findExchangeRate != null)
                            {
                                price = (itemProduct.Price ?? 0) * Convert.ToDouble(findExchangeRate.ExRateUSD);
                            }
                        }
                        else
                        {
                            // ko co ti gia => = 1
                            price = itemProduct.Price ?? 0;
                        }
                        double oneStockPartNo = itemProduct.Invqy ?? 0 * price;
                        totalStock += oneStockPartNo;
                        /*
                         * require thang tiep theo
                         * chi lay kho 1
                         */
                        if (itemProduct.House == "1")
                        {
                            double oneRequirePartNo = itemProduct.Rqrqy ?? 0 * price;
                            totalRequire += oneRequirePartNo;
                        }
                    }
                    var dataChartDetail = new DataChartModel();
                    dataChartDetail.Year = yearCurrent;
                    dataChartDetail.Month = itemMonth.Value;
                    dataChartDetail.CatName = itemCat.Value;
                    dataChartDetail.Stock = totalStock;
                    dataChartDetail.Require = totalRequire;
                    dataChartDetail.TotalWorkingDays = totalWorkingDayInMonth;
                    listDataChartModel.Add(dataChartDetail);
                    /*
                     * them du lieu theo thang, nam
                     * truyen tham so de lay explaination
                     * **/
                    if (dicMonthYear.ContainsKey(itemMonth.Value) == false)
                    {
                        dicMonthYear.Add(itemMonth.Value, yearCurrent);
                    }
                }// endmonth
                var seriItem = new SeriesModel<double>();
                seriItem.name = $"{itemCat.Value}";
                seriItem.type = "column";
                //
                series.Add(seriItem);
            }//end cat
            foreach (var item in series)
            {
                var dataItem = new List<double>();
                foreach (var item2 in dicMonthYear)
                {
                    var totalRequire = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key).Sum(s => s.Require);
                    var totalStock = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.Stock);
                    var totalWDays = listDataChartModel.Where(w => w.Year == item2.Value && w.Month == item2.Key && w.CatName == item.name).Sum(s => s.TotalWorkingDays);
                    double doh = (totalStock / totalRequire) * totalWDays;
                    doh = (Double.IsNaN(doh) || Double.IsInfinity(doh)) ? 0 : doh;
                    dataItem.Add(Math.Round(doh, 2, MidpointRounding.ToEven));
                }
                item.data = dataItem;
            }
            //
            model.Series = series;
            model.Evaluations = null;
            return model;
        }
    }
}
