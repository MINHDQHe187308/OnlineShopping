namespace StockControl.Models.Front.Reports
{
    public interface ReportRepositoryInterface
    {
        public HighChartModel? GetChartFactory(int year);
        public ReportFactoryModel? GetStockFactory(int year);
        public HighChartModel? GetChartCategory(int year, string cat);
        public ReportFactoryModel? GetStockCategory(int year, string cat);
        public HighChartModel? GetChartSupplier(int year, string fSupplier);
    }
}
