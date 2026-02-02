namespace StockControl.Models.Front.Reports
{
    public class ReportFactoryModel
    {
        public Dictionary<string, Dictionary<int, int>> ListMonths { get; set; }
        public Dictionary<string, List<double>>  ListValues { get; set; }
        // lay data danh gia
        public ICollection<EvaluationModel> Evaluations { get; set; }
    }
}
