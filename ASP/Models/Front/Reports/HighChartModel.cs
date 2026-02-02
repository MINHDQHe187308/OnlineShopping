namespace StockControl.Models.Front.Reports
{
    public class HighChartModel
    {
        public ICollection<SeriesModel<double>> Series { get; set; }
        public ICollection<EvaluationModel> Evaluations { get; set; }
    }
    public class SeriesModel<T>
    {
        public string name { get; set; }
        public string type { get; set; }
        public List<double> data { get; set; }
    }
    public class DataChartModel
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string CatName { get; set; }
        public double Stock { get; set; }
        public double Require { get; set; }
        public int TotalWorkingDays { get; set; }

    }
    public class EvaluationModel
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public string CatName { get; set; }
        public double TotalDoh { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double Target { get; set; }
        public string Judge { get; set; }
    }
}
