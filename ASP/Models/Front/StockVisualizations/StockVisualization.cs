namespace StockControl.Models.Front.StockVisualizations
{
    public class StockVisualization
    {
        public List<string> KindOfPart { get; set; }
        //public Dictionary<string,int> Stocks { get; set; }
        public List<StockDetail> Stocks { get; set; }
        public List<double> Dohs { get; set; }
        public List<string> DohPolicy { get; set; }
        public List<string> Pics { get; set; }
        public List<double> Policy1 { get; set; }
        public List<double> Actual1 { get; set; }
        public List<double> PolicyA { get; set; }
        public List<double> ActualA { get; set; }
        public List<double> ActualR { get; set; }
        public List<string> Lines { get; set; }
    }
    public class ArrayDetail<T>
    {
        public string Key { get; set; }
        public string Title { get; set; }
    }
    public class StockDetail
    {
        public string Key { get; set; }
        public double Value { get; set; }
        public bool CheckRemark { get; set; }
    }
}
