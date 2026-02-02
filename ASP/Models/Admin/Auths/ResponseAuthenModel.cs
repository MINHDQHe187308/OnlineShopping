namespace ASP.Models.Admin.Auths
{
    public class ResponseAuthenModel
    {
        public bool result { get; set; }
        public DataDetail data { get; set; }
        public string error_code { get; set; }
        public string error_message { get; set; }
    }
    public class DataDetail
    {
        public string employeeId { get; set; }
        public string email { get; set; }
        public string fullName { get; set; }
    }
}
