namespace ASP.BaseCommon
{
    public static class EnumTypeLog
    {
        public const string APP_LOG_USER = "APP_LOG_USER";
        public const string APP_LOG_ROLE = "APP_LOG_ROLE";
        public const string APP_LOG_THEMEOPTION = "APP_LOG_THEMEOPTION";
        public const string APP_LOG_MENU = "APP_LOG_MENU";

        public const string APP_LOG_CUSTOMER = "APP_LOG_CUSTOMER";
        public const string APP_LOG_PRODUCTS = "APP_LOG_PRODUCTS";
        public const string APP_LOG_DECODING_CONFIGS = "APP_LOG_DECODING_CONFIGS";

        public const string APP_LOG_ORDER = "APP_LOG_ORDER";
        public const string APP_LOG_SHOPPINGLIST = "APP_LOG_SHOPPINGLIST";
        public const string APP_LOG_TAGASSIGNMENT = "APP_LOG_TAGASSIGNMENT";

        public const string APP_LOG_FEEDBACKS = "APP_LOG_FEEDBACKS";
        public const string APP_LOG_EXCEPTIONS = "APP_LOG_EXCEPTIONS";

        public static string LogTypeDescription(string key)
        {
            var _dicMessage = new Dictionary<string, string>() {
                { "APP_LOG_USER","Thành viên" },
                { "APP_LOG_ROLE","Vai trò" },
                { "APP_LOG_THEMEOPTION","Thông tin chung" },
                { "APP_LOG_MENU","Menu" },
                { "APP_LOG_PRODUCTS", "Sản phẩm" },
                { "APP_LOG_CUSTOMER", "Công ty" },
                { "APP_LOG_ORDER", "Xử lý đơn hàng" },
                { "APP_LOG_DECODING_CONFIGS", "Cấu hình giải mã QR" },
                { "APP_LOG_SHOPPINGLIST", "Thu thập theo đơn hàng" },
                { "APP_LOG_TAGASSIGNMENT", "Đính Tag Khách hàng" },
                { "APP_LOG_FEEDBACKS","Feedback" },
                { "APP_LOG_EXCEPTIONS","Lỗi EXCEPTIONS" },
            };
            return _dicMessage[key];
        }
        public static Dictionary<string, string> GetLogTypeDescription()
        {
            var _dicMessage = new Dictionary<string, string>() {
                { "APP_LOG_USER","Thành viên" },
                { "APP_LOG_ROLE","Vai trò" },
                { "APP_LOG_THEMEOPTION","Thông tin chung" },
                { "APP_LOG_MENU","Menu" },
                { "APP_LOG_PRODUCTS", "Sản phẩm" },
                { "APP_LOG_CUSTOMER", "Công ty" },
                { "APP_LOG_ORDER", "Xử lý đơn hàng" },
                { "APP_LOG_DECODING_CONFIGS", "Cấu hình giải mã QR" },
                { "APP_LOG_SHOPPINGLIST", "Thu thập theo đơn hàng" },
                { "APP_LOG_TAGASSIGNMENT", "Đính Tag Khách hàng" },
                { "APP_LOG_FEEDBACKS","Feedback" },
                { "APP_LOG_EXCEPTIONS","Lỗi EXCEPTIONS" },
            };
            return _dicMessage;
        }
        public static string SetLogTitle(string title)
        {
            return "<h3 class='log-title'>" + title + "</h3>";
        }
        public static string SetLogLine(string title, string old, string news = null)
        {
            if (old == news)
            {
                return null;
            }
            if (old == null)
            {
                return "<br><span class='log-key'>" + title + ": </span>" + news;
            }
            else
            {
                return "<br><span class='log-key'>" + title + ": </span>" + old + "<span class='log-s'> => </span>" + news;
            }
        }
    }
}
