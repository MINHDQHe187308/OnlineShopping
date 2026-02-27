namespace ASP.ConfigCommon
{
    public class Permissions
    {
        public static string APP_NAME = "ASP";
        public static List<PermissionDetail> GetPermissions()
        {
            List<PermissionDetail> permissions = new List<PermissionDetail>();

            // Dashboard
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Dashboard",
                Pname = APP_NAME + "Dashboard",
                Picon = "nav-icon fas fa-tachometer-alt",
                Pcan = "ASPDashboardView",
                Proute = "admin.dashboard",
                Pcontroller = "Dashboard",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" }
                }
            });
            
            // Account
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Thành viên",
                Pname = APP_NAME + "Users",
                Picon = "fas fa-user",
                Pcan = "ASPUsersView",
                Proute = "admin.accounts",
                Pcontroller = "Account",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Create", Pvalue = "Tạo mới" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" },
                    new ActionDetail() { Pkey = "Banned", Pvalue = "Khóa" },
                    new ActionDetail() { Pkey = "Delete", Pvalue = "Xóa" }
                }
            });

            // Roles
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Vai trò",
                Pname = APP_NAME + "Roles",
                Picon = "fas fa-users-cog",
                Pcan = "ASPRolesView",
                Proute = "admin.roles",
                Pcontroller = "Role",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Create", Pvalue = "Tạo mới" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" },
                    new ActionDetail() { Pkey = "Banned", Pvalue = "Khóa" },
                    new ActionDetail() { Pkey = "Delete", Pvalue = "Xóa" }
                }
            });
            
            // Menus
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Menu",
                Pname = APP_NAME + "Menus",
                Picon = "fas fa-bars",
                Pcan = "ASPMenusView",
                Proute = "admin.menus",
                Pcontroller = "Menu",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Create", Pvalue = "Tạo mới" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" },
                    new ActionDetail() { Pkey = "Delete", Pvalue = "Xóa" }
                }
            });
            
            // Theme Options
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Thông tin chung",
                Pname = APP_NAME + "Themoptions",
                Picon = "fas fa-info-circle",
                Pcan = "ASPThemoptionsView",
                Proute = "admin.themoptions",
                Pcontroller = "ThemeOption",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" }
                }
            });

            // Logs
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Log",
                Pname = APP_NAME + "Logs",
                Picon = "fas fa-clipboard-check",
                Pcan = "ASPLogsView",
                Proute = "admin.logs",
                Pcontroller = "Log",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" }
                }
            });

            // Customers
            permissions.Add(new PermissionDetail()
            {
                Ptitle = "Customer",
                Pname = APP_NAME + "Customer",
                Picon = "fas fa-list-ul",
                Pcan = "ASPCutomerView",
                Proute = "admin.customers",
                Pcontroller = "Customer",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Create", Pvalue = "Tạo mới" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" },
                    new ActionDetail() { Pkey = "Delete", Pvalue = "Xóa" },
                }
            });

            // Leadtime
            permissions.Add(new PermissionDetail() {
                Ptitle = "Leadtime",
                Pname = APP_NAME + "Leadtime",
                Picon = "fas fa-list-ul",
                Pcan = "ASPLeadtimeView",
                Proute = "admin.leadtimes",
                Pcontroller = "Leadtime",
                Pdetail = new List<ActionDetail>() {
                    new ActionDetail() { Pkey = "View", Pvalue = "Xem" },
                    new ActionDetail() { Pkey = "Create", Pvalue = "Tạo mới" },
                    new ActionDetail() { Pkey = "Update", Pvalue = "Cập nhật" },
                    new ActionDetail() { Pkey = "Delete", Pvalue = "Xóa" },
                }
            });

            return permissions;
        }
    }
    public class PermissionDetail
    {
        public string Ptitle { get; set; }
        public string Pname { get; set; }
        public string Picon { get; set; }
        public string Pcan { get; set; }
        public string Proute { get; set; }
        public string Pcontroller { get; set; }
        public List<ActionDetail> Pdetail { get; set; }
    }
    public class ActionDetail
    {
        public string Pkey { get; set; }
        public string Pvalue { get; set; }
    }
}
