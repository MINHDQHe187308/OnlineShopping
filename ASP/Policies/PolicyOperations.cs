using Microsoft.AspNetCore.Authorization.Infrastructure;

namespace ASP.Policies
{
    public class PolicyOperations
    {
        #region Dashboard policies
        public static OperationAuthorizationRequirement ASPDashboardView =
          new OperationAuthorizationRequirement { Name = "ASPDashboardView" };
        #endregion

        #region Menus policies
        public static OperationAuthorizationRequirement ASPMenusView =
          new OperationAuthorizationRequirement { Name = "ASPMenusView" };
        public static OperationAuthorizationRequirement ASPMenusCreate =
          new OperationAuthorizationRequirement { Name = "ASPMenusCreate" };
        public static OperationAuthorizationRequirement ASPMenusUpdate =
          new OperationAuthorizationRequirement { Name = "ASPMenusUpdate" };
        public static OperationAuthorizationRequirement ASPMenusDelete =
          new OperationAuthorizationRequirement { Name = "ASPMenusDelete" };
        #endregion

        #region Theme Options policies
        public static OperationAuthorizationRequirement ASPThemoptionsView =
          new OperationAuthorizationRequirement { Name = "ASPThemoptionsView" };
        public static OperationAuthorizationRequirement ASPThemoptionsUpdate =
          new OperationAuthorizationRequirement { Name = "ASPThemoptionsUpdate" };
        #endregion

        #region Users policies
        public static OperationAuthorizationRequirement ASPUsersView =
          new OperationAuthorizationRequirement { Name = "ASPUsersView" };
        public static OperationAuthorizationRequirement ASPUsersCreate =
          new OperationAuthorizationRequirement { Name = "ASPUsersCreate" };
        public static OperationAuthorizationRequirement ASPUsersUpdate =
          new OperationAuthorizationRequirement { Name = "ASPUsersUpdate" };
        public static OperationAuthorizationRequirement ASPUsersBanned =
          new OperationAuthorizationRequirement { Name = "ASPUsersBanned" };
        public static OperationAuthorizationRequirement ASPUsersDelete =
          new OperationAuthorizationRequirement { Name = "ASPUsersDelete" };
        #endregion

        #region Roles policies
        public static OperationAuthorizationRequirement ASPRolesView =
          new OperationAuthorizationRequirement { Name = "ASPRolesView" };
        public static OperationAuthorizationRequirement ASPRolesCreate =
          new OperationAuthorizationRequirement { Name = "ASPRolesCreate" };
        public static OperationAuthorizationRequirement ASPRolesUpdate =
          new OperationAuthorizationRequirement { Name = "ASPRolesUpdate" };
        public static OperationAuthorizationRequirement ASPRolesBanned =
          new OperationAuthorizationRequirement { Name = "ASPRolesBanned" };
        public static OperationAuthorizationRequirement ASPRolesDelete =
          new OperationAuthorizationRequirement { Name = "ASPRolesDelete" };
        #endregion

        #region Logs policies
        public static OperationAuthorizationRequirement ASPLogsView =
          new OperationAuthorizationRequirement { Name = "ASPLogsView" };
        #endregion

        #region Customer policies
        public static OperationAuthorizationRequirement ASPCustomerView =
            new OperationAuthorizationRequirement { Name = "ASPCustomerView" };
        public static OperationAuthorizationRequirement ASPCustomerCreate =
            new OperationAuthorizationRequirement { Name = "ASPCustomerCreate" };
        public static OperationAuthorizationRequirement ASPCustomerUpdate =
            new OperationAuthorizationRequirement { Name = "ASPCustomerUpdate" };
        public static OperationAuthorizationRequirement ASPCustomerDelete =
            new OperationAuthorizationRequirement { Name = "ASPCustomerDelete" };
        #endregion

        #region Leadtime policies
        public static OperationAuthorizationRequirement ASPLeadtimeView =
            new OperationAuthorizationRequirement { Name = "ASPLeadtimeView" };
        public static OperationAuthorizationRequirement ASPLeadtimeCreate =
            new OperationAuthorizationRequirement { Name = "ASPLeadtimeCreate" };
        public static OperationAuthorizationRequirement ASPLeadtimeUpdate =
            new OperationAuthorizationRequirement { Name = "ASPLeadtimeUpdate" };
        public static OperationAuthorizationRequirement ASPLeadtimeDelete =
            new OperationAuthorizationRequirement { Name = "ASPLeadtimeDelete" };
        #endregion
    }
}
