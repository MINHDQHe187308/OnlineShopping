using Microsoft.AspNetCore.Identity;
using System.Security.Claims;
using ASP.BaseCommon;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Roles;

namespace ASP.SeedData
{
    public class ApplicationUsersSeeder
    {
        public static async Task SeedRolesAndAdminAsyn(IServiceProvider serviceProvider)
        {
            var roleManager = serviceProvider.GetRequiredService<RoleManager<Role>>();
            var userManager = serviceProvider.GetRequiredService<UserManager<ApplicationUser>>();

            var roles = new[] { "Admin" };
            var claims = new List<Claim>()
            {
                new Claim("ASPDashboardView", "1"),
                new Claim("ASPUsersView", "1"),
                new Claim("ASPUsersCreate", "1"),
                new Claim("ASPUsersUpdate", "1"),
                new Claim("ASPUsersDelete", "1"),
                new Claim("ASPRolesView", "1"),
                new Claim("ASPRolesCreate", "1"),
                new Claim("ASPRolesUpdate", "1"),
                new Claim("ASPRolesBanned", "1"),
                new Claim("ASPRolesDelete", "1"),
                new Claim("ASPMenusView", "1"),
                new Claim("ASPMenusCreate", "1"),
                new Claim("ASPMenusUpdate", "1"),
                new Claim("ASPMenusDelete", "1"),
                new Claim("ASPThemoptionsView", "1"),
                new Claim("ASPThemoptionsUpdate", "1"),
                new Claim("ASPLogsView", "1"),
            };

            foreach (var role in roles) { 
                if(!await roleManager.RoleExistsAsync(role))
                {
                    var model = new Role();
                    model.Id = Guid.NewGuid().ToString();
                    model.Name = "Admin";
                    model.Content = "[{\"Pkey\":\"ASPDashboardView\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPUsersView\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPUsersCreate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPUsersUpdate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPUsersDelete\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPRolesView\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPRolesCreate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPRolesUpdate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPRolesBanned\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPRolesDelete\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPMenusView\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPMenusCreate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPMenusUpdate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPMenusDelete\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPThemoptionsView\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPThemoptionsUpdate\",\"Pvalue\":\"1\"},{\"Pkey\":\"ASPLogsView\",\"Pvalue\":\"1\"}]";
                    model.Status = (short)EnumStatusUser.Active;
                    model.DefaultRole = false;
                    model.CreatedDate = DateTime.Now;
                    model.UpdatedDate = DateTime.Now;
                    await roleManager.CreateAsync(model);
                }
            }

            var adminUser = "User";
            var adminPass = "Minh031124";

            var checkExist = await userManager.FindByNameAsync(adminUser);
            if (checkExist == null) {
                var userModel = new ApplicationUser();
                userModel.UserName = adminUser;
                userModel.FullName = "";
                userModel.NormalizedUserName = adminUser;
                userModel.LevelManage = (short)EnumLevelManage.Administrator;
                userModel.Status = (short)EnumStatusUser.Active;
                userModel.EmailConfirmed = false;
                userModel.PhoneNumberConfirmed = false;
                userModel.TwoFactorEnabled = false;
                userModel.LockoutEnabled = true;
                userModel.AccessFailedCount = 0;

                var res = await userManager.CreateAsync(userModel, adminPass);
                if (res.Succeeded) {
                    res = await userManager.AddToRoleAsync(userModel, "Admin");

                    res = await userManager.AddClaimsAsync(userModel, claims);
                }
            }
        }
    }
}
