using ASP.Models;
using Microsoft.AspNetCore.Identity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASP.Models.Admin.Accounts;

namespace ASP.BaseCommon
{
    public class Md5PasswordHasher : IPasswordHasher<ApplicationUser>
    {
        public string HashPassword(ApplicationUser user, string password)
        {
            return password;
        }

        public PasswordVerificationResult VerifyHashedPassword(ApplicationUser user, string hashedPassword, string providedPassword)
        {
            return hashedPassword.Equals(providedPassword) ? PasswordVerificationResult.Success : PasswordVerificationResult.Failed;
        }
    }
}
