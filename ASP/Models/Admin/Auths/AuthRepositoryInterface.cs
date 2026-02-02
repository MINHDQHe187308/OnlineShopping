using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Auths
{
    public interface AuthRepositoryInterface
    {
        public Task<IActionResult> CreateAccount(AuthListenerInterface listener, Register request);
        public bool CheckAccount(Register request);
        public Task<bool> CheckAccountAD(Register request);
    }
}
