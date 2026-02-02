using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Auths
{
   public interface AuthListenerInterface
    {
        public IActionResult RegisterAccountSuccess();
        public IActionResult RegisterAccountFails(Register user);
    }
}
