using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Auths
{
    public class Login
    {
        [Required]
        public string UserName { get; set; }

        [Required]
        public string PassWord { get; set; }

        public string? ReturnUrl { get; set; }
    }
}
