using System;
using System.Collections.Generic;

namespace ASP.Models.ASPModel
{
    public partial class AspNetUserLogin
    {
        public string LoginProvider { get; set; } = null!;
        public string ProviderKey { get; set; } = null!;
        public string? ProviderDisplayName { get; set; }
        public string UserId { get; set; } = null!;
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }

        public virtual AspNetUser User { get; set; } = null!;
    }
}
