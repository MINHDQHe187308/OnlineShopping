using System;
using System.Collections.Generic;

namespace ASP.Models.ASPModel
{
    public partial class AspNetRoleClaim
    {
        public int Id { get; set; }
        public string RoleId { get; set; } = null!;
        public string? ClaimType { get; set; }
        public string? ClaimValue { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }

        public virtual AspNetRole Role { get; set; } = null!;
    }
}
