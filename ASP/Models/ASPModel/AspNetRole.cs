using System;
using System.Collections.Generic;

namespace ASP.Models.ASPModel
{
    public partial class AspNetRole
    {
        public AspNetRole()
        {
            AspNetRoleClaims = new HashSet<AspNetRoleClaim>();
            AspNetUserRoles = new HashSet<AspNetUserRole>();
        }

        public string Id { get; set; } = null!;
        public string Name { get; set; } = null!;
        public string Content { get; set; } = null!;
        public short Status { get; set; }
        public bool? DefaultRole { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
        public string? NormalizedName { get; set; }
        public string? ConcurrencyStamp { get; set; }

        public virtual ICollection<AspNetRoleClaim> AspNetRoleClaims { get; set; }
        public virtual ICollection<AspNetUserRole> AspNetUserRoles { get; set; }
    }
}
