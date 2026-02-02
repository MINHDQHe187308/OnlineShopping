using System;
using System.Collections.Generic;

namespace ASP.Models.ASPModel
{
    public partial class AspNetUserRole
    {
        public string UserId { get; set; } = null!;
        public string RoleId { get; set; } = null!;
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }

        public virtual AspNetRole Role { get; set; } = null!;
        public virtual AspNetUser User { get; set; } = null!;
    }
}
