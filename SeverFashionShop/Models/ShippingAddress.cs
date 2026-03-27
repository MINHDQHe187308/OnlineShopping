using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class ShippingAddress
    {
        public ShippingAddress()
        {
            Orders = new HashSet<Order>();
        }

        public int AddressId { get; set; }
        public string UserId { get; set; } = null!;
        public string FullName { get; set; } = null!;
        public string Phone { get; set; } = null!;
        public string AddressLine { get; set; } = null!;
        public string City { get; set; } = null!;
        public string District { get; set; } = null!;
        public string Ward { get; set; } = null!;
        public bool IsDefault { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }

        public virtual AspNetUser User { get; set; } = null!;
        public virtual ICollection<Order> Orders { get; set; }
    }
}
