using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class Order
    {
        public Order()
        {
            OrderDetails = new HashSet<OrderDetail>();
        }

        public int OrderId { get; set; }
        public string CreatedBy { get; set; } = null!;
        public DateTime OrderDate { get; set; }
        public decimal TotalAmount { get; set; }
        public string Status { get; set; } = null!;
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
        public string UserId { get; set; } = null!;
        public int ShippingAddressId { get; set; }

        public virtual ShippingAddress ShippingAddress { get; set; } = null!;
        public virtual AspNetUser User { get; set; } = null!;
        public virtual ICollection<OrderDetail> OrderDetails { get; set; }
    }
}
