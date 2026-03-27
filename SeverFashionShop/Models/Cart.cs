using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class Cart
    {
        public Cart()
        {
            CartItems = new HashSet<CartItem>();
        }

        public int CartId { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
        public string UserId { get; set; } = null!;

        public virtual AspNetUser User { get; set; } = null!;
        public virtual ICollection<CartItem> CartItems { get; set; }
    }
}
