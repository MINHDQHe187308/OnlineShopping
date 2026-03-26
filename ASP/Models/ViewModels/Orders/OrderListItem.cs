using System;

namespace ASP.Models.ViewModels.Orders
{
    public class OrderListItem
    {
        public int OrderId { get; set; }
        public string? UserName { get; set; }
        public DateTime OrderDate { get; set; }
        public decimal TotalAmount { get; set; }
        public string Status { get; set; } = string.Empty;
    }
}

