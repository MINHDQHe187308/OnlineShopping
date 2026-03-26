namespace ASP.Models.ViewModels.OrderDetails
{
    /// <summary>
    /// Đại diện cho một dòng sản phẩm trong đơn hàng
    /// </summary>
    public class OrderDetailItemViewModel
    {
        public int OrderDetailId { get; set; }
        public int OrderId { get; set; }
        public int VariantId { get; set; }
        public string ProductName { get; set; } = string.Empty;
        public string? VariantName { get; set; }
        public string? ProductImage { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal TotalPrice { get; set; }
    }

    /// <summary>
    /// Đại diện cho toàn bộ thông tin chi tiết của một đơn hàng
    /// </summary>
    public class OrderDetailViewModel
    {
        // Order Information
        public int OrderId { get; set; }
        public string UserName { get; set; } = string.Empty;
        public string UserEmail { get; set; } = string.Empty;
        public string UserPhone { get; set; } = string.Empty;
        public DateTime OrderDate { get; set; }
        public decimal TotalAmount { get; set; }
        public string Status { get; set; } = string.Empty;
        public string CreatedBy { get; set; } = string.Empty;
        
        // Shipping Address Information
        public string? ShippingAddress { get; set; }
        public string? ShippingCity { get; set; }
        public string? ShippingCountry { get; set; }
        public string? ShippingPostalCode { get; set; }
        
        // Order Details (danh sách sản phẩm trong đơn hàng)
        public List<OrderDetailItemViewModel> OrderDetails { get; set; } = new();
    }
}
