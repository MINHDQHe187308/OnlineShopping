using ASP.Models.Domains;

namespace ASP.Models.ViewModels
{
    public class CheckoutViewModel
    {
        public List<CartItem> CartItems { get; set; }

        public string FullName { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }

        public string PaymentMethod { get; set; } // COD, VNPay

        public decimal TotalAmount { get; set; }
    }
}
