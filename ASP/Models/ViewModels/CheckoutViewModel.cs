using ASP.Models.Admin.Accounts;
using ASP.Models.ASPModel;
using ASP.Models.Domains;
using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;

namespace ASP.Models.ViewModels
{
    public class CheckoutViewModel
    {
        public List<CartItem>? CartItems { get; set; }

        public List<ShippingAddress>? Addresses { get; set; }

        public ShippingAddress? Address { get; set; }

        public ApplicationUser? user { get; set; }

        public decimal TotalAmount { get; set; }
    }
    
}
