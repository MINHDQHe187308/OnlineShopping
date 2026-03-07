using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Domains
{
    public class Cart
    {
        [Key]
        public int CartId { get; set; }

        public int CustomerId { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;

     
        [ForeignKey("CustomerId")]
        public Customer Customer { get; set; }

        public ICollection<CartItem> CartItems { get; set; }
    }
}