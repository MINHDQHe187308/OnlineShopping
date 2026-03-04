using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Domains
{
    public class Order
    {
        [Key]
        public int OrderId { get; set; }

        [Required]
        public int CustomerId { get; set; }

        [Required]
        public int CreatedBy { get; set; }

        public DateTime OrderDate { get; set; } = DateTime.Now;

        [Column(TypeName = "decimal(18,2)")]
        public decimal TotalAmount { get; set; }

        [Required]
        [StringLength(20)]
        public string Status { get; set; }   // Completed, Cancelled

        // Navigation
        [ForeignKey("CustomerId")]
        public Customer Customer { get; set; }

        [ForeignKey("CreatedBy")]
        public User User { get; set; }

        public ICollection<OrderDetail> OrderDetails { get; set; }
    }
}
