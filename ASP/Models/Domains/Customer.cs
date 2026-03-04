using Spire.Presentation;
using System.ComponentModel.DataAnnotations;

namespace ASP.Models.Domains
{
    public class Customer
    {
        [Key]
        public int CustomerId { get; set; }

        [Required]
        [StringLength(150)]
        public string FullName { get; set; }

        [Required]
        [StringLength(20)]
        public string Phone { get; set; }

        public ICollection<Order> Orders { get; set; }
    }
}
