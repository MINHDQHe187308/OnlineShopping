using Spire.Presentation;
using System.ComponentModel.DataAnnotations;

namespace ASP.Models.Domains
{
    public class User
    {
        [Key]
        public int UserId { get; set; }

        [Required]
        [StringLength(50)]
        public string Username { get; set; }

        [Required]
        [MaxLength]
        public string PasswordHash { get; set; }

        [Required]
        [StringLength(100)]
        public string FullName { get; set; }

        [Required]
        [StringLength(20)]
        public string Role { get; set; }   // Admin, Manager, Staff

        public bool IsActive { get; set; } = true;


        public ICollection<Order> Orders { get; set; }
        public ICollection<Notification> Notifications { get; set; }
    }
}
