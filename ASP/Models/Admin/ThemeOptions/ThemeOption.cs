using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.ThemeOptions
{
    public class ThemeOption:BaseEntity
    {
        [Key]
        public int ID { get; set; }
        [Required]
        [MaxLength(500)]
        public string Name { get; set; }
        [Column(TypeName = "ntext")]
        public string? Value { get; set; }
        [MaxLength(200)]
        public string? TypeData { get; set; }
        [MaxLength(1000)]
        public string? Language { get; set; }
        [NotMapped]
        public List<int> _optCategory { get; set; }
        
    }
}
