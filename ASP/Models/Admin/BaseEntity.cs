using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin
{
    public class BaseEntity
    {
        [NotMapped]
        [Display(Name = "lbl_created_date", ResourceType = typeof(Resources.SharedResource))]
        public DateTime CreatedDate { get; set; }
        [NotMapped]
        [Display(Name = "lbl_updated_date", ResourceType = typeof(Resources.SharedResource))]
        public DateTime UpdatedDate { get; set; }
    }
}
