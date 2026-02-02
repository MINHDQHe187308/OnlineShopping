using Microsoft.AspNetCore.Identity;
using ASP.Resources;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Admin.Roles
{
    public class Role : IdentityRole
    {
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_role", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(256, ErrorMessageResourceName = "msg_err_role_name_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        public override string Name { get; set; }

        [Display(Name = "lbl_role_permission", ResourceType = typeof(Resources.SharedResource))]
        [Column(TypeName = "ntext")]
        public string Content { get; set; }

        [Display(Name = "lbl_status", ResourceType = typeof(Resources.SharedResource))]
        public short Status { get; set; }

        [Display(Name = "lbl_role_default", ResourceType = typeof(Resources.SharedResource))]   
        public bool DefaultRole { get; set; }

        [NotMapped]
        [Display(Name = "lbl_role_count_user", ResourceType = typeof(Resources.SharedResource))]
        public int CountUser { get; set; }

        [NotMapped] 
        [Display(Name = "lbl_issued_by", ResourceType = typeof(Resources.SharedResource))]
        public DateTime CreatedDate { get; set; }

        [NotMapped]
        [Display(Name = "lbl_updated_by", ResourceType = typeof(Resources.SharedResource))]
        public DateTime UpdatedDate { get; set; }
    }
}
