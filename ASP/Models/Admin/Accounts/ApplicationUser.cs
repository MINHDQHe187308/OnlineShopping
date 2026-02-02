using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Admin.Accounts
{
    public class ApplicationUser : IdentityUser
    {
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_account", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(20, ErrorMessageResourceName = "msg_err_account_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(6, ErrorMessageResourceName = "msg_err_account_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public override string UserName { get; set; }
        
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_fullname", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(50, ErrorMessageResourceName = "msg_err_fullname_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public string FullName { get; set; }

        [Display(Name = "lbl_avatar", ResourceType = typeof(Resources.SharedResource))]
        public string? Avatar { get; set; }
        
        [NotMapped]
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MaxLength(30, ErrorMessageResourceName = "msg_err_password_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(8, ErrorMessageResourceName = "msg_err_password_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_password", ResourceType = typeof(Resources.SharedResource))]
        public string? PassWord { get; set; }
        
        [NotMapped]
        [MaxLength(30, ErrorMessageResourceName = "msg_err_re_password_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(8, ErrorMessageResourceName = "msg_err_re_password_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_re_password", ResourceType = typeof(Resources.SharedResource))]
        [Compare("PassWord", ErrorMessageResourceName = "msg_err_re_password_not_match", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public string? RePassWord { get; set; }
        
        [Display(Name = "lbl_account_type", ResourceType = typeof(Resources.SharedResource))]
        [Required]
        public short LevelManage { get; set; }

        [NotMapped]        
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_role", ResourceType = typeof(Resources.SharedResource))]
        public string RoleId { get; set; }

        [NotMapped]
        [Display(Name = "lbl_role", ResourceType = typeof(Resources.SharedResource))]
        public string? RoleName { get; set; }

        [Display(Name = "lbl_status", ResourceType = typeof(Resources.SharedResource))]
        public short Status { get; set; }

        [Display(Name = "lbl_phone_number", ResourceType = typeof(Resources.SharedResource))]
        public override string? PhoneNumber { get; set; }

        [NotMapped]
        public string? rmavatar { get; set; }

        [NotMapped]
        [Display(Name = "lbl_created_date", ResourceType = typeof(Resources.SharedResource))]
        public DateTime CreatedDate { get; set; }

        [NotMapped]
        [Display(Name = "lbl_updated_date", ResourceType = typeof(Resources.SharedResource))]
        public DateTime UpdatedDate { get; set; }
    }
}
