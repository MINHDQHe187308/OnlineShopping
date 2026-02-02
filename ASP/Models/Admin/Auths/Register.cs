using Microsoft.Extensions.Localization;
using ASP.ConfigCommon;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Auths
{
    public class Register   
    {       
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_account", ResourceType = typeof(Resources.Message_Shared))]
        [MaxLength(20, ErrorMessageResourceName = "msg_err_account_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(6, ErrorMessageResourceName = "msg_err_account_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public string UserName { get; set; }

        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_fullname", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(50, ErrorMessageResourceName = "msg_err_fullname_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public string FullName { get; set; }

        [Display(Name = "lbl_email", ResourceType = typeof(Resources.SharedResource))]
        public string? Email { get; set; }

        [NotMapped]
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MaxLength(30, ErrorMessageResourceName = "msg_err_password_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(8, ErrorMessageResourceName = "msg_err_password_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_password", ResourceType = typeof(Resources.SharedResource))]
        public string PassWord { get; set; }
        //
        [NotMapped]
        [MaxLength(30, ErrorMessageResourceName = "msg_err_re_password_max_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [MinLength(8, ErrorMessageResourceName = "msg_err_re_password_min_length", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        [Display(Name = "lbl_re_password", ResourceType = typeof(Resources.Message_Shared))]
        [Compare("PassWord", ErrorMessageResourceName = "msg_err_re_password_not_match", ErrorMessageResourceType = typeof(Resources.Message_Shared))]
        public string RePassWord { get; set; }
    }
}
