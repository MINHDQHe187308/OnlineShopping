using ASP.Resources;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Logs
{
    public class Log : BaseEntity
    {
        [Key]
        public long ID { get; set; }

        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Message_Shared))]
        [MaxLength(500, ErrorMessageResourceName = "msg_err_log_type_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_log_type", ResourceType = typeof(Resources.SharedResource))]
        public string LogType { get; set; }

        [Column(TypeName = "ntext")]
        [Display(Name = "lbl_content", ResourceType = typeof(Resources.SharedResource))]
        public string Content { get; set; }

        [MaxLength(500, ErrorMessageResourceName = "msg_err_author_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_author", ResourceType = typeof(Resources.SharedResource))]
        public string Author { get; set; }

        [MaxLength(50, ErrorMessageResourceName = "msg_err_ip_address_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_ip_address", ResourceType = typeof(Resources.SharedResource))]
        public string IP { get; set; }
    }
}
