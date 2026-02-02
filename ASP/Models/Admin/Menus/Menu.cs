using ASP.Resources;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ASP.Models.Admin.Menus
{
    public class Menu : BaseEntity
    {
        [Key]
        public int ID { get; set; }
        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Message_Shared))]
        [MaxLength(50, ErrorMessageResourceName = "msg_err_menu_title_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_title", ResourceType = typeof(Resources.SharedResource))]
        public string Name { get; set; }

        [Display(Name = "lbl_description", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(100, ErrorMessageResourceName = "msg_err_menu_description_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        public string? Description { get; set; }

        [Column(TypeName = "ntext")]
        [Display(Name = "lbl_content", ResourceType = typeof(Resources.SharedResource))]
        [MaxLength(1000, ErrorMessageResourceName = "msg_err_menu_content_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        public string? Content { get; set; }

        [Required(ErrorMessageResourceName = "msg_err_string_required", ErrorMessageResourceType = typeof(Message_Shared))]
        [MaxLength(20, ErrorMessageResourceName = "msg_err_menu_language_max_length", ErrorMessageResourceType = typeof(Message_Shared))]
        [Display(Name = "lbl_language", ResourceType = typeof(Resources.SharedResource))]
        public string Language { get; set; }

        [NotMapped]
        public string? out_menu { get; set; }

        public static implicit operator Menu?(Models.Menu? v)
        {
            throw new NotImplementedException();
        }
    }
    public class MenuDetail
    {
        [NotMapped]
        public string name { get; set; }
        [NotMapped]
        public string excerpt { get; set; }
        [NotMapped]
        public string thumbnail { get; set; }
        [NotMapped]
        public string url { get; set; }
        [NotMapped]
        public string target { get; set; }
        [NotMapped]
        public string languagekey { get; set; }
        [NotMapped]
        public List<MenuDetail> children { get; set; }
    }
}
