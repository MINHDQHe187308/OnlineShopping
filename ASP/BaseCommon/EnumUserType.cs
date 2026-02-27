using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumUserType
    {
        [Display(Name = "lbl_enum_user_type_none", Description = "lbl_enum_user_type_none", ResourceType = typeof(Resources.SharedResource))]
        None = 0,
        [Display(Name = "lbl_enum_user_type_pc", Description = "lbl_enum_user_type_pc", ResourceType = typeof(Resources.SharedResource))]
        PC = 5,
        [Display(Name = "lbl_enum_user_type_acc", Description = "lbl_enum_user_type_acc", ResourceType = typeof(Resources.SharedResource))]
        ACC = 10,
        [Display(Name = "lbl_enum_user_type_pro", Description = "lbl_enum_user_type_pro", ResourceType = typeof(Resources.SharedResource))]
        PRO = 15
    }
}
