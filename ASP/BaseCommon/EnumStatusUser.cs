using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumStatusUser
    {
        [Display(Name = "lbl_enum_status_user_in_active", Description = "lbl_enum_status_user_in_active", ResourceType = typeof(Resources.SharedResource))]
        InActive = -5,
        [Display(Name = "lbl_enum_status_user_pending", Description = "lbl_enum_status_user_pending", ResourceType = typeof(Resources.SharedResource))]
        Pending = 5,
        [Display(Name = "lbl_enum_status_user_active", Description = "lbl_enum_status_user_active", ResourceType = typeof(Resources.SharedResource))]
        Active = 10
    }
}
