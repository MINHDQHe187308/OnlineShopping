using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumStatusActive
    {
        [Display(Name = "lbl_enum_status_active_disable", Description = "lbl_enum_status_active_disable", ResourceType = typeof(Resources.SharedResource))]
        Disable = 5,
        [Display(Name = "lbl_enum_status_active_active", Description = "lbl_enum_status_active_active", ResourceType = typeof(Resources.SharedResource))]
        Active = 10
    }
}
