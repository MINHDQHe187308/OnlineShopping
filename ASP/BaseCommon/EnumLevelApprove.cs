using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumLevelApprove
    {
        [Display(Name = "lbl_enum_level_approve_none", Description = "lbl_enum_level_approve_none", ResourceType = typeof(Resources.SharedResource))]
        None = 0,
        [Display(Name = "lbl_enum_level_approve_check", Description = "lbl_enum_level_approve_check", ResourceType = typeof(Resources.SharedResource))]
        Check = 5,
        [Display(Name = "lbl_enum_level_approve_approve", Description = "lbl_enum_level_approve_approve", ResourceType = typeof(Resources.SharedResource))]
        Approve = 10
    }
}
