using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumAprroveStatus
    {
        [Display(Name = "lbl_enum_approval_pending", Description = "lbl_enum_approval_pending", ResourceType = typeof(Resources.SharedResource))]
        Pending = 5,
        [Display(Name = "lbl_enum_approval_approved", Description = "lbl_enum_approval_approved", ResourceType = typeof(Resources.SharedResource))]
        Approved = 10,
        [Display(Name = "lbl_enum_approval_rejected", Description = "lbl_enum_approval_rejected", ResourceType = typeof(Resources.SharedResource))]
        Rejected = 25,
        [Display(Name = "lbl_enum_approval_cancel", Description = "lbl_enum_approval_cancel", ResourceType = typeof(Resources.SharedResource))]
        Cancelled = 30
    }
}
