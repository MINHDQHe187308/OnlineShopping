using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Security.AccessControl;

namespace ASP.BaseCommon
{
    public enum EnumLocationType
    {
        [Display(Name = "lbl_enum_location_wip", Description = "lbl_enum_location_wip", ResourceType = typeof(Resources.SharedResource))]
        WIP = 1,
        [Display(Name = "lbl_enum_location_warehouse", Description = "lbl_enum_location_warehouse", ResourceType = typeof(Resources.SharedResource))]
        WH = 2,
        [Display(Name = "lbl_enum_location_mix", Description = "lbl_enum_location_mix", ResourceType = typeof(Resources.SharedResource))]
        MIX = 3
    }
}
