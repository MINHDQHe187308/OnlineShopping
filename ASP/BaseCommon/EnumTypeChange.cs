using System.ComponentModel.DataAnnotations;

namespace ASP.BaseCommon
{
    public enum EnumTypeChange
    {
        [Display(Name = "lbl_enum_type_change_create", Description = "lbl_enum_type_change_create", ResourceType = typeof(Resources.SharedResource))]
        Create = 5,
        [Display(Name = "lbl_enum_type_change_update", Description = "lbl_enum_type_change_update", ResourceType = typeof(Resources.SharedResource))]
        Update = 10,
        [Display(Name = "lbl_enum_type_change_delete", Description = "lbl_enum_type_change_delete", ResourceType = typeof(Resources.SharedResource))]
        Delete = 15
    }
}
