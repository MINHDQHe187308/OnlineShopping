using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace ODCS.BaseCommon
{
    public enum EnumStatusInventory
    {
        [Description("Low")]
        [Display(Name = "Low")]
        InActive = -5,
        [Description("High")]
        [Display(Name = "High")]
        Pending = 5,
        [Description("Normal")]
        [Display(Name = "Normal")]
        Active = 10,
        [Description("Other")]
        [Display(Name = "Other")]
        Other = 15
    }
}
