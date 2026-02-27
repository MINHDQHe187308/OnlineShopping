using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace ODCS.BaseCommon
{
    public enum EnumTaxonomy
    {
        [Description("PIC")]
        Pic = 5,
        [Description("Lý do")]
        Reason = 10,
        [Description("Loại sản phẩm")]
        Category = 15,
        [Description("Kho")]
        WareHouse = 20,
        [Description("Tin tức")]
        News = 25
    }
}
