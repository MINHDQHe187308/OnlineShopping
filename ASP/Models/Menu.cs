using System;
using System.Collections.Generic;

namespace ASP.Models
{
    public partial class Menu : ASP.Models.Admin.BaseEntity
    {
        public int Id { get; set; }
      
        public string Name { get; set; } = null!;
        public string? Description { get; set; }
        public string? Content { get; set; }
        public string Language { get; set; } = null!;
      
    }   
}
