using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class ThemeOption
    {
        public int Id { get; set; }
        public string Name { get; set; } = null!;
        public string? Value { get; set; }
        public string? TypeData { get; set; }
        public string? Language { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
    }
}
