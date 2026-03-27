using System;
using System.Collections.Generic;

namespace SeverFashionShop.Models
{
    public partial class Log
    {
        public long Id { get; set; }
        public string LogType { get; set; } = null!;
        public string Content { get; set; } = null!;
        public string Author { get; set; } = null!;
        public string Ip { get; set; } = null!;
        public DateTime CreatedDate { get; set; }
        public DateTime UpdatedDate { get; set; }
    }
}
