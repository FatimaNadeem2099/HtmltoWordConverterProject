using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Guidelines
    {
        [Key]
        public long ID { get; set; }
        public string Name { get; set; }
        public string Heading { get; set; }
        public string Link { get; set; }
        public DateTime? Created_At { get; set; }
        public DateTime? Updated_At { get; set; }
    }
}
