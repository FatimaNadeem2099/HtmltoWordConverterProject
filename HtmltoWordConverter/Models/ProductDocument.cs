using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class ProductDocument
    {
        [Key]
        public long Id { get; set; }

        public string Name { get; set; }

        public string NameAr { get; set; }

        public int? Pages { get; set; }

        public byte? Status { get; set; }

        public DateTime? CreatedAt { get; set; }

        public DateTime? UpdatedAt { get; set; }
    }
}
