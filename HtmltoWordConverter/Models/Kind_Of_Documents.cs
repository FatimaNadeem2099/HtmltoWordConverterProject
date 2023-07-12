using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Kind_Of_Documents
    {
        [Key]
        public int ID { get; set; }
        public string Name { get; set; }
        public int Status { get; set; }
        public DateTime Created_At { get; set; }
        public int Created_By { get; set; }
        public DateTime Updated_At { get; set; }
        public int Updated_By { get; set; }
    }
}
