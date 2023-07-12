using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class languages
    {
        [Key]
        public int id { get; set; }
        public string name { get; set; }
        public int status { get; set; }
        public DateTime created_at { get; set; }
        public int created_by { get; set; }
        public DateTime updated_at { get; set; }
        public int updated_by { get; set; }
    }
}
