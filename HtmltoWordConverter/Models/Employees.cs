using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Employees
    {
        [Key]
        public int ID { get; set; }
        public string First_Name { get; set; }
        public string Last_Name { get; set; }
        public string First_Name_AR { get; set; }
        public string Last_Name_AR { get; set; }
        public string Short_Name { get; set; }
        public string Picture { get; set; }
        public string Description { get; set; }
        public int Status { get; set; }
        public DateTime Created_at { get; set; }
        public int Created_by { get; set; }
        public DateTime? Updated_at { get; set; }
        public int Updated_by { get; set; }
    }
}
