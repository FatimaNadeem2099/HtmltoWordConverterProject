using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Companies
    {
        [Key]
        public int CompanyId { get; set; }
        public string Name { get; set; }
        public string Name_AR { get; set; }
        public string Short_Name { get; set; }
        public string Picture { get; set; }
        public string Description { get; set; }
        public int Status { get; set; }
        public DateTime Created_at { get; set; }
        public int? Created_by { get; set; }
        public DateTime? Updated_at { get; set; }
        public int? Updated_by { get; set; }
        public long? Shariyah_Company_ID { get; set; }
    }
}
