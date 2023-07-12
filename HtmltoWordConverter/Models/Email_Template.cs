using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Email_Template
    {
        [Key]
        public int Email_Template_ID { get; set; }
        public string Column1 { get; set; }
        public string Image1 { get; set; }
        public string Column2 { get; set; }
        public string Image2 { get; set; }
        public int? Created_by { get; set; }
        public string Link { get; set; }
    }
}
