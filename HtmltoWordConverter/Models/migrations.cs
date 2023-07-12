using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class migrations
    {
        [Key]
        public int id { get; set; }
        public string migration { get; set; }
        public int batch { get; set; }
    }
}
