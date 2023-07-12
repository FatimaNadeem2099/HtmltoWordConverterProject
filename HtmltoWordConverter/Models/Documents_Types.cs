using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Documents_Types
    {
        [Key]
        public int Document_ID { get; set; }
        public int Type_ID { get; set; }
    }
}
