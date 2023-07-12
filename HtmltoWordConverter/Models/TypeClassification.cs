using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class TypeClassification
    {
        [Key]
        public int TypeClassificationId { get; set; }
        [Required]
        public int TypeId { get; set; }

        [Required]
        public int ClassificationId { get; set; }
    }
}
