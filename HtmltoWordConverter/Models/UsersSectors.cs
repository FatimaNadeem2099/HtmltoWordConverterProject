using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class UsersSectors
    {
        [Key]
        public int UsersSectorsId { get; set; }
        [Required]
        public short UserId { get; set; }

        [Required]
        public int SectorId { get; set; }
    }
}
