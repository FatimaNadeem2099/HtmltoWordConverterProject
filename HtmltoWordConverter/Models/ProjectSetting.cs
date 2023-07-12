using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class ProjectSetting
    {
        [Key]
        public int SettingId { get; set; }

        public string ProjectName { get; set; }

        [Required]
        public int Product { get; set; }

        [Required]
        public int Query { get; set; }

        [Required]
        public int ProductQuery { get; set; }
    }
}
