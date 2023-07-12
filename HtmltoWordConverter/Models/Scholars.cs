using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class Scholars
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(255)]
        public string FirstName { get; set; }

        [Required]
        [MaxLength(255)]
        public string LastName { get; set; }

        [Required]
        [MaxLength(255)]
        public string FirstNameAr { get; set; }

        [Required]
        [MaxLength(255)]
        public string LastNameAr { get; set; }

        [MaxLength(255)]
        public string ShortName { get; set; }

        [MaxLength(255)]
        public string Picture { get; set; }

        public string Description { get; set; }

        [Required]
        public int Status { get; set; }

        [Required]
        public DateTime CreatedAt { get; set; }

        public int? CreatedBy { get; set; }

        public DateTime? UpdatedAt { get; set; }

        public int? UpdatedBy { get; set; }

        [Required]
        [MaxLength(255)]
        public string Email { get; set; }

        public string Signature { get; set; }
    }
}
