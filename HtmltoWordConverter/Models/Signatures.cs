using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class Signatures
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public long Id { get; set; }

        [Required]
        public int ProjectId { get; set; }

        [Required]
        [MaxLength(255)]
        public string CertificateNumber { get; set; }

        [Required]
        public int ScholarId { get; set; }

        [Required]
        public string Files { get; set; }

        public DateTime? CreatedAt { get; set; }

        public DateTime? UpdatedAt { get; set; }
    }
}
