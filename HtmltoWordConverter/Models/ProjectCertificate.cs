using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class ProjectCertificate
    {
        [Key]
        public int ProjectCertificateId { get; set; }

        public string CertificateNumber { get; set; }

        public string Certificate1 { get; set; }

        public string Certificate2 { get; set; }

        [Required]
        public int CreatedBy { get; set; }

        [Required]
        public DateTime CreatedOn { get; set; }
    }
}
