using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class ProjectAudit
    {
        [Key]
        public int AuditId { get; set; }

        public int? ProjectId { get; set; }

        public string ProjectName { get; set; }

        public string PersonalInCharge { get; set; }

        public int? Year { get; set; }

        public int? NumberOfEmails { get; set; }

        public string AuditStart { get; set; }

        public string AuditEnd { get; set; }

        public string NumberProductsTransactionsDepartments { get; set; }

        public string PepoleInvolved { get; set; }

        public string Hours { get; set; }

        public string Comments { get; set; }

        [Required]
        public DateTime LastUpdated { get; set; }
    }
}
