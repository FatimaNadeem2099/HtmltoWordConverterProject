using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class Observation
    {
       
            [Key]
            public long ObservationId { get; set; }

            public string Year { get; set; }

            public string Function { get; set; }

            public string Description { get; set; }

            public string ObservationDetails { get; set; }

            public string Risk { get; set; }

            public string Recommendation { get; set; }

            public string ManagementResponse { get; set; }

            public DateTime? DateOfAudit { get; set; }

            public byte? Status { get; set; }

            public string Reason { get; set; }

        [Display(Name = "Companies")]
            public int CompanyId { get; set; }
        [ForeignKey("CompanyId")]
        public virtual Companies Companies { get; set; }

        public int? ProjectId { get; set; }

            public DateTime? CreatedAt { get; set; }

            public DateTime? UpdatedAt { get; set; }

        
    }
}
