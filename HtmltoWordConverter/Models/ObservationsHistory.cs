using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class ObservationsHistory
    {
        [Key]
        public long Id { get; set; }

        public string Field { get; set; }

        public string OriginalValue { get; set; }

        public string ChangedValue { get; set; }

        [Display(Name = "Observation")]
        public long ObservationId { get; set; }
        [ForeignKey("ObservationId")]
        public virtual Observation Observation { get; set; }

        public DateTime? CreatedAt { get; set; }

        public DateTime? UpdatedAt { get; set; }

        
    }
}
