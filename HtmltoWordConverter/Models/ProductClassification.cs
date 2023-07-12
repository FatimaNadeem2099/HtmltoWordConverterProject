using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class ProductClassification
    {
        [Key]
        public int Id { get; set; }

        public string Name { get; set; }

        [Display(Name = "Sector")]
        public int? SectorId { get; set; }
        [ForeignKey("SectorId")]
        public Sector Sector { get; set; }

        public int Status { get; set; }

        public int? CreatedBy { get; set; }

        public DateTime? CreatedAt { get; set; }

        public int? UpdatedBy { get; set; }

        public DateTime? UpdatedAt { get; set; }

        
    }
}
