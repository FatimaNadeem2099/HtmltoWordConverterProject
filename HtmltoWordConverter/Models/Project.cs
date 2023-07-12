using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Project
    {
        [Key]
        public int ProjectId { get; set; }

        [Required]
        public string ClientCode { get; set; }

        [Required]
        public string CompanyClient { get; set; }

        public string CertificateStatus { get; set; }

        public string ContactPerson { get; set; }

        public string Email { get; set; }

        public string ContactNumber { get; set; }

        public string CertificateNumber { get; set; }

        public string Year { get; set; }

        public string Month { get; set; }

        public string NameOfFund { get; set; }

        public string TypeOfFund { get; set; }

        public string FundSize { get; set; }

        public string FundCurrency { get; set; }

        public string Domicile { get; set; }

        public string WorkType { get; set; }

        public string Query { get; set; }

        public string LaunchStatus { get; set; }

        public string NotLaunched { get; set; }

        public string Product { get; set; }

        public string NoOfDocuments { get; set; }

        public string DateReceived { get; set; }

        public string DateCompleted { get; set; }

        public string NoOfDays { get; set; }

        public string HoursForReview { get; set; }

        public string Pages { get; set; }

        public string Language { get; set; }

        public string ScholarsName { get; set; }

        public string OtherEmployees { get; set; }

        public string NoOfTouchpoints { get; set; }

        public string EstimatedCalls { get; set; }

        public string CallTiming { get; set; }

        public string Audit { get; set; }

        public string Screening { get; set; }

        public string Certificate1 { get; set; }

        public string Certificate2 { get; set; }

        public string Status { get; set; }

        public string Remarks1 { get; set; }

        public string Remarks2 { get; set; }

        [Required]
        public string XXs { get; set; }

        [Required]
        public string YXs { get; set; }
    }
}
