using System;
using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Failed_Jobs
    {
        [Key]
        public long ID { get; set; }
        public string UUID { get; set; }
        public string Connection { get; set; }
        public string Queue { get; set; }
        public string Payload { get; set; }
        public string Exception { get; set; }
        public DateTime Failed_At { get; set; }
    }
}
