using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class SessionAdmin
    {
        [Key]
        public int SessionAdminID { get; set; }
        [MaxLength(32)]
        public string SessionID { get; set; }

        public int? SessionTime { get; set; }

        public short? AdminID { get; set; }
    }
}
