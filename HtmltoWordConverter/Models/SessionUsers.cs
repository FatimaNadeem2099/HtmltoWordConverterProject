using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class SessionUsers
    {
        public int SessionUsersID { get; set; }
        [MaxLength(32)]
        public string SessionID { get; set; }

        public int? SessionTime { get; set; }

        public short? UserID { get; set; }
    }
}
