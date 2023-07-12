using System.ComponentModel.DataAnnotations;

namespace HtmltoWordConverter.Models
{
    public class Admin
    {
        [Key]
        public int AdminID { get; set; }
        public string UserName { get; set; }
        public string Pwd { get; set; }
    }
}
