using System.ComponentModel.DataAnnotations;
using System;

namespace HtmltoWordConverter.Models
{
    public class PasswordReset
    {
        [Key]
        public string Email { get; set; }

        public string Token { get; set; }

        public DateTime? CreatedAt { get; set; }
    }
}
