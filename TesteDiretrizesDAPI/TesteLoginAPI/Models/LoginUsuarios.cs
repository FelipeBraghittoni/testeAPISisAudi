using System.ComponentModel.DataAnnotations;

namespace TesteDiretrizesDAPI.Models
{
    public class LoginUsuarios
    {
        [Key]
        public decimal id { get; set; }
        public string usuario { get; set; }
        public string password { get; set; }
    }
}
