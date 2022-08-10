using Microsoft.EntityFrameworkCore;
using System.Diagnostics.CodeAnalysis;

namespace TesteDiretrizesDAPI.Models
{
    public class LoginUsuariosContext : DbContext
    {
        public LoginUsuariosContext(DbContextOptions<LoginUsuariosContext> options)
           : base(options)
        {
        }

        public DbSet<LoginUsuarios> Login { get; set; }
    }
}
