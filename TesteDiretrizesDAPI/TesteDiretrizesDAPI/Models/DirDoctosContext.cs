using Microsoft.EntityFrameworkCore;
using System.Diagnostics.CodeAnalysis;

namespace TesteDiretrizesDAPI.Models
{
    public class DirDoctosContext : DbContext
    {
        public DirDoctosContext(DbContextOptions<DirDoctosContext> options)
            : base(options)
        {
        }

    public DbSet<DirDoctosDir> DirDoctos { get; set; }
    }
}
