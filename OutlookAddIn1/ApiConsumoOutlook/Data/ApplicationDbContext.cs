using ApiConsumoOutlook.Models;
using Microsoft.EntityFrameworkCore;

namespace ApiConsumoOutlook.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {

        }


        public DbSet<Proyectos> Proyectos { get; set; }
    }
}
