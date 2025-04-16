using ManagerPdf.Models;
using Microsoft.EntityFrameworkCore;

namespace ManagerPdf.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<ApplicationUser> Users { get; set; }
    }
}
