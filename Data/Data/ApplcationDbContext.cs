using Data.Models;
using Microsoft.EntityFrameworkCore;

namespace Data.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<Student> students { get; set; }
    }
}