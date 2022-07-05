using Microsoft.EntityFrameworkCore;

namespace CaisseWorksheet
{
    public class CaisseContext : DbContext
    {
        //      public DbSet<Student> Student { get; set; }
        public DbSet<Rapport> Rapport { get; set; }
        public DbSet<Article> Article { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=C:\pc\perso\marmite\Caisse\CaisseWorksheet\softcaisse.mdf;Integrated Security=True");
        
        }
    }
}
