using DataAccessLibrary.Model;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;


namespace DataAccessLibrary
{
    public partial class PgoDBContext : IdentityDbContext<ApplicationUser>
    {
        public PgoDBContext(DbContextOptions<PgoDBContext> options) : base(options)
        {
            Database.EnsureCreated();
        }
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
        }
        public virtual DbSet<Analysis> Analysis { get; set; }
        public virtual DbSet<AnalysisClass> AnalysisClass { get; set; }
        public virtual DbSet<AnalysisClassCategory> AnalysisClassCategory  { get; set; }
        public virtual DbSet<Class> Class { get; set; }
        public virtual DbSet<Company> Company { get; set; }
        public virtual DbSet<CompanyType> CompanyType { get; set; }
        public virtual DbSet<ProjectionMethod> ProjectionMethod { get; set; }
        public virtual DbSet<RevenueExpenseData> RevenueExpenseData { get; set; }
    }
}
