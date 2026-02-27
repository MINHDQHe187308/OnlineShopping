using ASP.Models.Admin;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Logs;
using ASP.Models.Admin.Menus;
using ASP.Models.Admin.Roles;
using ASP.Models.Admin.ThemeOptions;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace ASP.Models.ASPModel
{
    public class ASPDbContext : IdentityDbContext<ApplicationUser, Role, string>
    {
        public ASPDbContext(DbContextOptions<ASPDbContext> options) : base(options) { }

        public override int SaveChanges()
        {
            SetModifiedInformation();
            return base.SaveChanges();
        }

        public override async Task<int> SaveChangesAsync(CancellationToken cancellationToken = new CancellationToken())
        {
            SetModifiedInformation();
            return await base.SaveChangesAsync(cancellationToken);
        }

        private void SetModifiedInformation()
        {
            var entries = ChangeTracker
                .Entries()
                .Where(e => e.Entity is BaseEntity && (
                        e.State == EntityState.Added
                        || e.State == EntityState.Modified));

            foreach (var entityEntry in entries)
            {
                // Use a single timestamp for consistency
                var now = DateTime.Now;

                // Update CLR properties for in-memory use (keeps existing behavior)
                try
                {
                    ((BaseEntity)entityEntry.Entity).UpdatedDate = now;
                    if (entityEntry.State == EntityState.Added)
                    {
                        ((BaseEntity)entityEntry.Entity).CreatedDate = now;
                    }
                }
                catch
                {
                    // If for some reason casting fails, continue to set shadow properties below
                }

                // Also set EF shadow properties so values are persisted even though CLR properties are [NotMapped]
                var updatedProp = entityEntry.Property("UpdatedDate");
                if (updatedProp != null)
                {
                    updatedProp.CurrentValue = now;
                }

                if (entityEntry.State == EntityState.Added)
                {
                    var createdProp = entityEntry.Property("CreatedDate");
                    if (createdProp != null)
                    {
                        createdProp.CurrentValue = now;
                    }
                }
            }
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Apply CreatedDate and UpdatedDate to all entities (giữ nguyên)
            var allEntities = modelBuilder.Model.GetEntityTypes();
            foreach (var entity in allEntities)
            {
                entity.AddProperty("CreatedDate", typeof(DateTime));
                entity.AddProperty("UpdatedDate", typeof(DateTime));
            }

          
        }

        public DbSet<Log> Logs { get; set; }

        public DbSet<Admin.ThemeOptions.ThemeOption> ThemeOptions { get; set; }

        public DbSet<Menu> Menus { get; set; }

       
    }
}