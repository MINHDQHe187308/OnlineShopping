using ASP.Models.Admin;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Logs;
using ASP.Models.Admin.Menus;
using ASP.Models.Admin.Roles;
using ASP.Models.Admin.ThemeOptions;
using ASP.Models.Front;
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

            // Configure Role Default Value (giữ nguyên)
            modelBuilder.Entity<Role>()
                .Property(b => b.DefaultRole)
                .HasDefaultValue(false);

            // Cấu hình One-to-Many: Customer -> LeadtimeMaster (cập nhật với composite key)
            modelBuilder.Entity<LeadtimeMaster>()
                .HasKey(lm => new { lm.CustomerCode, lm.TransCd });  // Composite PK mới

            modelBuilder.Entity<LeadtimeMaster>()
                .HasOne(lm => lm.Customer)
                .WithMany(c => c.LeadtimeMasters)
                .HasForeignKey(lm => lm.CustomerCode)
                .OnDelete(DeleteBehavior.Cascade)  // Xóa cascade nếu xóa Customer
                .IsRequired();

            // Cấu hình One-to-Many: Customer -> ShippingSchedule (giữ nguyên)
            modelBuilder.Entity<ShippingSchedule>()
                .HasOne(ss => ss.Customer)
                .WithMany(c => c.ShippingSchedules)
                .HasForeignKey(ss => ss.CustomerCode)
                .OnDelete(DeleteBehavior.Cascade)
                .IsRequired();

            // Existing configurations (giữ nguyên)
            modelBuilder.Entity<ShoppingList>()
                .HasOne(s => s.ThreePointCheck)
                .WithOne(t => t.ShoppingList)
                .HasForeignKey<ThreePointCheck>(t => t.SPId)
                .OnDelete(DeleteBehavior.Cascade)
                .IsRequired(false);

            modelBuilder.Entity<ThreePointCheck>()
                .HasIndex(t => t.SPId)
                .IsUnique();

            modelBuilder.Entity<ShippingSchedule>()
                .HasKey(e => new { e.CustomerCode, e.TransCd, e.Weekday });

            modelBuilder.Entity<ShippingSchedule>()
                .Property(e => e.CutOffTime)
                .HasConversion(
                    v => v.ToTimeSpan(),
                    v => TimeOnly.FromTimeSpan(v)
                )
                .HasColumnType("time");
        }

        public DbSet<Log> Logs { get; set; }

        public DbSet<Admin.ThemeOptions.ThemeOption> ThemeOptions { get; set; }

        public DbSet<Menu> Menus { get; set; }

        public DbSet<Customer> Customers { get; set; }

        public DbSet<DelayHistory> DelayHistory { get; set; }

        public DbSet<LeadtimeMaster> LeadtimeMasters { get; set; }

        public DbSet<Order> Orders { get; set; }

        public DbSet<OrderDetail> OrderDetails { get; set; }

        public DbSet<ShoppingList> ShoppingLists { get; set; }

        public DbSet<ThreePointCheck> ThreePointChecks { get; set; }

        public DbSet<ShippingSchedule> ShippingSchedules { get; set; }
    }
}