using ASP.Models.Admin;
using ASP.Models.Admin.Accounts;
using ASP.Models.Admin.Logs;
using ASP.Models.Admin.Menus;
using ASP.Models.Admin.Roles;
using ASP.Models.Admin.ThemeOptions;
using ASP.Models.Domains;
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

            modelBuilder.Entity<Category>().HasIndex(c => c.CategoryName).IsUnique();

            modelBuilder.Entity<ProductVariant>()
                .HasIndex(v => v.SKU)
                .IsUnique();

            // Computed column
            modelBuilder.Entity<OrderDetail>()
                .Property(p => p.TotalPrice)
                .HasComputedColumnSql("[Quantity] * [UnitPrice]");

            // Relationships

            modelBuilder.Entity<Product>()
                .HasOne(p => p.Category)
                .WithMany(c => c.Products)
                .HasForeignKey(p => p.CategoryId)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<ProductImage>()
                .HasOne(pi => pi.Product)
                .WithMany(p => p.ProductImages)
                .HasForeignKey(pi => pi.ProductId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<ProductVariant>()
                .HasOne(v => v.Product)
                .WithMany(p => p.ProductVariants)
                .HasForeignKey(v => v.ProductId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<Order>()
                .HasOne(o => o.Customer)
                .WithMany(c => c.Orders)
                .HasForeignKey(o => o.CustomerId)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<Order>()
                .HasOne(o => o.User)
                .WithMany(u => u.Orders)
                .HasForeignKey(o => o.CreatedBy)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<OrderDetail>()
                .HasOne(od => od.Order)
                .WithMany(o => o.OrderDetails)
                .HasForeignKey(od => od.OrderId)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<OrderDetail>()
                .HasOne(od => od.ProductVariant)
                .WithMany()
                .HasForeignKey(od => od.VariantId)
                .OnDelete(DeleteBehavior.Restrict);

            modelBuilder.Entity<Notification>()
                .HasOne(n => n.User)
                .WithMany(u => u.Notifications)
                .HasForeignKey(n => n.UserId)
                .OnDelete(DeleteBehavior.Cascade);


        }

        public DbSet<Log> Logs { get; set; }

        public DbSet<Admin.ThemeOptions.ThemeOption> ThemeOptions { get; set; }

        public DbSet<Menu> Menus { get; set; }

        public DbSet<User> ShopUsers { get; set; }
        public DbSet<Category> Categories { get; set; }
        public DbSet<Product> Products { get; set; }
        public DbSet<ProductImage> ProductImages { get; set; }
        public DbSet<ProductVariant> ProductVariants { get; set; }
        public DbSet<Customer> Customers { get; set; }
        public DbSet<Order> Orders { get; set; }
        public DbSet<OrderDetail> OrderDetails { get; set; }
        public DbSet<Notification> Notifications { get; set; }







    }
}