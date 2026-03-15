using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using Microsoft.AspNetCore.SignalR;
using ASP.Hubs;

namespace ASP.Models.Domains
{
    public class CategoryRepository : CategoryRepositoryInterface   
    {
        private readonly ASPDbContext _context;
        private readonly IHubContext<AdminHub> _hubContext;

        public CategoryRepository(ASPDbContext context, IHubContext<AdminHub> hubContext)
        {
            _context = context;
            _hubContext = hubContext;
        }
        public IEnumerable<Category> GetAllCategories()
        {
            return _context.Categories
                .OrderBy(c => c.CategoryName)
                .ToList();
        }

        public async Task<IEnumerable<Category>> GetAllCategoriesAsync()
        {
            return await _context.Categories
                .OrderByDescending(c => c.CategoryId)
                .ToListAsync();
        }

        public async Task<Category?> GetCategoryByIdAsync(int id)
        {
            return await _context.Categories
                .FirstOrDefaultAsync(c => c.CategoryId == id);
        }

        public async Task<bool> CreateCategoryAsync(Category category)
        {
            try
            {
                _context.Categories.Add(category);
                await _context.SaveChangesAsync();

                // Real-time notification
                await _hubContext.Clients.All.SendAsync("CategoryCreated", category);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> UpdateCategoryAsync(Category category)
        {
            try
            {
                var existing = await _context.Categories.FindAsync(category.CategoryId);
                if (existing == null) return false;

                existing.CategoryName = category.CategoryName;
                
                await _context.SaveChangesAsync();

                // Real-time notification
                await _hubContext.Clients.All.SendAsync("CategoryUpdated", existing);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<bool> DeleteCategoryAsync(int id)
        {
            try
            {
                var category = await _context.Categories.FindAsync(id);
                if (category == null) return false;

                bool isReferenced = await _context.Products.AnyAsync(p => p.CategoryId == id);

                if (isReferenced)
                {
                    // If referenced, we might not be able to delete it due to restrict constraint
                    return false;
                }
                
                _context.Categories.Remove(category);
                await _context.SaveChangesAsync();

                // Real-time notification
                await _hubContext.Clients.All.SendAsync("CategoryDeleted", id);
                
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
