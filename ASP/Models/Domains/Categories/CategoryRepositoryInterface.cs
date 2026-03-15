using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;
namespace ASP.Models.Domains
{
    public interface CategoryRepositoryInterface
    {
        IEnumerable<Category> GetAllCategories(); 
        Task<IEnumerable<Category>> GetAllCategoriesAsync();
        Task<Category?> GetCategoryByIdAsync(int id);
        Task<bool> CreateCategoryAsync(Category category);
        Task<bool> UpdateCategoryAsync(Category category);
        Task<bool> DeleteCategoryAsync(int id);
    }
}
