using Microsoft.EntityFrameworkCore;
using ASP.Models.ASPModel;
using ReflectionIT.Mvc.Paging;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public interface CategoryRepositoryInterface
    {
        IEnumerable<Category> GetAllCategories(); 
        Task<IEnumerable<Category>> GetAllCategoriesAsync();
        Task<PagingList<Category>> GetAllByFilterAsync(string? searchString, int pageSize, int page, string? sort);
        Task<Category?> GetCategoryByIdAsync(int id);
        Task<bool> CreateCategoryAsync(Category category);
        Task<bool> UpdateCategoryAsync(Category category);
        Task<bool> DeleteCategoryAsync(int id);
    }
}
