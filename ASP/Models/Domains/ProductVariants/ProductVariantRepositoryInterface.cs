using ReflectionIT.Mvc.Paging;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public interface ProductVariantRepositoryInterface
    {
        Task<PagingList<ProductVariant>> GetAllByFilterAsync(string? searchString, int pageSize, int page, string? sort);
        Task<ProductVariant?> GetVariantByIdAsync(int id);
        Task<bool> CreateVariantAsync(ProductVariant variant);
        Task<(bool Success, string? Message)> UpdateVariantAsync(ProductVariant variant);
        Task<bool> DeleteVariantAsync(int id);
    }
}
