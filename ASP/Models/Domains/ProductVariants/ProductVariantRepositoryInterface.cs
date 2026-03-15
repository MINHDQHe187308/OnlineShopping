using System.Collections.Generic;
using System.Threading.Tasks;

namespace ASP.Models.Domains
{
    public interface ProductVariantRepositoryInterface
    {
        Task<IEnumerable<ProductVariant>> GetAllVariantsAsync();
        Task<ProductVariant?> GetVariantByIdAsync(int id);
        Task<bool> CreateVariantAsync(ProductVariant variant);
        Task<bool> UpdateVariantAsync(ProductVariant variant);
        Task<bool> DeleteVariantAsync(int id);
    }
}
