using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.ThemeOptions
{
    public interface ThemeOptionRepositoryInterface
    {
        public Task<IActionResult> UpdateThemeOptionById(ThemeOptionListenerInterface listener, IFormCollection request, List<IFormFile> formFile);
        //
        public List<ThemeOption> GetAll(string typeData);
        public ThemeOption GetThemeOption(string strName, string typeData);

    }
}
