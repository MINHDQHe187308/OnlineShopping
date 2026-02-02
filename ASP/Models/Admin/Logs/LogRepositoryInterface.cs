using Microsoft.AspNetCore.Mvc;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Models.Admin.Logs
{
    public interface LogRepositoryInterface
    {
        public void CreateLog(string type, string content, string userId = null);
        public Task<PagingList<Log>> GetAllByLimit(string filter, int pagesize, int limit, string sort);
        public int CountTotalAll(string logType = "");
    }
}
