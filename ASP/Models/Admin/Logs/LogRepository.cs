using ASP.BaseCommon;
using ASP.Models.Admin.Logs;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Routing;
using ReflectionIT.Mvc.Paging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Security.Claims;
using System.Threading.Tasks;
using ASP.Models.ASPModel;

namespace ASP.Models.Admin.Logs
{
    public class LogRepository : Controller, LogRepositoryInterface
    {
        protected readonly ASPDbContext _context;
        protected readonly IHttpContextAccessor httpContextAccessor;
        public LogRepository(ASPDbContext context, IHttpContextAccessor httpContextAccessor)
        {
            _context = context;
            this.httpContextAccessor = httpContextAccessor;
        }
        public void CreateLog(string type, string content, string userId = null)
        {
            #region get IP
            string ipAddress = null;
            var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    ipAddress = ip.ToString();
                }
            }
            #endregion
            //
            _context.Add(new Log
            {
                LogType = type,
                Content = content,
                Author = httpContextAccessor.HttpContext.User.Identity.Name != null ? httpContextAccessor.HttpContext.User.Identity.Name : "Initial",
                IP = ipAddress
            });
        }
        public async Task<PagingList<Log>> GetAllByLimit(string filter = "", int pagesize = 10, int page = 0, string sort = null)
        {
            var qry = _context.Logs.Select(s => new Log()
            {
                ID = s.ID,
                LogType = s.LogType,
                Content = s.Content,
                Author = s.Author,
                IP = s.IP,
                CreatedDate = s.CreatedDate
            }).OrderByDescending(f => f.ID).AsQueryable();
            if (!string.IsNullOrWhiteSpace(filter))
            {
                qry = qry.Where(p => p.LogType.Contains(filter));
            }
            var objs = await PagingList.CreateAsync(qry, pagesize, page, sort, "ID");
            objs.RouteValue = new RouteValueDictionary { { "filter", filter }, { "psize", pagesize } };
            return objs;
        }
        public int CountTotalAll(string logType = "")
        {
            var qry = _context.Logs.Count(c => c.LogType == logType);
            return qry;
        }
    }
}
