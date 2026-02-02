using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ASP.Controllers
{
    public class Page404Controller : Controller
    {
        // GET: Page404Controller
        [Route("Page404")]
        [Route("Page404/Index")]
        public ActionResult Index()
        {
            return View("Index");
        }
    }
}
