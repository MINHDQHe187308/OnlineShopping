using Microsoft.AspNetCore.Mvc;

namespace ASP.Controllers.Front
{
    public class CheckoutController : Controller
    {
        public IActionResult Index()
        {
            return View("~/Views/Front/Checkout/Index.cshtml");
        }
    }
}
