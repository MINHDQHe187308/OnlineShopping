using Microsoft.AspNetCore.Mvc;
using ASP.Models.Front;

namespace ASP.Controllers.Front
{
    public class Statisticv1Controller : Controller
    {
        private readonly OrderRepositoryInterface _orderRepository;

        public Statisticv1Controller(OrderRepositoryInterface orderRepository)
        {
            _orderRepository = orderRepository;
        }

        public IActionResult Statisticv1()
        {
            return View("~/Views/Front/DensoWareHouse/Statisticv1.cshtml");
        }

        [HttpGet]
        public async Task<IActionResult> GetOrderStatistics()
        {
            var orders = await _orderRepository.GetAllOrdersAsync();

        
            var uniqueCustomersInOrder = orders
                .Select(o => o.CustomerCode)
                .Where(code => code != null) 
                .Distinct()
                .ToList();


            var stats = uniqueCustomersInOrder
                .Select(code =>
                {
                    var group = orders.Where(o => o.CustomerCode == code);
                    return new
                    {
                        CustomerCode = code ?? "Unknown",
                        Plan = group.Count(),
                        Progress = group.Count(o => o.OrderStatus == 1),
                        Completed = group.Count(o => o.OrderStatus == 2 || (o.OrderStatus == 3 && !(o.IsAdvance ?? false))),
                        Delay = group.Count(o => o.OrderStatus == 4),
                        Advance = group.Count(o => o.OrderStatus == 3 && (o.IsAdvance ?? false))
                    };
                })
                .Where(s => s.Plan + s.Progress + s.Completed + s.Delay + s.Advance > 0)
               // .Take(14)
                .ToList();
           

            return Json(stats);
        }
    }
}