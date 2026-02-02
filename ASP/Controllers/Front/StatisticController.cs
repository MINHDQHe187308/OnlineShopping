using Microsoft.AspNetCore.Mvc;
using ASP.Models.ASPModel;
using Microsoft.EntityFrameworkCore;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace ASP.Controllers.Front
{
    [Route("Front/[controller]")]
    [Route("[controller]")]
    public class StatisticController : Controller
    {
        private readonly ASPDbContext _context;

        public StatisticController(ASPDbContext context)
        {
            _context = context;
        }

        [HttpGet("DelayHistory")]
        [HttpGet("Statistic")]
        public IActionResult DelayHistory()
        {
            return View("~/Views/Front/DensoWareHouse/Statistic.cshtml");
        }

        // GET /api/statistics/overview?from=yyyy-MM-dd&to=yyyy-MM-dd&customerCode=XXX
        [HttpGet("/api/statistics/overview")]
        public async Task<IActionResult> Overview([FromQuery] string? from, [FromQuery] string? to, [FromQuery] string? customerCode, [FromQuery] bool includeAll = false)
        {
            try
            {
                // Default: last 30 days if not provided
                DateTime dtTo = string.IsNullOrEmpty(to) ? DateTime.Today : DateTime.Parse(to);
                DateTime dtFrom = string.IsNullOrEmpty(from) ? dtTo.AddDays(-30) : DateTime.Parse(from);
                var start = dtFrom.Date;
                var end = dtTo.Date.AddDays(1); // exclusive

                var ordersQuery = _context.Orders.AsQueryable();
                if (!string.IsNullOrEmpty(customerCode))
                {
                    ordersQuery = ordersQuery.Where(o => o.CustomerCode == customerCode);
                }

                // Compare by Date to avoid excluding orders due to time component
                long totalOrders;
                int shipped;
                int pending;
                if (includeAll)
                {
                    var allOrders = ordersQuery; // no ShipDate filter
                    totalOrders = await allOrders.CountAsync();
                    shipped = await allOrders.Where(o => o.OrderStatus == 2 || o.OrderStatus == 3).CountAsync();
                    pending = await allOrders.Where(o => o.OrderStatus == 0 || o.OrderStatus == 1).CountAsync();
                }
                else
                {
                    var ordersInRange = ordersQuery.Where(o => o.ShipDate.Date >= start && o.ShipDate.Date <= dtTo.Date);
                    totalOrders = await ordersInRange.CountAsync();
                    // Shipped/Completed statuses: 2 (Completed) and 3 (Shipped)
                    shipped = await ordersInRange.Where(o => o.OrderStatus == 2 || o.OrderStatus == 3).CountAsync();
                    // Pending: 0 (Planned) and 1 (Pending)
                    pending = await ordersInRange.Where(o => o.OrderStatus == 0 || o.OrderStatus == 1).CountAsync();
                }

                // Orders with delay history in the period
                var delayQuery = _context.DelayHistory.AsQueryable();
                if (!string.IsNullOrEmpty(customerCode))
                {
                    // join to orders to filter by customer
                    delayQuery = from d in _context.DelayHistory
                                 join o in _context.Orders on d.OId equals o.UId
                                 where o.CustomerCode == customerCode
                                 select d;
                }

                var delayedOrdersCount = 0;
                if (includeAll)
                {
                    delayedOrdersCount = await delayQuery
                        .Select(d => d.OId)
                        .Distinct()
                        .CountAsync();
                }
                else
                {
                    delayedOrdersCount = await delayQuery
                        .Where(d => d.StartTime >= start && d.StartTime < end)
                        .Select(d => d.OId)
                        .Distinct()
                        .CountAsync();
                }

                // Orders flagged as advance either on Order or DelayHistory
                var advanceFromOrders = new System.Collections.Generic.List<Guid>();
                var advanceFromDelay = new System.Collections.Generic.List<Guid>();
                if (includeAll)
                {
                    advanceFromOrders = await ordersQuery.Where(o => o.IsAdvance == true).Select(o => o.UId).Distinct().ToListAsync();
                    advanceFromDelay = await delayQuery.Where(d => d.IsAdvance == true).Select(d => d.OId).Distinct().ToListAsync();
                }
                else
                {
                    var ordersInRange = ordersQuery.Where(o => o.ShipDate.Date >= start && o.ShipDate.Date <= dtTo.Date);
                    advanceFromOrders = await ordersInRange.Where(o => o.IsAdvance == true).Select(o => o.UId).Distinct().ToListAsync();
                    advanceFromDelay = await delayQuery.Where(d => d.StartTime >= start && d.StartTime < end && d.IsAdvance == true).Select(d => d.OId).Distinct().ToListAsync();
                }
                var advanceOrders = advanceFromOrders.Union(advanceFromDelay).Count();

                double totalDelayMinutes;
                if (includeAll)
                {
                    totalDelayMinutes = await delayQuery.SumAsync(d => (double?)d.DelayTime) ?? 0.0;
                }
                else
                {
                    totalDelayMinutes = await delayQuery.Where(d => d.StartTime >= start && d.StartTime < end).SumAsync(d => (double?)d.DelayTime) ?? 0.0;
                }

                var result = new
                {
                    totalOrders,
                    shipped,
                    pending,
                    delayedOrders = delayedOrdersCount,
                    advanceOrders,
                    totalDelayMinutes
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // GET /api/statistics/overview/customers?from=yyyy-MM-dd&to=yyyy-MM-dd&customers=code1,code2&top=5
        [HttpGet("/api/statistics/overview/customers")]
        public async Task<IActionResult> OverviewByCustomers([FromQuery] string? from, [FromQuery] string? to, [FromQuery] string? customers, [FromQuery] int top = 5, [FromQuery] bool includeAll = false)
        {
            try
            {
                DateTime dtTo = string.IsNullOrEmpty(to) ? DateTime.Today : DateTime.Parse(to);
                DateTime dtFrom = string.IsNullOrEmpty(from) ? dtTo.AddDays(-30) : DateTime.Parse(from);
                var start = dtFrom.Date;
                var end = dtTo.Date.AddDays(1); // exclusive

                List<string>? customerList = null;
                if (!string.IsNullOrEmpty(customers))
                {
                    customerList = customers.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
                }

                // If no explicit customers provided, return ALL distinct customer codes in the range
                // (Previously this returned only top N by order count.)
                if (customerList == null || customerList.Count == 0)
                {
                    var qCust = _context.Orders.AsQueryable();
                    if (!includeAll)
                    {
                        qCust = qCust.Where(o => o.ShipDate.Date >= start && o.ShipDate.Date <= dtTo.Date);
                    }
                    customerList = await qCust
                        .Where(o => !string.IsNullOrEmpty(o.CustomerCode))
                        .Select(o => o.CustomerCode)
                        .Distinct()
                        .OrderBy(c => c)
                        .ToListAsync();
                }

                if (customerList == null || customerList.Count == 0)
                {
                    return Ok(new List<object>());
                }

                // Prepare orders grouped metrics
                var ordersInRange = _context.Orders.Where(o => customerList.Contains(o.CustomerCode));
                if (!includeAll)
                    ordersInRange = ordersInRange.Where(o => o.ShipDate.Date >= start && o.ShipDate.Date <= dtTo.Date);

                var orderGroups = await ordersInRange
                    .GroupBy(o => o.CustomerCode)
                    .Select(g => new
                    {
                        CustomerCode = g.Key,
                        TotalOrders = g.Count(),
                        Shipped = g.Count(o => o.OrderStatus == 2 || o.OrderStatus == 3),
                        Pending = g.Count(o => o.OrderStatus == 0 || o.OrderStatus == 1),
                        AdvanceFromOrders = g.Count(o => o.IsAdvance == true)
                    }).ToListAsync();

                // Prepare delay history grouped metrics (join to orders to get customer code)
                var delayJoin = from d in _context.DelayHistory
                                join o in _context.Orders on d.OId equals o.UId
                                where customerList.Contains(o.CustomerCode)
                                select new { o.CustomerCode, d.OId, d.IsAdvance, d.DelayTime, d.StartTime };

                if (!includeAll)
                {
                    delayJoin = from x in delayJoin where x.StartTime >= start && x.StartTime < end select x;
                }

                var delayGroups = await delayJoin
                    .GroupBy(x => x.CustomerCode)
                    .Select(g => new
                    {
                        CustomerCode = g.Key,
                        DelayedOrders = g.Select(x => x.OId).Distinct().Count(),
                        AdvanceFromDelay = g.Where(x => x.IsAdvance == true).Select(x => x.OId).Distinct().Count(),
                        TotalDelayMinutes = g.Sum(x => (double?)x.DelayTime) ?? 0.0
                    }).ToListAsync();

                // Merge results by customerList order to keep consistent ordering
                var results = new List<object>();
                foreach (var cust in customerList)
                {
                    var og = orderGroups.FirstOrDefault(x => x.CustomerCode == cust);
                    var dg = delayGroups.FirstOrDefault(x => x.CustomerCode == cust);
                    var totalOrders = og?.TotalOrders ?? 0;
                    var shipped = og?.Shipped ?? 0;
                    var pending = og?.Pending ?? 0;
                    var advanceOrders = (og?.AdvanceFromOrders ?? 0) + (dg?.AdvanceFromDelay ?? 0);
                    var delayed = dg?.DelayedOrders ?? 0;
                    var totalDelayMinutes = dg?.TotalDelayMinutes ?? 0.0;

                    results.Add(new
                    {
                        customerCode = cust,
                        totalOrders,
                        shipped,
                        pending,
                        delayedOrders = delayed,
                        advanceOrders,
                        totalDelayMinutes
                    });
                }

                return Ok(results);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        // GET /api/statistics/overview/customers/monthly?from=yyyy-MM-dd&to=yyyy-MM-dd&customers=code1,code2
        [HttpGet("/api/statistics/overview/customers/monthly")]
        public async Task<IActionResult> OverviewByCustomersMonthly([FromQuery] string? from, [FromQuery] string? to, [FromQuery] string? customers, [FromQuery] bool includeAll = false)
        {
            try
            {
                DateTime dtTo = string.IsNullOrEmpty(to) ? DateTime.Today : DateTime.Parse(to);
                DateTime dtFrom = string.IsNullOrEmpty(from) ? dtTo.AddDays(-30) : DateTime.Parse(from);

                // Normalize to month boundaries
                DateTime monthStart;
                DateTime monthEnd;
                if (includeAll)
                {
                    var minShip = await _context.Orders.MinAsync(o => (DateTime?)o.ShipDate);
                    var maxShip = await _context.Orders.MaxAsync(o => (DateTime?)o.ShipDate);
                    if (minShip == null || maxShip == null)
                    {
                        return Ok(new { labels = new string[0], datasets = new object[0] });
                    }
                    monthStart = new DateTime(minShip.Value.Year, minShip.Value.Month, 1);
                    monthEnd = new DateTime(maxShip.Value.Year, maxShip.Value.Month, 1);
                }
                else
                {
                    monthStart = new DateTime(dtFrom.Year, dtFrom.Month, 1);
                    monthEnd = new DateTime(dtTo.Year, dtTo.Month, 1);
                }

                // Build months list
                var months = new System.Collections.Generic.List<(int Year, int Month)>();
                var cur = monthStart;
                while (cur <= monthEnd)
                {
                    months.Add((cur.Year, cur.Month));
                    cur = cur.AddMonths(1);
                }

                System.Collections.Generic.List<string>? customerList = null;
                if (!string.IsNullOrEmpty(customers))
                {
                    customerList = customers.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
                }

                if (customerList == null || customerList.Count == 0)
                {
                    // get all distinct customer codes in the range (or all if includeAll)
                    var qCust = _context.Orders.AsQueryable();
                    if (!includeAll)
                        qCust = qCust.Where(o => o.ShipDate.Date >= monthStart && o.ShipDate.Date < monthEnd.AddMonths(1));
                    customerList = await qCust
                        .Where(o => !string.IsNullOrEmpty(o.CustomerCode))
                        .Select(o => o.CustomerCode)
                        .Distinct()
                        .OrderBy(c => c)
                        .ToListAsync();
                }

                if (customerList == null || customerList.Count == 0)
                {
                    return Ok(new { labels = new string[0], datasets = new object[0] });
                }

                var startInclusive = monthStart;
                var endExclusive = monthEnd.AddMonths(1);

                var q = _context.Orders.AsQueryable();
                q = q.Where(o => customerList.Contains(o.CustomerCode));
                q = q.Where(o => o.ShipDate.Date >= startInclusive.Date && o.ShipDate.Date <= endExclusive.AddDays(-1));
                var q2 = q.GroupBy(o => new { o.CustomerCode, Year = o.ShipDate.Year, Month = o.ShipDate.Month })
                    .Select(g => new { g.Key.CustomerCode, g.Key.Year, g.Key.Month, Count = g.Count() });

                var raw = await q2.ToListAsync();

                var labels = months.Select(m => $"{m.Year}-{m.Month:D2}").ToArray();

                var datasets = new System.Collections.Generic.List<object>();
                foreach (var cust in customerList)
                {
                    var counts = months.Select(m =>
                    {
                        var found = raw.FirstOrDefault(r => r.CustomerCode == cust && r.Year == m.Year && r.Month == m.Month);
                        return found?.Count ?? 0;
                    }).ToArray();

                    datasets.Add(new { label = cust, data = counts });
                }

                return Ok(new { labels, datasets });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
