using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.Extensions.Configuration;
using SeverFashionShop.Models;
using SeverFashionShop.Models;
using System;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace ASP.ProductServer
{
    class Program
    {
        private static IConfiguration _configuration;
        private static IDbContextFactory<FaShionShopContext> _dbContextFactory;

        static async Task Main(string[] args)
        {

            _configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            string connectionString = _configuration.GetConnectionString("MyCnn");
            if (string.IsNullOrEmpty(connectionString))
            {
                Console.WriteLine(" Khong tim thay ConnectionString 'MyCnn' trong appsettings.json!");
                Console.WriteLine("Vui long kiem tra tra file appsettings.json");
                return;
            }


            var options = new DbContextOptionsBuilder<FaShionShopContext>()
                .UseSqlServer(connectionString)
                .Options;

            _dbContextFactory = new PooledDbContextFactory<FaShionShopContext>(options);


            int port = 5000;
            TcpListener listener = new TcpListener(IPAddress.Any, port);

            try
            {
                listener.Start();
                Console.WriteLine("==================================================");
                Console.WriteLine($"TCP Product Server dang chay tren port {port}");
                Console.WriteLine("IP: 127.0.0.1 ");
                Console.WriteLine("Port: 5000");
                Console.WriteLine("==================================================\n");

                while (true)
                {
                    TcpClient client = await listener.AcceptTcpClientAsync();
                    Console.WriteLine($"[{DateTime.Now}] Client ket noi tu: {client.Client.RemoteEndPoint}");


                    _ = Task.Run(() => HandleClientAsync(client));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($" Loi server: {ex.Message}");
            }
            finally
            {
                listener.Stop();
            }
        }


        private static async Task HandleClientAsync(TcpClient client)
        {
            using NetworkStream stream = client.GetStream();
            byte[] buffer = new byte[8192];

            try
            {
                while (true)
                {
                    int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);
                    if (bytesRead == 0)
                    {
                        Console.WriteLine($"[{DateTime.Now}] Client ngat ket noi.");
                        break;
                    }

                    string request = Encoding.UTF8.GetString(buffer, 0, bytesRead).Trim();
                    Console.WriteLine($"[{DateTime.Now}] Nhan lenh: {request}");

                    string response = await ProcessRequestAsync(request);

                    byte[] responseBytes = Encoding.UTF8.GetBytes(response + "\n");
                    await stream.WriteAsync(responseBytes, 0, responseBytes.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($" Loi xu ly client: {ex.Message}");
            }
            finally
            {
                client.Close();
            }
        }


        private static async Task<string> ProcessRequestAsync(string request)
        {
            if (!request.Equals("GET_PRODUCTS", StringComparison.OrdinalIgnoreCase))
            {
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    message = "Lenh khong hop le chi ho tro: GET_PRODUCTS"
                }, new JsonSerializerOptions { WriteIndented = true });
            }

            try
            {
                using var context = _dbContextFactory.CreateDbContext();


                var products = await context.Products
                    .Include(p => p.Category)
                    .Include(p => p.ProductImages)
                    .Include(p => p.ProductVariants)
                    .OrderByDescending(p => p.ProductId)
                    .ToListAsync();


                var data = products.Select(p => new
                {
                    p.ProductId,
                    p.ProductName,
                    Category = p.Category?.CategoryName ?? "N/A",
                    Price = p.ProductVariants?.FirstOrDefault()?.Price ?? 0m,
                    MainImage = p.ProductImages?.FirstOrDefault(x => x.IsMain)?.ImageUrl
                             ?? p.ProductImages?.FirstOrDefault()?.ImageUrl
                             ?? "/images/no-image.jpg",
                    p.Quantity
                }).ToList();

                var responseObj = new
                {
                    success = true,
                    command = "GET_PRODUCTS",
                    count = data.Count,
                    products = data,
                    timestamp = DateTime.Now
                };

                return JsonSerializer.Serialize(responseObj, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Loi khi lay du lieu Product: {ex.Message}");
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    message = "Loi server: " + ex.Message
                });
            }
        }
    }
}