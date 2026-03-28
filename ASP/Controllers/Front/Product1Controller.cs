using ASP.Models.Domains;
using ASP.Models.ViewModels;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore; 
using System.Threading.Tasks;
using OfficeOpenXml;
using Microsoft.AspNetCore.Authorization;
namespace ASP.Controllers.Front
{
    public class Product1Controller : Controller
    {
        private readonly ProductRepositoryInterface _productRepository;
        private readonly ProductImageRepositoryInterface _productImageRepository;
        private readonly ASP.Models.ASPModel.ASPDbContext _context;
        public Product1Controller(ProductRepositoryInterface productRepository, ProductImageRepositoryInterface productImageRepository, ASP.Models.ASPModel.ASPDbContext context)
        {
            _productRepository = productRepository;
            _productImageRepository = productImageRepository;
            _context = context;
        }


        public IActionResult RemoteStore()
        {
            try
            {
          
                var products = _productRepository.GetAllProducts();

              
                if (products == null) products = new List<Product>();

                return View("~/Views/Front/Products/RemoteStore.cshtml", products);
            }
            catch (Exception ex)
            {           
                ViewBag.ErrorMessage = "Không thể kết nối tới Socket Server: " + ex.Message;
                return View("~/Views/Front/Products/RemoteStore.cshtml", new List<Product>());
            }
        }

        public IActionResult Index(
   string? keyword,
   int? category,
   string? size,
   string? color,
   decimal? minPrice,
   decimal? maxPrice,
   string? sort,
   int page = 1)
        {
            const int pageSize = 9;

            if (page < 1) page = 1;

            var query = _productRepository.QueryProducts()
                .Where(p => p.IsActive)
                .Where(p => p.ProductVariants.Any(v => v.IsActive));

            if (!string.IsNullOrWhiteSpace(keyword))
            {
                keyword = keyword.Trim();
                query = query.Where(p =>
                    p.ProductName.Contains(keyword) ||
                    (p.Description != null && p.Description.Contains(keyword)));
            }

            if (category.HasValue)
            {
                query = query.Where(p => p.CategoryId == category.Value);
            }

            if (!string.IsNullOrWhiteSpace(size))
            {
                query = query.Where(p => p.ProductVariants.Any(v => v.IsActive && v.Size == size));
            }

            if (!string.IsNullOrWhiteSpace(color))
            {
                query = query.Where(p => p.ProductVariants.Any(v => v.IsActive && v.Color == color));
            }

            if (minPrice.HasValue)
            {
                query = query.Where(p => p.ProductVariants.Any(v => v.IsActive && v.Price >= minPrice.Value));
            }

            if (maxPrice.HasValue)
            {
                query = query.Where(p => p.ProductVariants.Any(v => v.IsActive && v.Price <= maxPrice.Value));
            }

            switch (sort)
            {
                case "price_asc":
                    query = query.OrderBy(p => p.ProductVariants
                        .Where(v => v.IsActive)
                        .Min(v => (decimal?)v.Price) ?? 0);
                    break;

                case "price_desc":
                    query = query.OrderByDescending(p => p.ProductVariants
                        .Where(v => v.IsActive)
                        .Max(v => (decimal?)v.Price) ?? 0);
                    break;

                case "newest":
                    query = query.OrderByDescending(p => p.ProductId);
                    break;

                case "popular":
                    query = query.OrderByDescending(p => _context.OrderDetails
                        .Where(od => od.ProductVariant.ProductId == p.ProductId)
                        .Sum(od => (int?)od.Quantity) ?? 0);
                    break;

                default:
                    query = query.OrderByDescending(p => p.ProductId);
                    break;
            }

            var totalItems = query.Count();
            var totalPages = (int)Math.Ceiling((double)totalItems / pageSize);

            if (totalPages == 0) totalPages = 1;
            if (page > totalPages) page = totalPages;

            var products = query
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToList();

            var baseQueryForOptions = _productRepository.QueryProducts()
                .Where(p => p.IsActive)
                .Where(p => p.ProductVariants.Any(v => v.IsActive));

            var sizeOptions = baseQueryForOptions
                .SelectMany(p => p.ProductVariants.Where(v => v.IsActive))
                .Where(v => !string.IsNullOrWhiteSpace(v.Size))
                .Select(v => v.Size!)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            var colorOptions = baseQueryForOptions
                .SelectMany(p => p.ProductVariants.Where(v => v.IsActive))
                .Where(v => !string.IsNullOrWhiteSpace(v.Color))
                .Select(v => v.Color!)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            ViewBag.Keyword = keyword;
            ViewBag.SelectedCategory = category;
            ViewBag.SelectedSize = size;
            ViewBag.SelectedColor = color;
            ViewBag.MinPrice = minPrice;
            ViewBag.MaxPrice = maxPrice;
            ViewBag.Sort = sort;

            ViewBag.SizeOptions = sizeOptions;
            ViewBag.ColorOptions = colorOptions;

            ViewBag.CurrentPage = page;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalItems = totalItems;
            ViewBag.TotalPages = totalPages;

            return View("~/Views/Front/Products/Index.cshtml", products);
        }

        public async Task<IActionResult> Details(int id)
        {
            var product = await _productRepository.GetProductByIdAsync(id);

            if (product == null || !product.IsActive)
            {
                return NotFound();
            }

            product.ProductVariants = product.ProductVariants?
                .Where(v => v.IsActive)
                .OrderBy(v => v.VariantId)
                .ToList();

            if (product.ProductVariants == null || !product.ProductVariants.Any())
            {
                return NotFound();
            }

            var defaultVariant = product.ProductVariants.FirstOrDefault();

            var viewModel = new ProductDetailViewModel
            {
                Product = product,
                Images = product.ProductImages?.ToList() ?? new List<ProductImage>(),
                DefaultVariant = defaultVariant,
                CurrentColor = defaultVariant?.Color ?? "Chưa có màu",
                CurrentPrice = defaultVariant?.Price ?? 0
            };

            viewModel.MainImageUrl = product.ProductImages?
                .FirstOrDefault(x => x.IsMain)?.ImageUrl
                ?? product.ProductImages?.FirstOrDefault()?.ImageUrl
                ?? "/images/no-image.jpg";

            return View("~/Views/Front/Products/ProductDetail.cshtml", viewModel);
        }

        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> ImportProductsFromFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Error"] = "Vui lòng chọn file Excel!";
                return RedirectToAction("Index");
            }

            var list = new List<Product>();

            try
            {
                using (var stream = new MemoryStream())
                {
                    await file.CopyToAsync(stream);
              
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rowCount = worksheet.Dimension?.Rows ?? 0;

                        if (rowCount < 2)
                        {
                            TempData["Error"] = "File Excel không có dữ liệu hoặc sai định dạng!";
                            return RedirectToAction("Index");
                        }

                        for (int row = 2; row <= rowCount; row++)
                        {
                         
                            var productName = worksheet.Cells[row, 1].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(productName)) continue;

                            var product = new Product
                            {
                                ProductName = productName,
                                CategoryId = int.Parse(worksheet.Cells[row, 2].Value?.ToString() ?? "1"),
                                Description = worksheet.Cells[row, 3].Value?.ToString() ?? "",
                                Quantity = int.Parse(worksheet.Cells[row, 4].Value?.ToString() ?? "0"),
                                IsActive = true,
                                ProductVariants = new List<ProductVariant>
                        {
                            new ProductVariant
                            {
                                SKU = "SKU-" + Guid.NewGuid().ToString().Substring(0, 8).ToUpper(),
                                Price = decimal.Parse(worksheet.Cells[row, 5].Value?.ToString() ?? "0"),
                                QuantityVariants = int.Parse(worksheet.Cells[row, 4].Value?.ToString() ?? "0"),
                                IsActive = true
                            }
                        }
                            };
                            list.Add(product);
                        }
                    }
                }

                if (list.Any())
                {
                    await _productRepository.ImportProductsAsync(list);
                    TempData["Success"] = $"Đã nhập thành công {list.Count} sản phẩm!";
                }
                else
                {
                    TempData["Error"] = "Không tìm thấy dữ liệu hợp lệ trong file!";
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Có lỗi xảy ra khi xử lý file: " + ex.Message;
            }

            return RedirectToAction("Index");
        }

        [Authorize(Roles = "Admin")]
        public IActionResult DownloadTemplate()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Template");

                // Header
                worksheet.Cells[1, 1].Value = "ProductName";
                worksheet.Cells[1, 2].Value = "CategoryID";
                worksheet.Cells[1, 3].Value = "Description";
                worksheet.Cells[1, 4].Value = "Quantity";
                worksheet.Cells[1, 5].Value = "Price";

                // Định dạng Header đen trắng
                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }

                worksheet.Cells.AutoFitColumns();
                var fileContents = package.GetAsByteArray();
                return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Template_Import.xlsx");
            }
        }
        [Authorize(Roles = "Admin")]
        public IActionResult ExportProductsToExcel()
        {
            var products = _productRepository.GetAllProducts1();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Products List");

                // 1. Tạo Header
                string[] headers = { "ID", "Tên sản phẩm", "Danh mục", "Mô tả", "Tổng số lượng", "Giá (Variant đầu)", "SKU" };
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                // Định dạng Header
                using (var range = worksheet.Cells[1, 1, 1, headers.Length])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkBlue);
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

                // 2. Đổ dữ liệu
                int currentRow = 2;
                foreach (var p in products)
                {
                    var firstVariant = p.ProductVariants?.FirstOrDefault();

                    worksheet.Cells[currentRow, 1].Value = p.ProductId;
                    worksheet.Cells[currentRow, 2].Value = p.ProductName;
                    worksheet.Cells[currentRow, 3].Value = p.Category?.CategoryName ?? "N/A";
                    worksheet.Cells[currentRow, 4].Value = p.Description;
                    worksheet.Cells[currentRow, 5].Value = p.Quantity;
                    worksheet.Cells[currentRow, 6].Value = firstVariant?.Price ?? 0;
                    worksheet.Cells[currentRow, 7].Value = firstVariant?.SKU ?? "";

                    // Định dạng số cho cột Giá
                    worksheet.Cells[currentRow, 6].Style.Numberformat.Format = "#,##0";

                    currentRow++;
                }

                worksheet.Cells.AutoFitColumns();

                var fileContents = package.GetAsByteArray();
                string fileName = $"Danh_Sach_San_Pham_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

                return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }
        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> UpdateProductImage(int productId, IFormFile imageFile)
        {
            if (imageFile == null || imageFile.Length == 0) return BadRequest();

         
            var fileName = Guid.NewGuid().ToString() + Path.GetExtension(imageFile.FileName);
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/images", fileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await imageFile.CopyToAsync(stream);
            }

           
            var newImage = new ProductImage
            {
                ProductId = productId,
                ImageUrl = "/images/" + fileName,
                IsMain = false 
            };

          
            await _productImageRepository.AddImageAsync(newImage);

            return Json(new { success = true, newImageUrl = newImage.ImageUrl });
        }
    }
}
