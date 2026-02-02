using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ASP.BaseCommon;
using ASP.Models.Admin.Logs;
using ASP.Models.ASPModel;

namespace ASP.Controllers.Admin
{
    public class CkGalleryController : Controller
    {
        // GET: CkGalleryController
        protected readonly ASPDbContext _context;
        protected LogRepositoryInterface log;
        private readonly IWebHostEnvironment env;
        protected string photosPath;
        protected string photosPathSummerNoted;
        private object userManager;
        protected readonly IHttpContextAccessor httpContextAccessor;
        public CkGalleryController(ASPDbContext context, IWebHostEnvironment env, LogRepositoryInterface log, IHttpContextAccessor httpContextAccessor)
        {
            _context = context;
            this.env = env;
            this.log = log;
            this.photosPath = this.env.WebRootPath + "/assets/ckgallery";
            this.photosPathSummerNoted = this.env.WebRootPath + "/assets/summernote";
            this.httpContextAccessor = httpContextAccessor;
        }
        public ActionResult Index()
        {
            return View("../Admin/CkGallerys/Index");
        }

        [HttpPost]
        [Route("UploadImage", Name = "admin.ckgallery.uploadimage")]
        public IActionResult UploadImage(IFormFile upload, string CKEditorFuncNum, string CKEditor, string langCode)
        {
            if (upload.Length <= 0) return null;
            // full path to file in temp location
            if (!Directory.Exists(this.photosPath))
            {
                Directory.CreateDirectory(this.photosPath);
            }
            //
            string image_name = null;
            if (upload != null)
            {
                // add new file
                /* var a = Path.GetFileNameWithoutExtension(upload.FileName);
                 var b = Path.GetFileName(upload.FileName);
                 var c = Path.GetExtension(upload.FileName);*/
                image_name = DateTime.Now.Ticks + "_" + FriendlyUrlHelper.GetFriendlyTitle(Path.GetFileNameWithoutExtension(upload.FileName), true, 200) + Path.GetExtension(upload.FileName);
                var filePath = Path.Combine(this.photosPath, image_name);
                var extension = Path.GetExtension(filePath);
                //
                using (var stream = System.IO.File.Create(filePath))
                {
                    upload.CopyTo(stream);
                }
            }
            //
            var url = $"{Url.Content("~/")}{"assets/ckgallery/"}{image_name}";
            return Json(new { uploaded = true, url });
        }
        [HttpGet]
        [Route("RemoveImage", Name = "admin.ckgallery.removeimage")]
        public IActionResult RemoveImage(string fileUrl)
        {
            // full path to file in temp location
            string message = "image ???";
            if (System.IO.File.Exists(Path.Combine(this.photosPath, fileUrl)))
            {
                System.IO.File.Delete(Path.Combine(this.photosPath, fileUrl));
                message = "Xóa file thành công!";
            }
            else
            {
                message = "Không tồn tại file trên hệ thống!";
            }
            return Json(new { result = true, message = message });
        }
        [Route("FileBrowse", Name = "admin.ckgallery.filebrowse")]
        public IActionResult FileBrowse(string CKEditorFuncNum = null, string CKEditor = null, string langCode = null, string fname = "")
        {
            if (fname == null) fname = "";
            var dir = new DirectoryInfo(Path.Combine(this.photosPath));
            ViewBag.fileInfos = dir.GetFiles().Where(w => w.Name.Contains(fname)).OrderByDescending(o => o.Name).ToList();
            ViewBag.fname = fname;
            return View("../Admin/CkGallerys/FileBrowse");
        }
        // summernote editor
        [HttpPost]
        [Route("UploadImageSummerNote", Name = "admin.ckgallery.uploadimagesummernote")]
        public IActionResult UploadImageSummerNote(IFormFile file)
        {
            if (file.Length <= 0) return Json(new { uploaded = false, url = "" });
            // full path to file in temp location
            if (!Directory.Exists(this.photosPathSummerNoted))
            {
                Directory.CreateDirectory(this.photosPathSummerNoted);
            }
            //
            string image_name = null;
            if (file != null)
            {
                // add new file
                image_name = DateTime.Now.Ticks + "_" + FriendlyUrlHelper.GetFriendlyTitle(Path.GetFileNameWithoutExtension(file.FileName), true, 200) + Path.GetExtension(file.FileName);
                var filePath = Path.Combine(this.photosPathSummerNoted, image_name);
                var extension = Path.GetExtension(filePath);
                //
                using (var stream = System.IO.File.Create(filePath))
                {
                    file.CopyTo(stream);
                }
            }
            //
            var url = $"{Url.Content("~/")}{"assets/summernote/"}{image_name}";
            return Json(new { uploaded = true, url });
        }

    }
}
