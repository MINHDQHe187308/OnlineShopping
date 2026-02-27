using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using System.Reflection;

namespace ASP.BaseCommon
{
    public class DeleteFileAttribute : ActionFilterAttribute
    {
        public string rootsPath { get; set; }
        public string fileName { get; set; }
        public DeleteFileAttribute()
        {
            rootsPath = null;
            fileName = null;
        }
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext != null)
            {
                rootsPath = filterContext.ActionArguments["rootsPath"].ToString();
                fileName = filterContext.ActionArguments["file"].ToString();
            }
            base.OnActionExecuting(filterContext);
        }
        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {
            if (!string.IsNullOrEmpty(rootsPath) && !string.IsNullOrEmpty(fileName))
            {
                int count = 0;
                var dir = new DirectoryInfo(rootsPath);
                foreach (var file in dir.EnumerateFiles("Excel_*.xlsx"))
                {
                    count++;
                    if (count < dir.EnumerateFiles("Excel_*.xlsx").Count())
                        file.Delete();
                }
            }
            //if (!string.IsNullOrEmpty(rootsPath) && !string.IsNullOrEmpty(fileName))
            //{
            //    var url = Path.Combine(this.rootsPath, this.fileName);
            //    if (File.Exists(url))
            //    {
            //        File.Delete(url);
            //    }
            //}
        }
        //
    }
}
