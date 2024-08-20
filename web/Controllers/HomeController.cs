using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;

namespace MohwEmail.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult Quiz()
        {
            return View();
        }

        #region 使用者手冊下載

        public ActionResult DownloadUserManual()
        {
            string fileName = "衛生福利部部長信箱_使用者手冊V2.pdf";              
            string path = HostingEnvironment.MapPath($"/App_Data/{fileName}");
            
            try
            {
                FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                return File(stream, "application/pdf", fileName);
            }
            catch (System.Exception)
            {
                return Content("<script>alert('查無此檔案');</script>");
            }
        }

        #endregion
    }
}