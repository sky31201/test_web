using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace MohwEmail
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            // read webconfig to check if TaskActivation value is active.
            if (bool.TryParse(ConfigurationManager.AppSettings["TaskActivation"], out bool isActive)) 
            {
                // if the value is set to true, run the task.
                if (isActive)
                {
                    // 執行背景程式 - 信件檢查
                    var task = new Service.Services.EmailServices();
                    task.Run();
                }
            };
            

            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }
    }
}
