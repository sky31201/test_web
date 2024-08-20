using MohwEmail.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MohwEmail.Filters
{
    public class LogFilter : ActionFilterAttribute
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            var controller = filterContext.Controller.ControllerContext;
            var userId = ((User)filterContext.HttpContext.Session["User"]).UserDetail.UserId;
            var ip = controller.HttpContext.Request.UserHostAddress;
            var actionName = controller.RouteData.Values["action"];
            var appealNo = controller.RouteData.Values["id"];            

            logger
                //.WithProperty("Property1", userName)
                //.WithProperty("Property2", ip)
                .Info($"{userId}=>{ actionName }[{appealNo}]");

            base.OnActionExecuting(filterContext);
        }
    }
}