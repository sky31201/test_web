using MohwEmail.Models;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MohwEmail.Filters
{
    public class ErrorAttr: HandleErrorAttribute
    {
        private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        
        public override void OnException(ExceptionContext filterContext)
        {
            base.OnException(filterContext);

            User user = filterContext.HttpContext.Session["User"] as User;
            string userId = user.UserDetail.UserId;
            string userName = user.UserDetail.UserName;

            string controller = (string)filterContext.RouteData.Values["Controller"];
            string action = (string)filterContext.RouteData.Values["Action"];

            string errMessage = $"{userName}({userId})=> {controller}.{action} ";

            Exception ex = filterContext.Exception;
            logger.Warn(errMessage + ex.Message);
            logger.Error(ex.ToString());
        }
    }
}