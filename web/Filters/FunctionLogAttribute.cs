using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MohwEmail.Models;
using MohwEmail.ViewModels;
//using MohwEmail.Models;

namespace MohwEmail.Filters
{
    public class FunctionLogAttribute : FilterAttribute, IActionFilter, IResultFilter
    {
        public void OnActionExecuted(ActionExecutedContext filterContext)
        {
        }

        /// <summary>
        /// 紀錄使用功能
        /// </summary>
        /// <param name="filterContext"></param>
        public void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext.HttpContext.Session["User"] != null)
            {
                string ControllerName = filterContext.RouteData.Values["Controller"].ToString();
                string ActionName = filterContext.RouteData.Values["Action"].ToString();

                User user = filterContext.HttpContext.Session["User"] as User;

                //功能Log              
                using (MOHWEntities db = new MOHWEntities())
                {
                    FunctionInfo MainFunction = db.FunctionInfo.SingleOrDefault(x => x.FunctionController.Equals(ControllerName) && x.FunctionAction.Equals(ActionName) );

                    if (MainFunction != null)
                    {
                        db.LogFunction.Add(new LogFunction() { FunctionId= MainFunction.FunctionId, LogMessage = ControllerName + "-" + ActionName, LogTime = DateTime.Now, UserId = user.UserDetail.UserId, });
                        db.SaveChanges();
                    }
                }               
            }
        }

        public void OnResultExecuted(ResultExecutedContext filterContext)
        {
        }

        public void OnResultExecuting(ResultExecutingContext filterContext)
        {
        }
    }
}