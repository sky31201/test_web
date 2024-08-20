using MohwEmail.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace MohwEmail.Filters
{
    /// <summary>
    /// 使用者權限檢查，參數為限制權限 (Ex: 155 總窗口, 156 單位窗口..)；
    /// 若使用者無該功能使用權限，返回搜尋畫面。
    /// </summary>
    public class RoleAttr : ActionFilterAttribute
    {
        readonly List<string> _role;
        public RoleAttr(string role)
        {
            _role = role.Split(',').ToList<string>();
        }

        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            // 取得要求權限

            // 取得Session["User"]
            User user = (User)filterContext.HttpContext.Session["User"];
            // 取得Role
            List<string> userRole = user.UserDetail.Role.Split(',').ToList();

            // 比較限制權限 & 使用者權限
            if (userRole.Where(x=> _role.Contains(x.Trim())).Count() == 0)
            {
                filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary ( new { Action = "CaseQuery", Controller= "CaseManagement" } ));
            }
        }

        public override void OnActionExecuted(ActionExecutedContext filterContext)
        {

        }

        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {

        }

        public override void OnResultExecuted(ResultExecutedContext filterContext)
        {

        }
    }
}