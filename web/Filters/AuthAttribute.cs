using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Filters;

namespace MohwEmail.Filters
{
    public class AuthAttribute : IAuthenticationFilter
    {
        /// <summary>
        /// 是否登入認證
        /// </summary>
        /// <param name="filterContext"></param>
        public void OnAuthentication(AuthenticationContext filterContext)
        {
            if (filterContext.RouteData.GetRequiredString("controller").Equals("Security") && (filterContext.RouteData.GetRequiredString("action").Equals("Login") || filterContext.RouteData.GetRequiredString("action").Equals("ForgetLogin")))
            {
                return;
            }
            else if (filterContext.RouteData.GetRequiredString("controller").ToLower().Equals("financial") && filterContext.RouteData.GetRequiredString("action").ToLower().Equals("financial"))
            {
                return;
            }
            else if (filterContext.RouteData.GetRequiredString("controller").ToLower().Equals("_base"))
            {
                return;
            }
            else if (filterContext.RouteData.GetRequiredString("controller").ToLower().Equals("api"))
            {
                return;
            }
            else if (filterContext.HttpContext.Session["User"] == null)
            {
                Controller controller = filterContext.Controller as Controller;
                controller.Response.Write("<script>location.href = '/Security/Login';</script>");
                controller.Response.End();
                filterContext.Result = new EmptyResult();
            }
            
        }
        public void OnAuthenticationChallenge(AuthenticationChallengeContext filterContext)
        {
            //throw new NotImplementedException();
        }
    }
}