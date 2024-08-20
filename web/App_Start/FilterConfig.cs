using System.Web;
using System.Web.Mvc;
using MohwEmail.Filters;

namespace MohwEmail
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
            filters.Add(new AuthAttribute());
            filters.Add(new FunctionLogAttribute());
        }
    }
}
