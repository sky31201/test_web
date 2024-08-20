using MohwEmail.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MohwEmail.Controllers
{
    [AllowAnonymous]
    public class ApiController : Controller
    {
        [System.Web.Mvc.HttpPost]
        /// <summary>
        /// 根據EmpID取得待處理總案件數
        /// </summary>
        /// <returns></returns>
        public string GetCaseTotal(GetCaseTotalModel request)
        {
            // 員工編號
            var empNo = request.EmpNo;
            
            int num = 0;
            using (MOHWEntities db = new MOHWEntities())
            {
                try
                {
                    // 員工資料
                    var undertaker_detail = db.vw_UserDetail.Where(x => x.employeeID == empNo);
                    // 避免查無對應資料產生NULL，檢查筆數是否 > 0
                    if (undertaker_detail.Count() > 0)
                    {
                        // 員工名稱
                        var undertaker = undertaker_detail.FirstOrDefault().UserId;
                        num = db.Appeal.Where(x => x.Undertaker == undertaker && x.Status == 26).Count();
                    }
                }
                catch (Exception ex)
                {
                    return JsonConvert.SerializeObject(ex);
                }
            }

            AppealResults result = new AppealResults()
            {
                Count = num
            };
            
            return JsonConvert.SerializeObject(result);
        }

        public class GetCaseTotalModel
        {
            public string EmpNo { get; set; }
        }

        public class AppealResults
        {
            public int Count { get; set; }
        }
    }
}