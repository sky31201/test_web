using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using MohwEmail.Helpers;
using MohwEmail.Models;
using MohwEmail.Services;
using NLog;

namespace MohwEmail.Controllers
{
    public class _BaseController : JsonNetController
    {
        public _BaseController()
        {
            NLogService nLogServices = new NLogService();
        }

        public Logger logger = LogManager.GetCurrentClassLogger();



        /// <summary>
        /// 登入使用者
        /// </summary>
        protected User user
        {
            get
            {
                User result = null;
                if (Session["User"] != null)
                {
                    result = Session["User"] as User;
                }
                return result;
            }
            set
            {
                Session["User"] = value;
            }
        }

        public JsonResult GetDataDictinary(int DataType)
        {
            List<Models.SysSubCode> CaseSourceList = new CommonService().GetSourceList(DataType);
            return Json(CaseSourceList.Select(x => new { x.SerialNo, x.ZHName, x.MainCode }));
        }


        /// <summary>
        /// 驗證結果
        /// </summary>
        /// <param name="id"></param>
        /// <param name="verifyCode"></param>
        /// <returns></returns>
        public async Task<ActionResult> EmailVerification(string verifyCode)
        {
            Service.Services.EmailServices emailServices = new Service.Services.EmailServices();

            ViewBag.Message = "確認失敗";

            if (verifyCode != null)
            {

                bool isVerifyCodeValid = emailServices.ValidateVerifyCode(verifyCode);

                if (isVerifyCodeValid)
                {
                    // 建立 案號
                    string appealNo = emailServices.CreateAppealNo();                    
                    // 更新 Appeal表
                    int res1 = await emailServices.UpdateAppealWithVerifyCodeAsync(appealNo, verifyCode);
                    // 更新 AppendFile 表
                    int res2 = await emailServices.UpdateAppendFileWithVerifyCodeAsync(appealNo, verifyCode);

                    // 更新 資料夾名稱
                    // emailServices.ChangeFileName(appealNo, verifyCode);

                    // 寄出案件受理通知信
                    emailServices.SendAcceptNotification(appealNo);

                    ViewBag.Message = "確認成功";
                }
            }

            return View();
        }
    }
}