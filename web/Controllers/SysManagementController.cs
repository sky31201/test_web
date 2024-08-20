using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MohwEmail.Filters;
using MohwEmail.Models;
using MohwEmail.Services;
using MohwEmail.ViewModels.SysManagement;


namespace MohwEmail.Controllers
{
    [ErrorAttr]
    public class SysManagementController : Controller
    {
        readonly SysManagementService _sysManagementSer;
        readonly CaseManagementService _caseManagementSer;

        public SysManagementController()
        {
            _caseManagementSer = new CaseManagementService();
            _sysManagementSer = new SysManagementService();
        }


        #region  帳號管理
        /// <summary>
        /// 帳號管理
        /// </summary>
        /// <returns>View</returns>
        public ActionResult AccountManagement()
        {
            var userInfo = Session["User"] as User;
            AccountManagementViewModel viewModel = new AccountManagementViewModel();
            viewModel.ConditionModel = new AccountConditionModel();
            viewModel.Units = _sysManagementSer.GetAssignUnit(userInfo);
            viewModel.Units.Add(new SelectListItem() { Value = "", Text = "請選擇", Selected = true });
            viewModel.Roles = _sysManagementSer.GetRoleType();
            viewModel.Roles.Add(new SelectListItem() { Value = "", Text = "請選擇", Selected = true });
            viewModel.Ugroups = new List<SelectListItem>();
            viewModel.Divisions = new List<SelectListItem>();
            viewModel.Internal = userInfo.UserDetail.Role.Contains("155")? "N" : userInfo.UserDetail.Internal;

            return View(viewModel);
        }

        /// <summary> 
        ///  查詢結果
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GetAccountList(AccountManagementViewModel viewModel)
        {
            var userInfo = Session["User"] as MohwEmail.Models.User;

            if ((userInfo.UserDetail.Role.Contains("161") || userInfo.UserDetail.Role.Contains("156")) && !userInfo.UserDetail.Role.Contains("155"))
            {
                if (viewModel.ConditionModel.Unit == null)
                {
                    viewModel.ConditionModel.Unit = userInfo.UserDetail.TopUnit.ToString();

                    if (userInfo.UserDetail.Level3Unit != null)
                    {
                        viewModel.ConditionModel.Internal = "3";
                    }
                    else if (userInfo.UserDetail.Level2Unit != null)
                    {
                        viewModel.ConditionModel.Internal = "2";
                    }
                    else
                    {
                        viewModel.ConditionModel.Internal = "1";
                    }
                }else
                {
                    if (viewModel.ConditionModel.Ugroup != null)
                    {
                        viewModel.ConditionModel.Internal = "2";
                    }
                    else
                    {
                        viewModel.ConditionModel.Internal = "1";
                    }
                }
            }
            else if (userInfo.UserDetail.Role.Contains("162") && !userInfo.UserDetail.Role.Contains("155") && !userInfo.UserDetail.Role.Contains("161"))
            {
                //true 代表有兩層
                bool checkLevel2 = userInfo.UserDetail.Level2Unit != null ? true : false;

                viewModel.ConditionModel.Unit = userInfo.UserDetail.TopUnit.ToString();
                viewModel.ConditionModel.Ugroup = checkLevel2? userInfo.UserDetail.Level2Unit.ToString() : null;
                viewModel.ConditionModel.Internal = "2";


            }
            else if(userInfo.UserDetail.Role.Contains("155"))
            {
                if (viewModel.ConditionModel.Unit == null)
                {
                    viewModel.ConditionModel.Unit = "1";
                    viewModel.ConditionModel.Internal = "3";
                }
                else
                {
                    if (viewModel.ConditionModel.Ugroup != null)
                    {
                        viewModel.ConditionModel.Internal = "2";
                    }
                    else
                    {
                        viewModel.ConditionModel.Internal = "1";
                    }
                }
            }

            var result = _sysManagementSer.GetAccountList(viewModel);

            return PartialView("_accountList_Template", result);
        }

        /// <summary> 
        ///  新增或修改
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SaveAccountInfo(AccountManagementViewModel model)
        {
            
            var userInfo = Session["User"] as MohwEmail.Models.User;
            var result = _sysManagementSer.SaveAccountData(model, userInfo);


            return Json(new { Status = result.Item1, Message =result.Item2 });
        }

        [HttpPost]
        public JsonResult ResetAccount(string UserId)
        {
            var result = _sysManagementSer.ResetPassword(UserId);
            return Json(new { Status = result.Item1, Message = result.Item2 });

        }


        /// <summary>
        /// 取得使用者資料
        /// </summary>
        /// <param name="UserId"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetAccountInfo(string UserId)
        {
            var result = _sysManagementSer.GetAccountDetail(UserId);
            return Json(result);
        }

        /// <summary>
        /// 下拉選單-組名
        /// </summary>
        /// <param name="unit"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetUgroup(string unit)
        {
            var result = new List<SelectListItem>();
            var userInfo = Session["User"] as User;
            result = _sysManagementSer.GetUgroupType(int.Parse(unit), userInfo);

            return Json(result);
        }

        #endregion

        /// <summary>
        /// 警訊排程通報紀錄
        /// </summary>
        /// <returns>View</returns>
        public ActionResult AlarmNotificationRecord()
        {
            return View();
        }
        /// <summary>
        /// 警訊通報設定
        /// </summary>
        /// <returns>View</returns>
        public ActionResult AlertNotificationSettings()
        {
            return View();
        }
        /// <summary>
        /// 角色管理
        /// </summary>
        /// <returns>View</returns>
        public ActionResult RoleManagement()
        {
            return View();
        }

        #region 工作行事曆
        /// <summary>
        /// 工作行事曆
        /// </summary>
        /// <returns>View</returns>
        public ActionResult WorkCalender()
        {
            WorkCalenderViewModel viewModel = new WorkCalenderViewModel();
            viewModel.ConditionModel = new WCConditionModel();
            return View(viewModel);
        }

        /// <summary> 
        ///  查詢結果
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult GetWorkCalenderList(WorkCalenderViewModel viewModel)
        {
            var userInfo = Session["User"] as MohwEmail.Models.User;
            var result = _sysManagementSer.GetWorkCalenderList(viewModel);
            return PartialView("_workCalenderList_Template", result);
        }

        /// <summary> 
        ///  新增
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SaveWorkCalenderInfo(WorkCalenderViewModel model)
        {

            var userInfo = Session["User"] as MohwEmail.Models.User;
            var result = _sysManagementSer.SaveWorkCalenderData(model, userInfo);
            return Json(new { Status = result.Item1, Message = result.Item2 });
        }

        /// <summary> 
        ///  刪除
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult DelWorkCalenderData(string date)
        {
            var result = _sysManagementSer.DelWorkCalenderData(date);
            return Json(new { Status = result.Item1, Message = result.Item2 });
        }

        #endregion

        #region 修改密碼 
        [HttpPost]
        public JsonResult ChangePass(string OriginalPass, string NewPass)
        {
            var userInfo = Session["User"] as MohwEmail.Models.User;
            var result = _sysManagementSer.ChangePass(OriginalPass, NewPass, userInfo);
            return Json(new { Status = result.Item1, Message = result.Item2 });

        }
        #endregion

    }
}