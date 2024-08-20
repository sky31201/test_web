using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MohwEmail.Models;
using MohwEmail.Services;
using MohwEmail.ViewModels.CaseReport;
using MohwEmail.ViewModels.CaseManagement;
using MohwEmail.Helpers;
using Dapper;
using System.Data.SqlClient;
using System.Data;
using Newtonsoft.Json;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.SqlServer;
using System.Web.Http;
using System.Runtime.Caching;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Zip;
using System.IO;
using System.IO.Compression;
using System.Text;
using MohwEmail.Filters;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using MohwEmail.ViewModels.CaseMangement;
using MohwEmail.Service.Services;

namespace MohwEmail.Controllers
{
    [LogFilter]
    public class CaseManagementController : _BaseController
    {
        static MemoryCache cache = MemoryCache.Default;
        readonly FileService _fileService;
        readonly CaseManagementService _caseManagementService;

        User user;
        private readonly CommonService _commonService;
        private User _currentUser;
        private XmlService _xmlService;

        public CaseManagementController()
        {            
            _fileService = new FileService();
            _caseManagementService = new CaseManagementService();
            _commonService = new CommonService();
            _xmlService = new XmlService();
        }


        #region 案件查詢
        /// <summary>
        /// 案件查詢
        /// </summary>
        /// <returns>View</returns>        
        public ActionResult CaseQuery()
        {
            //頁面初始化
            _currentUser = Session["User"] as User;
            bool isSupervisor = false;
            foreach(var item in _currentUser.rolesList)
            {
                if(item.SerialNo == 155)
                {
                    isSupervisor = true;
                }
            }
            string memberType = _currentUser.UserDetail.Internal == "Y" ? "1" : "2"; //部內 = 1 , 所屬 = 2;
            List<SysSubCode> CaseSourceList = new CommonService().GetCaseSourceList();
            List<Organization> caseOrganizerList = new CommonService().GetCaseOrganizerList(memberType, isSupervisor);
            CaseQueryViewModel caseQueryViewModel = new CaseQueryViewModel();

            //陳情類別
            foreach (SysSubCode CaseSource in CaseSourceList.Where(x => x.MainCode == 2))
            {
                caseQueryViewModel.CasePetitionTypeList.Add(new SelectListItem() { Text = CaseSource.ZHName, Value = CaseSource.SerialNo.ToString() });
            }
            //來源管道
            foreach (SysSubCode CaseSource in CaseSourceList.Where(x => x.MainCode == 3))
            {
                caseQueryViewModel.CaseSourceList.Add(new SelectListItem() { Text = CaseSource.ZHName, Value = CaseSource.SerialNo.ToString() });
            }
            //案件狀態
            foreach (SysSubCode CaseSource in CaseSourceList.Where(x => x.MainCode == 4 && x.SerialNo < 30))
            {
                bool supervisor = false; //是否為總窗口 
                foreach(var role in _currentUser.rolesList)
                {
                    if(role.SerialNo == 155)
                    {
                        supervisor = true;
                    }
                }
                if (supervisor)
                {
                    caseQueryViewModel.CaseStatusList.Add(new SelectListItem() { Text = CaseSource.ZHName, Value = CaseSource.SerialNo.ToString() });
                }
                else
                {
                    if(CaseSource.SerialNo != 25)
                    {
                        caseQueryViewModel.CaseStatusList.Add(new SelectListItem() { Text = CaseSource.ZHName, Value = CaseSource.SerialNo.ToString() });
                    }
                }
            }
            // 議題類別
            caseQueryViewModel.StatisticCaseTypeList = _caseManagementService.GetStatCaseList();

            foreach (var caseOrganizer in caseOrganizerList.Select(x => new { x.SerialNo, x.ZHName }).OrderBy(x => x.SerialNo))
            {
                caseQueryViewModel.CaseOrganizerList.Add(new SelectListItem() { Text = caseOrganizer.ZHName, Value = caseOrganizer.SerialNo.ToString() });
            }
            ViewBag.User = _currentUser;
            //新增預設密碼提醒視窗
            SysManagementService sysManagementService = new SysManagementService();
            if (sysManagementService.Decrypt(_currentUser.UserDetail.Password, _currentUser.UserDetail.CreateDate.Value) == "emailmg@1234" && _currentUser.UserDetail.Internal == "N")
            {
                TempData["msg"] = "使用預設密碼登入系統，請修改密碼";
            }
            return View(caseQueryViewModel);
        }
        public JsonResult CaseQueryforList(CaseQueryViewModel caseQueryViewModel)
        {
            caseQueryViewModel.InitialNull();
            Dictionary<string, object> Result = new Dictionary<string, object>() { { "message", "" }, { "data", null } };

            List<Dictionary<string, object>> ResultList = new List<Dictionary<string, object>>();
            //Dictionary<string, object> Result = new Dictionary<string, object>();
            string SourceList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.SourceList);
            string PetitionTypeList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.PetitionTypeList);
            string CaseTypeList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.CaseTypeList);

            string sqlCommand = @"EXEC	[sp_QueryAppeal]
                                    @AppealNoS,
                              @AppealNoE,
                              @ODNoS,
                              @ODNoE,
                              @VerifyDateS,
                              @VerifyDateE,
                              @DeadLineS,
                              @DeadLineE,
                              @Name,
                              @Phone,
                              @Subject,
                              @EMail,
                              @Contents,
                              @OfficialDocumentSystem,
                              @OfficialDocumentStatus,
                              @SourceList,
                              @PetitionTypeList,
                              @StatusList,
                              @CaseTypeList";
            Dictionary<string, object> Parameters = new Dictionary<string, object>()
            {
                { "@AppealNoS", caseQueryViewModel.AppealNoS },
                { "@AppealNoE", caseQueryViewModel.AppealNoE },
                { "@ODNoS", caseQueryViewModel.ODNoS },
                { "@ODNoE", caseQueryViewModel.ODNoE },
                { "@VerifyDateS", caseQueryViewModel.VerifyDateS },
                { "@VerifyDateE", caseQueryViewModel.VerifyDateE },
                { "@DeadLineS", caseQueryViewModel.DeadLineS },
                { "@DeadLineE", caseQueryViewModel.DeadLineE },
                { "@Name", caseQueryViewModel.Name },
                { "@Phone", caseQueryViewModel.Phone },
                { "@Subject", caseQueryViewModel.Subject },
                { "@EMail", caseQueryViewModel.EMail },
                { "@Contents", caseQueryViewModel.Contents },
                { "@OfficialDocumentSystem", caseQueryViewModel.OfficialDocumentSystem },
                { "@OfficialDocumentStatus", caseQueryViewModel.OfficialDocumentStatus },
                { "@SourceList", SourceList },
                { "@PetitionTypeList", PetitionTypeList },
                { "@StatusList", caseQueryViewModel.StatusList },
                { "@CaseTypeList", CaseTypeList },
            };

            using (MOHWEntities db = new MOHWEntities())
            {
                ResultList = db.SqlQuery(sqlCommand, Parameters);
            }

            //if (ResultList.Count > 0)
            //{
            //    Result = ResultList[0];
            //}
            Result["data"] = ResultList;

            if (ResultList.Count.Equals(0))
            {
                Result["message"] = "查無資料";
            }
            //return View(caseQueryViewModel);            
            return Json(Result);
        }

        [System.Web.Mvc.HttpPost]
        /// <summary>
        /// 案件查詢, 返回Json
        /// </summary>
        /// <param name="jsonModel"></param>
        /// <returns></returns>
        public ActionResult CaseListSearch([FromBody] CaseSearchVM jsonModel)
        {
            _currentUser = Session["User"] as User;
            string response = _caseManagementService.CaseListSearch(jsonModel, _currentUser);            
            return Content(response, "application/json", Encoding.UTF8);
        }

        /// <summary>
        /// 案件結案
        /// </summary>
        /// <param name="jsonModel"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult CaseClosed([FromBody] CaseClosedVM jsonModel)
        {
            _currentUser = Session["User"] as User;
            bool result = _caseManagementService.CaseClosed(jsonModel, _currentUser);
            CaseResponse response = new CaseResponse()
            {
                Success = result,
            };
            return Content(JsonConvert.SerializeObject(response), "application/json", Encoding.UTF8);
        }

        /// <summary>
        /// 根據已選取承辦單位, 取得對應人員
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult GetUnderTakerList([FromBody] List<string> model)
        {
            _currentUser = Session["User"] as User;
            int memberType = _currentUser.UserDetail.Internal == "Y" ? 1 : 2; //部內 = 1 , 所屬 = 2;
            object sqlResult;

            //List<PermissionPeople> res = new List<PermissionPeople>();
            //foreach(var item in model)
            //{
            //    res.AddRange(_commonService.GetUnderTakerList(Convert.ToInt32(item), memberType));
            //}
            //sqlResult = res.ToList();

            List<vw_UserDetail> res = new List<vw_UserDetail>();
            foreach(string item in model)
            {
                res.AddRange(_caseManagementService.LookupUnitMembers(int.Parse(item)));
            }
            sqlResult = res.ToList();
            string response = JsonConvert.SerializeObject(sqlResult);
            return Content(response, "application/json", Encoding.UTF8);
        }

        #region Entity Framework
        [System.Web.Http.HttpPost]
        public JsonResult getDataTable()
        {
            JsonResult result;

            result = Json(GetList(""));

            return result;
        }
        public List<Models.Append_View> GetList(string Creator = "")
        {

            List<Models.Append_View> result = new List<Models.Append_View>();
            using (MOHWEntities db = new MOHWEntities())
            {
                //result = (from a in db.Append_View orderby a.SerialNo descending select a).ToList();
                result = db.Append_View.OrderByDescending(x => x.SerialNo).ToList();
            }

            return result;
        }
        #endregion

        #region Dapper
        /// <summary>
        /// Return Json
        /// </summary>
        /// <returns></returns>
        //[HttpPost]
        //public object getDataTable()
        //{
        //    var result = JsonConvert.SerializeObject(GetRsultList());
        //    return result;
        //}

        public object GetRsultList()
        {
            object result;
            var sql = @"SELECT 
                        a.SerialNo,
                        a.AppealNo,
                        b.ODNo,
                        a.VerifyDate,
                        b.DeadLine,
                        a.Name,
                        a.Subject,
                        c.ZHName,
                        (SELECT COUNT(AppealNo) FROM AppendFile WHERE AppealNo = a.AppealNo) AS AppealFileNo,
                        (SELECT IIF(DATEDIFF(day, b.DeadLine, SYSDATETIME()) < 0, 0, DATEDIFF(day, b.DeadLine, SYSDATETIME()))) AS DeadLineDay
                        FROM
                        Appeal AS a
                        LEFT JOIN
                        AppealDetail AS b 
                        ON a.AppealNo = b.AppealNo
                        LEFT JOIN
                        SysSubCode AS c 
                        ON b.Status = c.SerialNo
                        ORDER BY 
                        a.SerialNo DESC
                        ";
            using (var connection = new SqlConnection(StringExtension.GetConnectionString()))
            {
                connection.Open();
                result = connection.Query(sql);
            }

            return result;
        }
        #endregion


        /// <summary>
        /// 案件匯出Excel
        /// </summary>
        /// <returns>File</returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult ExcelList([FromBody] CaseSearchVM jsonModel)
        {
            _currentUser = Session["User"] as User;
            string jsonString = _caseManagementService.CaseListSearch(jsonModel, _currentUser);
            List<dynamic> result = JsonConvert.DeserializeObject<List<dynamic>>(jsonString);

            var tc = new System.Globalization.TaiwanCalendar();
            var ttime = $"{tc.GetYear(DateTime.Now)}/{DateTime.Now.Month}/{DateTime.Now.Day} {DateTime.Now.Hour}:{DateTime.Now.Minute}";

            // 建立Excel
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hSSFWorkbook.CreateSheet("sheet"); //建立sheet

            // 設定樣式
            ICellStyle headerStyle = hSSFWorkbook.CreateCellStyle();
            IFont headerFont = hSSFWorkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerFont.FontName = "Microsoft JhengHei";
            headerFont.FontHeightInPoints = 20;
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerFont);

            // 標題列
            sheet.CreateRow(0);
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 11));
            sheet.GetRow(0).CreateCell(0).SetCellValue("陳情信件列表");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle;
            sheet.CreateRow(2);
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 2));
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 9, 11));
            sheet.GetRow(2).CreateCell(0).SetCellValue($"匯出筆數: {result.Count()}");
            sheet.GetRow(2).CreateCell(9).SetCellValue($"列印時間: {ttime}");
            // 表頭Header
            sheet.CreateRow(3).CreateCell(0).SetCellValue("序號");
            sheet.GetRow(3).CreateCell(1).SetCellValue("案件編號");
            sheet.GetRow(3).CreateCell(2).SetCellValue("陳情時間");
            sheet.GetRow(3).CreateCell(3).SetCellValue("限辦時間");
            sheet.GetRow(3).CreateCell(4).SetCellValue("民眾姓名");
            sheet.GetRow(3).CreateCell(5).SetCellValue("陳情主旨");
            sheet.GetRow(3).CreateCell(6).SetCellValue("陳情內容");
            sheet.GetRow(3).CreateCell(7).SetCellValue("業務單位");
            sheet.GetRow(3).CreateCell(8).SetCellValue("承辦人");
            sheet.GetRow(3).CreateCell(9).SetCellValue("是否逾期");
            sheet.GetRow(3).CreateCell(10).SetCellValue("已結案");

            //填入資料
            int rowIndex = 4;
            foreach(var item in result)
            {
                string deadLine = "";

                if (item.DeadLine != null)
                {
                    if (item.StatusMain == 26)
                    {
                        deadLine = DateTime.Compare(DateTime.Parse((string)item.DeadLine), DateTime.Now) == -1 ? "是" : "否";
                    }
                    if (item.StatusMain == 27)
                    {
                        if(item.CloseDate != null)
                        {
                            deadLine = DateTime.Compare(DateTime.Parse((string)item.DeadLine), DateTime.Parse((string)item.CloseDate)) == -1 ? "是" : "否";
                        }
                    }
                }

                sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue((int)item.SerialNo);
                sheet.GetRow(rowIndex).CreateCell(1).SetCellValue((string)item.AppealNo);
                sheet.GetRow(rowIndex).CreateCell(2).SetCellValue(item.VerifyDate != null ? Helpers.StringExtension.ToCalendarRC(item.VerifyDate.ToString("yyyy-MM-dd")) : "");
                sheet.GetRow(rowIndex).CreateCell(3).SetCellValue(item.DeadLine != null ? Helpers.StringExtension.ToCalendarRC(item.DeadLine.ToString("yyyy-MM-dd")) : "");
                sheet.GetRow(rowIndex).CreateCell(4).SetCellValue($"{(string)item.Name} 先生/小姐");
                sheet.GetRow(rowIndex).CreateCell(5).SetCellValue((string)item.Subject);
                sheet.GetRow(rowIndex).CreateCell(6).SetCellValue((string)item.Contents);
                sheet.GetRow(rowIndex).CreateCell(7).SetCellValue((string)item.Runit);
                sheet.GetRow(rowIndex).CreateCell(8).SetCellValue((string)item.UnderTakerName);
                sheet.GetRow(rowIndex).CreateCell(9).SetCellValue((string)item.DeadLine != null ? deadLine : "");
                sheet.GetRow(rowIndex).CreateCell(10).SetCellValue((int?)item.StatusMain != null ? ((int?)item.StatusMain == 27 ? "是" : "否") : "否");
                rowIndex++;
            }

            for(int i = 1; i < 12; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            MemoryStream ms = new MemoryStream();
            hSSFWorkbook.Write(ms);            

            return File(ms.ToArray(), "application/vnd.ms-excel", string.Format($"陳情信件列表.xls"));
        }

        /// <summary>
        /// 案件明細匯出Excel
        /// </summary>
        /// <returns>File</returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult ExcelDetail([FromBody] CaseSearchVM jsonModel)
        {
            _currentUser = Session["User"] as User;
            string jsonString = _caseManagementService.CaseListSearch(jsonModel, _currentUser);
            List<dynamic> result = JsonConvert.DeserializeObject<List<dynamic>>(jsonString);

            var tc = new System.Globalization.TaiwanCalendar();
            var ttime = $"{tc.GetYear(DateTime.Now)}/{DateTime.Now.Month}/{DateTime.Now.Day} {DateTime.Now.Hour}:{DateTime.Now.Minute}";

            // 建立Excel
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(); //建立活頁簿
            ISheet sheet = hSSFWorkbook.CreateSheet("sheet"); //建立sheet

            // 設定樣式
            ICellStyle headerStyle = hSSFWorkbook.CreateCellStyle();
            IFont headerFont = hSSFWorkbook.CreateFont();
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerFont.FontName = "Microsoft JhengHei";
            headerFont.FontHeightInPoints = 20;
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerFont);

            // 標題列
            sheet.CreateRow(0);
            sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 6));
            sheet.GetRow(0).CreateCell(0).SetCellValue("陳情信件明細表");
            sheet.GetRow(0).GetCell(0).CellStyle = headerStyle;
            sheet.CreateRow(2);
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 1));
            sheet.AddMergedRegion(new CellRangeAddress(2, 2, 4, 6));
            sheet.GetRow(2).CreateCell(0).SetCellValue($"匯出筆數: {result.Count()}");
            sheet.GetRow(2).CreateCell(4).SetCellValue($"列印時間: {ttime}");

            int rowIndex = 3;
            foreach(var item in result)
            {
                #region 帶入資料
                sheet.CreateRow(rowIndex + 1).CreateCell(0).SetCellValue("序號");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 1, rowIndex + 1, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 1, rowIndex + 1, 2, 6));
                sheet.GetRow(rowIndex + 1).CreateCell(2).SetCellValue((int)item.SerialNo);
                sheet.CreateRow(rowIndex + 2).CreateCell(0).SetCellValue("案件編號");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 2, rowIndex + 2, 2, 6));
                sheet.GetRow(rowIndex + 2).CreateCell(2).SetCellValue((string)item.AppealNo);
                sheet.CreateRow(rowIndex + 3).CreateCell(0).SetCellValue("民眾姓名");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 3, rowIndex + 3, 2, 6));
                sheet.GetRow(rowIndex + 3).CreateCell(2).SetCellValue($"{(string)item.Name} 先生/小姐");
                sheet.CreateRow(rowIndex + 4).CreateCell(0).SetCellValue("E-Mail");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 4, rowIndex + 4, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 4, rowIndex + 4, 2, 6));
                sheet.GetRow(rowIndex + 4).CreateCell(2).SetCellValue((string)item.EMail);
                sheet.CreateRow(rowIndex + 5).CreateCell(0).SetCellValue("電話");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 5, rowIndex + 5, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 5, rowIndex + 5, 2, 6));
                sheet.GetRow(rowIndex + 5).CreateCell(2).SetCellValue((string)item.Phone ?? "");
                sheet.CreateRow(rowIndex + 6).CreateCell(0).SetCellValue("手機號碼");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 6, rowIndex + 6, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 6, rowIndex + 6, 2, 6));
                sheet.GetRow(rowIndex + 6).CreateCell(2).SetCellValue((string)item.CellPhone ?? "");
                sheet.CreateRow(rowIndex + 7).CreateCell(0).SetCellValue("聯絡地址");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 7, rowIndex + 7, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 7, rowIndex + 7, 2, 6));
                sheet.GetRow(rowIndex + 7).CreateCell(2).SetCellValue((string)item.Adderss ?? "");
                sheet.CreateRow(rowIndex + 8).CreateCell(0).SetCellValue("陳情主旨");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 8, rowIndex + 8, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 8, rowIndex + 8, 2, 6));
                sheet.GetRow(rowIndex + 8).CreateCell(2).SetCellValue((string)item.Subject ?? "");
                sheet.CreateRow(rowIndex + 9).CreateCell(0).SetCellValue("陳情內容");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 9, rowIndex + 9, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 9, rowIndex + 9, 2, 6));
                sheet.GetRow(rowIndex + 9).CreateCell(2).SetCellValue((string)item.Contents ?? "");
                sheet.CreateRow(rowIndex + 10).CreateCell(0).SetCellValue("限辦時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 10, rowIndex + 10, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 10, rowIndex + 10, 2, 6));
                sheet.GetRow(rowIndex + 10).CreateCell(2).SetCellValue(item.DeadLine != null ? Helpers.StringExtension.ToCalendarRC(item.DeadLine.ToString("yyyy-MM-dd")) : "");
                sheet.CreateRow(rowIndex + 11).CreateCell(0).SetCellValue("受理時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 11, rowIndex + 11, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 11, rowIndex + 11, 2, 6));
                sheet.GetRow(rowIndex + 11).CreateCell(2).SetCellValue(item.VerifyDate != null ? Helpers.StringExtension.ToCalendarRC(item.VerifyDate.ToString("yyyy-MM-dd")) : "");
                sheet.CreateRow(rowIndex + 12).CreateCell(0).SetCellValue("轉入公文時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 12, rowIndex + 12, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 12, rowIndex + 12, 2, 6));
                sheet.GetRow(rowIndex + 12).CreateCell(2).SetCellValue($"");
                sheet.CreateRow(rowIndex + 13).CreateCell(0).SetCellValue("公文結案時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 13, rowIndex + 13, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 13, rowIndex + 13, 2, 6));
                sheet.GetRow(rowIndex + 13).CreateCell(2).SetCellValue($"");
                sheet.CreateRow(rowIndex + 14).CreateCell(0).SetCellValue("陳情案件處理情形");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 14, rowIndex + 14, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 14, rowIndex + 14, 2, 6));
                sheet.GetRow(rowIndex + 14).CreateCell(2).SetCellValue((string)item.CaseDetailStatus ?? "");                         
                sheet.CreateRow(rowIndex + 15).CreateCell(0).SetCellValue("陳情類別");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 15, rowIndex + 15, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 15, rowIndex + 15, 2, 6));
                sheet.GetRow(rowIndex + 15).CreateCell(2).SetCellValue((string)item.AppealName ?? "");
                sheet.CreateRow(rowIndex + 16).CreateCell(0).SetCellValue("結案時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 16, rowIndex + 16, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 16, rowIndex + 16, 2, 6));
                sheet.GetRow(rowIndex + 16).CreateCell(2).SetCellValue(item.CloseDate != null ? Helpers.StringExtension.ToCalendarRC(item.CloseDate.ToString("yyyy-MM-dd")) : "");
                sheet.CreateRow(rowIndex + 17).CreateCell(0).SetCellValue("承辦部門/人員");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 17, rowIndex + 17, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 17, rowIndex + 17, 2, 6));
                sheet.GetRow(rowIndex + 17).CreateCell(2).SetCellValue($"{(string)item.Runit ?? string.Empty}/{(string)item.UnderTakerName ?? string.Empty}");
                sheet.CreateRow(rowIndex + 18).CreateCell(0).SetCellValue("是否傳送至公文");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 18, rowIndex + 18, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 18, rowIndex + 18, 2, 6));
                sheet.GetRow(rowIndex + 18).CreateCell(2).SetCellValue($"");
                sheet.CreateRow(rowIndex + 19).CreateCell(0).SetCellValue("回復時間");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 19, rowIndex + 19, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 19, rowIndex + 19, 2, 6));
                sheet.GetRow(rowIndex + 19).CreateCell(2).SetCellValue((string)item.ReplyDate ?? "");
                sheet.CreateRow(rowIndex + 20).CreateCell(0).SetCellValue("回復單位");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 20, rowIndex + 20, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 20, rowIndex + 20, 2, 6));
                sheet.GetRow(rowIndex + 20).CreateCell(2).SetCellValue($"{(string)item.Runit}");
                sheet.CreateRow(rowIndex + 21).CreateCell(0).SetCellValue("回復內容");
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 21, rowIndex + 21, 0, 1));
                sheet.AddMergedRegion(new CellRangeAddress(rowIndex + 21, rowIndex + 21, 2, 6));
                sheet.GetRow(rowIndex + 21).CreateCell(2).SetCellValue((string)item.ReplyContents ?? "");
                #endregion

                rowIndex += 24;
            }

            MemoryStream ms = new MemoryStream();
            hSSFWorkbook.Write(ms);

            return File(ms.ToArray(), "application/vnd.ms-excel", string.Format($"陳情信件明細表.xls"));
        }

        /// <summary>
        /// 二次改分(總窗口)
        /// </summary>
        /// <param name="vm"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult ExportReDispatchReport([FromBody] ExportReDispatchVM vm)
        {
            var data = _fileService.ExportReDispatchForm(vm.AppealNo);
            MemoryStream ms = new MemoryStream(data);
            return File(ms.ToArray(), "application/vnd.ms-word", "二次改分請示單.doc");
        }

        /// <summary>
        /// 案件明細頁面
        /// </summary>
        /// <returns>View</returns>
        public ActionResult ViewDetail()
        {
            return View();
        }
        /// <summary>
        /// 案件編輯修改
        /// </summary>
        /// <returns>View</returns>
        public ActionResult ViewEdit()
        {
            return View();
        }
        /// <summary>
        /// 案件辦理
        /// </summary>
        /// <returns>View</returns>
        public ActionResult ViewHandle()
        {
            return View();
        }
        /// <summary>
        /// 案件結案
        /// </summary>
        /// <returns>View</returns>
        public ActionResult ViewExtend()
        {
            return View();
        }

        #endregion

        #region 取得案件歷程

        /// <summary>
        /// 案號(AppealNo)取得對應案件歷程(AppealDetail)
        /// </summary>
        /// <param name="dto"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult CaseDetails([FromBody] CaseDetailDto dto)
        {
            object result;
            try
            {
                using (MOHWEntities db = new MOHWEntities())
                {
                    result = (from a in db.AppealDetail
                              join b in db.SysSubCode on (int)a.Handle equals b.SerialNo into bb
                              from b in bb.DefaultIfEmpty()
                              join c in db.Organization on a.ResponsibleUnit equals c.SerialNo.ToString() into cc
                              from c in cc.DefaultIfEmpty()
                              join d in db.vw_UserDetail on a.Undertaker equals d.UserId into dd
                              from d in dd.DefaultIfEmpty()
                              select new
                              {
                                  a.SerialNo,
                                  a.AppealNo,
                                  a.Status,
                                  Action = (a.Handle != 244 && a.Handle != 248) ? b.Remark + b.ZHName : b.Remark,
                                  ResponsibleUnit = c.ZHName != null ? c.ZHName : string.Empty,
                                  Executor = d.UserName != null ? d.UserName : a.Undertaker,
                                  ExecuteTime = a.UpdateDate,
                                  a.ReplyContents,
                                  a.ReplyDate,
                                  a.IsTemp
                              }).Where(x => x.IsTemp == "N" && x.AppealNo == dto.AppealNo).OrderBy(x => x.SerialNo).ToList();
                }
            }
            catch(Exception ex)
            {
                result = ex.Message;
            }

            return Content(JsonConvert.SerializeObject(result), "application/json", Encoding.UTF8);
        }

        public class CaseDetailDto
        {
            public string AppealNo { get; set; }
        }

        #endregion

        #region 取得案件附加檔案資訊(案件查詢)
        [System.Web.Mvc.HttpGet]
        public ActionResult LookupAppendFiles(string appealNo)
        {
            List<AppendFile> appendFiles = new List<AppendFile>();
            using (MOHWEntities db = new MOHWEntities())
            {
                appendFiles = db.AppendFile.Where(x => x.AppealNo == appealNo).ToList();
            }
            return Content(JsonConvert.SerializeObject(appendFiles), "application/json", Encoding.UTF8) ;
        }
        #endregion

        #region 案件展延, 重啟
        [System.Web.Mvc.HttpPost]
        public ActionResult CaseExtend([FromBody] CaseExtendVM vm)
        {
            _currentUser = Session["User"] as User;
            bool result = _caseManagementService.CaseExtend(vm, _currentUser, out string returnMessage);
            CaseResponse response = new CaseResponse()
            {
                Success = result,
                Message = result ? "" : returnMessage
            };
            return Content(JsonConvert.SerializeObject(response), "application/json", Encoding.UTF8);
        }

        public ActionResult CaseExtend()
        {
            user = Session["User"] as User;

            string userId = user.UserDetail.UserId;           

            List<CaseExtendViewModel> viewModel = _caseManagementService.GetCaseExtends_WithUserId(userId);

            return View(viewModel);
        }

        /// <summary>
        /// 主管案件展延
        /// </summary>
        /// <param name="AppealNo"></param>
        /// <returns></returns>
        public ActionResult caseContent_supervisor_extend([Bind(Prefix = "id")] string CaseMainKey)
        {
            CaseManagementViewModel viewModel = new CaseManagementViewModel();

            if (string.IsNullOrEmpty(CaseMainKey))
            {
                // 查無則返回搜尋頁面
                // RedirectMessage("查無案件", "CaseQuery");
            }

            viewModel = _caseManagementService.GetAppealForView(CaseMainKey);

            viewModel.appeal = _caseManagementService.GetAppeal(CaseMainKey);
            viewModel.appealDetail = _caseManagementService.GetAppealDetail((int)viewModel.appeal.DetailNo);
            viewModel.appealExtension = _caseManagementService.GetAppealExtensions_WithAppealNo(CaseMainKey).OrderByDescending(x=>x.UpdateDate).FirstOrDefault();

            user = (User)Session["User"];

            // 取得主管資訊
            var supvEmpId = _caseManagementService.GetContactPerson(user.UserDetail.UserId).upepno;
            if (supvEmpId != null)
            {
                var userSupv = _caseManagementService.GetContactPersonFromEmpId(supvEmpId);
                if (userSupv != null)
                {
                    viewModel.Supv = userSupv.UserName;
                    viewModel.SupvID = userSupv.UserId;
                }
            }

            // 組合處理方式
            // 是否為最高階主管
            ViewBag.isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);
            
            viewModel.HanldingType = new List<SysSubCode>();

            if (!ViewBag.isTopSupv)
            {
                viewModel.HanldingType.Add(new SysSubCode() { SerialNo = 202, ZHName = "向上陳核" });
            }
            viewModel.HanldingType.AddRange( new List<SysSubCode>()
            {
                new SysSubCode(){ SerialNo = 203, ZHName = "展期決行" },
                new SysSubCode(){ SerialNo = 204, ZHName = "展期退回" }
            });            

            return View(viewModel);
        }

        /// <summary>
        /// 主管展延 送出
        /// </summary>
        /// <param name="viewModel"></param>
        /// <param name="isTemp"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult caseContent_supervisor_extend(CaseManagementViewModel viewModel, string isTemp)
        {
            try
            {
                User user = (User)Session["User"];
                DateTime timeNow = DateTime.Now;

                // 檢查案件狀態
                Appeal appeal = _caseManagementService.GetAppeal(viewModel.AppealNo);
                AppealExtension extension = _caseManagementService.GetAppealExtensions_WithAppealNo(appeal.AppealNo).OrderByDescending(x => x.SerialNo).FirstOrDefault();

                // 讀取主表查看使用者是否正確
                var assignedPerson = extension.Supv;

                // 若用戶名稱或角色對應不正確
                if (user.UserDetail.UserId != assignedPerson)
                {
                    //string handleZH = _caseManagementService.GetSysSubCode((short)viewModel.handle).ZHName;

                    return Json(new ResultModel
                    {
                        Success = false,
                        Messages = "案件流程有誤，將轉跳查詢畫面",
                        ReturnObject = new
                        {
                            controller = "CaseManagement",
                            action = "CaseQuery",
                        }
                    });
                }

                // 檢查是否為最高主管
                var isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);

                switch (viewModel.handle)
                {
                    // 若需要向上陳核，檢查主管欄位是否為空
                    case 202:
                        if (!isTopSupv && viewModel.SupvID == null)
                        {
                            return Json(new ResultModel { Success = false, Messages = "資料欄位錯誤：查無主管資料，請洽管理者。" });
                        }
                        extension.Supv = viewModel.SupvID;
                        extension.Approval = null;
                        break;

                    // 決行
                    case 203:
                        extension.Approval = true;
                        appeal.DeadLine = extension.RequestDeadline;
                        break;
                    // 退回
                    case 204:
                        extension.Approval = false;
                        break;
                }


                if (!string.IsNullOrEmpty(viewModel.appealDetail.RolSuggest))
                {
                    extension.SupvComment = viewModel.appealDetail.RolSuggest;
                }

                extension.SupvComment = viewModel.SupvSuggest;
                extension.UpdateDate = timeNow;
                extension.UpdateUser = user.UserDetail.UserId;

                _caseManagementService.CreateAppealExtention(extension);
                _caseManagementService.UpdateAppeal(appeal);

                return Json(new ResultModel { Success = true });
            }
            catch(Exception ex)
            {
                logger.Warn(ex.Message);
                logger.Error(ex.ToString());
                return Json(new ResultModel { Success = false, Messages = ex.Message, ReturnObject = ex }) ;
            }            
        }

        #endregion

        #region 總窗口或單位窗口, 重新寄信給承辦人
        [System.Web.Mvc.HttpPost]
        public ActionResult ReSendEmailToUnderTaker([FromBody] ReSendEmailDto dto)
        {
            var result = new EmailServices().SendNotification(dto.AppealNo, dto.UserId);
            return Content(JsonConvert.SerializeObject(result), "application/json", Encoding.UTF8);
        }

        public class ReSendEmailDto
        {
            public string AppealNo { get; set; }
            public string UserId { get; set; }
        }
        #endregion

        /// <summary>
        /// 案件內容
        /// </summary>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult AppealContent(string CaseMainKey)
        {
            CaseManagementViewModel viewModel = caseConetnt_GetCommonInfo(CaseMainKey);

            // 案件是否結案
            ViewBag.isCaseClosed = _caseManagementService.GetAppeal(CaseMainKey).Status == 27;

            if (ViewBag.isCaseClosed)
            {
                // 回覆內容
                ViewBag.ReplyContent = _caseManagementService.GetCaseLastStatus(CaseMainKey).ReplyContents;
            }


            return PartialView("~/Views/CaseManagement/SharedView/_caseContent.cshtml", viewModel);
        }
        
        #region 部內帳號簽核流程

        /// <summary>
        /// 案件維護 (總窗口/單位窗口)
        /// </summary>
        /// <param name="CaseMainKey"></param>
        /// <returns></returns>
        [RoleAttr("155,156")]
        public ActionResult caseContent([Bind(Prefix = "id")] string CaseMainKey)
        {
            user = (User)Session["User"];

            // 檢查是否有 CaseMainKey
            if (string.IsNullOrEmpty(CaseMainKey))
            {
                // 查無則返回搜尋頁面
                RedirectMessage("查無案件", "CaseQuery");                               
            }                                               
            
            // 讀取共用資訊
            CaseManagementViewModel caseManagementVM = caseConetnt_GetCommonInfo(CaseMainKey);

            // 讀取暫存資訊
            AppealDetail tmp = _caseManagementService.ReadTemp(CaseMainKey);
            if (tmp != null)
            {
                caseManagementVM.appealDetail = tmp;
            }

            // 判斷進度為何種窗口
            bool isMainContact = false;
            bool isUnitContact = false;
            switch (caseManagementVM.appeal.Status)
            {
                // 主檔紀錄狀態為
                case 25:
                    isMainContact = true;
                    break;

                case 26:
                    isUnitContact = true;
                    break;

                // 若條件皆不符合，返回搜尋畫面
                default:
                    RedirectMessage("身分不符，無法執行業務。", "CaseQuery");
                    break;
            }

            // 非總窗口亦非單位窗口業務
            if (!isMainContact && !isUnitContact)
            {
                RedirectMessage("身分不符，無法執行業務。", "CaseQuery");                
            }

            ViewBag.isMainContact = isMainContact;

            // 如果是單位窗口 直接給分派人員
            if (isMainContact)
            {
                // 後續處理方式 - 分派單位
                caseManagementVM.AssignUnit = _caseManagementService.GetAssignUnit();
                caseManagementVM.AssignPerson = _caseManagementService.GetAssignPerson("2", 155);
                if (tmp != null)
                {                    
                    // 選取暫存單位 sli: selectListItem
                    foreach(var sli in caseManagementVM.AssignUnit)
                    {
                        if (!string.IsNullOrEmpty(tmp.DisUnit.ToString()))
                        {
                            if(sli.Value == tmp.DisUnit.ToString())
                            {
                                sli.Selected = true;
                            }
                        }
                    }

                    // 選取暫存人員 sli: selectListItem
                    caseManagementVM.AssignPerson = _caseManagementService.GetAssignPerson(tmp.DisUnit.ToString(), 155);
                    foreach (var sli in caseManagementVM.AssignPerson)
                    {
                        if (!string.IsNullOrEmpty(tmp.DisUser.ToString()))
                        {
                            if (sli.Value == tmp.DisUser.ToString())
                            {
                                sli.Selected = true;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(tmp.DisSuggest))
                    {
                        caseManagementVM.DisSuggest = tmp.DisSuggest;
                    }

                    caseManagementVM.AppealCategory = tmp.AppealCategory;
                }
            }
            else if (isUnitContact)
            {
                // 取得單位窗口單位
                var unitContact = user.UserDetail.DepNO;

                var appeal = _caseManagementService.GetAppeal(caseManagementVM.AppealNo);
                // 取得Detail
                var detail = _caseManagementService.GetAppealDetail((int)appeal.DetailNo);

                // 下拉選單選取承辦人    
                caseManagementVM.AssignPerson = _caseManagementService.GetAssignPerson(detail.DisUnit.ToString(), 156);
                if (tmp != null)
                {
                    caseManagementVM.AssignPerson = _caseManagementService.GetAssignPerson(tmp.DisUnit.ToString(), 156);
                    foreach (var sli in caseManagementVM.AssignPerson)
                    {
                        if (!string.IsNullOrEmpty(tmp.DisUser.ToString()))
                        {
                            if (sli.Value == tmp.DisUser.ToString())
                            {
                                sli.Selected = true;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(tmp.DisSuggest))
                    {
                        caseManagementVM.DisSuggest = tmp.DisSuggest;
                    }
                    caseManagementVM.AppealCategory = tmp.AppealCategory;
                }
            }
            _caseManagementService.SetAppealCategoryAndCaseCategoryDropDownMenu(caseManagementVM);

            return View(caseManagementVM);
        }

        /// <summary>
        /// 案件維護 - 窗口表單送出
        /// </summary>
        /// <param name="caseManagementViewModel"></param>
        /// <param name="isTemp"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult caseContent(CaseManagementViewModel caseManagementViewModel, string isTemp)
        {
            // 先檢查此案件是否已有其他人辦理
            // 讀取最後一筆detail           
            AppealDetail appealDetail = _caseManagementService.GetCaseLastStatus(caseManagementViewModel.AppealNo);

            // 讀取主檔紀錄
            Appeal appeal = _caseManagementService.GetAppeal(caseManagementViewModel.AppealNo);

            // 取得案件狀態其對應承辦關卡
            // 總窗口
            if (appeal.Status == 25)
            {             
                if (appealDetail != null)
                {
                    if (appealDetail.Handle == 198)
                    {
                        return Json(new ResultModel 
                        {
                            Success = false,
                            Messages = "該案件已派給單位窗口，無須再分派，將轉跳查詢畫面",
                            ReturnObject = new 
                            {
                                controller = "CaseManagement",
                                action = "CaseQuery",
                            } 
                        });
                    }
                }
            }
            // 單位登記桌
            else if(appeal.Status == 26)
            {            
                if(appealDetail.Handle == 199)
                {
                    return Json(new ResultModel 
                    { 
                        Success = false,
                        Messages = "該案件已派給承辦人，無須再分派，將轉跳查詢畫面",
                        ReturnObject = new
                        {
                            controller = "CaseManagement",
                            action = "CaseQuery",
                        }
                    });
                }
            }

            User user = (User)Session["User"];
            var submitResult = _caseManagementService.SubmitCaseContent(caseManagementViewModel, user, isTemp);

            if (!submitResult.Item1)
            {
                return Json(new ResultModel { Success = false, Messages = submitResult.Item2 });
            }

            return Json(new ResultModel { Success = true }) ;
        }

        /// <summary>
        /// 承辦人 - 畫面
        /// </summary>
        /// <param name="CaseMainKey"></param>
        /// <returns></returns>
        [RoleAttr("157")]
        public ActionResult caseContent_contactPerson([Bind(Prefix = "id")] string CaseMainKey)
        {
            // 沒案號就送回去查詢頁面
            if (string.IsNullOrEmpty(CaseMainKey))
            {
                return RedirectToAction("CaseQuery");
            }

            // 讀取使用者資訊
            user = (User)Session["User"];

            // 讀取共用資訊
            CaseManagementViewModel caseManagementVM = caseConetnt_GetCommonInfo(CaseMainKey);
            if (string.IsNullOrEmpty(caseManagementVM.ReplyContents))
            {
                bool isChinese = true;
                caseManagementVM.ReplyContents = _caseManagementService.AddDefaultReplyContent(caseManagementVM.AppealNo, isChinese);
            }

            // 取得案件最後狀態
            AppealDetail lastCaseStatus = _caseManagementService.GetCaseLastStatus(CaseMainKey);

            // 檢查案件狀態：是否為「簽辦後主管決行」
            if (_caseManagementService.IsConfirmedSignoff(CaseMainKey))
            {
                ViewBag.supvConfirmSignOff = true;
                

                // 帶入回復內容
                caseManagementVM.ReplyContents = lastCaseStatus.ReplyContents;

                // 帶入承辦人資訊
                caseManagementVM.RName = lastCaseStatus.RName;
                caseManagementVM.RPhone = lastCaseStatus.RPhone;
            }
            // 若前一狀態為主管退回
            else if(lastCaseStatus.Status == 219)
            {
                caseManagementVM.ReplyContents = lastCaseStatus.ReplyContents;
            }
            
            // 讀取暫存資訊
            AppealDetail tmp = _caseManagementService.ReadTemp(CaseMainKey);
            // 若有暫存資訊
            if (tmp != null)
            {
                caseManagementVM.appealDetail = tmp;
                caseManagementVM.handle = tmp.Handle;

                if (!string.IsNullOrEmpty(tmp.RName))
                {
                    caseManagementVM.RName = tmp.RName;
                }
                if (!string.IsNullOrEmpty(tmp.RPhone))
                {
                    caseManagementVM.RPhone = tmp.RPhone;
                }
                if (tmp.AppealCategory.HasValue)
                {
                    caseManagementVM.AppealCategory = tmp.AppealCategory;
                }

                switch (caseManagementVM.handle)
                {
                    case 200:
                        if (!string.IsNullOrEmpty(tmp.ReplyContents))
                        {
                            caseManagementVM.ReplyContents = tmp.ReplyContents;
                        }
                        
                        break;
                    case 201:
                        if (!string.IsNullOrEmpty(tmp.CHReasion))
                        {
                            caseManagementVM.ReDisReason = tmp.CHReasion;
                        }
                        break;
                }
            }

            // 後續處理方式
            ViewBag.StaffFileSN = _caseManagementService.GetSerialNoList(caseManagementVM.StaffFiles);
            

            // 取得承辦人處理方式
            if(ViewBag.supvConfirmSignoff != null)
            {
                if (ViewBag.supvConfirmSignoff)
                {
                    caseManagementVM.HanldingType = caseManagementVM.HanldingType.Where(x => x.Remark == "承辦人" && x.ZHName == "結案" ).ToList();
                }
            }
            else
            {
                caseManagementVM.HanldingType = caseManagementVM.HanldingType.Where(x => x.Remark == "承辦人" && x.ZHName != "結案").ToList();
            }

            // 回復單位
            caseManagementVM.ResponsibleUnit = user.UserDetail.UserDep;
            ViewBag.ResponsibleUnitID = user.UserDetail.DepNO;

            // 回復人員
            caseManagementVM.ContactPerson = user.UserDetail.UserName;

            // 向上陳核
            var supvEmpId = _caseManagementService.GetContactPerson(user.UserDetail.UserId).upepno;
            if (supvEmpId != null)
            {
                var userSupv = _caseManagementService.GetContactPersonFromEmpId(supvEmpId);
                if (userSupv != null)
                {
                    caseManagementVM.Supv = userSupv.UserName;
                    caseManagementVM.SupvID = userSupv.UserId;
                }
            }

            var supvEmployeeID = _caseManagementService.GetContactPerson(user.UserDetail.UserId).upepno;
            caseManagementVM.SupvID = _caseManagementService.GetContactPersonFromEmpId(supvEmployeeID).UserId;

            ViewBag.CaseClosed = caseManagementVM.appeal.Status == 27;

            _caseManagementService.SetAppealCategoryAndCaseCategoryDropDownMenu(caseManagementVM);

            return View(caseManagementVM);
        }

        /// <summary>
        /// 承辦人 - 送出表單
        /// </summary>
        /// <param name="appealDetail"></param>
        /// <param name="caseCategory"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult caseContent_contactPerson(CaseManagementViewModel caseManagementViewModel, string isTemp, bool archiveOnly)
        {

            if (caseManagementViewModel.SupvID == null)
            {
                return Json(new ResultModel { Success = false, Messages = "資料欄位錯誤：查無主管資料，請洽管理者。" });
            }            

            try
            {
                user = (User)Session["User"];
                _caseManagementService.submitCase(caseManagementViewModel, user, isTemp, archiveOnly);

                return Json(new ResultModel { Success = true });
            }
            catch(Exception ex)
            {
                return Json(new ResultModel { Success = false, Messages = ex.Message, ReturnObject = ex }) ;
            }

        }

        /// <summary>
        /// 承辦⼈_簽辦 ->簽核主管 會到此⾴
        /// </summary>
        /// <returns></returns>
        //[RoleAttr]
        public ActionResult caseContent_supervisor_signOff([Bind(Prefix = "id")] string CaseMainKey)
        {
            user = (User)Session["User"];
            // 檢查主管是否為最高主管
            ViewBag.isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);

            // 檢查承辦人動作
            var contactPersonRecord = _caseManagementService.GetContactPersonRecord(CaseMainKey);

            // 若為201 轉跳 supervisor_assign
            if (contactPersonRecord.Handle == 201)
            {
                return RedirectToAction("caseContent_supervisor_assign", new { id = CaseMainKey });
            }

            // 讀取共用資訊
            CaseManagementViewModel caseManagementVM = _caseManagementService.GetAppealForView(CaseMainKey);
            ViewBag.PeitionerFileSN = _caseManagementService.GetSerialNoList(caseManagementVM.FrmPetitionerFiles);

            // 選擇下拉式選單 (議題類別)
            string caseCategory = caseManagementVM.CaseCategory.ToString();
            if (!string.IsNullOrEmpty(caseCategory))
            {
                SelectListItem selectListItem = caseManagementVM.StatisticCaseTypeList.FirstOrDefault(s => s.Value == caseCategory.ToString());
                if (selectListItem != null)
                {
                    selectListItem.Selected = true;
                }
                else
                {
                    caseManagementVM.StatisticCaseTypeList.First().Selected = true;
                }
            }

            // 選擇下拉式選單 (陳情類別)
            string appealCate = caseManagementVM.AppealCategory.ToString();
            if (!string.IsNullOrEmpty(appealCate))
            {
                caseManagementVM.PetitionType.Where(s => s.Value == appealCate.ToString()).First().Selected = true;
            }

            // 取得承辦人處理方式
            caseManagementVM.HanldingType = caseManagementVM.HanldingType.Where(x => x.Remark == "部內主管").ToList();

            // 最後一動紀錄
            var lastStat = _caseManagementService.GetCaseLastStatus(CaseMainKey);

            // 若為最高主管，移除向上陳核選項
            if (ViewBag.isTopSupv)
            {
                caseManagementVM.HanldingType.RemoveAll(x => x.SerialNo == 202);
            }


            // 回復內容
            if (!string.IsNullOrEmpty(lastStat.ReplyContents))
            {
                caseManagementVM.ReplyContents = lastStat.ReplyContents; 
            }
            // 改分理由
            if (!string.IsNullOrEmpty(lastStat.CHReasion))
            {
                caseManagementVM.ReDisReason = lastStat.CHReasion;
            }
            // 承辦人姓名
            if (!string.IsNullOrEmpty(lastStat.RName))
            {
                caseManagementVM.RName = lastStat.RName;
            }
            // 承辦人電話
            if (!string.IsNullOrEmpty(lastStat.RPhone))
            {
                caseManagementVM.RPhone = lastStat.RPhone;
            }

            // 回復單位
            caseManagementVM.ResponsibleUnit = user.UserDetail.UserDep;
            ViewBag.ResponsibleUnitID = user.UserDetail.DepNO;

            // 回復人員
            caseManagementVM.ContactPerson = user.UserDetail.UserName;

            // 向上陳核
            var supvEmpId = _caseManagementService.GetContactPerson(user.UserDetail.UserId).upepno;
            if (supvEmpId != null)
            {
                var userSupv = _caseManagementService.GetContactPersonFromEmpId(supvEmpId);
                if (userSupv != null)
                {
                    caseManagementVM.Supv = userSupv.UserName;
                    caseManagementVM.SupvID = userSupv.UserId;
                }
            }

            // 讀取暫存資訊
            AppealDetail tmp = _caseManagementService.ReadTemp(CaseMainKey);
            if (tmp != null)
            {
                caseManagementVM.appealDetail = tmp;

                caseManagementVM.handle = tmp.Handle;
                // 回復內容
                caseManagementVM.ReplyContents = tmp.ReplyContents;
                // 承辦人姓名 (不是回復人員
                caseManagementVM.RName = tmp.RName;
                // 承辦人電話
                caseManagementVM.RPhone = tmp.RPhone;

                caseManagementVM.AppealCategory = tmp.AppealCategory;
                // 主管意見
                caseManagementVM.SupvSuggest = tmp.ReplyRemark;
            }

            return View(caseManagementVM);
        }

        /// <summary>
        /// 簽辦 -> 主管送出
        /// </summary>
        /// <param name="caseManagementVM"></param>
        /// <param name="isTemp"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult caseContent_supervisor_signoff(CaseManagementViewModel caseManagementVM, string isTemp)
        {
            var user = (User)Session["User"];

            // 檢查是否為最高主管
            var isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);

            // 若需要向上改分，檢查主管欄位是否為空
            if (caseManagementVM.handle == 202)
            {
                if (!isTopSupv && caseManagementVM.SupvID == null)
                {
                    return Json(new ResultModel { Success = false, Messages = "資料欄位錯誤：查無主管資料，請洽管理者。" });
                }
            }

            try
            {
                //user = (User)Session["User"];
                _caseManagementService.submitSupervisorForm(caseManagementVM, user, isTemp);
                return Json(new ResultModel { Success = true });
            }
            catch (Exception ex)
            {
                return Json(new ResultModel { Success = false, Messages = ex.Message, ReturnObject = ex });
            }

        }

        /// <summary>
        /// 承辦⼈_改分 ->簽核主管 會到此⾴
        /// </summary>
        /// <returns></returns>
        public ActionResult caseContent_supervisor_assign([Bind(Prefix = "id")] string CaseMainKey)
        {

            user = (User)Session["User"];

            // 檢查主管是否為最高主管
            ViewBag.isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);            

            // 檢查承辦人動作
            var contactPersonRecord = _caseManagementService.GetContactPersonRecord(CaseMainKey);

            // 若為201 轉跳 supervisor_assign
            if (contactPersonRecord.Handle == 200)
            {
                return RedirectToAction("caseContent_supervisor_signOff", new { id = CaseMainKey});
            }

            // 讀取共用資訊
            CaseManagementViewModel caseManagementVM = _caseManagementService.GetAppealForView(CaseMainKey);
            ViewBag.PeitionerFileSN = _caseManagementService.GetSerialNoList(caseManagementVM.FrmPetitionerFiles);


            // 取得承辦人處理方式
            caseManagementVM.HanldingType = caseManagementVM.HanldingType.Where(x => x.Remark == "部內主管").ToList();

            // 最後一動紀錄
            var lastStat = _caseManagementService.GetCaseLastStatus(CaseMainKey);

            // 若為最高主管，移除向上陳核選項
            if (ViewBag.isTopSupv)
            {
                caseManagementVM.HanldingType.RemoveAll(x => x.SerialNo == 202);
            }

            // 回復內容
            if (!string.IsNullOrEmpty(lastStat.ReplyContents))
            {
                caseManagementVM.ReplyContents = lastStat.ReplyContents;
            }
            // 改分理由
            if (!string.IsNullOrEmpty(lastStat.CHReasion))
            {
                caseManagementVM.ReDisReason = lastStat.CHReasion;
            }
            // 承辦人姓名
            if (!string.IsNullOrEmpty(lastStat.RName))
            {
                caseManagementVM.RName = lastStat.RName;
            }
            // 承辦人電話
            if (!string.IsNullOrEmpty(lastStat.RPhone))
            {
                caseManagementVM.RPhone = lastStat.RPhone;
            }

            // 回復單位
            caseManagementVM.ResponsibleUnit = user.UserDetail.UserDep;
            ViewBag.ResponsibleUnitID = user.UserDetail.DepNO;

            // 回復人員
            caseManagementVM.ContactPerson = user.UserDetail.UserName;

            // 向上陳核
            var supvEmpId = _caseManagementService.GetContactPerson(user.UserDetail.UserId).upepno;
            if (supvEmpId != null)
            {
                var userSupv = _caseManagementService.GetContactPersonFromEmpId(supvEmpId);
                if (userSupv != null)
                {
                    caseManagementVM.Supv = userSupv.UserName;
                    caseManagementVM.SupvID = userSupv.UserId;
                }
            }

            // 讀取暫存資訊
            AppealDetail tmp = _caseManagementService.ReadTemp(CaseMainKey);
            if (tmp != null)
            {
                caseManagementVM.appealDetail = tmp;

                caseManagementVM.handle = tmp.Handle;
                // 回復內容
                caseManagementVM.ReplyContents = tmp.ReplyContents;
                // 承辦人姓名 (不是回復人員
                caseManagementVM.RName = tmp.RName;
                // 承辦人電話
                caseManagementVM.RPhone = tmp.RPhone;

                caseManagementVM.AppealCategory = tmp.AppealCategory;
                // 主管意見
                caseManagementVM.SupvSuggest = tmp.ReplyRemark;
            }


            return View(caseManagementVM);
        }

        /// <summary>
        /// 改分 -> 主管 送出
        /// </summary>
        /// <param name="caseManagementVM"></param>
        /// <param name="isTemp"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult caseContent_supervisor_assign(CaseManagementViewModel caseManagementVM, string isTemp)
        {
            var user = (User)Session["User"];

            // 檢查是否為最高主管
            var isTopSupv = _caseManagementService.isTopSupv(user.UserDetail.UserId);

            // 若需要向上改分，檢查主管欄位是否為空
            if (caseManagementVM.handle == 202)
            {
                if (!isTopSupv && caseManagementVM.SupvID == null)
                {
                    return Json(new ResultModel { Success = false, Messages = "資料欄位錯誤：查無主管資料，請洽管理者。" });
                }
            }

            try
            {
                //user = (User)Session["User"];
                _caseManagementService.submitSupervisorForm(caseManagementVM, user, isTemp);
                return Json(new ResultModel { Success = true });
            }
            catch (Exception ex)
            {
                return Json(new ResultModel { Success = false, Messages = ex.Message, ReturnObject = ex });
            }

        }

        /// <summary>
        ///  取得分派人員
        /// </summary>
        /// <param name="deptNo"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult GetAssignPerson(string deptNo, int operatorRole)
        {
            CaseManagementService caseManagementService = new CaseManagementService();
            var result = caseManagementService.GetAssignPerson(deptNo, operatorRole);
            return Json(result);
        }

        /// <summary>
        /// 案件維護共用資訊
        /// </summary>
        private CaseManagementViewModel caseConetnt_GetCommonInfo(string CaseMainKey)
        {
            // 讀取人員資訊
            User user = (User)Session["User"];            

            // 讀取共用資訊             
            CaseManagementViewModel caseManagementVM = _caseManagementService.GetAppealForView(CaseMainKey);

            // 陳情類別及議題類別下拉選單
            _caseManagementService.SetAppealCategoryAndCaseCategoryDropDownMenu(caseManagementVM);            

            ViewBag.PeitionerFileSN = _caseManagementService.GetSerialNoList(caseManagementVM.FrmPetitionerFiles);

            // 是否為總統府來信
            PresidentEmail presidentEmail = _caseManagementService.GetPresidentEmail(caseManagementVM.AppealNo);
            ViewBag.isPresident = presidentEmail != null;
            if (ViewBag.isPresident)
            {
                ViewBag.PresVerifyCode = presidentEmail.VerifyCode;
            }

            return caseManagementVM;
        }

       
        #endregion

        #region 外部機關簽核
        /// <summary>
        /// 案件維護-外部機關簽核
        /// </summary>
        /// <param name="CaseMainKey"></param>
        /// <returns></returns>    
        [System.Web.Mvc.HttpGet]
        public ActionResult CaseContent_subordinateUnits([Bind(Prefix = "id")] string CaseMainKey)
        {
            // 檢查是否有 CaseMainKey
            if (string.IsNullOrEmpty(CaseMainKey))
            {
                // 查無則返回搜尋頁面
                return RedirectToAction("CaseQuery");
            }

            CaseManagementViewModel caseManagementVM = new CaseManagementViewModel();

            var userInfo = Session["User"] as MohwEmail.Models.User;

            // 讀取共用資訊
            caseManagementVM = caseConetnt_GetCommonInfo(CaseMainKey);
            caseManagementVM.AppealNo = CaseMainKey;

            // 取得後續處理方式
            // 派員(依照上個派案人員去判斷這層是什麼角色)

            int detailMo = _caseManagementService.GetAppealNewDetail(CaseMainKey);
            AppealDetail detail = _caseManagementService.GetAppealDetail(detailMo);

            #region 後續處理資料帶入
            SignOff signOff = new SignOff();
            signOff.ResponsibleUnit = userInfo.UserDetail.UserDep;
            signOff.Undertaker = userInfo.UserDetail.UserName;
            signOff.UndertakerName = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.RName) ? "" : detail.RName;
            signOff.UndertakerPhone = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.RPhone) ? "" : detail.RPhone;
            signOff.ReplyContents = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.ReplyContents) ? "" : detail.ReplyContents;
            if (string.IsNullOrEmpty(signOff.ReplyContents))
            {
                bool isChinese = true;
                signOff.ReplyContents = _caseManagementService.AddDefaultReplyContent(caseManagementVM.AppealNo, isChinese);
            }
            Assign assign = new Assign();
            assign.DisSuggest = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.DisSuggest) ? "" : detail.DisSuggest;
            assign.DisUser = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.DisUser) ? "" : detail.DisUser;
            assign.DisUnit = detail.IsTemp.Equals("N") ? "" : detail.DisUnit == null ? "" : detail.DisUnit.Value.ToString();

            ReClass reClass = new ReClass();
            reClass.RJReasion = detail.IsTemp.Equals("N") ? "" : string.IsNullOrEmpty(detail.CHReasion) ? "" : detail.CHReasion;

            caseManagementVM.signOff = signOff;
            caseManagementVM.assign = assign;
            caseManagementVM.reClass = reClass;
            #endregion

            //取出組別名稱
            string userGroup = _caseManagementService.GetUnitGroupName(userInfo.UserDetail.UserId);

            #region 判斷是否暫存與角色
            if (detail.IsTemp.Equals("Y"))
            {
                switch (detail.Role)
                {
                    case "161": //機關單位窗口 
                        if (string.IsNullOrEmpty(detail.DisUser) || detail.DisUser== "ContactPerson")
                        {
                            caseManagementVM.OrganUnit = _caseManagementService.GetOrganUnitGroupType(int.Parse(userInfo.UserDetail.DepNO));
                            caseManagementVM.OrganName = new List<SelectListItem>();
                        }
                        else
                        {
                            var selectUgroup = _caseManagementService.GetOrganUnitGroupType(int.Parse(userInfo.UserDetail.DepNO));
                            caseManagementVM.OrganUnit = _caseManagementService.GetSelectItem(detail.DisUser, selectUgroup, "Ugroup");
                            caseManagementVM.assign.DisUgroup = _caseManagementService.GetUnitGroupName(detail.DisUser);
                            var selectUser = _caseManagementService.GetOrganUnitGroupContactPersonType(caseManagementVM.assign.DisUgroup);
                            caseManagementVM.OrganName = _caseManagementService.GetSelectItem(detail.DisUser, selectUser,"User");
                        }
                        caseManagementVM.Role = "機關單位窗口";
                        break;
                    case "162": //機關組窗口
                        caseManagementVM.OrganUnit = new List<SelectListItem>();
                        if (string.IsNullOrEmpty(detail.DisUser) || detail.DisUser == userInfo.UserDetail.UserId)
                        {
                            caseManagementVM.OrganName = _caseManagementService.GetOrganUnitGroupPersonType(int.Parse(userInfo.UserDetail.DepNO));
                        }
                        else
                        {
                            var selectUser = _caseManagementService.GetOrganUnitGroupPersonType(int.Parse(userInfo.UserDetail.DepNO));
                            caseManagementVM.OrganName = _caseManagementService.GetSelectItem(detail.DisUser, selectUser, "User");
                        }
                        caseManagementVM.Role = "機關組窗口";
                        break;
                    case "163": //機關組承辦人
                        caseManagementVM.OrganName = new List<SelectListItem>();
                        caseManagementVM.OrganUnit = new List<SelectListItem>();
                        caseManagementVM.Role = "機關組承辦人";
                        break;
                    default:
                        break;
                }
            }
            else
            {
                switch (detail.Role)
                {
                    case "155": //機關單位窗口 
                        caseManagementVM.OrganUnit = _caseManagementService.GetOrganUnitGroupType(int.Parse(userInfo.UserDetail.DepNO));
                        caseManagementVM.OrganName = new List<SelectListItem>();  
                        caseManagementVM.Role = "機關單位窗口";
                        break;
                    case "161": //機關組窗口
                        caseManagementVM.OrganUnit = new List<SelectListItem>();
                        caseManagementVM.OrganName = _caseManagementService.GetOrganUnitGroupPersonType(int.Parse(userInfo.UserDetail.DepNO));
                        caseManagementVM.Role = "機關組窗口";
                        break;
                    case "162": //機關組承辦人
                        caseManagementVM.OrganName = new List<SelectListItem>();
                        caseManagementVM.OrganUnit = new List<SelectListItem>();
                        caseManagementVM.Role = "機關組承辦人";
                        caseManagementVM.dealWayRadio = "回復";
                        break;
                    default:
                        break;
                }
            }
            #endregion

            //後續檔案
            ViewBag.StaffFileSN = _caseManagementService.GetSerialNoList(caseManagementVM.StaffFiles);

            return View(caseManagementVM);
        }


        /// <summary>
        /// 案件維護-送出案件
        /// </summary>
        /// <param name="CaseManagementViewModel"></param>
        /// <returns></returns>    
        [System.Web.Mvc.HttpPost]
        public ActionResult CreateOrganUnitCase(CaseManagementViewModel model)
        {
            try
            {
                var userInfo = Session["User"] as MohwEmail.Models.User;
                CaseManagementService caseManagementSer = new CaseManagementService();
                caseManagementSer.CreateOrganUnitCase(model, userInfo);
                return Redirect("CaseQuery");
            }
            catch (Exception ex)
            {
                logger.Warn(ex.Message);
                logger.Error(ex.ToString());
                return View(model);
            }
        }
        /// <summary>
        ///  取得分派人員 (組窗口)
        /// </summary>
        /// <param name="Ugroup"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult GetContactPersonUnitId(string Ugroup)
        {
            //判斷使用者角色
            //var userInfo = Session["User"] as MohwEmail.Models.User;
            var result = new List<SelectListItem>();
            CaseManagementService caseManagementSer = new CaseManagementService();        
            result= caseManagementSer.GetOrganUnitGroupContactPersonType(Ugroup);
                  
            return Json(result);
        }

        /// <summary>
        /// 產生陳核單(送出)
        /// </summary>
        /// <param name="vm"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult ExportRepReport([FromBody] ExportRepVM vm)
        {
            var data = _fileService.ExportSubReForm(vm.AppealNo, vm.Contents);
            MemoryStream ms = new MemoryStream(data);
            return File(ms.ToArray(), "application/vnd.ms-word", "案件陳核單.doc");
        }

        /// <summary>
        /// 產生改分單
        /// </summary>
        /// <param name="vm"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult ExportRJReport([FromBody] ExportRepVM vm)
        {
            var userInfo = Session["User"] as MohwEmail.Models.User;

            var data = _fileService.ExportSubRJForm(vm.AppealNo, vm.Contents, userInfo);
            MemoryStream ms = new MemoryStream(data);
            return File(ms.ToArray(), "application/vnd.ms-word", "案件陳核單.doc");
        }

        /// <summary>
        /// 檢核是否有上傳word
        /// </summary>
        /// <param name="vm"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult CheckFileForSubmit(CaseManagementViewModel model)
        {
            CaseManagementService caseManagementSer = new CaseManagementService();
            var userInfo = Session["User"] as MohwEmail.Models.User;

            var result = caseManagementSer.checkFile( model, userInfo);

            return Json(new { Status = result.Item1, Message = result.Item2 });
        }

        #endregion

        public ActionResult Case()
        {
            CaseEmailViewModel viewModel = new CaseEmailViewModel();
            viewModel.ConditionModel = new EmailConditionModel();

            return View(viewModel);
        }

        #region 案件尚未確認電子信箱查詢
        /// <summary>
        /// 顯示清單 
        /// </summary>
        /// <returns>View</returns>
        public ActionResult CaseEmail()
        {
            return View();
        }

        /// <summary> 
        ///  查詢結果
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public ActionResult GetEmailList(CaseEmailViewModel viewModel)
        {
            var result = _caseManagementService.GetEmailList(viewModel);
            return PartialView("_caseEmailList_Template", result);
        }

        /// <summary>
        /// 取得使用者EMail
        /// </summary>
        /// <param name="UserId"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult GetEmail(string CaseNo)
        {
            var result = _caseManagementService.GetEmail(CaseNo);
            return Json(result);
        }

        /// <summary> 
        ///  修改Email
        /// </summary>
        /// <param name="viewModel"></param>
        /// <returns></returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult SaveEmail(CaseEmailViewModel model)
        {
            var result = _caseManagementService.SaveEmailData(model);

            return Json(new { Status = result.Item1, Message = result.Item2 });
        }

        /// <summary>
        /// 重發認證信
        /// </summary>
        /// <returns>View</returns>
        [System.Web.Mvc.HttpPost]
        public JsonResult CaseEnailAgain(string verifyCode)
        {
            var result = _caseManagementService.SendEmailAgain(verifyCode);

            return Json(new { Status = result.Item1, Message = result.Item2 });
        }

        #endregion

        #region 二次回復
        /// <summary>
        /// 顯示清單(只顯示已結案的案件) 
        /// </summary>
        /// <returns>View</returns>
        public ActionResult CaseReply()
        {
            return View();
        }
        /// <summary>
        /// 重啟案件
        /// </summary>
        /// <returns>View</returns>
        public ActionResult CaseReplyEdit()
        {
            return View();
        }
        #endregion

        void RedirectMessage(string msg, string action)
        {
            ViewBag.msg = msg;
            ViewBag.RedirectToAction = action;
        }

        /* 檔案處理
         * UploadFile(上傳附件)、
         * DeleteFile(移除附件)、
         * GetFile(取得單筆檔案)、
         * GetFiles(壓縮檔形式下載多筆檔案)
         */
        #region 檔案處理
        /// <summary>
        /// 上傳附件
        /// </summary>
        /// <param name="fileInfo"></param>
        /// <param name="file"></param>
        /// <returns></returns>
        public JsonResult UploadFile(AppendFile fileInfo, HttpPostedFileBase[] files)
        {
            List<ResultModel> rsModel = new List<ResultModel>();                                    

            foreach(HttpPostedFileBase file in files)
            {
                AppendFile appendFile = new AppendFile();

                appendFile.Path = Path.Combine(new FilePath().ConvertToPath(fileInfo.Path), fileInfo.AppealNo) + "/" + file.FileName;
                appendFile.AppealNo = fileInfo.AppealNo;
                appendFile.FileName = file.FileName;
                appendFile.Account = Session["User"] == null ? throw new Exception("Invalid User!!") : ((User)Session["User"]).UserDetail.UserId;
                appendFile.Size = file.ContentLength;

                // 檢查檔名是否重複
                bool duplicated = _fileService.CheckDuplicateName(appendFile);

                // 若重複，紀錄重複檔名
                if (duplicated)
                {
                    rsModel.Add(new ResultModel{ Success = false, Messages = "已有重複檔名 !!", ReturnObject = appendFile });
                    continue;
                }

                try
                {
                    // 寫入db                    
                    appendFile = _fileService.CreateFileRecord(appendFile);

                    // 上傳至資料夾
                    _fileService.UploadFile(appendFile, file);

                    rsModel.Add(new ResultModel { Messages = "上傳成功 !!", ReturnObject = appendFile });
                }
                catch(Exception ex)
                {
                    rsModel.Add(new ResultModel { Success = false, Messages = ex.Message, ReturnObject = appendFile });                    
                }


            }

            var result = rsModel.ToArray();
            return Json(result);
        }

        /// <summary>
        /// 移除附件
        /// </summary>
        /// <param name="serialNo"></param>
        /// <returns></returns>
        public JsonResult DeleteFile(int serialNo)
        {
            FileService fileService = new FileService();

            // 移除db紀錄
            AppendFile file = fileService.DeleteFileRecord(serialNo);

            // 檢查檔案是否存在
            if(System.IO.File.Exists(System.Web.HttpContext.Current.Server.MapPath($"/{file.Path}")))
            {
                // 從資料夾移除附件
                fileService.DeleteFile(file);
            }


            return Json(file);
        }

        /// <summary>
        /// 取得單筆檔案
        /// </summary>
        /// <param name="serialNo"></param>
        /// <returns></returns>        
        public ActionResult GetFile(int serialNo)
        {
            try
            {
                byte[] bytes = _fileService.DownloadFile(serialNo);

                AppendFile appendFile = new FileService().GetFileInfo(serialNo);
                // 檔名：案號_編號_檔名
                string fileName = appendFile.FileName;

                string path = Path.Combine(System.Web.HttpContext.Current.Server.MapPath(appendFile.Path));
                if (appendFile.Path[0].ToString() != "/")
                {
                    path = System.Web.HttpContext.Current.Server.MapPath("/" + appendFile.Path);
                }

                if (!System.IO.File.Exists(path))
                {
                    fileName = @"(查無檔案) [" + fileName + "].txt";                    
                }

                FileData file = new FileData
                {
                    Name = fileName,
                    Content = bytes,
                };

                //產生取檔Token
                var token = Guid.NewGuid();

                //存入Cache待下載(限30秒內有效)
                cache.Add(token.ToString(), file, DateTime.Now.AddSeconds(30));

                return Content(token.ToString());
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 壓縮檔形式下載多筆檔案
        /// </summary>
        /// <param name="serialNos"></param>
        /// <returns></returns>
        public ActionResult GetFiles(int[] serialNos)
        {
            var appealNo = string.Empty;

            using(var compressedfileStream = new MemoryStream())
            {
                // 建立壓縮檔
                using (var zipArc = new ZipArchive(compressedfileStream, ZipArchiveMode.Create, false))
                {
                    // create zip entry for each attachment
                    foreach (int sn in serialNos)
                    {
                        // 取得附件檔案
                        byte[] bytes = _fileService.DownloadFile(sn);

                        AppendFile appendFile = new FileService().GetFileInfo(sn);
                        appealNo = appendFile.AppealNo;

                        // 檔名：案號_編號_檔名
                        string fileName = string.Format("{0}", appendFile.FileName);

                        string fullPath = System.Web.HttpContext.Current.Server.MapPath($"~/{appendFile.Path}");
                        // 檢查檔案是否存在
                        if (!System.IO.File.Exists(fullPath))
                        {
                            logger.Warn($"查無檔案 {fullPath}");
                            fileName = @"(查無檔案) [" + fileName + "].log";
                        }                        

                        var zipEntry = zipArc.CreateEntry(fileName);

                        // get stream of the attachment
                        using(var origin = new MemoryStream(bytes))
                        {
                            using(var zipEntryStream = zipEntry.Open())
                            {
                                //copy attachment stream to the zip entry stream
                                origin.CopyTo(zipEntryStream);
                            }
                        }
                    }
                }

                FileData file = new FileData
                {
                    Name = appealNo + ".zip",
                    Content = compressedfileStream.ToArray(),
                };

                //產生取檔Token
                var token = Guid.NewGuid();

                //存入Cache待下載(限30秒內有效)
                cache.Add(token.ToString(), file, DateTime.Now.AddSeconds(30));

                return Content(token.ToString());
            }
            
        }

        /// <summary>
        /// 下載檔案
        /// </summary>
        /// <param name="token"></param>
        /// <returns></returns>
        public ActionResult DownloadFile(string token)
        {          
            var file = cache[token] as FileData;            
            return File(file.Content, "application/octet-stream", file.Name);
        }


        class FileData
        {
            public string Name;
            public byte[] Content;
        }

        #endregion
    }
}