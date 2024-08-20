using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using NPOI.XSSF.UserModel;
using MohwEmail.Models;
using MohwEmail.Services;
using MohwEmail.ViewModels.CaseReport;
using MohwEmail.Helpers;
using MohwEmail.Filters;

namespace MohwEmail.Controllers
{
    [LogFilter]
    [ErrorAttr]
    public class CaseReportController : Controller
    {
        readonly CaseManagementService _caseManagementService;

        private User _currentUser;

        public CaseReportController()
        {
            _caseManagementService = new CaseManagementService();
        }

        // GET: Report
        public ActionResult CaseReport()
        {
            return View();
        }

        /// <summary>
        /// 受理案件統計表
        /// </summary>
        /// <returns>View</returns>
        public ActionResult AcceptReport()
        {
            //頁面初始化
            _currentUser = Session["User"] as User;
            bool isSupervisor = false;
            foreach (var item in _currentUser.rolesList)
            {
                if (item.SerialNo == 155)
                {
                    isSupervisor = true;
                }
            }
            string memberType = _currentUser.UserDetail.Internal == "Y" ? "1" : "2"; //部內 = 1 , 所屬 = 2;
            List<SysSubCode> CaseSourceList = new CommonService().GetCaseSourceList();
            List<Organization> caseOrganizerList = new CommonService().GetCaseOrganizerList(memberType, isSupervisor);
            CaseQueryViewModel caseQueryViewModel = new CaseQueryViewModel();
            string NowD = DateTime.Now.ToString("yyyy/MM/dd");
            caseQueryViewModel.VerifyDateS = NowD;
            caseQueryViewModel.VerifyDateE = NowD;

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
            caseQueryViewModel.CaseStatusList.Add(new SelectListItem() { Text = "請選擇", Value = "0" });

            foreach (SysSubCode CaseSource in CaseSourceList.Where(x => x.MainCode == 4 && x.SerialNo < 30))
            {
                caseQueryViewModel.CaseStatusList.Add(new SelectListItem() { Text = CaseSource.ZHName, Value = CaseSource.SerialNo.ToString() });
            }
            // 議題類別
            caseQueryViewModel.StatisticCaseTypeList = _caseManagementService.GetStatCaseList();

            foreach (var caseOrganizer in caseOrganizerList.Select(x => new { x.SerialNo, x.ZHName }).OrderBy(x => x.SerialNo))
            {
                caseQueryViewModel.CaseOrganizerList.Add(new SelectListItem() { Text = caseOrganizer.ZHName, Value = caseOrganizer.SerialNo.ToString() });
            }
            //統計類型
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "日", Value = "1" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "周", Value = "2" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "月", Value = "3" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "季", Value = "4" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "年", Value = "5" });

            return View(caseQueryViewModel);
        }
        /// <summary>
        /// 受理案件統計表匯出Excel
        /// 統計類型 日/周/月/季/年 
        /// 條件 日期區間 
        /// 匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>File</returns>
        public ActionResult AcceptExcel(CaseQueryViewModel caseQueryViewModel)
        {
            //取得資料
            string vds = caseQueryViewModel.VerifyDateS;
            string vde = caseQueryViewModel.VerifyDateE;
            string ReportList = caseQueryViewModel.ReportList;
            string StatusList = caseQueryViewModel.StatusList;
            string AppealCategory = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.CaseTypeList);
            string PetitionTypeList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.PetitionTypeList);
            string SourceList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.SourceList);
            string CaseOrganizerList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.OrganizerList);

            //條件轉中文
            string AppealCategory_ZH = "";
            string PetitionTypeList_ZH = " ";
            string SourceList_ZH = "";
            List <SysSubCode> CaseSourceList = new CommonService().GetCaseSourceList();
            if (caseQueryViewModel.CaseTypeList != null)
            {
                AppealCategory_ZH += "議題類別： ";
                for (int i = 0; i < caseQueryViewModel.CaseTypeList.Count(); i++)
                {
                    AppealCategory_ZH += CaseSourceList.Where(x => x.SerialNo == Int32.Parse(caseQueryViewModel.CaseTypeList[i])).FirstOrDefault().ZHName;
                    AppealCategory_ZH += "; ";
                }
            }
            if (caseQueryViewModel.PetitionTypeList != null)
            {
                PetitionTypeList_ZH += "陳情類別： ";
                for (int i = 0; i < caseQueryViewModel.PetitionTypeList.Count(); i++)
                {
                    PetitionTypeList_ZH += CaseSourceList.Where(x => x.SerialNo == Int32.Parse(caseQueryViewModel.PetitionTypeList[i])).FirstOrDefault().ZHName;
                    PetitionTypeList_ZH += "; ";
                }
            }
            if (caseQueryViewModel.SourceList != null)
            {
                 SourceList_ZH += "來源管道：";
                for (int i = 0; i < caseQueryViewModel.SourceList.Count(); i++)
                {
                    SourceList_ZH += CaseSourceList.Where(x => x.SerialNo == Int32.Parse(caseQueryViewModel.SourceList[i])).FirstOrDefault().ZHName;
                    SourceList_ZH += "; ";
                }
            }
            
            //因統計類型為動態所以無法載入模板 需要自己刻
            XSSFWorkbook wk = new XSSFWorkbook();
            ICellStyle style = wk.CreateCellStyle();//設定儲存格的樣式：水平對齊置中style.Alignment = HorizontalAlignment.CENTER;//建立一個字型樣式對象
            style.Alignment = HorizontalAlignment.Center;
            XSSFSheet sheet1 = (XSSFSheet)wk.CreateSheet("受理案件統計表");
            //查詢條件
            for (int i = 0; i < 4; i++)
            {
                XSSFRow cells = (XSSFRow)sheet1.CreateRow(i);
                if (i == 0)
                {
                    cells.CreateCell(0).SetCellValue("部長信箱受理案件統計數");
                }
                if (i == 1)
                {
                    _currentUser = Session["User"] as User;

                    cells.CreateCell(0).SetCellValue("列印時間： " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    cells.CreateCell(2).SetCellValue("列印人員： " + _currentUser.UserDetail.UserName);
                    cells.CreateCell(4).SetCellValue("報表編號：SWSC01016");
                }
                if (i == 2)
                {
                    cells.CreateCell(0).SetCellValue("受理日期： " + vds + " ~ " + vde);
                    string ReportList_ = "";
                    switch (ReportList)
                    {
                        case "1":
                            ReportList_ = "日";
                            break;
                        case "2":
                            ReportList_ = "週";
                            break;
                        case "3":
                            ReportList_ = "月";
                            break;
                        case "4":
                            ReportList_ = "季";
                            break;
                        case "5":
                            ReportList_ = "年";
                            break;
                        default:
                            break;
                    }
                    cells.CreateCell(2).SetCellValue("統計類型： " + ReportList_);
                }
                if (i == 3)
                {
                    cells.CreateCell(0).SetCellValue(AppealCategory_ZH);
                    cells.CreateCell(2).SetCellValue(PetitionTypeList_ZH);
                    cells.CreateCell(4).SetCellValue(SourceList_ZH);
                }
            }
            ////欄位title
            //XSSFRow cells2 = (XSSFRow)sheet1.CreateRow(4);
            //cells2.CreateCell(0).SetCellValue("主辦單位");
            //cells2.CreateCell(1).SetCellValue("日期");
            ////XSSFRow cells3 = (XSSFRow)sheet1.CreateRow(4);
            //cells2.CreateCell(2).SetCellValue("受理");
            //cells2.CreateCell(3).SetCellValue("未結案");
            //cells2.CreateCell(4).SetCellValue("已結案");
            //插入資料
            //需要很多邏輯判斷 (總計加總 欄位有多少 本部小計 所屬小計)
            int nRow = 5;//開始插入的行（第五行）
                         //for (int i = 0; i < 99; i++)
                         //{
                         //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                         //    cells4.CreateCell(0).SetCellValue("單位名稱");
                         //    cells4.CreateCell(1).SetCellValue(1);
                         //    cells4.CreateCell(2).SetCellValue(2);
                         //    cells4.CreateCell(3).SetCellValue(3);
                         //}
            switch (ReportList)
            {
                case "1":
                    List<sp_AcceptExcel_1_Result> result = new List<sp_AcceptExcel_1_Result>();
                    using (MOHWEntities db = new MOHWEntities())
                    {
                        result = db.sp_AcceptExcel_1(vds, vde, ReportList, StatusList, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
                        //result = result.Where(x => x.VerifyDate == "2022/06/27").ToList();                    
                    }
                    //找出有幾天資料
                    var ListD = (from p in result select p.VerifyDate).Distinct().ToList();
                    //欄位title
                    XSSFRow cells2 = (XSSFRow)sheet1.CreateRow(4);
                    cells2.CreateCell(0).SetCellValue("主辦單位");
                    //合併
                    sheet1.AddMergedRegion(new CellRangeAddress(4, 5, 0, 0));
                    //
                    sheet1.CreateRow(6).CreateCell(0).SetCellValue("本部小計");
                    sheet1.CreateRow(7).CreateCell(0).SetCellValue("部長室");
                    sheet1.CreateRow(8).CreateCell(0).SetCellValue("主秘室");
                    sheet1.CreateRow(9).CreateCell(0).SetCellValue("綜規司");
                    sheet1.CreateRow(10).CreateCell(0).SetCellValue("社保司");
                    sheet1.CreateRow(11).CreateCell(0).SetCellValue("社工司");
                    sheet1.CreateRow(12).CreateCell(0).SetCellValue("保護司");
                    sheet1.CreateRow(13).CreateCell(0).SetCellValue("照護司");
                    sheet1.CreateRow(14).CreateCell(0).SetCellValue("醫事司");
                    sheet1.CreateRow(15).CreateCell(0).SetCellValue("心理健康司");
                    sheet1.CreateRow(16).CreateCell(0).SetCellValue("口腔健康司");
                    sheet1.CreateRow(17).CreateCell(0).SetCellValue("中醫藥司");
                    sheet1.CreateRow(18).CreateCell(0).SetCellValue("秘書處");
                    sheet1.CreateRow(19).CreateCell(0).SetCellValue("人事處");
                    sheet1.CreateRow(20).CreateCell(0).SetCellValue("政風處");
                    sheet1.CreateRow(21).CreateCell(0).SetCellValue("會計處");
                    sheet1.CreateRow(22).CreateCell(0).SetCellValue("統計處");
                    sheet1.CreateRow(23).CreateCell(0).SetCellValue("資訊處");
                    sheet1.CreateRow(24).CreateCell(0).SetCellValue("法規會");
                    sheet1.CreateRow(25).CreateCell(0).SetCellValue("國際合作組");
                    sheet1.CreateRow(26).CreateCell(0).SetCellValue("醫福會");
                    sheet1.CreateRow(27).CreateCell(0).SetCellValue("健保會");
                    sheet1.CreateRow(28).CreateCell(0).SetCellValue("爭審會");
                    sheet1.CreateRow(29).CreateCell(0).SetCellValue("訓練中心");
                    sheet1.CreateRow(30).CreateCell(0).SetCellValue("監理會");
                    sheet1.CreateRow(31).CreateCell(0).SetCellValue("科發組");
                    sheet1.CreateRow(32).CreateCell(0).SetCellValue("公關室");
                    sheet1.CreateRow(33).CreateCell(0).SetCellValue("國會組");
                    sheet1.CreateRow(34).CreateCell(0).SetCellValue("C肝辦");
                    sheet1.CreateRow(35).CreateCell(0).SetCellValue("長照司");
                    sheet1.CreateRow(36).CreateCell(0).SetCellValue("疾管署");
                    sheet1.CreateRow(37).CreateCell(0).SetCellValue("食藥署");
                    sheet1.CreateRow(38).CreateCell(0).SetCellValue("國健署");
                    sheet1.CreateRow(39).CreateCell(0).SetCellValue("健保署");
                    sheet1.CreateRow(40).CreateCell(0).SetCellValue("國衛院");
                    sheet1.CreateRow(41).CreateCell(0).SetCellValue("中醫藥所");
                    sheet1.CreateRow(42).CreateCell(0).SetCellValue("社家署");
                    sheet1.CreateRow(43).CreateCell(0).SetCellValue("所屬機關小計");
                    sheet1.CreateRow(44).CreateCell(0).SetCellValue("總計");
                    int CellRange1 = 1;
                    int CellRange2 = 3;
                    int cell0 = 1;
                    int cell1 = 1;
                    int cell2 = 2;
                    int cell3 = 3;
                  
                    //合計
                    double Total_sum1 = 0;
                    double Total_sum2 = 0;
                    double Total_sum3 = 0;
                    double Total_sum4 = 0;
                    double Total_sum5 = 0;
                    double Total_sum6 = 0;


                    XSSFRow cells3 = (XSSFRow)sheet1.CreateRow(5);
                    for (int i = 0; i < ListD.Count(); i++)
                    {
                        //本部小計
                        double Total_in1 = 0;
                        double Total_in2 = 0;
                        double Total_in3 = 0;
                        //所屬小計
                        double Total_out1 = 0;
                        double Total_out2 = 0;
                        double Total_out3 = 0;
                        //總計
                        double Total1 = 0;
                        double Total2 = 0;
                        double Total3 = 0;
                        //日期
                        sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1, CellRange2));
                        cells2.CreateCell(cell0).SetCellValue(ListD[i].ToString());
                        cells2.GetCell(cell0).CellStyle = style;
                        cell0 = cell0 + 3;
                        CellRange1 = CellRange1 + 3;
                        CellRange2 = CellRange2 + 3;
                        //狀態
                        cells3.CreateCell(cell1).SetCellValue("受理");
                        cells3.CreateCell(cell2).SetCellValue("未結案");
                        cells3.CreateCell(cell3).SetCellValue("已結案");

                        List<sp_AcceptExcel_1_Result> result_ = new List<sp_AcceptExcel_1_Result>();
                        result_ = result.Where(x => x.VerifyDate == ListD[i].ToString()).ToList();
                        for (int j = 0; j < result_.Count(); j++)
                        {
                            Total_in1 = Total_in1 + (double)result_[j].受理;
                            Total_in2 = Total_in2 + (double)result_[j].未結案;
                            Total_in3 = Total_in3 + (double)result_[j].已結案;
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell3).SetCellValue((double)result_[j].已結案);

                            ////部長室
                            //if (result_[j].主辦單位 == "部長室")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(7).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(7).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(7).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////主任秘書室
                            //if (result_[j].主辦單位 == "主任秘書室")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(8).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(8).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(8).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////綜合規劃司
                            //if (result_[j].主辦單位 == "綜合規劃司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(9).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(9).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(9).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會保險司
                            //if (result_[j].主辦單位 == "社會保險司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(10).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(10).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(10).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會救助及社工司
                            //if (result_[j].主辦單位 == "社會救助及社工司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(11).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(11).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(11).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////保護服務司
                            //if (result_[j].主辦單位 == "保護服務司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(12).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(12).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(12).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////護理及健康照護司
                            //if (result_[j].主辦單位 == "護理及健康照護司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(13).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(13).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(13).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////醫事司
                            //if (result_[j].主辦單位 == "醫事司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(14).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(14).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(14).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////心理健康司
                            //if (result_[j].主辦單位 == "心理健康司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(15).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(15).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(15).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////中醫藥司
                            //if (result_[j].主辦單位 == "中醫藥司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(16).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(16).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(16).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////秘書處
                            //if (result_[j].主辦單位 == "秘書處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(17).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(17).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(17).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////人事處
                            //if (result_[j].主辦單位 == "人事處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(18).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(18).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(18).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////政風處
                            //if (result_[j].主辦單位 == "政風處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(19).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(19).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(19).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////會計處
                            //if (result_[j].主辦單位 == "會計處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(20).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(20).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(20).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////統計處
                            //if (result_[j].主辦單位 == "統計處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(21).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(21).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(21).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////資訊處
                            //if (result_[j].主辦單位 == "資訊處")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(22).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(22).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(22).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////法規會
                            //if (result_[j].主辦單位 == "法規會")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(23).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(23).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(23).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國際合作組
                            //if (result_[j].主辦單位 == "國際合作組")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(24).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(24).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(24).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////附屬醫療及社會福利機構管理會
                            //if (result_[j].主辦單位 == "附屬醫療及社會福利機構管理會")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(25).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(25).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(25).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險會
                            //if (result_[j].主辦單位 == "全民健康保險會")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(26).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(26).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(26).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險爭議審議會
                            //if (result_[j].主辦單位 == "全民健康保險爭議審議會")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(27).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(27).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(27).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////衛生福利人員訓練中心
                            //if (result_[j].主辦單位 == "衛生福利人員訓練中心")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(28).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(28).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(28).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國民年金監理會
                            //if (result_[j].主辦單位 == "國民年金監理會")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(29).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(29).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(29).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////科技發展組
                            //if (result_[j].主辦單位 == "科技發展組")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(30).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(30).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(30).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////公共關係室
                            //if (result_[j].主辦單位 == "公共關係室")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(31).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(31).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(31).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國會聯絡組
                            //if (result_[j].主辦單位 == "國會聯絡組")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(32).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(32).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(32).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國家消除C肝辦公室
                            //if (result_[j].主辦單位 == "國家消除C肝辦公室")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(33).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(33).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(33).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            ////長期照顧司
                            //if (result_[j].主辦單位 == "長期照顧司")
                            //{
                            //    Total_in1 = Total_in1 + (double)result_[j].受理;
                            //    Total_in2 = Total_in2 + (double)result_[j].未結案;
                            //    Total_in3 = Total_in3 + (double)result_[j].已結案;
                            //    sheet1.GetRow(34).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(34).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(34).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            //}
                            //疾病管制署
                            if (result_[j].主辦單位 == "疾病管制署")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(35).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(35).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(35).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //食品藥物管理署
                            if (result_[j].主辦單位 == "食品藥物管理署")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(36).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(36).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(36).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //國民健康署
                            if (result_[j].主辦單位 == "國民健康署")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(37).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(37).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(37).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //中央健康保險署
                            if (result_[j].主辦單位 == "中央健康保險署")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(38).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(38).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(38).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //國家衛生研究院
                            if (result_[j].主辦單位 == "國家衛生研究院")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(39).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(39).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(39).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //國家中醫藥研究所
                            if (result_[j].主辦單位 == "國家中醫藥研究所")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(40).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(40).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(40).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //社會及家庭署
                            if (result_[j].主辦單位 == "社會及家庭署")
                            {
                                Total_out1 = Total_out1 + (double)result_[j].受理;
                                Total_out2 = Total_out2 + (double)result_[j].未結案;
                                Total_out3 = Total_out3 + (double)result_[j].已結案;
                                //sheet1.GetRow(41).CreateCell(cell1).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(41).CreateCell(cell2).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(41).CreateCell(cell3).SetCellValue((double)result_[j].已結案);
                            }
                            //本部小計
                            sheet1.GetRow(6).CreateCell(cell1).SetCellValue(Total_in1 - Total_out1);
                            sheet1.GetRow(6).CreateCell(cell2).SetCellValue(Total_in2 - Total_out2);
                            sheet1.GetRow(6).CreateCell(cell3).SetCellValue(Total_in3 - Total_out2);
                            //所屬小計
                            sheet1.GetRow(43).CreateCell(cell1).SetCellValue(Total_out1);
                            sheet1.GetRow(43).CreateCell(cell2).SetCellValue(Total_out2);
                            sheet1.GetRow(43).CreateCell(cell3).SetCellValue(Total_out3);
                            //總計
                            sheet1.GetRow(44).CreateCell(cell1).SetCellValue(Total_in1);
                            sheet1.GetRow(44).CreateCell(cell2).SetCellValue(Total_in2);
                            sheet1.GetRow(44).CreateCell(cell3).SetCellValue(Total_in3);
                                                   
                        }
                        //合計
                        Total_sum1 = Total_sum1 + Total_in1 - Total_out1;
                        Total_sum2 = Total_sum2 + Total_in2 - Total_out2;
                        Total_sum3 = Total_sum3 + Total_in3 - Total_out3;
                        Total_sum4 = Total_sum4 + Total_out1;
                        Total_sum5 = Total_sum5 + Total_out2;
                        Total_sum6 = Total_sum6 + Total_out3;

                        cell1 = cell1 + 3;
                        cell2 = cell2 + 3;
                        cell3 = cell3 + 3;

                        if (i == (ListD.Count() - 1))
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1, CellRange2));
                            cells2.CreateCell(cell0).SetCellValue("合計");
                            cells2.GetCell(cell0).CellStyle = style;
                            cells3.CreateCell(cell1).SetCellValue("受理");
                            cells3.CreateCell(cell2).SetCellValue("未結案");
                            cells3.CreateCell(cell3).SetCellValue("已結案");
                            sheet1.GetRow(6).CreateCell(cell1).SetCellValue(Total_sum1);
                            sheet1.GetRow(6).CreateCell(cell2).SetCellValue(Total_sum2);
                            sheet1.GetRow(6).CreateCell(cell3).SetCellValue(Total_sum3);
                            sheet1.GetRow(43).CreateCell(cell1).SetCellValue(Total_sum4);
                            sheet1.GetRow(43).CreateCell(cell2).SetCellValue(Total_sum5);
                            sheet1.GetRow(43).CreateCell(cell3).SetCellValue(Total_sum6);
                            sheet1.GetRow(44).CreateCell(cell1).SetCellValue(Total_sum1 + Total_sum4);
                            sheet1.GetRow(44).CreateCell(cell2).SetCellValue(Total_sum2 + Total_sum5);
                            sheet1.GetRow(44).CreateCell(cell3).SetCellValue(Total_sum3 + Total_sum6);
                            var linqStament = from p in result
                                              group p by new { p.主辦單位} into g
                                              select new { Code = g.Key.主辦單位, QTY1 = g.Sum(p => p.受理), QTY2 = g.Sum(p => p.未結案), QTY3 = g.Sum(p => p.已結案) };
                            foreach (var item in linqStament)
                            {                               
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                ////部長室
                                //if (item.Code == "部長室")
                                //{
                                //    sheet1.GetRow(7).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(7).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(7).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////主任秘書室
                                //if (item.Code == "主任秘書室")
                                //{
                                //    sheet1.GetRow(8).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(8).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(8).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////綜合規劃司
                                //if (item.Code == "綜合規劃司")
                                //{
                                  
                                //    sheet1.GetRow(9).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(9).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(9).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////社會保險司
                                //if (item.Code == "社會保險司")
                                //{
                                   
                                //    sheet1.GetRow(10).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(10).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(10).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////社會救助及社工司
                                //if (item.Code == "社會救助及社工司")
                                //{
                                    
                                //    sheet1.GetRow(11).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(11).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(11).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////保護服務司
                                //if (item.Code == "保護服務司")
                                //{
                                    
                                //    sheet1.GetRow(12).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(12).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(12).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////護理及健康照護司
                                //if (item.Code == "護理及健康照護司")
                                //{
                                 
                                //    sheet1.GetRow(13).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(13).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(13).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////醫事司
                                //if (item.Code == "醫事司")
                                //{
                                   
                                //    sheet1.GetRow(14).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(14).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(14).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////心理健康司
                                //if (item.Code == "心理健康司")
                                //{
                                 
                                //    sheet1.GetRow(15).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(15).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(15).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////中醫藥司
                                //if (item.Code == "中醫藥司")
                                //{
                                  
                                //    sheet1.GetRow(16).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(16).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(16).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////秘書處
                                //if (item.Code == "秘書處")
                                //{
                                  
                                //    sheet1.GetRow(17).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(17).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(17).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////人事處
                                //if (item.Code == "人事處")
                                //{
                                   
                                //    sheet1.GetRow(18).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(18).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(18).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////政風處
                                //if (item.Code == "政風處")
                                //{
                                   
                                //    sheet1.GetRow(19).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(19).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(19).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////會計處
                                //if (item.Code == "會計處")
                                //{
                                    
                                //    sheet1.GetRow(20).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(20).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(20).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////統計處
                                //if (item.Code == "統計處")
                                //{
                                   
                                //    sheet1.GetRow(21).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(21).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(21).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////資訊處
                                //if (item.Code == "資訊處")
                                //{
                                   
                                //    sheet1.GetRow(22).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(22).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(22).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////法規會
                                //if (item.Code == "法規會")
                                //{
                                   
                                //    sheet1.GetRow(23).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(23).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(23).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國際合作組
                                //if (item.Code == "國際合作組")
                                //{
                                    
                                //    sheet1.GetRow(24).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(24).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(24).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////附屬醫療及社會福利機構管理會
                                //if (item.Code == "附屬醫療及社會福利機構管理會")
                                //{
                                   
                                //    sheet1.GetRow(25).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(25).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(25).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險會
                                //if (item.Code == "全民健康保險會")
                                //{
                                   
                                //    sheet1.GetRow(26).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(26).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(26).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險爭議審議會
                                //if (item.Code == "全民健康保險爭議審議會")
                                //{
                                   
                                //    sheet1.GetRow(27).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(27).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(27).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////衛生福利人員訓練中心
                                //if (item.Code == "衛生福利人員訓練中心")
                                //{
                                   
                                //    sheet1.GetRow(28).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(28).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(28).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國民年金監理會
                                //if (item.Code == "國民年金監理會")
                                //{
                                   
                                //    sheet1.GetRow(29).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(29).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(29).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////科技發展組
                                //if (item.Code == "科技發展組")
                                //{
                                  
                                //    sheet1.GetRow(30).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(30).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(30).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////公共關係室
                                //if (item.Code == "公共關係室")
                                //{
                                   
                                //    sheet1.GetRow(31).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(31).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(31).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國會聯絡組
                                //if (item.Code == "國會聯絡組")
                                //{
                                    
                                //    sheet1.GetRow(32).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(32).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(32).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國家消除C肝辦公室
                                //if (item.Code == "國家消除C肝辦公室")
                                //{
                                  
                                //    sheet1.GetRow(33).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(33).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(33).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////長期照顧司
                                //if (item.Code == "長期照顧司")
                                //{
                                 
                                //    sheet1.GetRow(34).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(34).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(34).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////疾病管制署
                                //if (item.Code == "疾病管制署")
                                //{
                                   
                                //    sheet1.GetRow(35).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(35).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(35).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////食品藥物管理署
                                //if (item.Code == "食品藥物管理署")
                                //{
        
                                //    sheet1.GetRow(36).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(36).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(36).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國民健康署
                                //if (item.Code == "國民健康署")
                                //{

                                //    sheet1.GetRow(37).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(37).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(37).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////中央健康保險署
                                //if (item.Code == "中央健康保險署")
                                //{

                                //    sheet1.GetRow(38).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(38).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(38).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國家衛生研究院
                                //if (item.Code == "國家衛生研究院")
                                //{
       
                                //    sheet1.GetRow(39).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(39).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(39).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////國家中醫藥研究所
                                //if (item.Code == "國家中醫藥研究所")
                                //{

                                //    sheet1.GetRow(40).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(40).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(40).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                                ////社會及家庭署
                                //if (item.Code == "社會及家庭署")
                                //{
                
                                //    sheet1.GetRow(41).CreateCell(cell1).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(41).CreateCell(cell2).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(41).CreateCell(cell3).SetCellValue((double)item.QTY3);
                                //}
                            }
                        }

                    }

                    //for (int i = 0; i < result.Count(); i++)
                    //{
                    //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                    //    cells4.CreateCell(0).SetCellValue(result[i].主辦單位);
                    //    cells4.CreateCell(1).SetCellValue(result[i].VerifyDate);
                    //    cells4.CreateCell(2).SetCellValue((double)result[i].受理);
                    //    cells4.CreateCell(3).SetCellValue((double)result[i].未結案);
                    //    cells4.CreateCell(4).SetCellValue((double)result[i].已結案);
                    //}
                    break;
                case "2":
                    List<sp_AcceptExcel_2_Result> result2 = new List<sp_AcceptExcel_2_Result>();
                    using (MOHWEntities db = new MOHWEntities())
                    {
                        result2 = db.sp_AcceptExcel_2(vds, vde, ReportList, StatusList, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
                    }
                    //找出有幾天資料
                    //var ListD2 = (from p in result2 select p.VerifyWeek).Distinct().ToList();
                    var ListD2 = result2.Select(p => new { p.VerifyWeek, p.VerifyYear }).Distinct().ToList();

                    //欄位title
                    XSSFRow cells4_ = (XSSFRow)sheet1.CreateRow(4);
                    cells4_.CreateCell(0).SetCellValue("主辦單位");
                    //合併
                    sheet1.AddMergedRegion(new CellRangeAddress(4, 5, 0, 0));
                    //
                    sheet1.CreateRow(6).CreateCell(0).SetCellValue("本部小計");
                    sheet1.CreateRow(7).CreateCell(0).SetCellValue("部長室");
                    sheet1.CreateRow(8).CreateCell(0).SetCellValue("主秘室");
                    sheet1.CreateRow(9).CreateCell(0).SetCellValue("綜規司");
                    sheet1.CreateRow(10).CreateCell(0).SetCellValue("社保司");
                    sheet1.CreateRow(11).CreateCell(0).SetCellValue("社工司");
                    sheet1.CreateRow(12).CreateCell(0).SetCellValue("保護司");
                    sheet1.CreateRow(13).CreateCell(0).SetCellValue("照護司");
                    sheet1.CreateRow(14).CreateCell(0).SetCellValue("醫事司");
                    sheet1.CreateRow(15).CreateCell(0).SetCellValue("心理健康司");
                    sheet1.CreateRow(16).CreateCell(0).SetCellValue("口腔健康司");
                    sheet1.CreateRow(17).CreateCell(0).SetCellValue("中醫藥司");
                    sheet1.CreateRow(18).CreateCell(0).SetCellValue("秘書處");
                    sheet1.CreateRow(19).CreateCell(0).SetCellValue("人事處");
                    sheet1.CreateRow(20).CreateCell(0).SetCellValue("政風處");
                    sheet1.CreateRow(21).CreateCell(0).SetCellValue("會計處");
                    sheet1.CreateRow(22).CreateCell(0).SetCellValue("統計處");
                    sheet1.CreateRow(23).CreateCell(0).SetCellValue("資訊處");
                    sheet1.CreateRow(24).CreateCell(0).SetCellValue("法規會");
                    sheet1.CreateRow(25).CreateCell(0).SetCellValue("國際合作組");
                    sheet1.CreateRow(26).CreateCell(0).SetCellValue("醫福會");
                    sheet1.CreateRow(27).CreateCell(0).SetCellValue("健保會");
                    sheet1.CreateRow(28).CreateCell(0).SetCellValue("爭審會");
                    sheet1.CreateRow(29).CreateCell(0).SetCellValue("訓練中心");
                    sheet1.CreateRow(30).CreateCell(0).SetCellValue("監理會");
                    sheet1.CreateRow(31).CreateCell(0).SetCellValue("科發組");
                    sheet1.CreateRow(32).CreateCell(0).SetCellValue("公關室");
                    sheet1.CreateRow(33).CreateCell(0).SetCellValue("國會組");
                    sheet1.CreateRow(34).CreateCell(0).SetCellValue("C肝辦");
                    sheet1.CreateRow(35).CreateCell(0).SetCellValue("長照司");
                    sheet1.CreateRow(36).CreateCell(0).SetCellValue("疾管署");
                    sheet1.CreateRow(37).CreateCell(0).SetCellValue("食藥署");
                    sheet1.CreateRow(38).CreateCell(0).SetCellValue("國健署");
                    sheet1.CreateRow(39).CreateCell(0).SetCellValue("健保署");
                    sheet1.CreateRow(40).CreateCell(0).SetCellValue("國衛院");
                    sheet1.CreateRow(41).CreateCell(0).SetCellValue("中醫藥所");
                    sheet1.CreateRow(42).CreateCell(0).SetCellValue("社家署");
                    sheet1.CreateRow(43).CreateCell(0).SetCellValue("所屬機關小計");
                    sheet1.CreateRow(44).CreateCell(0).SetCellValue("總計");
                    int CellRange1_ = 1;
                    int CellRange2_ = 3;
                    int cell0_ = 1;
                    int cell1_ = 1;
                    int cell2_ = 2;
                    int cell3_ = 3;
                   
                    //合計
                    double Total_sum1_ = 0;
                    double Total_sum2_ = 0;
                    double Total_sum3_ = 0;
                    double Total_sum4_ = 0;
                    double Total_sum5_ = 0;
                    double Total_sum6_ = 0;


                    XSSFRow cells3_ = (XSSFRow)sheet1.CreateRow(5);
                    for (int i = 0; i < ListD2.Count(); i++)
                    {
                        //本部小計
                        double Total_in1_ = 0;
                        double Total_in2_ = 0;
                        double Total_in3_ = 0;
                        //所屬小計
                        double Total_out1_ = 0;
                        double Total_out2_ = 0;
                        double Total_out3_ = 0;
                        //總計
                        double Total1_ = 0;
                        double Total2_ = 0;
                        double Total3_ = 0;
                        //日期
                        sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1_, CellRange2_));
                        string gwst = ListD2[i].VerifyYear + ListD2[i].VerifyWeek.ToString().PadLeft(3, '0');
                        string d = StringExtension.GetWeekStartTime(gwst).ToString("yyyy/MM/dd");
                        string dd = StringExtension.GetWeekEndTime(gwst).ToString("yyyy/MM/dd");
                        cells4_.CreateCell(cell0_).SetCellValue(d + " ~ " + dd);
                        //cells4_.CreateCell(cell0_).SetCellValue(ListD2[i].ToString());
                        cells4_.GetCell(cell0_).CellStyle = style;
                        cell0_ = cell0_ + 3;
                        CellRange1_ = CellRange1_ + 3;
                        CellRange2_ = CellRange2_ + 3;
                        //狀態
                        cells3_.CreateCell(cell1_).SetCellValue("受理");
                        cells3_.CreateCell(cell2_).SetCellValue("未結案");
                        cells3_.CreateCell(cell3_).SetCellValue("已結案");

                        List<sp_AcceptExcel_2_Result> result_ = new List<sp_AcceptExcel_2_Result>();
                        result_ = result2.Where(x => x.VerifyYear == ListD2[i].VerifyYear.ToString() && x.VerifyWeek == ListD2[i].VerifyWeek).ToList();
                        for (int j = 0; j < result_.Count(); j++)
                        {
                            Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);

                            ////部長室
                            //if (result_[j].主辦單位 == "部長室")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(7).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(7).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(7).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////主任秘書室
                            //if (result_[j].主辦單位 == "主任秘書室")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(8).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(8).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(8).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////綜合規劃司
                            //if (result_[j].主辦單位 == "綜合規劃司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(9).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(9).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(9).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會保險司
                            //if (result_[j].主辦單位 == "社會保險司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(10).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(10).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(10).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會救助及社工司
                            //if (result_[j].主辦單位 == "社會救助及社工司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(11).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(11).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(11).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////保護服務司
                            //if (result_[j].主辦單位 == "保護服務司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(12).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(12).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(12).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////護理及健康照護司
                            //if (result_[j].主辦單位 == "護理及健康照護司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(13).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(13).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(13).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////醫事司
                            //if (result_[j].主辦單位 == "醫事司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(14).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(14).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(14).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////心理健康司
                            //if (result_[j].主辦單位 == "心理健康司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(15).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(15).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(15).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////中醫藥司
                            //if (result_[j].主辦單位 == "中醫藥司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(16).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(16).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(16).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////秘書處
                            //if (result_[j].主辦單位 == "秘書處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(17).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(17).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(17).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////人事處
                            //if (result_[j].主辦單位 == "人事處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(18).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(18).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(18).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////政風處
                            //if (result_[j].主辦單位 == "政風處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(19).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(19).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(19).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////會計處
                            //if (result_[j].主辦單位 == "會計處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(20).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(20).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(20).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////統計處
                            //if (result_[j].主辦單位 == "統計處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(21).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(21).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(21).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////資訊處
                            //if (result_[j].主辦單位 == "資訊處")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(22).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(22).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(22).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////法規會
                            //if (result_[j].主辦單位 == "法規會")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(23).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(23).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(23).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國際合作組
                            //if (result_[j].主辦單位 == "國際合作組")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(24).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(24).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(24).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////附屬醫療及社會福利機構管理會
                            //if (result_[j].主辦單位 == "附屬醫療及社會福利機構管理會")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(25).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(25).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(25).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險會
                            //if (result_[j].主辦單位 == "全民健康保險會")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(26).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(26).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(26).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險爭議審議會
                            //if (result_[j].主辦單位 == "全民健康保險爭議審議會")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(27).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(27).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(27).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////衛生福利人員訓練中心
                            //if (result_[j].主辦單位 == "衛生福利人員訓練中心")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(28).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(28).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(28).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國民年金監理會
                            //if (result_[j].主辦單位 == "國民年金監理會")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(29).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(29).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(29).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////科技發展組
                            //if (result_[j].主辦單位 == "科技發展組")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(30).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(30).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(30).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////公共關係室
                            //if (result_[j].主辦單位 == "公共關係室")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(31).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(31).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(31).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國會聯絡組
                            //if (result_[j].主辦單位 == "國會聯絡組")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(32).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(32).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(32).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國家消除C肝辦公室
                            //if (result_[j].主辦單位 == "國家消除C肝辦公室")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(33).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(33).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(33).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            ////長期照顧司
                            //if (result_[j].主辦單位 == "長期照顧司")
                            //{
                            //    Total_in1_ = Total_in1_ + (double)result_[j].受理;
                            //    Total_in2_ = Total_in2_ + (double)result_[j].未結案;
                            //    Total_in3_ = Total_in3_ + (double)result_[j].已結案;
                            //    sheet1.GetRow(34).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(34).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(34).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            //}
                            //疾病管制署
                            if (result_[j].主辦單位 == "疾病管制署")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(35).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(35).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(35).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //食品藥物管理署
                            if (result_[j].主辦單位 == "食品藥物管理署")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(36).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(36).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(36).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //國民健康署
                            if (result_[j].主辦單位 == "國民健康署")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(37).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(37).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(37).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //中央健康保險署
                            if (result_[j].主辦單位 == "中央健康保險署")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(38).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(38).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(38).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //國家衛生研究院
                            if (result_[j].主辦單位 == "國家衛生研究院")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(39).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(39).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(39).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //國家中醫藥研究所
                            if (result_[j].主辦單位 == "國家中醫藥研究所")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(40).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(40).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(40).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //社會及家庭署
                            if (result_[j].主辦單位 == "社會及家庭署")
                            {
                                Total_out1_ = Total_out1_ + (double)result_[j].受理;
                                Total_out2_ = Total_out2_ + (double)result_[j].未結案;
                                Total_out3_ = Total_out3_ + (double)result_[j].已結案;
                                //sheet1.GetRow(41).CreateCell(cell1_).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(41).CreateCell(cell2_).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(41).CreateCell(cell3_).SetCellValue((double)result_[j].已結案);
                            }
                            //本部小計
                            sheet1.GetRow(6).CreateCell(cell1_).SetCellValue(Total_in1_ - Total_out1_);
                            sheet1.GetRow(6).CreateCell(cell2_).SetCellValue(Total_in2_ - Total_out2_);
                            sheet1.GetRow(6).CreateCell(cell3_).SetCellValue(Total_in3_ - Total_out3_);
                            //所屬小計
                            sheet1.GetRow(43).CreateCell(cell1_).SetCellValue(Total_out1_);
                            sheet1.GetRow(43).CreateCell(cell2_).SetCellValue(Total_out2_);
                            sheet1.GetRow(43).CreateCell(cell3_).SetCellValue(Total_out3_);
                            //總計
                            sheet1.GetRow(44).CreateCell(cell1_).SetCellValue(Total_in1_);
                            sheet1.GetRow(44).CreateCell(cell2_).SetCellValue(Total_in2_);
                            sheet1.GetRow(44).CreateCell(cell3_).SetCellValue(Total_in3_);

                        }
                        //合計
                        Total_sum1_ = Total_sum1_ + Total_in1_ - Total_out1_;
                        Total_sum2_ = Total_sum2_ + Total_in2_ - Total_out2_;
                        Total_sum3_ = Total_sum3_ + Total_in3_ - Total_out3_;
                        Total_sum4_ = Total_sum4_ + Total_out1_;
                        Total_sum5_ = Total_sum5_ + Total_out2_;
                        Total_sum6_ = Total_sum6_ + Total_out3_;

                        cell1_ = cell1_ + 3;
                        cell2_ = cell2_ + 3;
                        cell3_ = cell3_ + 3;

                        if (i == (ListD2.Count() - 1))
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1_, CellRange2_));
                            cells4_.CreateCell(cell0_).SetCellValue("合計");
                            cells4_.GetCell(cell0_).CellStyle = style;
                            cells3_.CreateCell(cell1_).SetCellValue("受理");
                            cells3_.CreateCell(cell2_).SetCellValue("未結案");
                            cells3_.CreateCell(cell3_).SetCellValue("已結案");
                            sheet1.GetRow(6).CreateCell(cell1_).SetCellValue(Total_sum1_);
                            sheet1.GetRow(6).CreateCell(cell2_).SetCellValue(Total_sum2_);
                            sheet1.GetRow(6).CreateCell(cell3_).SetCellValue(Total_sum3_);
                            sheet1.GetRow(43).CreateCell(cell1_).SetCellValue(Total_sum4_);
                            sheet1.GetRow(43).CreateCell(cell2_).SetCellValue(Total_sum5_);
                            sheet1.GetRow(43).CreateCell(cell3_).SetCellValue(Total_sum6_);
                            sheet1.GetRow(44).CreateCell(cell1_).SetCellValue(Total_sum1_ + Total_sum4_);
                            sheet1.GetRow(44).CreateCell(cell2_).SetCellValue(Total_sum2_ + Total_sum5_);
                            sheet1.GetRow(44).CreateCell(cell3_).SetCellValue(Total_sum3_ + Total_sum6_);
                            var linqStament = from p in result_
                                              group p by new { p.主辦單位 } into g
                                              select new { Code = g.Key.主辦單位, QTY1 = g.Sum(p => p.受理), QTY2 = g.Sum(p => p.未結案), QTY3 = g.Sum(p => p.已結案) };
                            foreach (var item in linqStament)
                            {
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                ////部長室
                                //if (item.Code == "部長室")
                                //{
                                //    sheet1.GetRow(7).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(7).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(7).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////主任秘書室
                                //if (item.Code == "主任秘書室")
                                //{
                                //    sheet1.GetRow(8).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(8).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(8).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////綜合規劃司
                                //if (item.Code == "綜合規劃司")
                                //{

                                //    sheet1.GetRow(9).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(9).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(9).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////社會保險司
                                //if (item.Code == "社會保險司")
                                //{

                                //    sheet1.GetRow(10).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(10).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(10).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////社會救助及社工司
                                //if (item.Code == "社會救助及社工司")
                                //{

                                //    sheet1.GetRow(11).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(11).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(11).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////保護服務司
                                //if (item.Code == "保護服務司")
                                //{

                                //    sheet1.GetRow(12).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(12).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(12).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////護理及健康照護司
                                //if (item.Code == "護理及健康照護司")
                                //{

                                //    sheet1.GetRow(13).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(13).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(13).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////醫事司
                                //if (item.Code == "醫事司")
                                //{

                                //    sheet1.GetRow(14).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(14).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(14).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////心理健康司
                                //if (item.Code == "心理健康司")
                                //{

                                //    sheet1.GetRow(15).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(15).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(15).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////中醫藥司
                                //if (item.Code == "中醫藥司")
                                //{

                                //    sheet1.GetRow(16).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(16).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(16).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////秘書處
                                //if (item.Code == "秘書處")
                                //{

                                //    sheet1.GetRow(17).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(17).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(17).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////人事處
                                //if (item.Code == "人事處")
                                //{

                                //    sheet1.GetRow(18).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(18).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(18).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////政風處
                                //if (item.Code == "政風處")
                                //{

                                //    sheet1.GetRow(19).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(19).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(19).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////會計處
                                //if (item.Code == "會計處")
                                //{

                                //    sheet1.GetRow(20).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(20).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(20).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////統計處
                                //if (item.Code == "統計處")
                                //{

                                //    sheet1.GetRow(21).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(21).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(21).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////資訊處
                                //if (item.Code == "資訊處")
                                //{

                                //    sheet1.GetRow(22).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(22).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(22).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////法規會
                                //if (item.Code == "法規會")
                                //{

                                //    sheet1.GetRow(23).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(23).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(23).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國際合作組
                                //if (item.Code == "國際合作組")
                                //{

                                //    sheet1.GetRow(24).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(24).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(24).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////附屬醫療及社會福利機構管理會
                                //if (item.Code == "附屬醫療及社會福利機構管理會")
                                //{

                                //    sheet1.GetRow(25).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(25).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(25).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險會
                                //if (item.Code == "全民健康保險會")
                                //{

                                //    sheet1.GetRow(26).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(26).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(26).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險爭議審議會
                                //if (item.Code == "全民健康保險爭議審議會")
                                //{

                                //    sheet1.GetRow(27).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(27).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(27).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////衛生福利人員訓練中心
                                //if (item.Code == "衛生福利人員訓練中心")
                                //{

                                //    sheet1.GetRow(28).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(28).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(28).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國民年金監理會
                                //if (item.Code == "國民年金監理會")
                                //{

                                //    sheet1.GetRow(29).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(29).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(29).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////科技發展組
                                //if (item.Code == "科技發展組")
                                //{

                                //    sheet1.GetRow(30).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(30).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(30).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////公共關係室
                                //if (item.Code == "公共關係室")
                                //{

                                //    sheet1.GetRow(31).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(31).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(31).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國會聯絡組
                                //if (item.Code == "國會聯絡組")
                                //{

                                //    sheet1.GetRow(32).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(32).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(32).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國家消除C肝辦公室
                                //if (item.Code == "國家消除C肝辦公室")
                                //{

                                //    sheet1.GetRow(33).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(33).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(33).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////長期照顧司
                                //if (item.Code == "長期照顧司")
                                //{

                                //    sheet1.GetRow(34).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(34).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(34).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////疾病管制署
                                //if (item.Code == "疾病管制署")
                                //{

                                //    sheet1.GetRow(35).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(35).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(35).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////食品藥物管理署
                                //if (item.Code == "食品藥物管理署")
                                //{

                                //    sheet1.GetRow(36).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(36).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(36).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國民健康署
                                //if (item.Code == "國民健康署")
                                //{

                                //    sheet1.GetRow(37).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(37).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(37).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////中央健康保險署
                                //if (item.Code == "中央健康保險署")
                                //{

                                //    sheet1.GetRow(38).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(38).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(38).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國家衛生研究院
                                //if (item.Code == "國家衛生研究院")
                                //{

                                //    sheet1.GetRow(39).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(39).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(39).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////國家中醫藥研究所
                                //if (item.Code == "國家中醫藥研究所")
                                //{

                                //    sheet1.GetRow(40).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(40).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(40).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                                ////社會及家庭署
                                //if (item.Code == "社會及家庭署")
                                //{

                                //    sheet1.GetRow(41).CreateCell(cell1_).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(41).CreateCell(cell2_).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(41).CreateCell(cell3_).SetCellValue((double)item.QTY3);
                                //}
                            }
                        }

                    }

                    //for (int i = 0; i < result2.Count(); i++)
                    //{
                    //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                    //    cells4.CreateCell(0).SetCellValue(result2[i].主辦單位);
                    //    string gwst = result2[i].VerifyYear + result2[i].VerifyWeek.ToString().PadLeft(3, '0');
                    //    string d = StringExtension.GetWeekStartTime(gwst).ToString("yyyy/MM/dd");
                    //    string dd = StringExtension.GetWeekEndTime(gwst).ToString("yyyy/MM/dd");
                    //    cells4.CreateCell(1).SetCellValue(d + " ~ " + dd);
                    //    cells4.CreateCell(2).SetCellValue((double)result2[i].受理);
                    //    cells4.CreateCell(3).SetCellValue((double)result2[i].未結案);
                    //    cells4.CreateCell(4).SetCellValue((double)result2[i].已結案);
                    //}
                    break;
                case "3":
                    List<sp_AcceptExcel_3_Result> result3 = new List<sp_AcceptExcel_3_Result>();
                    using (MOHWEntities db = new MOHWEntities())
                    {
                        result3 = db.sp_AcceptExcel_3(vds, vde, ReportList, StatusList, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
                    }
                    //找出有幾天資料
                    //var ListD__ = (from p in result3 select p.VerifyDate).Distinct().ToList().OrderBy(x => x);
                    var ListD__ = (from p in result3 select p.VerifyDate).Distinct().ToList();
                    //欄位title
                    XSSFRow cells2__ = (XSSFRow)sheet1.CreateRow(4);
                    cells2__.CreateCell(0).SetCellValue("主辦單位");
                    //合併
                    sheet1.AddMergedRegion(new CellRangeAddress(4, 5, 0, 0));
                    //
                    sheet1.CreateRow(6).CreateCell(0).SetCellValue("本部小計");
                    sheet1.CreateRow(7).CreateCell(0).SetCellValue("部長室");
                    sheet1.CreateRow(8).CreateCell(0).SetCellValue("主秘室");
                    sheet1.CreateRow(9).CreateCell(0).SetCellValue("綜規司");
                    sheet1.CreateRow(10).CreateCell(0).SetCellValue("社保司");
                    sheet1.CreateRow(11).CreateCell(0).SetCellValue("社工司");
                    sheet1.CreateRow(12).CreateCell(0).SetCellValue("保護司");
                    sheet1.CreateRow(13).CreateCell(0).SetCellValue("照護司");
                    sheet1.CreateRow(14).CreateCell(0).SetCellValue("醫事司");
                    sheet1.CreateRow(15).CreateCell(0).SetCellValue("心理健康司");
                    sheet1.CreateRow(16).CreateCell(0).SetCellValue("口腔健康司");
                    sheet1.CreateRow(17).CreateCell(0).SetCellValue("中醫藥司");
                    sheet1.CreateRow(18).CreateCell(0).SetCellValue("秘書處");
                    sheet1.CreateRow(19).CreateCell(0).SetCellValue("人事處");
                    sheet1.CreateRow(20).CreateCell(0).SetCellValue("政風處");
                    sheet1.CreateRow(21).CreateCell(0).SetCellValue("會計處");
                    sheet1.CreateRow(22).CreateCell(0).SetCellValue("統計處");
                    sheet1.CreateRow(23).CreateCell(0).SetCellValue("資訊處");
                    sheet1.CreateRow(24).CreateCell(0).SetCellValue("法規會");
                    sheet1.CreateRow(25).CreateCell(0).SetCellValue("國際合作組");
                    sheet1.CreateRow(26).CreateCell(0).SetCellValue("醫福會");
                    sheet1.CreateRow(27).CreateCell(0).SetCellValue("健保會");
                    sheet1.CreateRow(28).CreateCell(0).SetCellValue("爭審會");
                    sheet1.CreateRow(29).CreateCell(0).SetCellValue("訓練中心");
                    sheet1.CreateRow(30).CreateCell(0).SetCellValue("監理會");
                    sheet1.CreateRow(31).CreateCell(0).SetCellValue("科發組");
                    sheet1.CreateRow(32).CreateCell(0).SetCellValue("公關室");
                    sheet1.CreateRow(33).CreateCell(0).SetCellValue("國會組");
                    sheet1.CreateRow(34).CreateCell(0).SetCellValue("C肝辦");
                    sheet1.CreateRow(35).CreateCell(0).SetCellValue("長照司");
                    sheet1.CreateRow(36).CreateCell(0).SetCellValue("疾管署");
                    sheet1.CreateRow(37).CreateCell(0).SetCellValue("食藥署");
                    sheet1.CreateRow(38).CreateCell(0).SetCellValue("國健署");
                    sheet1.CreateRow(39).CreateCell(0).SetCellValue("健保署");
                    sheet1.CreateRow(40).CreateCell(0).SetCellValue("國衛院");
                    sheet1.CreateRow(41).CreateCell(0).SetCellValue("中醫藥所");
                    sheet1.CreateRow(42).CreateCell(0).SetCellValue("社家署");
                    sheet1.CreateRow(43).CreateCell(0).SetCellValue("所屬機關小計");
                    sheet1.CreateRow(44).CreateCell(0).SetCellValue("總計");
                    int CellRange1__ = 1;
                    int CellRange2__ = 3;
                    int Cell0__ = 1;
                    int Cell1__ = 1;
                    int Cell2__ = 2;
                    int Cell3__ = 3;
                   
                    //合計
                    double Total_sum1__ = 0;
                    double Total_sum2__ = 0;
                    double Total_sum3__ = 0;
                    double Total_sum4__ = 0;
                    double Total_sum5__ = 0;
                    double Total_sum6__ = 0;


                    XSSFRow cells3__ = (XSSFRow)sheet1.CreateRow(5);
                    for (int i = 0; i < ListD__.Count(); i++)
                    {
                        //本部小計
                        double Total_in1__ = 0;
                        double Total_in2__ = 0;
                        double Total_in3__ = 0;
                        //所屬小計
                        double Total_out1__ = 0;
                        double Total_out2__ = 0;
                        double Total_out3__ = 0;
                        //總計
                        double Total1__ = 0;
                        double Total2__ = 0;
                        double Total3__ = 0;
                        //日期
                        sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1__, CellRange2__));
                        cells2__.CreateCell(Cell0__).SetCellValue(ListD__[i].ToString());
                        cells2__.GetCell(Cell0__).CellStyle = style;
                        Cell0__ = Cell0__ + 3;
                        CellRange1__ = CellRange1__ + 3;
                        CellRange2__ = CellRange2__ + 3;
                        //狀態
                        cells3__.CreateCell(Cell1__).SetCellValue("受理");
                        cells3__.CreateCell(Cell2__).SetCellValue("未結案");
                        cells3__.CreateCell(Cell3__).SetCellValue("已結案");

                        List<sp_AcceptExcel_3_Result> result_ = new List<sp_AcceptExcel_3_Result>();
                        result_ = result3.Where(x => x.VerifyDate == ListD__[i].ToString()).ToList();
                        for (int j = 0; j < result_.Count(); j++)
                        {
                            Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            ////部長室
                            //if (result_[j].主辦單位 == "部長室")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(7).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(7).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(7).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////主任秘書室
                            //if (result_[j].主辦單位 == "主任秘書室")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(8).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(8).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(8).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////綜合規劃司
                            //if (result_[j].主辦單位 == "綜合規劃司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(9).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(9).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(9).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會保險司
                            //if (result_[j].主辦單位 == "社會保險司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(10).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(10).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(10).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會救助及社工司
                            //if (result_[j].主辦單位 == "社會救助及社工司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(11).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(11).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(11).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////保護服務司
                            //if (result_[j].主辦單位 == "保護服務司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(12).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(12).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(12).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////護理及健康照護司
                            //if (result_[j].主辦單位 == "護理及健康照護司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(13).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(13).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(13).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////醫事司
                            //if (result_[j].主辦單位 == "醫事司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(14).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(14).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(14).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////心理健康司
                            //if (result_[j].主辦單位 == "心理健康司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(15).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(15).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(15).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////中醫藥司
                            //if (result_[j].主辦單位 == "中醫藥司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(16).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(16).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(16).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////秘書處
                            //if (result_[j].主辦單位 == "秘書處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(17).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(17).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(17).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////人事處
                            //if (result_[j].主辦單位 == "人事處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(18).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(18).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(18).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////政風處
                            //if (result_[j].主辦單位 == "政風處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(19).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(19).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(19).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////會計處
                            //if (result_[j].主辦單位 == "會計處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(20).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(20).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(20).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////統計處
                            //if (result_[j].主辦單位 == "統計處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(21).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(21).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(21).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////資訊處
                            //if (result_[j].主辦單位 == "資訊處")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(22).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(22).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(22).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////法規會
                            //if (result_[j].主辦單位 == "法規會")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(23).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(23).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(23).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國際合作組
                            //if (result_[j].主辦單位 == "國際合作組")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(24).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(24).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(24).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////附屬醫療及社會福利機構管理會
                            //if (result_[j].主辦單位 == "附屬醫療及社會福利機構管理會")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(25).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(25).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(25).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險會
                            //if (result_[j].主辦單位 == "全民健康保險會")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(26).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(26).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(26).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險爭議審議會
                            //if (result_[j].主辦單位 == "全民健康保險爭議審議會")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(27).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(27).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(27).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////衛生福利人員訓練中心
                            //if (result_[j].主辦單位 == "衛生福利人員訓練中心")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(28).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(28).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(28).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國民年金監理會
                            //if (result_[j].主辦單位 == "國民年金監理會")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(29).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(29).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(29).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////科技發展組
                            //if (result_[j].主辦單位 == "科技發展組")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(30).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(30).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(30).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////公共關係室
                            //if (result_[j].主辦單位 == "公共關係室")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(31).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(31).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(31).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國會聯絡組
                            //if (result_[j].主辦單位 == "國會聯絡組")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(32).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(32).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(32).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國家消除C肝辦公室
                            //if (result_[j].主辦單位 == "國家消除C肝辦公室")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(33).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(33).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(33).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            ////長期照顧司
                            //if (result_[j].主辦單位 == "長期照顧司")
                            //{
                            //    Total_in1__ = Total_in1__ + (double)result_[j].受理;
                            //    Total_in2__ = Total_in2__ + (double)result_[j].未結案;
                            //    Total_in3__ = Total_in3__ + (double)result_[j].已結案;
                            //    sheet1.GetRow(34).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(34).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(34).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            //}
                            //疾病管制署
                            if (result_[j].主辦單位 == "疾病管制署")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(35).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(35).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(35).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //食品藥物管理署
                            if (result_[j].主辦單位 == "食品藥物管理署")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(36).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(36).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(36).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //國民健康署
                            if (result_[j].主辦單位 == "國民健康署")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(37).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(37).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(37).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //中央健康保險署
                            if (result_[j].主辦單位 == "中央健康保險署")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(38).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(38).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(38).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //國家衛生研究院
                            if (result_[j].主辦單位 == "國家衛生研究院")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(39).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(39).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(39).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //國家中醫藥研究所
                            if (result_[j].主辦單位 == "國家中醫藥研究所")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(40).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(40).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(40).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //社會及家庭署
                            if (result_[j].主辦單位 == "社會及家庭署")
                            {
                                Total_out1__ = Total_out1__ + (double)result_[j].受理;
                                Total_out2__ = Total_out2__ + (double)result_[j].未結案;
                                Total_out3__ = Total_out3__ + (double)result_[j].已結案;
                                //sheet1.GetRow(41).CreateCell(Cell1__).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(41).CreateCell(Cell2__).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(41).CreateCell(Cell3__).SetCellValue((double)result_[j].已結案);
                            }
                            //本部小計
                            sheet1.GetRow(6).CreateCell(Cell1__).SetCellValue(Total_in1__ - Total_out1__);
                            sheet1.GetRow(6).CreateCell(Cell2__).SetCellValue(Total_in2__ - Total_out2__);
                            sheet1.GetRow(6).CreateCell(Cell3__).SetCellValue(Total_in3__ - Total_out3__);
                            //所屬小計
                            sheet1.GetRow(43).CreateCell(Cell1__).SetCellValue(Total_out1__);
                            sheet1.GetRow(43).CreateCell(Cell2__).SetCellValue(Total_out2__);
                            sheet1.GetRow(43).CreateCell(Cell3__).SetCellValue(Total_out3__);
                            //總計
                            sheet1.GetRow(44).CreateCell(Cell1__).SetCellValue(Total_in1__ );
                            sheet1.GetRow(44).CreateCell(Cell2__).SetCellValue(Total_in2__ );
                            sheet1.GetRow(44).CreateCell(Cell3__).SetCellValue(Total_in3__ );

                        }
                        //合計
                        Total_sum1__ = Total_sum1__ + Total_in1__ - Total_out1__;
                        Total_sum2__ = Total_sum2__ + Total_in2__ - Total_out2__;
                        Total_sum3__ = Total_sum3__ + Total_in3__ - Total_out3__;
                        Total_sum4__ = Total_sum4__ + Total_out1__;
                        Total_sum5__ = Total_sum5__ + Total_out2__;
                        Total_sum6__ = Total_sum6__ + Total_out3__;

                        Cell1__ = Cell1__ + 3;
                        Cell2__ = Cell2__ + 3;
                        Cell3__ = Cell3__ + 3;

                        if (i == (ListD__.Count() - 1))
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1__, CellRange2__));
                            cells2__.CreateCell(Cell0__).SetCellValue("合計");
                            cells2__.GetCell(Cell0__).CellStyle = style;
                            cells3__.CreateCell(Cell1__).SetCellValue("受理");
                            cells3__.CreateCell(Cell2__).SetCellValue("未結案");
                            cells3__.CreateCell(Cell3__).SetCellValue("已結案");
                            sheet1.GetRow(6).CreateCell(Cell1__).SetCellValue(Total_sum1__);
                            sheet1.GetRow(6).CreateCell(Cell2__).SetCellValue(Total_sum2__);
                            sheet1.GetRow(6).CreateCell(Cell3__).SetCellValue(Total_sum3__);
                            sheet1.GetRow(43).CreateCell(Cell1__).SetCellValue(Total_sum4__);
                            sheet1.GetRow(43).CreateCell(Cell2__).SetCellValue(Total_sum5__);
                            sheet1.GetRow(43).CreateCell(Cell3__).SetCellValue(Total_sum6__);
                            sheet1.GetRow(44).CreateCell(Cell1__).SetCellValue(Total_sum1__ + Total_sum4__);
                            sheet1.GetRow(44).CreateCell(Cell2__).SetCellValue(Total_sum2__ + Total_sum5__);
                            sheet1.GetRow(44).CreateCell(Cell3__).SetCellValue(Total_sum3__ + Total_sum6__);
                            var linqStament = from p in result3
                                              group p by new { p.主辦單位 } into g
                                              select new { Code = g.Key.主辦單位, QTY1 = g.Sum(p => p.受理), QTY2 = g.Sum(p => p.未結案), QTY3 = g.Sum(p => p.已結案) };
                            foreach (var item in linqStament)
                            {
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                ////部長室
                                //if (item.Code == "部長室")
                                //{
                                //    sheet1.GetRow(7).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(7).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(7).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////主任秘書室
                                //if (item.Code == "主任秘書室")
                                //{
                                //    sheet1.GetRow(8).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(8).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(8).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////綜合規劃司
                                //if (item.Code == "綜合規劃司")
                                //{

                                //    sheet1.GetRow(9).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(9).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(9).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////社會保險司
                                //if (item.Code == "社會保險司")
                                //{

                                //    sheet1.GetRow(10).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(10).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(10).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////社會救助及社工司
                                //if (item.Code == "社會救助及社工司")
                                //{

                                //    sheet1.GetRow(11).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(11).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(11).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////保護服務司
                                //if (item.Code == "保護服務司")
                                //{

                                //    sheet1.GetRow(12).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(12).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(12).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////護理及健康照護司
                                //if (item.Code == "護理及健康照護司")
                                //{

                                //    sheet1.GetRow(13).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(13).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(13).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////醫事司
                                //if (item.Code == "醫事司")
                                //{

                                //    sheet1.GetRow(14).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(14).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(14).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////心理健康司
                                //if (item.Code == "心理健康司")
                                //{

                                //    sheet1.GetRow(15).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(15).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(15).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////中醫藥司
                                //if (item.Code == "中醫藥司")
                                //{

                                //    sheet1.GetRow(16).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(16).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(16).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////秘書處
                                //if (item.Code == "秘書處")
                                //{

                                //    sheet1.GetRow(17).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(17).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(17).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////人事處
                                //if (item.Code == "人事處")
                                //{

                                //    sheet1.GetRow(18).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(18).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(18).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////政風處
                                //if (item.Code == "政風處")
                                //{

                                //    sheet1.GetRow(19).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(19).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(19).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////會計處
                                //if (item.Code == "會計處")
                                //{

                                //    sheet1.GetRow(20).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(20).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(20).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////統計處
                                //if (item.Code == "統計處")
                                //{

                                //    sheet1.GetRow(21).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(21).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(21).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////資訊處
                                //if (item.Code == "資訊處")
                                //{

                                //    sheet1.GetRow(22).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(22).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(22).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////法規會
                                //if (item.Code == "法規會")
                                //{

                                //    sheet1.GetRow(23).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(23).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(23).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國際合作組
                                //if (item.Code == "國際合作組")
                                //{

                                //    sheet1.GetRow(24).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(24).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(24).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////附屬醫療及社會福利機構管理會
                                //if (item.Code == "附屬醫療及社會福利機構管理會")
                                //{

                                //    sheet1.GetRow(25).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(25).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(25).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險會
                                //if (item.Code == "全民健康保險會")
                                //{

                                //    sheet1.GetRow(26).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(26).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(26).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險爭議審議會
                                //if (item.Code == "全民健康保險爭議審議會")
                                //{

                                //    sheet1.GetRow(27).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(27).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(27).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////衛生福利人員訓練中心
                                //if (item.Code == "衛生福利人員訓練中心")
                                //{

                                //    sheet1.GetRow(28).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(28).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(28).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國民年金監理會
                                //if (item.Code == "國民年金監理會")
                                //{

                                //    sheet1.GetRow(29).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(29).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(29).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////科技發展組
                                //if (item.Code == "科技發展組")
                                //{

                                //    sheet1.GetRow(30).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(30).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(30).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////公共關係室
                                //if (item.Code == "公共關係室")
                                //{

                                //    sheet1.GetRow(31).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(31).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(31).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國會聯絡組
                                //if (item.Code == "國會聯絡組")
                                //{

                                //    sheet1.GetRow(32).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(32).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(32).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國家消除C肝辦公室
                                //if (item.Code == "國家消除C肝辦公室")
                                //{

                                //    sheet1.GetRow(33).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(33).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(33).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////長期照顧司
                                //if (item.Code == "長期照顧司")
                                //{

                                //    sheet1.GetRow(34).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(34).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(34).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////疾病管制署
                                //if (item.Code == "疾病管制署")
                                //{

                                //    sheet1.GetRow(35).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(35).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(35).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////食品藥物管理署
                                //if (item.Code == "食品藥物管理署")
                                //{

                                //    sheet1.GetRow(36).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(36).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(36).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國民健康署
                                //if (item.Code == "國民健康署")
                                //{

                                //    sheet1.GetRow(37).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(37).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(37).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////中央健康保險署
                                //if (item.Code == "中央健康保險署")
                                //{

                                //    sheet1.GetRow(38).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(38).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(38).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國家衛生研究院
                                //if (item.Code == "國家衛生研究院")
                                //{

                                //    sheet1.GetRow(39).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(39).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(39).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////國家中醫藥研究所
                                //if (item.Code == "國家中醫藥研究所")
                                //{

                                //    sheet1.GetRow(40).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(40).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(40).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                                ////社會及家庭署
                                //if (item.Code == "社會及家庭署")
                                //{

                                //    sheet1.GetRow(41).CreateCell(Cell1__).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(41).CreateCell(Cell2__).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(41).CreateCell(Cell3__).SetCellValue((double)item.QTY3);
                                //}
                            }
                        }

                    }
                    //for (int i = 0; i < result3.Count(); i++)
                    //{
                    //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                    //    cells4.CreateCell(0).SetCellValue(result3[i].主辦單位);
                    //    cells4.CreateCell(1).SetCellValue(result3[i].VerifyDate);
                    //    cells4.CreateCell(2).SetCellValue((double)result3[i].受理);
                    //    cells4.CreateCell(3).SetCellValue((double)result3[i].未結案);
                    //    cells4.CreateCell(4).SetCellValue((double)result3[i].已結案);
                    //}
                    break;
                case "4":
                    List<sp_AcceptExcel_4_Result> result4 = new List<sp_AcceptExcel_4_Result>();
                    using (MOHWEntities db = new MOHWEntities())
                    {
                        result4 = db.sp_AcceptExcel_4(vds, vde, ReportList, StatusList, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
                    }
                    //for (int i = 0; i < result4.Count(); i++)
                    //{
                    //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                    //    cells4.CreateCell(0).SetCellValue(result4[i].主辦單位);
                    //    cells4.CreateCell(1).SetCellValue(result4[i].VerifyYear + "第" + result4[i].VerifyS + "季");
                    //    cells4.CreateCell(2).SetCellValue((double)result4[i].受理);
                    //    cells4.CreateCell(3).SetCellValue((double)result4[i].未結案);
                    //    cells4.CreateCell(4).SetCellValue((double)result4[i].已結案);
                    //}
                    //找出有幾天資料
                    //var ListD___ = (from p in result4 select p.VerifyDate).Distinct().ToList();
                    var ListD___ = result4.Select(p => new { p.VerifyYear, p.VerifyS }).Distinct().ToList();
                    //欄位title
                    XSSFRow cells2___ = (XSSFRow)sheet1.CreateRow(4);
                    cells2___.CreateCell(0).SetCellValue("主辦單位");
                    //合併
                    sheet1.AddMergedRegion(new CellRangeAddress(4, 5, 0, 0));
                    //
                    sheet1.CreateRow(6).CreateCell(0).SetCellValue("本部小計");
                    sheet1.CreateRow(7).CreateCell(0).SetCellValue("部長室");
                    sheet1.CreateRow(8).CreateCell(0).SetCellValue("主秘室");
                    sheet1.CreateRow(9).CreateCell(0).SetCellValue("綜規司");
                    sheet1.CreateRow(10).CreateCell(0).SetCellValue("社保司");
                    sheet1.CreateRow(11).CreateCell(0).SetCellValue("社工司");
                    sheet1.CreateRow(12).CreateCell(0).SetCellValue("保護司");
                    sheet1.CreateRow(13).CreateCell(0).SetCellValue("照護司");
                    sheet1.CreateRow(14).CreateCell(0).SetCellValue("醫事司");
                    sheet1.CreateRow(15).CreateCell(0).SetCellValue("心理健康司");
                    sheet1.CreateRow(16).CreateCell(0).SetCellValue("口腔健康司");
                    sheet1.CreateRow(17).CreateCell(0).SetCellValue("中醫藥司");
                    sheet1.CreateRow(18).CreateCell(0).SetCellValue("秘書處");
                    sheet1.CreateRow(19).CreateCell(0).SetCellValue("人事處");
                    sheet1.CreateRow(20).CreateCell(0).SetCellValue("政風處");
                    sheet1.CreateRow(21).CreateCell(0).SetCellValue("會計處");
                    sheet1.CreateRow(22).CreateCell(0).SetCellValue("統計處");
                    sheet1.CreateRow(23).CreateCell(0).SetCellValue("資訊處");
                    sheet1.CreateRow(24).CreateCell(0).SetCellValue("法規會");
                    sheet1.CreateRow(25).CreateCell(0).SetCellValue("國際合作組");
                    sheet1.CreateRow(26).CreateCell(0).SetCellValue("醫福會");
                    sheet1.CreateRow(27).CreateCell(0).SetCellValue("健保會");
                    sheet1.CreateRow(28).CreateCell(0).SetCellValue("爭審會");
                    sheet1.CreateRow(29).CreateCell(0).SetCellValue("訓練中心");
                    sheet1.CreateRow(30).CreateCell(0).SetCellValue("監理會");
                    sheet1.CreateRow(31).CreateCell(0).SetCellValue("科發組");
                    sheet1.CreateRow(32).CreateCell(0).SetCellValue("公關室");
                    sheet1.CreateRow(33).CreateCell(0).SetCellValue("國會組");
                    sheet1.CreateRow(34).CreateCell(0).SetCellValue("C肝辦");
                    sheet1.CreateRow(35).CreateCell(0).SetCellValue("長照司");
                    sheet1.CreateRow(36).CreateCell(0).SetCellValue("疾管署");
                    sheet1.CreateRow(37).CreateCell(0).SetCellValue("食藥署");
                    sheet1.CreateRow(38).CreateCell(0).SetCellValue("國健署");
                    sheet1.CreateRow(39).CreateCell(0).SetCellValue("健保署");
                    sheet1.CreateRow(40).CreateCell(0).SetCellValue("國衛院");
                    sheet1.CreateRow(41).CreateCell(0).SetCellValue("中醫藥所");
                    sheet1.CreateRow(42).CreateCell(0).SetCellValue("社家署");
                    sheet1.CreateRow(43).CreateCell(0).SetCellValue("所屬機關小計");
                    sheet1.CreateRow(44).CreateCell(0).SetCellValue("總計");
                    int CellRange1___ = 1;
                    int CellRange2___ = 3;
                    int cell0___ = 1;
                    int cell1___ = 1;
                    int cell2___ = 2;
                    int cell3___ = 3;
                   
                    //合計
                    double Total_sum1___ = 0;
                    double Total_sum2___ = 0;
                    double Total_sum3___ = 0;
                    double Total_sum4___ = 0;
                    double Total_sum5___ = 0;
                    double Total_sum6___ = 0;


                    XSSFRow cells3___ = (XSSFRow)sheet1.CreateRow(5);
                    for (int i = 0; i < ListD___.Count(); i++)
                    {
                        //本部小計
                        double Total_in1___ = 0;
                        double Total_in2___ = 0;
                        double Total_in3___ = 0;
                        //所屬小計
                        double Total_out1___ = 0;
                        double Total_out2___ = 0;
                        double Total_out3___ = 0;
                        //總計
                        double Total1___ = 0;
                        double Total2___ = 0;
                        double Total3___ = 0;
                        //日期
                        sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1___, CellRange2___));
                        //cells4.CreateCell(1).SetCellValue(result4[i].VerifyYear + "第" + result4[i].VerifyS + "季");
                        cells2___.CreateCell(cell0___).SetCellValue(ListD___[i].VerifyYear.ToString() + "第" + ListD___[i].VerifyS.ToString() + "季");
                        cells2___.GetCell(cell0___).CellStyle = style;
                        cell0___ = cell0___ + 3;
                        CellRange1___ = CellRange1___ + 3;
                        CellRange2___ = CellRange2___ + 3;
                        //狀態
                        cells3___.CreateCell(cell1___).SetCellValue("受理");
                        cells3___.CreateCell(cell2___).SetCellValue("未結案");
                        cells3___.CreateCell(cell3___).SetCellValue("已結案");

                        List<sp_AcceptExcel_4_Result> result_ = new List<sp_AcceptExcel_4_Result>();
                        result_ = result4.Where(x => x.VerifyYear == ListD___[i].VerifyYear.ToString() && x.VerifyS == ListD___[i].VerifyS).ToList();
                        for (int j = 0; j < result_.Count(); j++)
                        {
                            Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);

                            ////部長室
                            //if (result_[j].主辦單位 == "部長室")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(7).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(7).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(7).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////主任秘書室
                            //if (result_[j].主辦單位 == "主任秘書室")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(8).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(8).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(8).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////綜合規劃司
                            //if (result_[j].主辦單位 == "綜合規劃司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(9).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(9).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(9).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會保險司
                            //if (result_[j].主辦單位 == "社會保險司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(10).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(10).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(10).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會救助及社工司
                            //if (result_[j].主辦單位 == "社會救助及社工司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(11).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(11).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(11).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////保護服務司
                            //if (result_[j].主辦單位 == "保護服務司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(12).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(12).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(12).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////護理及健康照護司
                            //if (result_[j].主辦單位 == "護理及健康照護司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(13).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(13).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(13).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////醫事司
                            //if (result_[j].主辦單位 == "醫事司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(14).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(14).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(14).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////心理健康司
                            //if (result_[j].主辦單位 == "心理健康司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(15).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(15).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(15).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////中醫藥司
                            //if (result_[j].主辦單位 == "中醫藥司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(16).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(16).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(16).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////秘書處
                            //if (result_[j].主辦單位 == "秘書處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(17).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(17).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(17).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////人事處
                            //if (result_[j].主辦單位 == "人事處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(18).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(18).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(18).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////政風處
                            //if (result_[j].主辦單位 == "政風處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(19).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(19).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(19).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////會計處
                            //if (result_[j].主辦單位 == "會計處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(20).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(20).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(20).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////統計處
                            //if (result_[j].主辦單位 == "統計處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(21).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(21).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(21).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////資訊處
                            //if (result_[j].主辦單位 == "資訊處")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(22).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(22).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(22).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////法規會
                            //if (result_[j].主辦單位 == "法規會")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(23).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(23).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(23).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國際合作組
                            //if (result_[j].主辦單位 == "國際合作組")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(24).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(24).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(24).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////附屬醫療及社會福利機構管理會
                            //if (result_[j].主辦單位 == "附屬醫療及社會福利機構管理會")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(25).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(25).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(25).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險會
                            //if (result_[j].主辦單位 == "全民健康保險會")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(26).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(26).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(26).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險爭議審議會
                            //if (result_[j].主辦單位 == "全民健康保險爭議審議會")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(27).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(27).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(27).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////衛生福利人員訓練中心
                            //if (result_[j].主辦單位 == "衛生福利人員訓練中心")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(28).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(28).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(28).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國民年金監理會
                            //if (result_[j].主辦單位 == "國民年金監理會")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(29).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(29).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(29).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////科技發展組
                            //if (result_[j].主辦單位 == "科技發展組")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(30).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(30).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(30).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////公共關係室
                            //if (result_[j].主辦單位 == "公共關係室")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(31).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(31).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(31).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國會聯絡組
                            //if (result_[j].主辦單位 == "國會聯絡組")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(32).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(32).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(32).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國家消除C肝辦公室
                            //if (result_[j].主辦單位 == "國家消除C肝辦公室")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(33).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(33).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(33).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            ////長期照顧司
                            //if (result_[j].主辦單位 == "長期照顧司")
                            //{
                            //    Total_in1___ = Total_in1___ + (double)result_[j].受理;
                            //    Total_in2___ = Total_in2___ + (double)result_[j].未結案;
                            //    Total_in3___ = Total_in3___ + (double)result_[j].已結案;
                            //    sheet1.GetRow(34).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(34).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(34).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            //}
                            //疾病管制署
                            if (result_[j].主辦單位 == "疾病管制署")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(35).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(35).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(35).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //食品藥物管理署
                            if (result_[j].主辦單位 == "食品藥物管理署")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(36).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(36).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(36).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //國民健康署
                            if (result_[j].主辦單位 == "國民健康署")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(37).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(37).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(37).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //中央健康保險署
                            if (result_[j].主辦單位 == "中央健康保險署")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(38).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(38).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(38).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //國家衛生研究院
                            if (result_[j].主辦單位 == "國家衛生研究院")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(39).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(39).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(39).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //國家中醫藥研究所
                            if (result_[j].主辦單位 == "國家中醫藥研究所")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(40).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(40).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(40).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //社會及家庭署
                            if (result_[j].主辦單位 == "社會及家庭署")
                            {
                                Total_out1___ = Total_out1___ + (double)result_[j].受理;
                                Total_out2___ = Total_out2___ + (double)result_[j].未結案;
                                Total_out3___ = Total_out3___ + (double)result_[j].已結案;
                                //sheet1.GetRow(41).CreateCell(cell1___).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(41).CreateCell(cell2___).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(41).CreateCell(cell3___).SetCellValue((double)result_[j].已結案);
                            }
                            //本部小計
                            sheet1.GetRow(6).CreateCell(cell1___).SetCellValue(Total_in1___ - Total_out1___);
                            sheet1.GetRow(6).CreateCell(cell2___).SetCellValue(Total_in2___ - Total_out2___);
                            sheet1.GetRow(6).CreateCell(cell3___).SetCellValue(Total_in3___ - Total_out3___);
                            //所屬小計
                            sheet1.GetRow(43).CreateCell(cell1___).SetCellValue(Total_out1___);
                            sheet1.GetRow(43).CreateCell(cell2___).SetCellValue(Total_out2___);
                            sheet1.GetRow(43).CreateCell(cell3___).SetCellValue(Total_out3___);
                            //總計
                            sheet1.GetRow(44).CreateCell(cell1___).SetCellValue(Total_in1___);
                            sheet1.GetRow(44).CreateCell(cell2___).SetCellValue(Total_in2___);
                            sheet1.GetRow(44).CreateCell(cell3___).SetCellValue(Total_in3___);

                        }
                        //合計
                        Total_sum1___ = Total_sum1___ + Total_in1___ - Total_out1___;
                        Total_sum2___ = Total_sum2___ + Total_in2___ - Total_out2___;
                        Total_sum3___ = Total_sum3___ + Total_in3___ - Total_out3___;
                        Total_sum4___ = Total_sum4___ + Total_out1___;
                        Total_sum5___ = Total_sum5___ + Total_out2___;
                        Total_sum6___ = Total_sum6___ + Total_out3___;

                        cell1___ = cell1___ + 3;
                        cell2___ = cell2___ + 3;
                        cell3___ = cell3___ + 3;

                        if (i == (ListD___.Count() - 1))
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1___, CellRange2___));
                            cells2___.CreateCell(cell0___).SetCellValue("合計");
                            cells2___.GetCell(cell0___).CellStyle = style;
                            cells3___.CreateCell(cell1___).SetCellValue("受理");
                            cells3___.CreateCell(cell2___).SetCellValue("未結案");
                            cells3___.CreateCell(cell3___).SetCellValue("已結案");
                            sheet1.GetRow(6).CreateCell(cell1___).SetCellValue(Total_sum1___);
                            sheet1.GetRow(6).CreateCell(cell2___).SetCellValue(Total_sum2___);
                            sheet1.GetRow(6).CreateCell(cell3___).SetCellValue(Total_sum3___);
                            sheet1.GetRow(43).CreateCell(cell1___).SetCellValue(Total_sum4___);
                            sheet1.GetRow(43).CreateCell(cell2___).SetCellValue(Total_sum5___);
                            sheet1.GetRow(43).CreateCell(cell3___).SetCellValue(Total_sum6___);
                            sheet1.GetRow(44).CreateCell(cell1___).SetCellValue(Total_sum1___ + Total_sum4___);
                            sheet1.GetRow(44).CreateCell(cell2___).SetCellValue(Total_sum2___ + Total_sum5___);
                            sheet1.GetRow(44).CreateCell(cell3___).SetCellValue(Total_sum3___ + Total_sum6___);
                            var linqStament = from p in result4
                                              group p by new { p.主辦單位 } into g
                                              select new { Code = g.Key.主辦單位, QTY1 = g.Sum(p => p.受理), QTY2 = g.Sum(p => p.未結案), QTY3 = g.Sum(p => p.已結案) };
                            foreach (var item in linqStament)
                            {
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                ////部長室
                                //if (item.Code == "部長室")
                                //{
                                //    sheet1.GetRow(7).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(7).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(7).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////主任秘書室
                                //if (item.Code == "主任秘書室")
                                //{
                                //    sheet1.GetRow(8).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(8).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(8).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////綜合規劃司
                                //if (item.Code == "綜合規劃司")
                                //{

                                //    sheet1.GetRow(9).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(9).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(9).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////社會保險司
                                //if (item.Code == "社會保險司")
                                //{

                                //    sheet1.GetRow(10).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(10).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(10).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////社會救助及社工司
                                //if (item.Code == "社會救助及社工司")
                                //{

                                //    sheet1.GetRow(11).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(11).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(11).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////保護服務司
                                //if (item.Code == "保護服務司")
                                //{

                                //    sheet1.GetRow(12).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(12).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(12).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////護理及健康照護司
                                //if (item.Code == "護理及健康照護司")
                                //{

                                //    sheet1.GetRow(13).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(13).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(13).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////醫事司
                                //if (item.Code == "醫事司")
                                //{

                                //    sheet1.GetRow(14).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(14).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(14).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////心理健康司
                                //if (item.Code == "心理健康司")
                                //{

                                //    sheet1.GetRow(15).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(15).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(15).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////中醫藥司
                                //if (item.Code == "中醫藥司")
                                //{

                                //    sheet1.GetRow(16).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(16).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(16).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////秘書處
                                //if (item.Code == "秘書處")
                                //{

                                //    sheet1.GetRow(17).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(17).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(17).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////人事處
                                //if (item.Code == "人事處")
                                //{

                                //    sheet1.GetRow(18).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(18).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(18).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////政風處
                                //if (item.Code == "政風處")
                                //{

                                //    sheet1.GetRow(19).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(19).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(19).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////會計處
                                //if (item.Code == "會計處")
                                //{

                                //    sheet1.GetRow(20).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(20).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(20).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////統計處
                                //if (item.Code == "統計處")
                                //{

                                //    sheet1.GetRow(21).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(21).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(21).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////資訊處
                                //if (item.Code == "資訊處")
                                //{

                                //    sheet1.GetRow(22).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(22).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(22).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////法規會
                                //if (item.Code == "法規會")
                                //{

                                //    sheet1.GetRow(23).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(23).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(23).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國際合作組
                                //if (item.Code == "國際合作組")
                                //{

                                //    sheet1.GetRow(24).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(24).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(24).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////附屬醫療及社會福利機構管理會
                                //if (item.Code == "附屬醫療及社會福利機構管理會")
                                //{

                                //    sheet1.GetRow(25).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(25).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(25).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險會
                                //if (item.Code == "全民健康保險會")
                                //{

                                //    sheet1.GetRow(26).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(26).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(26).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險爭議審議會
                                //if (item.Code == "全民健康保險爭議審議會")
                                //{

                                //    sheet1.GetRow(27).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(27).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(27).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////衛生福利人員訓練中心
                                //if (item.Code == "衛生福利人員訓練中心")
                                //{

                                //    sheet1.GetRow(28).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(28).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(28).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國民年金監理會
                                //if (item.Code == "國民年金監理會")
                                //{

                                //    sheet1.GetRow(29).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(29).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(29).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////科技發展組
                                //if (item.Code == "科技發展組")
                                //{

                                //    sheet1.GetRow(30).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(30).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(30).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////公共關係室
                                //if (item.Code == "公共關係室")
                                //{

                                //    sheet1.GetRow(31).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(31).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(31).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國會聯絡組
                                //if (item.Code == "國會聯絡組")
                                //{

                                //    sheet1.GetRow(32).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(32).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(32).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國家消除C肝辦公室
                                //if (item.Code == "國家消除C肝辦公室")
                                //{

                                //    sheet1.GetRow(33).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(33).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(33).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////長期照顧司
                                //if (item.Code == "長期照顧司")
                                //{

                                //    sheet1.GetRow(34).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(34).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(34).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////疾病管制署
                                //if (item.Code == "疾病管制署")
                                //{

                                //    sheet1.GetRow(35).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(35).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(35).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////食品藥物管理署
                                //if (item.Code == "食品藥物管理署")
                                //{

                                //    sheet1.GetRow(36).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(36).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(36).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國民健康署
                                //if (item.Code == "國民健康署")
                                //{

                                //    sheet1.GetRow(37).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(37).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(37).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////中央健康保險署
                                //if (item.Code == "中央健康保險署")
                                //{

                                //    sheet1.GetRow(38).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(38).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(38).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國家衛生研究院
                                //if (item.Code == "國家衛生研究院")
                                //{

                                //    sheet1.GetRow(39).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(39).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(39).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////國家中醫藥研究所
                                //if (item.Code == "國家中醫藥研究所")
                                //{

                                //    sheet1.GetRow(40).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(40).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(40).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                                ////社會及家庭署
                                //if (item.Code == "社會及家庭署")
                                //{

                                //    sheet1.GetRow(41).CreateCell(cell1___).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(41).CreateCell(cell2___).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(41).CreateCell(cell3___).SetCellValue((double)item.QTY3);
                                //}
                            }
                        }

                    }

                    break;
                case "5":
                    List<sp_AcceptExcel_5_Result> result5 = new List<sp_AcceptExcel_5_Result>();
                    using (MOHWEntities db = new MOHWEntities())
                    {
                        result5 = db.sp_AcceptExcel_5(vds, vde, ReportList, StatusList, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
                    }
                    //for (int i = 0; i < result5.Count(); i++)
                    //{
                    //    XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                    //    cells4.CreateCell(0).SetCellValue(result5[i].主辦單位);
                    //    cells4.CreateCell(1).SetCellValue(result5[i].VerifyDate);
                    //    cells4.CreateCell(2).SetCellValue((double)result5[i].受理);
                    //    cells4.CreateCell(3).SetCellValue((double)result5[i].未結案);
                    //    cells4.CreateCell(4).SetCellValue((double)result5[i].已結案);
                    //}
                    //找出有幾天資料
                    var ListD____ = (from p in result5 select p.VerifyDate).Distinct().ToList();
                    //欄位title
                    XSSFRow cells2____ = (XSSFRow)sheet1.CreateRow(4);
                    cells2____.CreateCell(0).SetCellValue("主辦單位");
                    //合併
                    sheet1.AddMergedRegion(new CellRangeAddress(4, 5, 0, 0));
                    //
                    sheet1.CreateRow(6).CreateCell(0).SetCellValue("本部小計");
                    sheet1.CreateRow(7).CreateCell(0).SetCellValue("部長室");
                    sheet1.CreateRow(8).CreateCell(0).SetCellValue("主秘室");
                    sheet1.CreateRow(9).CreateCell(0).SetCellValue("綜規司");
                    sheet1.CreateRow(10).CreateCell(0).SetCellValue("社保司");
                    sheet1.CreateRow(11).CreateCell(0).SetCellValue("社工司");
                    sheet1.CreateRow(12).CreateCell(0).SetCellValue("保護司");
                    sheet1.CreateRow(13).CreateCell(0).SetCellValue("照護司");
                    sheet1.CreateRow(14).CreateCell(0).SetCellValue("醫事司");
                    sheet1.CreateRow(15).CreateCell(0).SetCellValue("心理健康司");
                    sheet1.CreateRow(16).CreateCell(0).SetCellValue("口腔健康司");
                    sheet1.CreateRow(17).CreateCell(0).SetCellValue("中醫藥司");
                    sheet1.CreateRow(18).CreateCell(0).SetCellValue("秘書處");
                    sheet1.CreateRow(19).CreateCell(0).SetCellValue("人事處");
                    sheet1.CreateRow(20).CreateCell(0).SetCellValue("政風處");
                    sheet1.CreateRow(21).CreateCell(0).SetCellValue("會計處");
                    sheet1.CreateRow(22).CreateCell(0).SetCellValue("統計處");
                    sheet1.CreateRow(23).CreateCell(0).SetCellValue("資訊處");
                    sheet1.CreateRow(24).CreateCell(0).SetCellValue("法規會");
                    sheet1.CreateRow(25).CreateCell(0).SetCellValue("國際合作組");
                    sheet1.CreateRow(26).CreateCell(0).SetCellValue("醫福會");
                    sheet1.CreateRow(27).CreateCell(0).SetCellValue("健保會");
                    sheet1.CreateRow(28).CreateCell(0).SetCellValue("爭審會");
                    sheet1.CreateRow(29).CreateCell(0).SetCellValue("訓練中心");
                    sheet1.CreateRow(30).CreateCell(0).SetCellValue("監理會");
                    sheet1.CreateRow(31).CreateCell(0).SetCellValue("科發組");
                    sheet1.CreateRow(32).CreateCell(0).SetCellValue("公關室");
                    sheet1.CreateRow(33).CreateCell(0).SetCellValue("國會組");
                    sheet1.CreateRow(34).CreateCell(0).SetCellValue("C肝辦");
                    sheet1.CreateRow(35).CreateCell(0).SetCellValue("長照司");
                    sheet1.CreateRow(36).CreateCell(0).SetCellValue("疾管署");
                    sheet1.CreateRow(37).CreateCell(0).SetCellValue("食藥署");
                    sheet1.CreateRow(38).CreateCell(0).SetCellValue("國健署");
                    sheet1.CreateRow(39).CreateCell(0).SetCellValue("健保署");
                    sheet1.CreateRow(40).CreateCell(0).SetCellValue("國衛院");
                    sheet1.CreateRow(41).CreateCell(0).SetCellValue("中醫藥所");
                    sheet1.CreateRow(42).CreateCell(0).SetCellValue("社家署");
                    sheet1.CreateRow(43).CreateCell(0).SetCellValue("所屬機關小計");
                    sheet1.CreateRow(44).CreateCell(0).SetCellValue("總計");
                    int CellRange1____ = 1;
                    int CellRange2____ = 3;
                    int cell0____ = 1;
                    int cell1____ = 1;
                    int cell2____ = 2;
                    int cell3____ = 3;
                  
                    //合計
                    double Total_sum1____ = 0;
                    double Total_sum2____ = 0;
                    double Total_sum3____ = 0;
                    double Total_sum4____ = 0;
                    double Total_sum5____ = 0;
                    double Total_sum6____ = 0;


                    XSSFRow cells3____ = (XSSFRow)sheet1.CreateRow(5);
                    for (int i = 0; i < ListD____.Count(); i++)
                    {
                        //本部小計
                        double Total_in1____ = 0;
                        double Total_in2____ = 0;
                        double Total_in3____ = 0;
                        //所屬小計
                        double Total_out1____ = 0;
                        double Total_out2____ = 0;
                        double Total_out3____ = 0;
                        //總計
                        double Total1____ = 0;
                        double Total2____ = 0;
                        double Total3____ = 0;
                        //日期
                        sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1____, CellRange2____));
                        cells2____.CreateCell(cell0____).SetCellValue(ListD____[i].ToString());
                        cells2____.GetCell(cell0____).CellStyle = style;
                        cell0____ = cell0____ + 3;
                        CellRange1____ = CellRange1____ + 3;
                        CellRange2____ = CellRange2____ + 3;
                        //狀態
                        cells3____.CreateCell(cell1____).SetCellValue("受理");
                        cells3____.CreateCell(cell2____).SetCellValue("未結案");
                        cells3____.CreateCell(cell3____).SetCellValue("已結案");

                        List<sp_AcceptExcel_5_Result> result_ = new List<sp_AcceptExcel_5_Result>();
                        result_ = result5.Where(x => x.VerifyDate == ListD____[i].ToString()).ToList();
                        for (int j = 0; j < result_.Count(); j++)
                        {
                            Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            sheet1.GetRow(_caseManagementService.UnitToRow(result_[j].主辦單位)).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);

                            ////部長室
                            //if (result_[j].主辦單位 == "部長室")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(7).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(7).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(7).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////主任秘書室
                            //if (result_[j].主辦單位 == "主任秘書室")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(8).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(8).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(8).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////綜合規劃司
                            //if (result_[j].主辦單位 == "綜合規劃司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(9).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(9).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(9).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會保險司
                            //if (result_[j].主辦單位 == "社會保險司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(10).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(10).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(10).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////社會救助及社工司
                            //if (result_[j].主辦單位 == "社會救助及社工司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(11).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(11).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(11).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////保護服務司
                            //if (result_[j].主辦單位 == "保護服務司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(12).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(12).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(12).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////護理及健康照護司
                            //if (result_[j].主辦單位 == "護理及健康照護司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(13).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(13).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(13).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////醫事司
                            //if (result_[j].主辦單位 == "醫事司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(14).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(14).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(14).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////心理健康司
                            //if (result_[j].主辦單位 == "心理健康司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(15).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(15).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(15).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////中醫藥司
                            //if (result_[j].主辦單位 == "中醫藥司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(16).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(16).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(16).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////秘書處
                            //if (result_[j].主辦單位 == "秘書處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(17).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(17).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(17).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////人事處
                            //if (result_[j].主辦單位 == "人事處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(18).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(18).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(18).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////政風處
                            //if (result_[j].主辦單位 == "政風處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(19).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(19).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(19).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////會計處
                            //if (result_[j].主辦單位 == "會計處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(20).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(20).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(20).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////統計處
                            //if (result_[j].主辦單位 == "統計處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(21).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(21).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(21).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////資訊處
                            //if (result_[j].主辦單位 == "資訊處")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(22).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(22).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(22).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////法規會
                            //if (result_[j].主辦單位 == "法規會")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(23).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(23).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(23).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國際合作組
                            //if (result_[j].主辦單位 == "國際合作組")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(24).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(24).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(24).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////附屬醫療及社會福利機構管理會
                            //if (result_[j].主辦單位 == "附屬醫療及社會福利機構管理會")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(25).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(25).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(25).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險會
                            //if (result_[j].主辦單位 == "全民健康保險會")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(26).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(26).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(26).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////全民健康保險爭議審議會
                            //if (result_[j].主辦單位 == "全民健康保險爭議審議會")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(27).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(27).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(27).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////衛生福利人員訓練中心
                            //if (result_[j].主辦單位 == "衛生福利人員訓練中心")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(28).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(28).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(28).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國民年金監理會
                            //if (result_[j].主辦單位 == "國民年金監理會")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(29).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(29).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(29).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////科技發展組
                            //if (result_[j].主辦單位 == "科技發展組")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(30).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(30).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(30).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////公共關係室
                            //if (result_[j].主辦單位 == "公共關係室")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(31).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(31).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(31).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國會聯絡組
                            //if (result_[j].主辦單位 == "國會聯絡組")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(32).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(32).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(32).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////國家消除C肝辦公室
                            //if (result_[j].主辦單位 == "國家消除C肝辦公室")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(33).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(33).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(33).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            ////長期照顧司
                            //if (result_[j].主辦單位 == "長期照顧司")
                            //{
                            //    Total_in1____ = Total_in1____ + (double)result_[j].受理;
                            //    Total_in2____ = Total_in2____ + (double)result_[j].未結案;
                            //    Total_in3____ = Total_in3____ + (double)result_[j].已結案;
                            //    sheet1.GetRow(34).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                            //    sheet1.GetRow(34).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                            //    sheet1.GetRow(34).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            //}
                            //疾病管制署
                            if (result_[j].主辦單位 == "疾病管制署")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(35).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(35).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(35).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //食品藥物管理署
                            if (result_[j].主辦單位 == "食品藥物管理署")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(36).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(36).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(36).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //國民健康署
                            if (result_[j].主辦單位 == "國民健康署")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(37).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(37).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(37).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //中央健康保險署
                            if (result_[j].主辦單位 == "中央健康保險署")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(38).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(38).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(38).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //國家衛生研究院
                            if (result_[j].主辦單位 == "國家衛生研究院")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(39).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(39).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(39).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //國家中醫藥研究所
                            if (result_[j].主辦單位 == "國家中醫藥研究所")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(40).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(40).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(40).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //社會及家庭署
                            if (result_[j].主辦單位 == "社會及家庭署")
                            {
                                Total_out1____ = Total_out1____ + (double)result_[j].受理;
                                Total_out2____ = Total_out2____ + (double)result_[j].未結案;
                                Total_out3____ = Total_out3____ + (double)result_[j].已結案;
                                //sheet1.GetRow(41).CreateCell(cell1____).SetCellValue((double)result_[j].受理);
                                //sheet1.GetRow(41).CreateCell(cell2____).SetCellValue((double)result_[j].未結案);
                                //sheet1.GetRow(41).CreateCell(cell3____).SetCellValue((double)result_[j].已結案);
                            }
                            //本部小計
                            sheet1.GetRow(6).CreateCell(cell1____).SetCellValue(Total_in1____ - Total_out1____);
                            sheet1.GetRow(6).CreateCell(cell2____).SetCellValue(Total_in2____ - Total_out2____);
                            sheet1.GetRow(6).CreateCell(cell3____).SetCellValue(Total_in3____ - Total_out3____);
                            //所屬小計
                            sheet1.GetRow(43).CreateCell(cell1____).SetCellValue(Total_out1____);
                            sheet1.GetRow(43).CreateCell(cell2____).SetCellValue(Total_out2____);
                            sheet1.GetRow(43).CreateCell(cell3____).SetCellValue(Total_out3____);
                            //總計
                            sheet1.GetRow(44).CreateCell(cell1____).SetCellValue(Total_in1____);
                            sheet1.GetRow(44).CreateCell(cell2____).SetCellValue(Total_in2____);
                            sheet1.GetRow(44).CreateCell(cell3____).SetCellValue(Total_in3____);

                        }
                        //合計
                        Total_sum1____ = Total_sum1____ + Total_in1____;
                        Total_sum2____ = Total_sum2____ + Total_in2____;
                        Total_sum3____ = Total_sum3____ + Total_in3____;
                        Total_sum4____ = Total_sum4____ + Total_out1____;
                        Total_sum5____ = Total_sum5____ + Total_out2____;
                        Total_sum6____ = Total_sum6____ + Total_out3____;

                        cell1____ = cell1____ + 3;
                        cell2____ = cell2____ + 3;
                        cell3____ = cell3____ + 3;

                        if (i == (ListD____.Count() - 1))
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(4, 4, CellRange1____, CellRange2____));
                            cells2____.CreateCell(cell0____).SetCellValue("合計");
                            cells2____.GetCell(cell0____).CellStyle = style;
                            cells3____.CreateCell(cell1____).SetCellValue("受理");
                            cells3____.CreateCell(cell2____).SetCellValue("未結案");
                            cells3____.CreateCell(cell3____).SetCellValue("已結案");
                            sheet1.GetRow(6).CreateCell(cell1____).SetCellValue(Total_sum1____);
                            sheet1.GetRow(6).CreateCell(cell2____).SetCellValue(Total_sum2____);
                            sheet1.GetRow(6).CreateCell(cell3____).SetCellValue(Total_sum3____);
                            sheet1.GetRow(43).CreateCell(cell1____).SetCellValue(Total_sum4____);
                            sheet1.GetRow(43).CreateCell(cell2____).SetCellValue(Total_sum5____);
                            sheet1.GetRow(43).CreateCell(cell3____).SetCellValue(Total_sum6____);
                            sheet1.GetRow(44).CreateCell(cell1____).SetCellValue(Total_sum1____ + Total_sum4____);
                            sheet1.GetRow(44).CreateCell(cell2____).SetCellValue(Total_sum2____ + Total_sum5____);
                            sheet1.GetRow(44).CreateCell(cell3____).SetCellValue(Total_sum3____ + Total_sum6____);
                            var linqStament = from p in result5
                                              group p by new { p.主辦單位 } into g
                                              select new { Code = g.Key.主辦單位, QTY1 = g.Sum(p => p.受理), QTY2 = g.Sum(p => p.未結案), QTY3 = g.Sum(p => p.已結案) };
                            foreach (var item in linqStament)
                            {
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                sheet1.GetRow(_caseManagementService.UnitToRow(item.Code)).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                ////部長室
                                //if (item.Code == "部長室")
                                //{
                                //    sheet1.GetRow(7).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(7).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(7).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////主任秘書室
                                //if (item.Code == "主任秘書室")
                                //{
                                //    sheet1.GetRow(8).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(8).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(8).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////綜合規劃司
                                //if (item.Code == "綜合規劃司")
                                //{

                                //    sheet1.GetRow(9).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(9).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(9).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////社會保險司
                                //if (item.Code == "社會保險司")
                                //{

                                //    sheet1.GetRow(10).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(10).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(10).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////社會救助及社工司
                                //if (item.Code == "社會救助及社工司")
                                //{

                                //    sheet1.GetRow(11).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(11).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(11).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////保護服務司
                                //if (item.Code == "保護服務司")
                                //{

                                //    sheet1.GetRow(12).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(12).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(12).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////護理及健康照護司
                                //if (item.Code == "護理及健康照護司")
                                //{

                                //    sheet1.GetRow(13).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(13).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(13).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////醫事司
                                //if (item.Code == "醫事司")
                                //{

                                //    sheet1.GetRow(14).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(14).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(14).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////心理健康司
                                //if (item.Code == "心理健康司")
                                //{

                                //    sheet1.GetRow(15).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(15).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(15).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////中醫藥司
                                //if (item.Code == "中醫藥司")
                                //{

                                //    sheet1.GetRow(16).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(16).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(16).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////秘書處
                                //if (item.Code == "秘書處")
                                //{

                                //    sheet1.GetRow(17).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(17).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(17).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////人事處
                                //if (item.Code == "人事處")
                                //{

                                //    sheet1.GetRow(18).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(18).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(18).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////政風處
                                //if (item.Code == "政風處")
                                //{

                                //    sheet1.GetRow(19).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(19).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(19).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////會計處
                                //if (item.Code == "會計處")
                                //{

                                //    sheet1.GetRow(20).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(20).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(20).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////統計處
                                //if (item.Code == "統計處")
                                //{

                                //    sheet1.GetRow(21).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(21).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(21).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////資訊處
                                //if (item.Code == "資訊處")
                                //{

                                //    sheet1.GetRow(22).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(22).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(22).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////法規會
                                //if (item.Code == "法規會")
                                //{

                                //    sheet1.GetRow(23).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(23).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(23).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國際合作組
                                //if (item.Code == "國際合作組")
                                //{

                                //    sheet1.GetRow(24).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(24).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(24).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////附屬醫療及社會福利機構管理會
                                //if (item.Code == "附屬醫療及社會福利機構管理會")
                                //{

                                //    sheet1.GetRow(25).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(25).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(25).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險會
                                //if (item.Code == "全民健康保險會")
                                //{

                                //    sheet1.GetRow(26).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(26).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(26).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////全民健康保險爭議審議會
                                //if (item.Code == "全民健康保險爭議審議會")
                                //{

                                //    sheet1.GetRow(27).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(27).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(27).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////衛生福利人員訓練中心
                                //if (item.Code == "衛生福利人員訓練中心")
                                //{

                                //    sheet1.GetRow(28).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(28).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(28).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國民年金監理會
                                //if (item.Code == "國民年金監理會")
                                //{

                                //    sheet1.GetRow(29).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(29).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(29).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////科技發展組
                                //if (item.Code == "科技發展組")
                                //{

                                //    sheet1.GetRow(30).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(30).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(30).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////公共關係室
                                //if (item.Code == "公共關係室")
                                //{

                                //    sheet1.GetRow(31).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(31).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(31).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國會聯絡組
                                //if (item.Code == "國會聯絡組")
                                //{

                                //    sheet1.GetRow(32).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(32).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(32).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國家消除C肝辦公室
                                //if (item.Code == "國家消除C肝辦公室")
                                //{

                                //    sheet1.GetRow(33).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(33).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(33).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////長期照顧司
                                //if (item.Code == "長期照顧司")
                                //{

                                //    sheet1.GetRow(34).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(34).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(34).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////疾病管制署
                                //if (item.Code == "疾病管制署")
                                //{

                                //    sheet1.GetRow(35).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(35).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(35).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////食品藥物管理署
                                //if (item.Code == "食品藥物管理署")
                                //{

                                //    sheet1.GetRow(36).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(36).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(36).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國民健康署
                                //if (item.Code == "國民健康署")
                                //{

                                //    sheet1.GetRow(37).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(37).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(37).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////中央健康保險署
                                //if (item.Code == "中央健康保險署")
                                //{

                                //    sheet1.GetRow(38).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(38).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(38).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國家衛生研究院
                                //if (item.Code == "國家衛生研究院")
                                //{

                                //    sheet1.GetRow(39).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(39).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(39).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////國家中醫藥研究所
                                //if (item.Code == "國家中醫藥研究所")
                                //{

                                //    sheet1.GetRow(40).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(40).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(40).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                                ////社會及家庭署
                                //if (item.Code == "社會及家庭署")
                                //{

                                //    sheet1.GetRow(41).CreateCell(cell1____).SetCellValue((double)item.QTY1);
                                //    sheet1.GetRow(41).CreateCell(cell2____).SetCellValue((double)item.QTY2);
                                //    sheet1.GetRow(41).CreateCell(cell3____).SetCellValue((double)item.QTY3);
                                //}
                            }
                        }

                    }
                    break;
                default:
                    break;
            }


            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        /// <summary>
        /// 陳情案件年度統計
        /// </summary>
        /// <returns>View</returns>
        public ActionResult AppealReport()
        {
            //統計類型
            CaseQueryViewModel caseQueryViewModel = new CaseQueryViewModel();
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2015", Value = "2015" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2016", Value = "2016" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2017", Value = "2017" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2018", Value = "2018" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2019", Value = "2019" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2020", Value = "2020" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2021", Value = "2021" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2022", Value = "2022" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2023", Value = "2023" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2024", Value = "2024" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2025", Value = "2025" });
            caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2026", Value = "2026" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2027", Value = "2027" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2028", Value = "2028" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2029", Value = "2029" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2030", Value = "2030" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2031", Value = "2031" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2032", Value = "2032" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2033", Value = "2033" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2034", Value = "2034" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2035", Value = "2035" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2036", Value = "2036" });
            //caseQueryViewModel.ReportStatus.Add(new SelectListItem() { Text = "2037", Value = "2037" });
            return View(caseQueryViewModel);
        }
        /// <summary>
        /// 陳情案件年度統計匯出Excel
        /// 查詢(年月區間)匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>File</returns>
        public ActionResult AppealExcel(CaseQueryViewModel caseQueryViewModel)
        {
            //取得資料
            string yearstr = caseQueryViewModel.ReportList;
            string vds = yearstr + "/01/01";
            string vde = yearstr + "/12/31";
            //string vds = caseQueryViewModel.VerifyDateS;
            //string vde = caseQueryViewModel.VerifyDateE;

            //載入模板檔案路徑  
            string TempletFileName = Server.MapPath("../App_Data/ExcelTemplate/部長信箱年度統計表.xlsx");
            XSSFWorkbook wk = null;
            using (FileStream fs = System.IO.File.OpenRead(TempletFileName))
            {
                //把xlsx檔案讀入workbook變數裡，之後就可以關閉了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet1 = (XSSFSheet)wk.GetSheetAt(0);
            int nRow = 5;//開始插入的行（第二行）
            sheet1.GetRow(1).GetCell(0).SetCellValue("查詢月份：" + vds + " ~ " + vde);
            //開始塞資料
            List<sp_AppealExcel_Year_Result> result = new List<sp_AppealExcel_Year_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result = db.sp_AppealExcel_Year(vds, vde).ToList();
            }
            sheet1.GetRow(5).GetCell(1).SetCellValue((double)result[0].行政興革之建議);
            sheet1.GetRow(5).GetCell(2).SetCellValue((double)result[0].行政法令之查詢);
            sheet1.GetRow(5).GetCell(3).SetCellValue((double)result[0].行政違失之舉發);
            sheet1.GetRow(5).GetCell(4).SetCellValue((double)result[0].行政權益之維護);
            sheet1.GetRow(5).GetCell(5).SetCellValue((double)result[0].行政法令之釋疑);
            sheet1.GetRow(5).GetCell(6).SetCellValue((double)result[0].對主管業務資料之索取及查詢);
            sheet1.GetRow(5).GetCell(7).SetCellValue((double)result[0].其他);
            List<sp_AppealExcel_2_Result> result2 = new List<sp_AppealExcel_2_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result2 = db.sp_AppealExcel_2(vds, vde).ToList();
            }
            sheet1.GetRow(10).GetCell(1).SetCellValue((double)result2[0].轉請權責機關處理);
            sheet1.GetRow(10).GetCell(2).SetCellValue((double)result2[0].自行回復);
            sheet1.GetRow(10).GetCell(3).SetCellValue((double)result2[0].不予受理);

            List<sp_AppealExcel_3_Result> result3 = new List<sp_AppealExcel_3_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result3 = db.sp_AppealExcel_3(vds, vde).ToList();
            }
            sheet1.GetRow(15).GetCell(1).SetCellValue((double)result3[0].電子郵件);

            List<sp_AppealExcel_4_Result> result4 = new List<sp_AppealExcel_4_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result4 = db.sp_AppealExcel_4(vds, vde).ToList();
            }
            sheet1.GetRow(20).GetCell(1).SetCellValue((double)result4[0].總統府函轉);
            sheet1.GetRow(20).GetCell(2).SetCellValue((double)result4[0].行政院函轉);
            sheet1.GetRow(20).GetCell(3).SetCellValue((double)result4[0].機關自行受理);
            sheet1.GetRow(20).GetCell(4).SetCellValue((double)result4[0].其它);

            List<sp_AppealExcel_5_Result> result5 = new List<sp_AppealExcel_5_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result5 = db.sp_AppealExcel_5(vds, vde).ToList();
            }
            double r5_1 = 0;
            double r5_2 = 0;
            double r5_3 = 0;
            double r5_4 = 0;
            for (int i = 0; i < result5.Count(); i++)
            {
                if (result5[i].C6天以內 != 0)
                {
                    r5_1 = (double)result5[i].C6天以內;
                }
                if (result5[i].C6_15天 != 0)
                {
                    r5_2 = (double)result5[i].C6_15天;
                }
                if (result5[i].C15_30天 != 0)
                {
                    r5_3 = (double)result5[i].C15_30天;
                }
                if (result5[i].C30天以上 != 0)
                {
                    r5_4 = (double)result5[i].C30天以上;
                }
            }
            sheet1.GetRow(25).GetCell(1).SetCellValue(r5_1);
            sheet1.GetRow(25).GetCell(2).SetCellValue(r5_2);
            sheet1.GetRow(25).GetCell(3).SetCellValue(r5_3);
            sheet1.GetRow(25).GetCell(4).SetCellValue(r5_4);

            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        /// <summary>
        /// 部長信箱時效統計表
        /// </summary>
        /// <returns>View</returns>
        public ActionResult ObsoleteReport()
        {
            return View();
        }
        /// <summary>
        /// 部長信箱時效統計表匯出Excel
        /// 查詢(年月區間)匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>File</returns>
        public ActionResult ObsoleteExcel(CaseQueryViewModel caseQueryViewModel)
        {
            //取得資料
            string vds = caseQueryViewModel.VerifyDateS;
            string vde = caseQueryViewModel.VerifyDateE;

            //載入模板檔案路徑  
            string TempletFileName = Server.MapPath("../App_Data/ExcelTemplate/部長信箱時效統計表.xlsx");
            XSSFWorkbook wk = null;
            using (FileStream fs = System.IO.File.OpenRead(TempletFileName))
            {
                //把xlsx檔案讀入workbook變數裡，之後就可以關閉了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet1 = (XSSFSheet)wk.GetSheetAt(0);
            int nRow = 3;//開始插入的行（第四行）
            sheet1.GetRow(0).GetCell(2).SetCellValue("部長信箱時效統計表" + " 統計日期:" + vds + " ~ " + vde);
            //開始塞資料
            List<sp_OverdueExcel_Result> result = new List<sp_OverdueExcel_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result = db.sp_OverdueExcel(vds, vde).ToList();
            }

            for (int i = 0; i < result.Count(); i++)
            {
                //人事處
                if (result[i].TopUnit == 20)
                {
                    sheet1.GetRow(3).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(3).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(3).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(3).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(3).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(3).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(3).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //中醫藥司
                if (result[i].TopUnit == 18)
                {
                    sheet1.GetRow(4).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(4).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(4).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(4).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(4).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(4).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(4).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //公共關係室
                if (result[i].TopUnit == 31)
                {
                    sheet1.GetRow(5).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(5).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(5).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(5).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(5).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(5).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(5).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //心理健康司
                if (result[i].TopUnit == 17)
                {
                    sheet1.GetRow(6).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(6).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(6).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(6).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(6).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(6).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(6).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //主任秘書室
                if (result[i].TopUnit == 10)
                {
                    sheet1.GetRow(7).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(7).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(7).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(7).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(7).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(7).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(7).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //全民健康保險爭議審議會
                if (result[i].TopUnit == 28)
                {
                    sheet1.GetRow(8).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(8).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(8).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(8).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(8).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(8).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(8).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //全民健康保險會
                if (result[i].TopUnit == 27)
                {
                    sheet1.GetRow(9).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(9).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(9).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(9).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(9).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(9).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(9).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //法規會
                if (result[i].TopUnit == 25)
                {
                    sheet1.GetRow(10).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(10).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(10).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(10).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(10).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(10).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(10).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //社會保險司
                if (result[i].TopUnit == 12)
                {
                    sheet1.GetRow(11).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(11).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(11).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(11).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(11).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(11).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(11).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //社會救助及社工司
                if (result[i].TopUnit == 13)
                {
                    sheet1.GetRow(12).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(12).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(12).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(12).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(12).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(12).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(12).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //長期照顧司
                if (result[i].TopUnit == 34)
                {
                    sheet1.GetRow(13).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(13).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(13).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(13).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(13).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(13).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(13).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //附屬醫療及社會福利機構管理會
                if (result[i].TopUnit == 8)
                {
                    sheet1.GetRow(14).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(14).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(14).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(14).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(14).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(14).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(14).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //保護服務司
                if (result[i].TopUnit == 14)
                {
                    sheet1.GetRow(15).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(15).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(15).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(15).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(15).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(15).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(15).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //政風處
                if (result[i].TopUnit == 21)
                {
                    sheet1.GetRow(16).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(16).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(16).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(16).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(16).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(16).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(16).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //科技發展組
                if (result[i].TopUnit == 30)
                {
                    sheet1.GetRow(17).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(17).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(17).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(17).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(17).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(17).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(17).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //秘書處
                if (result[i].TopUnit == 19)
                {
                    sheet1.GetRow(18).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(18).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(18).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(18).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(18).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(18).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(18).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國民年金監理會
                if (result[i].TopUnit == 29)
                {
                    sheet1.GetRow(19).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(19).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(19).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(19).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(19).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(19).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(19).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國會聯絡組
                if (result[i].TopUnit == 32)
                {
                    sheet1.GetRow(20).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(20).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(20).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(20).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(20).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(20).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(20).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國際合作組
                if (result[i].TopUnit == 26)
                {
                    sheet1.GetRow(21).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(21).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(21).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(21).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(21).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(21).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(21).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //統計處
                if (result[i].TopUnit == 23)
                {
                    sheet1.GetRow(22).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(22).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(22).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(22).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(22).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(22).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(22).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //會計處
                if (result[i].TopUnit == 22)
                {
                    sheet1.GetRow(23).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(23).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(23).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(23).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(23).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(23).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(23).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //資訊處
                if (result[i].TopUnit == 24)
                {
                    sheet1.GetRow(24).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(24).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(24).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(24).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(24).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(24).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(24).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //綜合規劃司
                if (result[i].TopUnit == 11)
                {
                    sheet1.GetRow(25).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(25).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(25).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(25).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(25).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(25).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(25).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //衛生福利人員訓練中心
                if (result[i].TopUnit == 7)
                {
                    sheet1.GetRow(26).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(26).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(26).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(26).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(26).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(26).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(26).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //醫事司
                if (result[i].TopUnit == 16)
                {
                    sheet1.GetRow(27).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(27).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(27).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(27).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(27).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(27).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(27).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //護理及健康照護司
                if (result[i].TopUnit == 15)
                {
                    sheet1.GetRow(28).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(28).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(28).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(28).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(28).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(28).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(28).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國家消除C肝辦公室
                if (result[i].TopUnit == 33)
                {
                    sheet1.GetRow(29).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(29).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(29).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(29).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(29).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(29).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(29).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //疾病管制署
                if (result[i].TopUnit == 4)
                {
                    sheet1.GetRow(30).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(30).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(30).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(30).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(30).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(30).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(30).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //食品藥物管理署
                if (result[i].TopUnit == 2)
                {
                    sheet1.GetRow(31).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(31).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(31).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(31).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(31).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(31).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(31).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國民健康署
                if (result[i].TopUnit == 5)
                {
                    sheet1.GetRow(32).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(32).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(32).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(32).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(32).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(32).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(32).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //中央健康保險署
                if (result[i].TopUnit == 6)
                {
                    sheet1.GetRow(33).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(33).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(33).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(33).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(33).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(33).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(33).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國家衛生研究院
                if (result[i].TopUnit == 130)
                {
                    sheet1.GetRow(34).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(34).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(34).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(34).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(34).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(34).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(34).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //國家中醫藥研究所
                if (result[i].TopUnit == 210)
                {
                    sheet1.GetRow(35).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(35).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(35).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(35).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(35).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(35).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(35).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
                //社會及家庭署
                if (result[i].TopUnit == 3)
                {
                    sheet1.GetRow(36).GetCell(1).SetCellValue((double)result[i].本月份新收案件數);
                    sheet1.GetRow(36).GetCell(2).SetCellValue((double)result[i].待辦案件數);
                    sheet1.GetRow(36).GetCell(4).SetCellValue((double)result[i].依限辦結案件);
                    sheet1.GetRow(36).GetCell(6).SetCellValue((double)result[i].逾限辦結);
                    sheet1.GetRow(36).GetCell(8).SetCellValue((double)result[i].辦結案件合計);
                    sheet1.GetRow(36).GetCell(13).SetCellValue((double)result[i].未逾辦理期限案件數);
                    sheet1.GetRow(36).GetCell(14).SetCellValue((double)result[i].已逾辦理期限案件數);
                }
            }
            //sheet1.GetRow(3).GetCell(1).SetCellValue(99);
            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        /// <summary>
        /// 受理案件逾期天數統計表
        /// </summary>
        /// <returns>View</returns>
        public ActionResult OverdueReport()
        {
            return View();
        }
        /// <summary>
        /// 受理案件逾期天數統計表匯出Excel
        /// 查詢(年月日區間)匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>View</returns>
        public ActionResult OverdueExcel(CaseQueryViewModel caseQueryViewModel)
        {
            //取得資料
            string vds = caseQueryViewModel.VerifyDateS;
            string vde = caseQueryViewModel.VerifyDateE;
            _currentUser = Session["User"] as User;

            //載入模板檔案路徑  
            string TempletFileName = Server.MapPath("../App_Data/ExcelTemplate/受理案件逾期天數統計表.xlsx");
            XSSFWorkbook wk = null;
            using (FileStream fs = System.IO.File.OpenRead(TempletFileName))
            {
                //把xlsx檔案讀入workbook變數裡，之後就可以關閉了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet1 = (XSSFSheet)wk.GetSheetAt(0);
            int nRow = 5;//開始插入的行（第六行）
            sheet1.GetRow(1).GetCell(0).SetCellValue("列印時間：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            sheet1.GetRow(1).GetCell(2).SetCellValue("列印人員：" + _currentUser.UserDetail.UserName);
            sheet1.GetRow(2).GetCell(0).SetCellValue("查詢條件：受理日期期間：" + vds + " ~ " + vde + "，來源管道：部長信箱");

            //sheet1.GetRow(5).GetCell(1).SetCellValue(99);
            //開始塞資料
            List<sp_ObsoleteExcel_Result> result = new List<sp_ObsoleteExcel_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result = db.sp_ObsoleteExcel(vds, vde).ToList();
            }
            //sheet1.GetRow(5).GetCell(1).SetCellValue(99);
            for (int i = 0; i < result.Count(); i++)
            {
                //部長室
                if (result[i].TopUnit == 9)
                {

                    sheet1.GetRow(5).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(3).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(3).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(3).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(3).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //主任秘書室
                if (result[i].TopUnit == 10)
                {

                    sheet1.GetRow(6).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(6).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(6).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(6).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(6).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //綜合規劃司
                if (result[i].TopUnit == 11)
                {

                    sheet1.GetRow(7).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(7).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(7).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(7).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(7).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //社會保險司
                if (result[i].TopUnit == 12)
                {

                    sheet1.GetRow(8).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(8).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(8).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(8).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(8).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //社會救助及社工司
                if (result[i].TopUnit == 13)
                {

                    sheet1.GetRow(9).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(9).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(9).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(9).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(9).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //保護服務司
                if (result[i].TopUnit == 14)
                {

                    sheet1.GetRow(10).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(10).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(10).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(10).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(10).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //護理及健康照護司
                if (result[i].TopUnit == 15)
                {

                    sheet1.GetRow(11).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(11).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(11).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(11).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(11).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //醫事司
                if (result[i].TopUnit == 16)
                {

                    sheet1.GetRow(12).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(12).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(12).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(12).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(12).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //心理健康司
                if (result[i].TopUnit == 17)
                {

                    sheet1.GetRow(13).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(13).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(13).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(13).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(13).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //中醫藥司
                if (result[i].TopUnit == 18)
                {

                    sheet1.GetRow(14).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(14).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(14).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(14).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(14).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //秘書處
                if (result[i].TopUnit == 19)
                {

                    sheet1.GetRow(15).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(15).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(15).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(15).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(15).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //人事處
                if (result[i].TopUnit == 20)
                {

                    sheet1.GetRow(16).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(16).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(16).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(16).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(16).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //政風處
                if (result[i].TopUnit == 21)
                {

                    sheet1.GetRow(17).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(17).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(17).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(17).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(17).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //會計處
                if (result[i].TopUnit == 22)
                {

                    sheet1.GetRow(18).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(18).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(18).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(18).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(18).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //統計處
                if (result[i].TopUnit == 23)
                {

                    sheet1.GetRow(19).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(19).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(19).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(19).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(19).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //資訊處
                if (result[i].TopUnit == 24)
                {

                    sheet1.GetRow(20).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(20).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(20).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(20).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(20).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //法規會
                if (result[i].TopUnit == 25)
                {

                    sheet1.GetRow(21).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(21).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(21).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(21).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(21).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國際合作組
                if (result[i].TopUnit == 26)
                {

                    sheet1.GetRow(22).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(22).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(22).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(22).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(22).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //附屬醫療及社會福利機構管理會
                if (result[i].TopUnit == 8)
                {

                    sheet1.GetRow(23).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(23).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(23).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(23).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(23).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //全民健康保險會
                if (result[i].TopUnit == 27)
                {

                    sheet1.GetRow(24).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(24).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(24).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(24).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(24).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //全民健康保險爭議審議會
                if (result[i].TopUnit == 28)
                {

                    sheet1.GetRow(25).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(25).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(25).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(25).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(25).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //衛生福利人員訓練中心
                if (result[i].TopUnit == 7)
                {

                    sheet1.GetRow(26).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(26).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(26).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(26).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(26).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國民年金監理會
                if (result[i].TopUnit == 29)
                {

                    sheet1.GetRow(27).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(27).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(27).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(27).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(27).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //科技發展組
                if (result[i].TopUnit == 30)
                {

                    sheet1.GetRow(28).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(28).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(28).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(28).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(28).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //公共關係室
                if (result[i].TopUnit == 31)
                {

                    sheet1.GetRow(29).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(29).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(29).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(29).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(29).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國會聯絡組
                if (result[i].TopUnit == 32)
                {

                    sheet1.GetRow(30).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(30).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(30).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(30).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(30).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國家消除C肝辦公室
                if (result[i].TopUnit == 33)
                {

                    sheet1.GetRow(31).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(31).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(31).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(31).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(31).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //長期照顧司
                if (result[i].TopUnit == 34)
                {

                    sheet1.GetRow(32).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(32).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(32).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(32).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(32).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //疾病管制署
                if (result[i].TopUnit == 4)
                {

                    sheet1.GetRow(33).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(33).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(33).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(33).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(33).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //食品藥物管理署
                if (result[i].TopUnit == 2)
                {

                    sheet1.GetRow(34).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(34).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(34).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(34).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(34).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國民健康署
                if (result[i].TopUnit == 5)
                {

                    sheet1.GetRow(35).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(35).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(35).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(35).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(35).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //中央健康保險署
                if (result[i].TopUnit == 6)
                {

                    sheet1.GetRow(36).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(36).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(36).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(36).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(36).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國家衛生研究院
                if (result[i].TopUnit == 130)
                {

                    sheet1.GetRow(37).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(37).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(37).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(37).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(37).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //國家中醫藥研究所
                if (result[i].TopUnit == 210)
                {

                    sheet1.GetRow(38).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(38).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(38).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(38).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(38).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }
                //社會及家庭署
                if (result[i].TopUnit == 3)
                {

                    sheet1.GetRow(39).GetCell(1).SetCellValue((double)result[i].C1_3日);
                    sheet1.GetRow(39).GetCell(2).SetCellValue((double)result[i].C4_6日);
                    sheet1.GetRow(39).GetCell(3).SetCellValue((double)result[i].C7_15日);
                    sheet1.GetRow(39).GetCell(4).SetCellValue((double)result[i].C16天以上);
                    sheet1.GetRow(39).GetCell(6).SetCellValue((double)result[i].收文總件數);

                }

            }
            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        /// <summary>
        /// 滿意度問卷明細表
        /// </summary>
        /// <returns>View</returns>
        public ActionResult SatisfactionDetailReport()
        {
            //頁面初始化
            _currentUser = Session["User"] as User;
            bool isSupervisor = false;
            foreach (var item in _currentUser.rolesList)
            {
                if (item.SerialNo == 155)
                {
                    isSupervisor = true;
                }
            }
            string memberType = _currentUser.UserDetail.Internal == "Y" ? "1" : "2"; //部內 = 1 , 所屬 = 2;
            List<SysSubCode> CaseSourceList = new CommonService().GetCaseSourceList();
            List<Organization> caseOrganizerList = new CommonService().GetCaseOrganizerList(memberType, isSupervisor);
            CaseQueryViewModel caseQueryViewModel = new CaseQueryViewModel();
            string NowD = DateTime.Now.ToString("yyyy/MM/dd");
            caseQueryViewModel.VerifyDateS = NowD;
            caseQueryViewModel.VerifyDateE = NowD;
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
            //議題類別
            caseQueryViewModel.StatisticCaseTypeList = _caseManagementService.GetStatCaseList();
            //承辦單位
            foreach (var caseOrganizer in caseOrganizerList.Select(x => new { x.SerialNo, x.ZHName }).OrderBy(x => x.SerialNo))
            {
                caseQueryViewModel.CaseOrganizerList.Add(new SelectListItem() { Text = caseOrganizer.ZHName, Value = caseOrganizer.SerialNo.ToString() });
            }
            return View(caseQueryViewModel);
        }
        /// <summary>
        /// 滿意度問卷明細表匯出Excel
        /// 查詢(年月日區間)匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>View</returns>
        public ActionResult SatisfactionDetailExcel(CaseQueryViewModel caseQueryViewModel)
        {
            _currentUser = Session["User"] as User;
            //取得資料
            string vds = caseQueryViewModel.VerifyDateS;
            string vde = caseQueryViewModel.VerifyDateE;
            string dds = caseQueryViewModel.DeadLineS;
            string dde = caseQueryViewModel.DeadLineE;
            string AppealCategory = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.CaseTypeList);
            string PetitionTypeList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.PetitionTypeList);
            string SourceList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.SourceList);
            string CaseOrganizerList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.OrganizerList);
            //載入模板檔案路徑  
            string TempletFileName = Server.MapPath("../App_Data/ExcelTemplate/滿意度問卷明細表.xlsx");
            XSSFWorkbook wk = null;
            using (FileStream fs = System.IO.File.OpenRead(TempletFileName))
            {
                //把xlsx檔案讀入workbook變數裡，之後就可以關閉了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet1 = (XSSFSheet)wk.GetSheetAt(0);
            int nRow = 4;//開始插入的行（第五行）
            sheet1.GetRow(1).GetCell(0).SetCellValue("列印時間：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            sheet1.GetRow(1).GetCell(8).SetCellValue("列印人員：" + _currentUser.UserDetail.UserName);
            sheet1.GetRow(2).GetCell(0).SetCellValue("查詢條件：受理日期期間：" + vds + " ~ " + vde);
            //sheet1.GetRow(4).GetCell(0).SetCellValue(99);
            List<sp_SatisfactionDetail_Result> result = new List<sp_SatisfactionDetail_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result = db.sp_SatisfactionDetail(vds, vde, dds, dde, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList, null).ToList();
            }
            for (int i = 0; i < result.Count(); i++)
            {
                sheet1.GetRow(4 + i).GetCell(0).SetCellValue(i + 1);
                //XSSFRow cells4 = (XSSFRow)sheet1.CreateRow(i + nRow);
                //cells4.GetCell(0).SetCellValue(i + 1);
                //cells4.CreateCell(0).SetCellValue(i+1);
                //cells4.CreateCell(1).SetCellValue(result[i].AppealNo);
                sheet1.GetRow(4 + i).GetCell(1).SetCellValue(result[i].AppealNo);
                sheet1.GetRow(4 + i).GetCell(2).SetCellValue(result[i].Name);
                sheet1.GetRow(4 + i).GetCell(3).SetCellValue(StringExtension.ToCalendarRC(result[i].VerifyDate.ToString()));
                sheet1.GetRow(4 + i).GetCell(4).SetCellValue(result[i].Subject);
                sheet1.GetRow(4 + i).GetCell(5).SetCellValue(result[i].TopUnitName);
                sheet1.GetRow(4 + i).GetCell(6).SetCellValue(StringExtension.ToCalendarRC(result[i].ReceiveDate.ToString()));
                sheet1.GetRow(4 + i).GetCell(7).SetCellValue(StringExtension.ToCalendarRC(result[i].CreateDate.ToString()));
                sheet1.GetRow(4 + i).GetCell(8).SetCellValue(StringExtension.SatisfactionString(result[i].Q1.ToString(), "A"));
                sheet1.GetRow(4 + i).GetCell(9).SetCellValue(StringExtension.SatisfactionString(result[i].Q2.ToString(), "A"));
                sheet1.GetRow(4 + i).GetCell(10).SetCellValue(StringExtension.SatisfactionString(result[i].Q3.ToString(), "A"));
                sheet1.GetRow(4 + i).GetCell(11).SetCellValue(StringExtension.SatisfactionString(result[i].Q4.ToString(), "A"));
                sheet1.GetRow(4 + i).GetCell(12).SetCellValue(StringExtension.SatisfactionString(result[i].Q5.ToString(), "A"));
                sheet1.GetRow(4 + i).GetCell(13).SetCellValue(StringExtension.SatisfactionString(result[i].Q6.ToString(), "B") + result[i].Remark);
            }
            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }
        /// <summary>
        /// 滿意度問卷統計表
        /// </summary>
        /// <returns>View</returns>
        public ActionResult SatisfactionReport()
        {
            //頁面初始化
            _currentUser = Session["User"] as User;
            bool isSupervisor = false;
            foreach (var item in _currentUser.rolesList)
            {
                if (item.SerialNo == 155)
                {
                    isSupervisor = true;
                }
            }
            string memberType = _currentUser.UserDetail.Internal == "Y" ? "1" : "2"; //部內 = 1 , 所屬 = 2;
            List<SysSubCode> CaseSourceList = new CommonService().GetCaseSourceList();
            List<Organization> caseOrganizerList = new CommonService().GetCaseOrganizerList(memberType, isSupervisor);
            CaseQueryViewModel caseQueryViewModel = new CaseQueryViewModel();
            string NowD = DateTime.Now.ToString("yyyy/MM/dd");
            caseQueryViewModel.VerifyDateS = NowD;
            caseQueryViewModel.VerifyDateE = NowD;
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
            //議題類別
            caseQueryViewModel.StatisticCaseTypeList = _caseManagementService.GetStatCaseList();
            //承辦單位
            foreach (var caseOrganizer in caseOrganizerList.Select(x => new { x.SerialNo, x.ZHName }).OrderBy(x => x.SerialNo))
            {
                caseQueryViewModel.CaseOrganizerList.Add(new SelectListItem() { Text = caseOrganizer.ZHName, Value = caseOrganizer.SerialNo.ToString() });
            }
            return View(caseQueryViewModel);
        }
        /// <summary>
        /// 滿意度問卷統計表匯出Excel
        /// 查詢(年月日區間)匯出EXCEL OutReportQuery
        /// </summary>
        /// <returns>File</returns>        
        public ActionResult SatisfactionExcel(CaseQueryViewModel caseQueryViewModel)
        {
            _currentUser = Session["User"] as User;
            //取得資料
            string vds = caseQueryViewModel.VerifyDateS;
            string vde = caseQueryViewModel.VerifyDateE;
            string AppealCategory = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.CaseTypeList);
            string PetitionTypeList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.PetitionTypeList);
            string SourceList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.SourceList);
            string CaseOrganizerList = StringExtension.NormalizeSplitStringListToDB(caseQueryViewModel.OrganizerList);

            //載入模板檔案路徑  
            string TempletFileName = Server.MapPath("../App_Data/ExcelTemplate/滿意度問卷統計表.xlsx");
            XSSFWorkbook wk = null;
            using (FileStream fs = System.IO.File.OpenRead(TempletFileName))
            {
                //把xlsx檔案讀入workbook變數裡，之後就可以關閉了
                wk = new XSSFWorkbook(fs);
                fs.Close();
            }
            XSSFSheet sheet1 = (XSSFSheet)wk.GetSheetAt(0);
            int nRow = 2;//開始插入的行（第三行）
            sheet1.GetRow(2).GetCell(0).SetCellValue("列印時間：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            sheet1.GetRow(2).GetCell(4).SetCellValue("列印人員：" + _currentUser.UserDetail.UserName);
            sheet1.GetRow(3).GetCell(0).SetCellValue("查詢條件：案件結案期間：" + vds + " ~ " + vde);
            //sheet1.GetRow(5).GetCell(1).SetCellValue(99);
            List<sp_Satisfaction_Result> result = new List<sp_Satisfaction_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result = db.sp_Satisfaction(vds, vde, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
            }
            for (int i = 0; i < result.Count(); i++)
            {
                sheet1.GetRow(5).GetCell(1).SetCellValue((double)result[i].非常滿意);
                sheet1.GetRow(5).GetCell(2).SetCellValue((double)result[i].滿意);
                sheet1.GetRow(5).GetCell(3).SetCellValue((double)result[i].尚可);
                sheet1.GetRow(5).GetCell(4).SetCellValue((double)result[i].不滿意);
                sheet1.GetRow(5).GetCell(5).SetCellValue((double)result[i].非常不滿意);
                sheet1.GetRow(5).GetCell(6).SetCellValue((double)result[i].無效填寫案件);
                sheet1.GetRow(5).GetCell(7).SetCellValue((double)result[i].未填寫案件);
            }
            sheet1.GetRow(28).GetCell(0).SetCellValue("列印時間：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            sheet1.GetRow(28).GetCell(4).SetCellValue("列印人員：" + _currentUser.UserDetail.UserName);
            sheet1.GetRow(29).GetCell(0).SetCellValue("查詢條件：報表類別：滿意度統計表，案件結案期間：" + vds + " ~ " + vde);

            //Next Table
            List<sp_Satisfaction_1_Result> result1 = new List<sp_Satisfaction_1_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result1 = db.sp_Satisfaction_1(vds, vde, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
            }
            for (int i = 0; i < result1.Count(); i++)
            {
                //部長室
                if (result1[i].TopUnit == 9)
                {
                    sheet1.GetRow(32).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(32).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //主任秘書室
                if (result1[i].TopUnit == 10)
                {
                    sheet1.GetRow(33).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(33).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //綜合規劃司
                if (result1[i].TopUnit == 11)
                {
                    sheet1.GetRow(34).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(34).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //社會保險司
                if (result1[i].TopUnit == 12)
                {
                    sheet1.GetRow(35).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(35).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //社會救助及社工司
                if (result1[i].TopUnit == 13)
                {
                    sheet1.GetRow(36).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(36).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //保護服務司
                if (result1[i].TopUnit == 14)
                {

                    sheet1.GetRow(37).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(37).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //護理及健康照護司
                if (result1[i].TopUnit == 15)
                {

                    sheet1.GetRow(38).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(38).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //醫事司
                if (result1[i].TopUnit == 16)
                {

                    sheet1.GetRow(39).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(39).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //心理健康司
                if (result1[i].TopUnit == 17)
                {

                    sheet1.GetRow(40).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(40).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //中醫藥司
                if (result1[i].TopUnit == 18)
                {

                    sheet1.GetRow(41).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(41).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //秘書處
                if (result1[i].TopUnit == 19)
                {

                    sheet1.GetRow(42).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(42).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //人事處
                if (result1[i].TopUnit == 20)
                {

                    sheet1.GetRow(43).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(43).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //政風處
                if (result1[i].TopUnit == 21)
                {

                    sheet1.GetRow(44).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(44).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //會計處
                if (result1[i].TopUnit == 22)
                {

                    sheet1.GetRow(45).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(45).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //統計處
                if (result1[i].TopUnit == 23)
                {

                    sheet1.GetRow(46).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(46).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //資訊處
                if (result1[i].TopUnit == 24)
                {

                    sheet1.GetRow(47).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(47).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //法規會
                if (result1[i].TopUnit == 25)
                {

                    sheet1.GetRow(48).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(48).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國際合作組
                if (result1[i].TopUnit == 26)
                {

                    sheet1.GetRow(49).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(49).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //附屬醫療及社會福利機構管理會
                if (result1[i].TopUnit == 8)
                {

                    sheet1.GetRow(50).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(50).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //全民健康保險會
                if (result1[i].TopUnit == 27)
                {

                    sheet1.GetRow(51).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(51).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //全民健康保險爭議審議會
                if (result1[i].TopUnit == 28)
                {

                    sheet1.GetRow(52).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(52).GetCell(13).SetCellValue((double)result1[i].發送件數);
                }
                //衛生福利人員訓練中心
                if (result1[i].TopUnit == 7)
                {

                    sheet1.GetRow(53).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(53).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國民年金監理會
                if (result1[i].TopUnit == 29)
                {

                    sheet1.GetRow(54).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(54).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //科技發展組
                if (result1[i].TopUnit == 30)
                {

                    sheet1.GetRow(55).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(55).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //公共關係室
                if (result1[i].TopUnit == 31)
                {

                    sheet1.GetRow(56).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(56).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國會聯絡組
                if (result1[i].TopUnit == 32)
                {

                    sheet1.GetRow(57).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(57).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國家消除C肝辦公室
                if (result1[i].TopUnit == 33)
                {

                    sheet1.GetRow(58).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(58).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //長期照顧司
                if (result1[i].TopUnit == 34)
                {

                    sheet1.GetRow(59).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(59).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //疾病管制署
                if (result1[i].TopUnit == 4)
                {

                    sheet1.GetRow(61).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(61).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //食品藥物管理署
                if (result1[i].TopUnit == 2)
                {

                    sheet1.GetRow(62).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(62).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國民健康署
                if (result1[i].TopUnit == 5)
                {

                    sheet1.GetRow(63).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(63).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //防疫專線 No Unit GetRow(64)

                //中央健康保險署
                if (result1[i].TopUnit == 6)
                {

                    sheet1.GetRow(65).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(65).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國家衛生研究院
                if (result1[i].TopUnit == 130)
                {

                    sheet1.GetRow(66).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(66).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //國家中醫藥研究所
                if (result1[i].TopUnit == 210)
                {
                    sheet1.GetRow(67).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(67).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
                //社會及家庭署
                if (result1[i].TopUnit == 3)
                {

                    sheet1.GetRow(69).GetCell(12).SetCellValue((double)result1[i].回收件數);
                    sheet1.GetRow(68).GetCell(13).SetCellValue((double)result1[i].發送件數);

                }
            }
            List<sp_Satisfaction_2_Result> result2 = new List<sp_Satisfaction_2_Result>();
            using (MOHWEntities db = new MOHWEntities())
            {
                result2 = db.sp_Satisfaction_2(vds, vde, AppealCategory, PetitionTypeList, SourceList, CaseOrganizerList).ToList();
            }
            for (int i = 0; i < result2.Count(); i++)
            {
                //部長室
                if (result2[i].TopUnit == 9)
                {
                    sheet1.GetRow(32).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(32).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(32).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(32).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(32).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(32).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(32).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(32).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(32).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(32).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //主任秘書室
                if (result1[i].TopUnit == 10)
                {
                    sheet1.GetRow(33).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(33).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(33).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(33).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(33).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(33).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(33).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(33).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(33).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(33).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //綜合規劃司
                if (result1[i].TopUnit == 11)
                {
                    sheet1.GetRow(34).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(34).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(34).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(34).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(34).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(34).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(34).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(34).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(34).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(34).GetCell(9).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //社會保險司
                if (result1[i].TopUnit == 12)
                {
                    sheet1.GetRow(35).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(35).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(35).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(35).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(35).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(35).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(35).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(35).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(35).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(35).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //社會救助及社工司
                if (result1[i].TopUnit == 13)
                {
                    sheet1.GetRow(36).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(36).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(36).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(36).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(36).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(36).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(36).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(36).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(36).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(36).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //保護服務司
                if (result1[i].TopUnit == 14)
                {
                    sheet1.GetRow(37).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(37).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(37).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(37).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(37).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(37).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(37).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(37).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(37).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(37).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //護理及健康照護司
                if (result1[i].TopUnit == 15)
                {
                    sheet1.GetRow(38).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(38).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(38).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(38).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(38).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(38).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(38).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(38).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(38).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(38).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //醫事司
                if (result1[i].TopUnit == 16)
                {
                    sheet1.GetRow(39).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(39).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(39).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(39).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(39).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(39).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(39).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(39).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(39).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(39).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //心理健康司
                if (result1[i].TopUnit == 17)
                {
                    sheet1.GetRow(40).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(40).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(40).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(40).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(40).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(40).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(40).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(40).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(40).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(40).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //中醫藥司
                if (result1[i].TopUnit == 18)
                {
                    sheet1.GetRow(41).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(41).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(41).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(41).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(41).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(41).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(41).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(41).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(41).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(41).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //秘書處
                if (result1[i].TopUnit == 19)
                {

                    sheet1.GetRow(42).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(42).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(42).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(42).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(42).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(42).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(42).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(42).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(42).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(42).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);

                }
                //人事處
                if (result1[i].TopUnit == 20)
                {
                    sheet1.GetRow(43).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(43).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(43).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(43).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(43).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(43).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(43).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(43).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(43).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(43).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);

                }
                //政風處
                if (result1[i].TopUnit == 21)
                {
                    sheet1.GetRow(44).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(44).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(44).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(44).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(44).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(44).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(44).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(44).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(44).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(44).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //會計處
                if (result1[i].TopUnit == 22)
                {
                    sheet1.GetRow(45).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(45).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(45).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(45).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(45).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(45).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(45).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(45).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(45).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(45).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //統計處
                if (result1[i].TopUnit == 23)
                {
                    sheet1.GetRow(46).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(46).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(46).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(46).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(46).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(46).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(46).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(46).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(46).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(46).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //資訊處
                if (result1[i].TopUnit == 24)
                {
                    sheet1.GetRow(47).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(47).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(47).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(47).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(47).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(47).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(47).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(47).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(47).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(47).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //法規會
                if (result1[i].TopUnit == 25)
                {
                    sheet1.GetRow(48).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(48).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(48).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(48).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(48).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(48).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(48).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(48).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(48).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(48).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國際合作組
                if (result1[i].TopUnit == 26)
                {
                    sheet1.GetRow(49).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(49).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(49).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(49).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(49).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(49).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(49).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(49).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(49).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(49).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //附屬醫療及社會福利機構管理會
                if (result1[i].TopUnit == 8)
                {
                    sheet1.GetRow(50).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(50).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(50).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(50).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(50).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(50).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(50).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(50).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(50).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(50).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //全民健康保險會
                if (result1[i].TopUnit == 27)
                {
                    sheet1.GetRow(51).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(51).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(51).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(51).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(51).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(51).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(51).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(51).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(51).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(51).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //全民健康保險爭議審議會
                if (result1[i].TopUnit == 28)
                {
                    sheet1.GetRow(52).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(52).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(52).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(52).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(52).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(52).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(52).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(52).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(52).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(52).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //衛生福利人員訓練中心
                if (result1[i].TopUnit == 7)
                {
                    sheet1.GetRow(53).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(53).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(53).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(53).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(53).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(53).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(53).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(53).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(53).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(53).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);

                }
                //國民年金監理會
                if (result1[i].TopUnit == 29)
                {
                    sheet1.GetRow(54).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(54).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(54).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(54).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(54).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(54).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(54).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(54).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(54).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(54).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //科技發展組
                if (result1[i].TopUnit == 30)
                {
                    sheet1.GetRow(55).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(55).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(55).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(55).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(55).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(55).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(55).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(55).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(55).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(55).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //公共關係室
                if (result1[i].TopUnit == 31)
                {
                    sheet1.GetRow(56).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(56).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(56).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(56).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(56).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(56).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(56).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(56).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(56).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(56).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國會聯絡組
                if (result1[i].TopUnit == 32)
                {
                    sheet1.GetRow(57).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(57).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(57).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(57).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(57).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(57).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(57).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(57).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(57).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(57).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國家消除C肝辦公室
                if (result1[i].TopUnit == 33)
                {
                    sheet1.GetRow(58).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(58).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(58).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(58).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(58).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(58).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(58).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(58).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(58).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(58).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //長期照顧司
                if (result1[i].TopUnit == 34)
                {
                    sheet1.GetRow(59).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(59).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(59).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(59).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(59).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(59).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(59).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(59).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(59).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(59).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);

                }
                //疾病管制署
                if (result1[i].TopUnit == 4)
                {
                    sheet1.GetRow(61).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(61).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(61).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(61).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(61).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(61).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(61).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(61).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(61).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(61).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //食品藥物管理署
                if (result1[i].TopUnit == 2)
                {
                    sheet1.GetRow(62).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(62).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(62).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(62).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(62).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(62).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(62).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(62).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(62).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(62).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國民健康署
                if (result1[i].TopUnit == 5)
                {
                    sheet1.GetRow(63).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(63).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(63).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(63).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(63).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(63).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(63).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(63).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(63).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(63).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //防疫專線 No Unit GetRow(64)

                //中央健康保險署
                if (result1[i].TopUnit == 6)
                {
                    sheet1.GetRow(65).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(65).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(65).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(65).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(65).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(65).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(65).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(65).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(65).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(65).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國家衛生研究院
                if (result1[i].TopUnit == 130)
                {
                    sheet1.GetRow(66).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(66).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(66).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(66).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(66).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(66).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(66).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(66).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(66).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(66).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //國家中醫藥研究所
                if (result1[i].TopUnit == 210)
                {
                    sheet1.GetRow(67).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(67).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(67).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(67).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(67).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(67).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(67).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(67).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(67).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(67).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
                //社會及家庭署
                if (result1[i].TopUnit == 3)
                {
                    sheet1.GetRow(68).GetCell(1).SetCellValue((double)result2[i].態度滿意);
                    sheet1.GetRow(68).GetCell(2).SetCellValue((double)result2[i].態度不滿意);
                    sheet1.GetRow(68).GetCell(3).SetCellValue((double)result2[i].速度滿意);
                    sheet1.GetRow(68).GetCell(4).SetCellValue((double)result2[i].速度不滿意);
                    sheet1.GetRow(68).GetCell(5).SetCellValue((double)result2[i].專業度滿意);
                    sheet1.GetRow(68).GetCell(6).SetCellValue((double)result2[i].專業度不滿意);
                    sheet1.GetRow(68).GetCell(7).SetCellValue((double)result2[i].解決問題程度滿意);
                    sheet1.GetRow(68).GetCell(8).SetCellValue((double)result2[i].解決問題程度不滿意);
                    sheet1.GetRow(68).GetCell(9).SetCellValue((double)result2[i].整體滿意度滿意);
                    sheet1.GetRow(68).GetCell(10).SetCellValue((double)result2[i].整體滿意度不滿意);
                }
            }
            //sheet1.GetRow(32).GetCell(1).SetCellValue(99);

            //更新有公式的欄位
            sheet1.ForceFormulaRecalculation = true;

            MemoryStream file = new MemoryStream();
            wk.Write(file);
            byte[] FileByte = file.ToArray();
            file.Close();
            file.Dispose();

            return File(FileByte, "application/octet-stream", DateTime.Now.ToString("yyyyMMdd") + ".xlsx");
        }


    }
}