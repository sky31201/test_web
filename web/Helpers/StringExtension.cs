using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace MohwEmail.Helpers
{
    public static class StringExtension
    {
        /// <summary>
        /// 產生MD5Hash
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToMD5Hash(this string input)
        {
            // Create a new Stringbuilder to collect the bytes
            // and create a string.
            StringBuilder sBuilder = new StringBuilder();

            using (MD5 md5Hash = MD5.Create())
            {
                // Convert the input string to a byte array and compute the hash.
                byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input));

                // Loop through each byte of the hashed data 
                // and format each one as a hexadecimal string.
                for (int i = 0; i < data.Length; i++)
                {
                    sBuilder.Append(data[i].ToString("x2"));
                }
            }

            // Return the hexadecimal string.
            return sBuilder.ToString();
        }

        /// <summary>
        /// 正規化逗點分割字串
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string NormalizeSplitString(this string input)
        {
            string result = input;

            List<string> replaceString = new List<string> { "~", "`", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "=", "+", "[", "]", "\\", "|", "}", "{", "<", ">", ".", "?", "/", " ", "　" };
            foreach (string c in replaceString)
            {
                result = result.Replace(c, "");
            }
            string[] temp = result.Split(new string[] { ",", "\r\n" }, System.StringSplitOptions.RemoveEmptyEntries);

            result = string.Join(",", temp);

            return result;
        }

        /// <summary>
        /// 正規畫電話號碼
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToPhoneNumber(this string input)
        {
            string result = input;
            //去雜質
            List<string> replaceString = new List<string> { "~", "`", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "=", "+", "[", "]", "\\", "|", "}", "{", "<", ">", ".", "?", "/", " ", "　" };
            foreach (string c in replaceString)
            {
                result = result.Replace(c, "");
            }

            //多筆取一筆
            if (!string.IsNullOrEmpty(result))
            {
                result = result.Split(new string[] { ",", "\r\n" }, System.StringSplitOptions.RemoveEmptyEntries)[0];
            }

            //非數字
            if (!long.TryParse(result, out long temp))
            {
                return string.Empty;
            }

            //9xxxxxxxx => 09xxxxxxxx
            if (result.Length.Equals(9) && result.StartsWith("9"))
            {
                result = result.PadLeft(10, '0');
            }

            //8869xxxxxxxx => 09xxxxxxxx
            if (result.Length.Equals(12) && result.StartsWith("886"))
            {
                result = result.Substring(result.IndexOf("886") + 3).PadLeft(10, '0');
            }

            return result;
        }
        /// <summary>
        /// 西元年轉換民國年
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToCalendarRC(this string input)
        {
            string result = input;
            if (input != "")
            {
                //string sampleDate = "2012-2-29";
                DateTime dt = DateTime.Parse(input);
                CultureInfo culture = new CultureInfo("zh-TW");
                culture.DateTimeFormat.Calendar = new TaiwanCalendar();
                result = dt.ToString("yyy/MM/dd", culture);
            }
           

            return result;
        }
        /// <summary>
        /// 民國年轉換西元年
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToCalendarWest(this string input)
        {
            string result = input;
            //string sampleDate = "101/02/29";
            CultureInfo culture = new CultureInfo("zh-TW");
            culture.DateTimeFormat.Calendar = new TaiwanCalendar();
            result = DateTime.Parse(input, culture).ToString("yyyy/MM/dd");

            return result;
        }
        /// <summary>
        /// 幫字串加逗點跟''
        /// </summary>
        /// <param name="inputList"></param>
        /// <returns></returns>
        public static string NormalizeSplitStringListToDB(this List<string> inputList)
        {
            string result = null;
            string inputList_ = "";
            if (inputList != null)
            {
                foreach (string a in inputList)
                {
                    inputList_ += a + ",";
                }
                //改用TrimEnd清空最後一個分隔字元.
                if (inputList_ != "")
                {
                    inputList_ = inputList_.TrimEnd(',');
                }
                //添加''
                if (inputList_ != "")
                {
                    string toresult = inputList_.Replace("'", "");
                    result = toresult.Replace("'", "").Replace(",", "','");
                    if (!string.IsNullOrEmpty(result)) result = "'" + result + "'";
                }
            }


            //string IdentityNos = SourceList.Replace("'", "");
            //string IdentityNostring = IdentityNos.Replace("'", "").Replace(",", "','");
            //if (!string.IsNullOrEmpty(IdentityNostring)) IdentityNostring = "'" + IdentityNostring + "'";


            return result;

        }

        /// <summary>
        /// 取得資料庫連線字串(For Dapper使用)
        /// </summary>
        /// <returns></returns>
        public static string GetConnectionString()
        {
            var efConnString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["MOHWEntities"].ConnectionString;
            var dapperConnString = efConnString.Substring(efConnString.IndexOf("data source="));
            return dapperConnString.Substring(0, dapperConnString.Length - 1);
        }

        //public (DateTime startDate, DateTime endDate) GetDateByZC(TimeFrame entity, int zc)
        //{
        //    List & lt; DateTime & gt; list = new List& lt; DateTime & gt; ();
        //    for (DateTime d = entity.XQQZSJ.Value; d & lt;= entity.XQJSSJ.Value; d = d.AddDays(1))
        //    {
        //        list.Add(d);
        //    }
        //    while (list.First().DayOfWeek != DayOfWeek.Monday)
        //    {
        //        list.Insert(0, list.First().AddDays(-1));
        //    }

        //    list = list.OrderByDescending(x = &gt; x).ToList();
        //    while (list.First().DayOfWeek != DayOfWeek.Sunday)
        //    {
        //        list.Insert(0, list.First().AddDays(1));
        //    }
        //    list = list.OrderBy(x = &gt; x).ToList();

        //    DateTime startDate = list[(zc - 1) * 7];
        //    DateTime endDate = list[(zc - 1) * 7 + 6];

        //    while (startDate & lt; entity.XQQZSJ)
        //    {
        //        startDate = startDate.AddDays(1);
        //    }
        //    while (endDate & gt; entity.XQJSSJ)
        //    {
        //        endDate = endDate.AddDays(-1);
        //    }

        //    endDate = DateTime.Parse(endDate.ToString("yyyy-MM-dd") + " 23:59:59");

        //    return (startDate, endDate);
        //}

        /// <summary>
        /// 当前周的第一天(星期一)
        /// </summary>
        /// <param name="yearWeek">周数，格式：yyyywww</param>
        /// <returns></returns>
        public static DateTime GetWeekStartTime(string yearWeek)
        {
            int year = int.Parse(yearWeek.Substring(0, 4));
            //本年1月1日
            DateTime firstOfYear = new DateTime(year, 1, 1);
            //周数
            int weekNum = int.Parse(yearWeek.Substring(4));
            //本年1月1日与本周星期一相差的天数
            int dayDiff = (firstOfYear.DayOfWeek == DayOfWeek.Sunday ? 7 : Convert.ToInt32(firstOfYear.DayOfWeek)) - 1;
            //第一周的星期一
            DateTime firstDayOfFirstWeek = firstOfYear.AddDays(-dayDiff);
            //当前周的星期一
            DateTime firstDayOfThisWeek = firstDayOfFirstWeek.AddDays((weekNum - 1) * 7);
            return firstDayOfThisWeek;


        }
        /// <summary>
        /// 当前周的最后一天(星期天)
        /// </summary>
        /// <param name="yearWeek">周数，格式：yyyywww</param>
        /// <returns></returns>
        public static DateTime GetWeekEndTime(string yearWeek)
        {
            //当前周的星期一
            DateTime firstDayOfThisWeek = GetWeekStartTime(yearWeek);
            //当前周的星期天
            DateTime lastDayOfThisWeek = firstDayOfThisWeek.AddDays(6);
            return lastDayOfThisWeek;
        }
        /// <summary>
        /// 滿意度問卷中文顯示轉換有兩種選單
        /// </summary>
        /// <param inputQ="1~5"></param>
        /// <param QType="A or B"></param>
        /// <returns></returns>
        public static string SatisfactionString(string inputQ,string QType)
        {
            string result = null;

            if (inputQ != null)
            {
                if (QType == "A")
                {
                    switch (inputQ)
                    {
                        case "1":
                            result += "非常滿意";
                            break;
                        case "2":
                            result += "滿意";
                            break;
                        case "3":
                            result += "尚可";
                            break;
                        case "4":
                            result += "不滿意";
                            break;
                        case "5":
                            result += "非常不滿意";
                            break;
                        default:

                            break;
                    }
                }
                else if (QType == "B")
                {
                    List<string> list = new List<string>();
                    list = inputQ.Split(',').ToList();
                    foreach (var Q in list)
                    {
                        switch (Q)
                        {
                            case "1":
                                result += "已完全解決,";
                                break;
                            case "2":
                                result += "部份解決,";
                                break;
                            case "3":
                                result += "承辦人員態度不佳,";
                                break;
                            case "4":
                                result += "答覆內容沒有具體明確,";
                                break;
                            case "5":
                                result += "答覆內容為制式例稿，欠缺誠意,";
                                break;
                            case "6":
                                result += "處理結果沒有考量民眾需求,";
                                break;
                            default:

                                break;
                        }
                    }
                }
              
               
            }
            return result;

        }


    }
}