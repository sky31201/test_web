using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MohwEmail
{
    public enum CRUDMode { Insert, Update, Delete, Select }

    public class GlobalResource
    {
        public static class Strings
        {
            public static class LoginMessage
            {
                public const string Success = "登入成功";
                public const string AccountError = "帳號錯誤";
                public const string PasswordError = "密碼錯誤";
            }
            public static class CRUDMessage
            {
                public const string InsertSuccess = "新增成功";
                public const string UpdateSuccess = "修改成功";
                public const string DeleteSuccess = "刪除成功";
                public const string DisableSuccess = "停用帳號成功";
                public const string ModifyPWSuccess = "修改密碼成功";
                public const string ForgetLoginSuccess = "發送郵件成功";
                public const string InsertError = "新增失敗: {0}";
                public const string UpdateError = "修改失敗: {0}";
                public const string DeleteError = "刪除失敗: {0}";
            }
        }

    }
}