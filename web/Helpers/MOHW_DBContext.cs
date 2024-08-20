using MohwEmail.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MohwEmail.Helpers
{
    public class MOHW_DBContext : MOHWEntities
    {
        public MOHW_DBContext()
        {
            string Password = System.Configuration.ConfigurationManager.AppSettings.Get("MOHWEntities");
            string ConnectionString = string.Format("{0};Password={1}", base.Database.Connection.ConnectionString, Password);
            base.Database.Connection.ConnectionString = ConnectionString;
        }
    }
}