using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Web;

namespace MohwEmail.Helpers
{
    public static class DbContextExtension
    {
        /// <summary>
        /// 執行SQL 語法
        /// </summary>
        /// <param name="db"></param>
        /// <param name="Sql"></param>
        /// <param name="Parameters"></param>
        /// <returns></returns>
        public static List<Dictionary<string, object>> SqlQuery(this DbContext db, string Sql, Dictionary<string, object> Parameters = null)
        {
            List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();
            Dictionary<string, object> row = new Dictionary<string, object>();
            using (DbCommand cmd = db.Database.Connection.CreateCommand())
            {
                cmd.CommandText = Sql;
                cmd.CommandTimeout = 0;
                if (cmd.Connection.State != ConnectionState.Open) { cmd.Connection.Open(); }

                if (Parameters != null)
                {
                    foreach (KeyValuePair<string, object> p in Parameters)
                    {
                        DbParameter dbParameter = cmd.CreateParameter();
                        dbParameter.ParameterName = p.Key;
                        dbParameter.Value = p.Value;
                        cmd.Parameters.Add(dbParameter);
                    }
                }
                try
                {
                    using (DbDataReader dataReader = cmd.ExecuteReader())
                    {
                        while (dataReader.Read())
                        {
                            row = new Dictionary<string, object>();
                            for (var fieldCount = 0; fieldCount < dataReader.FieldCount; fieldCount++)
                            {
                                row.Add(dataReader.GetName(fieldCount), dataReader[fieldCount]);
                            }                                                         
                            result.Add(row);
                        }
                        dataReader.NextResult();
                        while (dataReader.Read())
                        {
                            row = new Dictionary<string, object>();
                            for (var fieldCount = 0; fieldCount < dataReader.FieldCount; fieldCount++)
                            {
                                row.Add(dataReader.GetName(fieldCount), dataReader[fieldCount]);
                            }
                            result.Add(row);

                        }
                    }
                }
                catch (Exception ex)
                {
                    string aaa = ex.Message;
                    throw;
                }

            }
            return result;
        }
    }
}