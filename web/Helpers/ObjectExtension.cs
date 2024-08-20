using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Web;

namespace MohwEmail.Helpers
{
    public static class ObjectExtension
    {
        public static void InitialNull(this object source)
        {
            foreach (PropertyInfo prop in source.GetType().GetProperties())
            {
                //string: null to string.Empty
                if (prop.PropertyType == typeof(string))
                {
                    string propValue = prop.GetValue(source, null) as string;
                    if (propValue == null)
                    {
                        prop.SetValue(source, string.Empty, null);
                    }
                }
                else if (prop.PropertyType == typeof(int?))
                {
                    int? propValue = prop.GetValue(source, null) as int?;
                    if (propValue == null)
                    {
                        prop.SetValue(source, 0, null);
                    }
                }
                else if (prop.PropertyType == typeof(bool?))
                {
                    bool? propValue = prop.GetValue(source, null) as bool?;
                    if (propValue == null)
                    {
                        prop.SetValue(source, false, null);
                    }
                }
                else if (prop.PropertyType == typeof(DateTime?))
                {
                    DateTime? propValue = prop.GetValue(source, null) as DateTime?;
                    if (propValue == null)
                    {
                        prop.SetValue(source, new Nullable<DateTime>(), null);
                    }                  
                }
            }
        }
    }
}