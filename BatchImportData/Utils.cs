using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace BatchImportData
{
    public class Utils
    {
        public static string GetReplaceMethod(string str)
        {
            if (string.IsNullOrEmpty(str))
                return str;

            str = str.Replace("\a", "");
            str = str.Replace("\r", "");
            str = str.Replace("\n", "");
            str = str.Replace("\t", "");
            str = str.Replace("\\y", "");
            str = str.Replace("\\l", "");
            str = str.Replace("","");

            return str.Trim();
        }

        /// <summary>
        /// 截取左边字符
        /// </summary>
        /// <param name="sSource"></param>
        /// <param name="iLength"></param>
        /// <returns></returns>
        public static string Left(string sSource, int iLength)
        {
            return sSource.Substring(0, iLength > sSource.Length ? sSource.Length : iLength);
        }

        /// <summary>
        /// 截取右边字符
        /// </summary>
        /// <param name="sSource"></param>
        /// <param name="iLength"></param>
        /// <returns></returns>
        public static string Right(string sSource, int iLength)
        {
            return sSource.Substring(iLength > sSource.Length ? 0 : sSource.Length - iLength);
        }
        /// <summary>
        /// 取字符串开始到结束之间到内容
        /// </summary>
        /// <param name="str"></param>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static Regex IsMatch(string s, string e)
        {
            Regex reg=new Regex("[.\\s\\S]*?(?<=(" + s + "))[.\\s\\S]*?(?<=(" + e + "))", RegexOptions.Multiline | RegexOptions.Singleline);
            return reg;   
        }

        public static DataTable ToDataTable<T>(List<T> items)
        {
            var tb = new DataTable(typeof(T).Name);

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (PropertyInfo prop in props)
            {
                Type t = GetCoreType(prop.PropertyType);
                tb.Columns.Add(prop.Name, t);
            }

            foreach (T item in items)
            {
                var values = new object[props.Length];

                for (int i = 0; i < props.Length; i++)
                {
                    values[i] = props[i].GetValue(item, null);
                }

                tb.Rows.Add(values);
            }

            return tb;
        }

        /// <summary>
        /// Determine of specified type is nullable
        /// </summary>
        public static bool IsNullable(Type t)
        {
            return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        public static Type GetCoreType(Type t)
        {
            if (t != null && IsNullable(t))
            {
                if (!t.IsValueType)
                {
                    return t;
                }
                else
                {
                    return Nullable.GetUnderlyingType(t);
                }
            }
            else
            {
                return t;
            }
        }


        // 读取配置文件
        public static string GetSettings(string key)
        {
            var builder = new ConfigurationBuilder()
                //.SetBasePath(Directory.GetCurrentDirectory() + "/publish/")
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfigurationRoot configuration = builder.Build();
            return configuration[key];
        }

        public static bool IsContains(string text, List<string> keys)
        {
            foreach(var key in keys)
            {
                if (text.Contains(key))
                    continue;
                else
                    return false;
            }

            return true;
        }

    }
}
