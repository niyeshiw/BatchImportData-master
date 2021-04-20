using System;
using System.IO;
using System.Text;

namespace BatchImportData
{
    public class LogFileHelper
    {
        public LogFileHelper()
        {
          
        }
        public static void WriteTextLog(string strMessage,string filename)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Log/";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            DateTime time = DateTime.Now;
            string fileFullPath = path + time.ToString("yyyy-MM-dd") + "." + filename + ".txt";
            StringBuilder str = new StringBuilder();

            str.Append("Time:" + time + ";Message: " + strMessage + "\r\n");
            StreamWriter sw;
            if (!File.Exists(fileFullPath))
            {
                sw = File.CreateText(fileFullPath);
            }
            else
            {
                sw = File.AppendText(fileFullPath);
            }
            sw.WriteLine(str.ToString());
            sw.Close();
        }

        public static void WriteInfo(string message)
        {
            WriteLog("info", message);
        }

        public static void WriteError(string message)
        {
            WriteLog("error", message);
        }

        public static void WriteLog(string category, string message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + $"Log\\{category}\\";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string fileName = DateTime.Today.ToString("yyyy-MM-dd") + ".txt";
            string fullPath = path + fileName;
            if (!File.Exists(fullPath))
            {
                File.Create(fullPath);
            }
            var sw = File.AppendText(fullPath);
            sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}");
            sw.Close();

        }

    }
}
