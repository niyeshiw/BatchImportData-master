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
    }
}
