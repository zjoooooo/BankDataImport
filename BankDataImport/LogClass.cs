using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankDataImport
{
    class LogClass
    {
        public static void WriteLog(string content)
        {
            string LogPath = Path.Combine(Environment.CurrentDirectory, "Log");
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }
            StreamWriter sw = new StreamWriter(LogPath +@"\"+ DateTime.Now.ToString("yyyy-MM-dd") + ".log", true);
            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss  ") + content);
            //sw.WriteLine(content);
            sw.Flush();
            sw.Close();
          //  Console.WriteLine(content);
        }


    }
}
