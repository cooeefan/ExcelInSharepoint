using System;
using System.Configuration;
using System.IO;

namespace ExcelInSharepoint
{
    class LogWriter
    {
        static public void Write(string logString)
        {
            string logFilePath;

            if (ConfigurationManager.AppSettings["LoggerWriter"] == "ON")
            {
                logFilePath = Directory.GetCurrentDirectory().ToString() + "\\LOG\\" + DateTime.Now.Date.ToString("yyyyMMdd") + ".log";

                if (File.Exists(logFilePath))
                {
                    using (StreamWriter file = File.AppendText(logFilePath))
                    {
                        file.WriteLine(DateTime.Now.ToString() + ": " + logString);
                    }
                }
                else
                {
                    using (FileStream fs = File.Create(logFilePath))
                    {
                        StreamWriter file = new StreamWriter(fs);
                        file.WriteLine(DateTime.Now.ToString() + ": " + logString);
                        file.Close();
                    }
                }
            }
            else
                Console.WriteLine(DateTime.Now.ToString() + ": " + logString);
        }
    }
}

