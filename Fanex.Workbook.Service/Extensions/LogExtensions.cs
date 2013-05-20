using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Fanex.Logging;

namespace Fanex.Workbook.Service.Extensions
{
    public static class LogExtensions
    {
        public static void Log(Exception ex)
        {

            string toAddress = System.Configuration.ConfigurationManager.AppSettings["errorreceiver"];

            MailExtension.SendMail(toAddress, "Fanex Workbook Error", "Fanex Workbook error", ex.Message + "<br> "+ ex.StackTrace +  "<br>" + "\n");
            FileExtension.WriteFile(ex.Message + "\n" + ex.StackTrace + "\n");
        }

        public static void Log(string message = "")
        {
            string toAddress = System.Configuration.ConfigurationManager.AppSettings["errorreceiver"];
            MailExtension.SendMail(toAddress, "Fanex Workbook Error", "Fanex Workbook error", message);

            FileExtension.WriteFile(message);
        }

        public enum LogType
        {
            Info,
            Error
        }


        
    }
}
