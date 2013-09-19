// <copyright file="FileExtension.cs" company="Nexcel Solutions Vietnam">
// Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Extensions
{
    using System;
    using System.IO;

    /// <summary>
    /// File extension.
    /// </summary>
    public class FileExtension
    {
        /// <summary>
        /// Write file.
        /// </summary>
        /// <param name="sContent"></param>
        public static void WriteFile(string sContent)
        {
            string logPath = AppDomain.CurrentDomain.BaseDirectory;

            ////get the folder that's in

            var outPath = string.Format(@"{0}\Logs\Error", logPath);

            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }

            outPath += string.Format(@"\{0:yyyy-MM-dd}.txt", DateTime.Now.Date);

            var fs = new FileStream(outPath, FileMode.Append, FileAccess.Write);
            var sw = new StreamWriter(fs);

            sw.WriteLine(DateTime.Now + ": " + sContent);

            sw.Close();
            fs.Close();
        }
    }
}
