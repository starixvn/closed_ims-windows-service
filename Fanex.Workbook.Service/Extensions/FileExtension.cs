// --------------------------------------------------------------------------
// <copyright file="FileExtension.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Extensions
{
    using System;
    using System.IO;

    /// <summary>
    /// Class FileExtension.
    /// </summary>
    public class FileExtension
    {
        /// <summary>
        /// Writes the file.
        /// </summary>
        /// <param name="errorContent">Content of the s.</param>
        public static void WriteFile(string errorContent)
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

            sw.WriteLine(DateTime.Now + ": " + errorContent);

            sw.Close();
            fs.Close();
        }
    }
}
