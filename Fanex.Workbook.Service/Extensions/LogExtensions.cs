// --------------------------------------------------------------------------
// <copyright file="LogExtensions.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Extensions
{
    using System;
    using System.Configuration;
    using System.Net;

    /// <summary>
    /// Class LogExtensions.
    /// </summary>
    public static class LogExtensions
    {
        /// <summary>
        /// Enumerate LogType.
        /// </summary>
        public enum LogType
        {
            /// <summary>
            /// The information type.
            /// </summary>
            Info,

            /// <summary>
            /// The error type.
            /// </summary>
            Error
        }

        /// <summary>
        /// Logs the specified exception.
        /// </summary>
        /// <param name="ex">The exception.</param>
        public static void Log(Exception ex)
        {
            var errorMessage = string.Format(
                "{0}<br>\r\nSOURCE: {1}<br>\r\nTARGETSITE: {2}<br>\r\nSTACKTRACE: {3}",
                ex.Message,
                ex.Source,
                ex.TargetSite,
                ex.StackTrace);
            Fanex.Logging.Logger.Log.Error(errorMessage);
            FileExtension.WriteFile(errorMessage + "\n");
        }

        /// <summary>
        /// Logs the specified message.
        /// </summary>
        /// <param name="message">The message.</param>
        public static void Log(string message = "")
        {
            Fanex.Logging.Logger.Log.Error(message);
            FileExtension.WriteFile(message);
        }
    }
}