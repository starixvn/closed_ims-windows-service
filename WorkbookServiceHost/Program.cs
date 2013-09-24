// --------------------------------------------------------------------------
// <copyright file="Program.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace WorkbookServiceHost
{
    using System;
    using System.ServiceProcess;
    using Fanex.Workbook.Service.Extensions;

    /// <summary>
    /// Class Program.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static void Main()
        {
            try
            {
                ServiceBase[] servicesToRun;
                servicesToRun = new ServiceBase[]
                {
                    new WorkbookServiceHost()
                };
                ServiceBase.Run(servicesToRun);
            }
            catch (Exception ex)
            {
                LogExtensions.Log(ex);
            }
        }
    }
}