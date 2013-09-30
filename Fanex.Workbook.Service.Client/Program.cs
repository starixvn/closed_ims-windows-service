// --------------------------------------------------------------------------
// <copyright file="Program.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Client
{
    using System.IO;
    using Fanex.Workbook.Service.Client.ServiceReference1;

    /// <summary>
    /// Class Program.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        private static void Main(string[] args)
        {
            var clientService = new ServiceReference1.WorkbookServiceClient();

            var test = clientService.CreateDocument(new DocumentSetting());
            var ms = new MemoryStream(test);

            //// convert stream to string
            var reader = new StreamReader(ms);
            string text = reader.ReadToEnd();
        }
    }
}
