using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using Fanex.Workbook.Service.Extensions;

namespace WorkbookServiceHost
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            try
            {

            
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
			    { 
				    new WorkbookServiceHost() 
			    };
                ServiceBase.Run(ServicesToRun);
            }
            catch (Exception ex)
            {
                FileExtension.WriteFile(ex.StackTrace);
            }
        }
    }
}
