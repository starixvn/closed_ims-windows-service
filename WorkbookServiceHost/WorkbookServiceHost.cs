using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.ServiceModel;
using Fanex.Workbook.Service.Extensions;

namespace WorkbookServiceHost
{
    public partial class WorkbookServiceHost : ServiceBase
    {
        internal static ServiceHost WorkBookServiceHost = null;

        public WorkbookServiceHost()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                if (WorkBookServiceHost != null)
                {
                    WorkBookServiceHost.Close();
                }
                WorkBookServiceHost = new ServiceHost(typeof(Fanex.Workbook.Service.WorkBookService));
                WorkBookServiceHost.Open();
            }
            catch (Exception ex)
            {
                FileExtension.WriteFile(ex.StackTrace);
                throw ex;
            }
           
        }

        protected override void OnStop()
        {
            if (WorkBookServiceHost != null)
            {
                WorkBookServiceHost.Close();
                WorkBookServiceHost = null;
            }
        }
    }
}
