using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Fanex.Workbook.Service.Client.WorkbookService;

namespace Fanex.Workbook.Service.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var myService = new Service1Client();

            var test = myService.GetStream(new ExportFileInput());
            var ms = new MemoryStream(test);

            // convert stream to string
            var reader = new StreamReader(ms);
            string text = reader.ReadToEnd();
        }
    }
}
