using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Fanex.Workbook.Service.Client.ServiceReference1;


namespace Fanex.Workbook.Service.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            
            var myService = new ServiceReference1.WorkbookServiceClient();

            var test = myService.CreateDocument(new DocumentSetting());
            var ms = new MemoryStream(test);

            // convert stream to string
            var reader = new StreamReader(ms);
            string text = reader.ReadToEnd();
        }
    }
}
