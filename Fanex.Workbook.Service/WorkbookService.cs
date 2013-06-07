using System;
using System.IO;
using System.Text;
using Fanex.Workbook.Service.Business;
using Fanex.Workbook.Service.Entities;
using Fanex.Workbook.Service.Extensions;
using System.Collections.Generic;

namespace Fanex.Workbook.Service
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in both code and config file together.
    public class WorkBookService : IWorkBookService
    {
        public string GetData(int value)
        {
            LogExtensions.Log("what the hell");
            return string.Format("You entered: {0}", value);
        }

        public byte[] CreateDocument(List<DocumentSetting> input)
        {
            if (input == null)
                return null;
            //return WorkBookManager.CreateDocument(input.HtmlContent, input.CssContent, input.DocumentType, input.PageOrientation);
            return WorkBookManager.CreateDocument(input);
        }

        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }
            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }
            return composite;
        }

    }
}
