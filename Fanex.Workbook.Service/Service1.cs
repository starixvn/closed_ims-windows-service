using System;
using System.IO;
using System.Text;

namespace Fanex.Workbook.Service
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in both code and config file together.
    public class Service1 : IService1
    {
        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        byte[] IService1.GetStream(int i)
        {
            try
            {
                var ms = new MemoryStream(ASCIIEncoding.Default.GetBytes("Your string here"));
                return ms.ToArray();
            }
            catch (Exception ex)
            {
                
                return new byte[6];
            }
           

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
