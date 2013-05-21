using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Fanex.Workbook.Service.Extensions;
using Microsoft.Office.Interop.Excel;

namespace Fanex.Workbook.Service.Business
{
    /// <summary>
    /// Excel, Word, PDF... extension
    /// </summary>
    public static class WorkBookManager
    {
        private static readonly string FileDirectory = AppDomain.CurrentDomain.BaseDirectory + "\\";

        private static readonly string TempDirectory = "C:\\";

        private static string _xlsFileFormat;

        public enum DocumentType:int
        {
            Xlsx = 0,
            Pdf = 1
        }

        public enum PageOrientation
        {
            Landscape, Portrait
        }
        
        /// <summary>
        /// Create Excel Memory Stream from html
        /// </summary>
        /// <param name="htmlContent"></param>
        /// <param name="cssContent"></param>
        /// <param name="format">Xlsx, Pdf </param>
        /// <param name="pageOrientation">Landscape, Portrait </param>
        /// <returns></returns>
        public static byte[] CreateDocument(string htmlContent, string cssContent, DocumentType format, PageOrientation pageOrientation)
        {
            try
            {
                var isExists = System.IO.Directory.Exists(FileDirectory);

                if (!isExists)
                    System.IO.Directory.CreateDirectory(FileDirectory);

                var excelTemplate = FileDirectory + "ExcelTemplate\\_ExcelTemplate.cshtml";

                if (string.IsNullOrEmpty(excelTemplate)) return null;

                var streamTemplate = new StreamReader(excelTemplate);

                var template = streamTemplate.ReadToEnd();

                //Assign css
                template = template.Replace("<!--style-->", cssContent);

                //Assign data
                template = template.Replace("<!--data-->", htmlContent);

                streamTemplate.Close();

                _xlsFileFormat = "." + format.ToString().ToLower();

                var excelFile = CreateDocFromHtml(template, format, pageOrientation);

                if (string.IsNullOrEmpty(excelFile)) return null;

                var ms = new MemoryStream();

                var file = new FileStream(excelFile, FileMode.Open);
                var bytes = new byte[file.Length];
                file.Read(bytes, 0, (int)file.Length);
                ms.Write(bytes, 0, (int)file.Length);
                file.Close();


                Thread.Sleep(100);
                if (File.Exists(excelFile))
                {
                    File.Delete(excelFile);
                }


                return ms.ToArray();
            }
            catch(Exception ex)
            {
                LogExtensions.Log(ex);
                return null;
            }

            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }


            
        }

        private static string LoadCss(List<string> listCssFilePath)
        {
            string css = string.Empty;

            if (listCssFilePath ==null || !listCssFilePath.Any()) return string.Empty;

            foreach (var cssFilePath in listCssFilePath)
            {
                var streamCss = new StreamReader(cssFilePath);
                css += streamCss.ReadToEnd();

                streamCss.Close();
            }
            
            
            return css;
        }

        private static string CreateDocFromHtml(string htmlContent, DocumentType format, PageOrientation pageOrientation)
        {

            var app = new Microsoft.Office.Interop.Excel.Application();
            // the reference to the workbook,
            // which is the xls document to read from.
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            // the reference to the worksheet,
            // we'll assume the first sheet in the book.
            Worksheet sheet = null;
            
            try
            {
                var htmlFile = CreateHtmlFile(htmlContent);

                if (string.IsNullOrEmpty(htmlFile)) return string.Empty;

                
                // the range object is used to hold the data
                // we'll be reading from and to find the range of data.
                //var workbooks = new Workbooks();

                //Binding html file to workbook
                
                object paramMissing = Type.Missing;
                Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(htmlFile, 0, true, 5, "", "", true,
                                                                                   Microsoft.Office.Interop.Excel.
                                                                                       XlPlatform.xlWindows, "\t", false,
                                                                                   false, 0, true, 1, 0);
                //Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Open(htmlFile, Type.Missing,

                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,

                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //save workbook to new real Excel file
                app.DisplayAlerts = false;
                sheet = (Worksheet)workBook.Worksheets[1];
                sheet.PageSetup.Zoom = false;
                sheet.PageSetup.FitToPagesWide = 1;

                //print option

                switch (pageOrientation)
                {
                    case PageOrientation.Landscape:
                        sheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                        break;
                    case PageOrientation.Portrait:
                        sheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                        break;
                }




                var docFile = TempDirectory + Guid.NewGuid() + _xlsFileFormat;
                sheet.Columns.AutoFit();
                sheet.Cells.WrapText = true;
                switch (format)
                {
                    case DocumentType.Xlsx:


                        

                        object oSaveAsFile = (object)(docFile);
                        workBook.SaveAs(oSaveAsFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                                        Type.Missing,
                                        false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        break;
                    case DocumentType.Pdf:
                        //




                        const XlFixedFormatType paramExportFormat = XlFixedFormatType.xlTypePDF;

                        const XlFixedFormatQuality paramExportQuality = XlFixedFormatQuality.xlQualityStandard;
                        const bool paramOpenAfterPublish = false;
                        const bool paramIncludeDocProps = true;
                        const bool paramIgnorePrintAreas = true;
                        var paramFromPage = Type.Missing;
                        var paramToPage = Type.Missing;
                        


                        workBook.ExportAsFixedFormat(paramExportFormat, docFile, paramExportQuality,
                                                     paramIncludeDocProps,
                                                     paramIgnorePrintAreas, paramFromPage, paramToPage,
                                                     paramOpenAfterPublish, paramMissing);

                        break;

                }



                workBook.Close();
                app.Workbooks.Close();


                //Delete temp html file

                if (File.Exists(htmlFile))
                {
                    File.Delete(htmlFile);
                }

                return docFile;
            }
            catch(Exception ex )
            {
                LogExtensions.Log(ex);

                return string.Empty;
            }

            finally
            {
                

                // Close the ApplicationClass object.
                if (app != null)
                {
                    app.Quit();
                   
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            

        }

        private static string CreateHtmlFile(string htmlContent)
        {
            //Open a memory stream that you can use to write back to the response
            byte[] byteArray = Encoding.ASCII.GetBytes(htmlContent);
            var stream = new MemoryStream(byteArray);

            if (stream.Length == 0) return string.Empty;


            var fileName = Guid.NewGuid() + ".xls";


            using (FileStream fileStream = System.IO.File.Create(FileDirectory + fileName, (int) stream.Length))
            {
                // Fill the bytes[] array with the stream data
                var bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, (int) bytesInStream.Length);

                // Use FileStream object to write to the specified file
                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }
            return FileDirectory + fileName;
        }
    }
}
