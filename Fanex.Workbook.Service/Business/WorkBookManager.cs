using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Fanex.Workbook.Service.Entities;
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

        private const double DocumentWidth = 1040;

        public enum DocumentType : int
        {
            Xlsx = 0,
            Pdf = 1
        }

        public enum PageOrientation
        {
            Landscape, Portrait
        }

        private static string CreateHtmFile(string htmlContent, string folderName, string fileName)
        {
            //Open a memory stream that you can use to write back to the response
            byte[] byteArray = Encoding.ASCII.GetBytes(htmlContent);
            var stream = new MemoryStream(byteArray);
            if (stream.Length == 0)
            {
                return null;
            }
            //var fileName = Guid.NewGuid() + ".xls";

            string filePath = string.IsNullOrEmpty(folderName) == true ? (FileDirectory) : (FileDirectory + "\\" + folderName);

            var isExists = System.IO.Directory.Exists(filePath);
            if (!isExists)
            {
                System.IO.Directory.CreateDirectory(filePath);
            }

            string fileNameTmp = string.IsNullOrEmpty(folderName) == true ? (FileDirectory + fileName + ".htm") : (FileDirectory + "\\" + folderName + "\\" + fileName + ".htm");

            using (FileStream fileStream = System.IO.File.Create(fileNameTmp, (int)stream.Length))
            {
                // Fill the bytes[] array with the stream data
                var bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, (int)bytesInStream.Length);

                // Use FileStream object to write to the specified file

                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }

            return fileNameTmp;
        }

        public static byte[] CreateDocument(DocumentFile documentFile)
        {
            try
            {
                List<PageSetting> sheetList = documentFile.PageDetailList;
                DocumentType format = documentFile.DocumentType;
                string folderName = documentFile.FileName;

                var isExists = System.IO.Directory.Exists(FileDirectory);
                if (!isExists)
                {
                    System.IO.Directory.CreateDirectory(FileDirectory);
                }

                var templatePath = FileDirectory + "ExcelTemplate\\_ExcelTemplate.cshtml";
                if (string.IsNullOrEmpty(templatePath))
                {
                    return null;
                }

                var templateStream = new StreamReader(templatePath);
                string template = templateStream.ReadToEnd();

                //Create folder that contains htm files with excel file functions. 
                for (int i = 0; i < sheetList.Count; i++)
                {
                    string fileContent = template;

                    //Assign css
                    fileContent = fileContent.Replace("<!--style-->", sheetList[i].CssContent);

                    //Assign data
                    fileContent = fileContent.Replace("<!--data-->", sheetList[i].HtmlContent);

                    string fileName = folderName + (i + 1);

                    CreateHtmFile(fileContent, folderName, fileName);
                }

                // Create main page htm to open excel file.
                string mainPageContent = template.Replace("<!--sheetRef-->", documentFile.MainPage.HtmlContent);

                templateStream.Close();

                _xlsFileFormat = "." + format.ToString().ToLower();


                string excelFile = CreateDocumentFromHtml(mainPageContent, documentFile.FileName, documentFile.DocumentType, documentFile);

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
            catch (Exception ex)
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

        private static string CreateDocumentFromHtml(string mainPage, string fileName, WorkBookManager.DocumentType documentType, DocumentFile documentFile)
        {
            DocumentType format = documentType;
            var app = new Microsoft.Office.Interop.Excel.Application();
            var workBooks = app.Workbooks;

            Microsoft.Office.Interop.Excel.Workbook workBook = null;

            var docFile = TempDirectory + Guid.NewGuid() + _xlsFileFormat;

            try
            {
                //Binding html file to workbook
                double marginLeftRight;
                double marginTopBottom;
                var prter = app.ActivePrinter;
                app.DisplayAlerts = false;
                object paramMissing = Type.Missing;

                string htmlFile = CreateHtmFile(mainPage, string.Empty, fileName);
                workBook = workBooks.Open(htmlFile, 0, true, 5, "", "", true,
                                                                                       Microsoft.Office.Interop.Excel.
                                                                                           XlPlatform.xlWindows, "\t", false,
                                                                                       false, 0, true, 1, 0);

                var length = documentFile.PageDetailList.Count;
                for (int i = 0; i < length; i++)
                {
                    //save workbook to new real Excel file
                    Worksheet sheetCurrent = (Worksheet)workBook.Worksheets[i + 1];
                    sheetCurrent.PageSetup.Zoom = false;
                    sheetCurrent.PageSetup.FitToPagesWide = 1;
                    sheetCurrent.PageSetup.CenterHorizontally = true;

                    //print option
                    switch (documentFile.PageDetailList[i].PageOrientation)
                    {
                        case PageOrientation.Landscape:
                            sheetCurrent.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                            break;
                        case PageOrientation.Portrait:
                            sheetCurrent.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                            break;
                    }

                    if (documentFile.PageDetailList[i].IsAutoFit)
                    {
                        sheetCurrent.Columns.AutoFit();
                        var usedRange = sheetCurrent.UsedRange;
                        var colCount = usedRange.Columns.Count;
                        var width = usedRange.Width;

                        // if the width of document smaller than default document width so the columns after the second column will be expanded
                        if (width < DocumentWidth)
                        {
                            double expandedWidth = DocumentWidth - width;
                            double expendedWidthForEachCol = expandedWidth / colCount - 2;
                            int count = 0;
                            foreach (Microsoft.Office.Interop.Excel.Range col in usedRange.Columns)
                            {
                                // ignore the first 2 columns A & B
                                if (count < 2)
                                {
                                    count++;
                                    continue;
                                }
                                col.ColumnWidth += expendedWidthForEachCol * (8.43 / 48);
                            }
                        }
                        Microsoft.Office.Interop.Excel.Range offset = sheetCurrent.UsedRange.Offset;
                    }

                    switch (format)
                    {
                        case DocumentType.Xlsx:
                            marginLeftRight = app.InchesToPoints(0.5);
                            marginTopBottom = app.InchesToPoints(0.75);
                            sheetCurrent.PageSetup.TopMargin = marginTopBottom;
                            sheetCurrent.PageSetup.BottomMargin = marginTopBottom;
                            sheetCurrent.PageSetup.LeftMargin = marginLeftRight;
                            sheetCurrent.PageSetup.RightMargin = marginLeftRight;
                            break;

                        case DocumentType.Pdf:
                            marginTopBottom = app.InchesToPoints(0.5);
                            sheetCurrent.PageSetup.TopMargin = marginTopBottom;
                            sheetCurrent.PageSetup.BottomMargin = marginTopBottom;
                            sheetCurrent.PageSetup.LeftMargin = 0;
                            sheetCurrent.PageSetup.RightMargin = 0;
                            break;
                    }
                }


                switch (format)
                {
                    case DocumentType.Xlsx:
                        object oSaveAsFile = (object)(docFile);
                        workBook.SaveAs(oSaveAsFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                                        Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        break;

                    case DocumentType.Pdf:
                        const XlFixedFormatType paramExportFormat = XlFixedFormatType.xlTypePDF;
                        const XlFixedFormatQuality paramExportQuality = XlFixedFormatQuality.xlQualityStandard;
                        const bool paramOpenAfterPublish = false;
                        const bool paramIncludeDocProps = true;
                        const bool paramIgnorePrintAreas = true;
                        var paramFromPage = Type.Missing;
                        var paramToPage = Type.Missing;
                        workBook.ExportAsFixedFormat(paramExportFormat, docFile, paramExportQuality, paramIncludeDocProps,
                                                     paramIgnorePrintAreas, paramFromPage, paramToPage,
                                                     paramOpenAfterPublish, paramMissing);
                        break;
                }

                //Delete temp html file
                string forderDirectory = FileDirectory + documentFile.FileName;
                Directory.Delete(@forderDirectory, true);

                if (File.Exists(htmlFile))
                {
                    File.Delete(htmlFile);
                }

                workBook.Close();
                app.Workbooks.Close();

                return docFile;
            }
            catch (Exception ex)
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
    }
}
