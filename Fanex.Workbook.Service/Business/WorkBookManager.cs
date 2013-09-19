using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Fanex.Workbook.Service.Extensions;
using Microsoft.Office.Interop.Excel;
using Fanex.Workbook.Service.Entities;

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

        private static string LoadCss(List<string> listCssFilePath)
        {
            string css = string.Empty;

            if (listCssFilePath == null || !listCssFilePath.Any()) return string.Empty;

            foreach (var cssFilePath in listCssFilePath)
            {
                var streamCss = new StreamReader(cssFilePath);
                css += streamCss.ReadToEnd();

                streamCss.Close();
            }

            return css;
        }

        private static string CreateHtmlFile(string htmlContent)
        {
            //Open a memory stream that you can use to write back to the response
            byte[] byteArray = Encoding.ASCII.GetBytes(htmlContent);
            var stream = new MemoryStream(byteArray);
            if (stream.Length == 0) return string.Empty;
            var fileName = Guid.NewGuid() + ".xls";
            using (FileStream fileStream = System.IO.File.Create(FileDirectory + fileName, (int)stream.Length))
            {
                // Fill the bytes[] array with the stream data
                var bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, (int)bytesInStream.Length);

                // Use FileStream object to write to the specified file
                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }

            return FileDirectory + fileName;
        }

        public static byte[] CreateDocument(List<DocumentSetting> sheetList)
        {
            try
            {
                DocumentType format = sheetList[0].DocumentType;
                var isExists = System.IO.Directory.Exists(FileDirectory);
                if (!isExists)
                    System.IO.Directory.CreateDirectory(FileDirectory);

                var excelTemplate = FileDirectory + "ExcelTemplate\\_ExcelTemplate.cshtml";
                if (string.IsNullOrEmpty(excelTemplate)) return null;

                var streamTemplate = new StreamReader(excelTemplate);

                var templateTmp = streamTemplate.ReadToEnd();

                for (int i = 0; i < sheetList.Count; i++)
                {
                    string template = string.Empty;
                    template = templateTmp;

                    //Assign css
                    template = template.Replace("<!--style-->", sheetList[i].CssContent);

                    //Assign data
                    template = template.Replace("<!--data-->", sheetList[i].HtmlContent);

                    sheetList[i].HtmlContent = string.Empty;
                    sheetList[i].HtmlContent = template;
                }

                streamTemplate.Close();
                _xlsFileFormat = "." + format.ToString().ToLower();
                string excelFile = CreateDocFromHtml(sheetList);
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
        private static string CreateDocFromHtml(List<DocumentSetting> sheetList, int beginRow = 0)
        {
            DocumentType format = sheetList[0].DocumentType;
            var app = new Microsoft.Office.Interop.Excel.Application();
            var workBooks = app.Workbooks;

            Microsoft.Office.Interop.Excel.Workbook workBook = workBooks.Add(XlWBATemplate.xlWBATWorksheet);

            var docFile = TempDirectory + Guid.NewGuid() + _xlsFileFormat;

            try
            {
                double marginLeftRight;
                double marginTopBottom;

                //Binding html file to workbook

                var prter = app.ActivePrinter;
                app.DisplayAlerts = false;
                object paramMissing = Type.Missing;

                int length = sheetList.Count;
                //for (int i = length - 1; i >= 0; i--)
                for (int i = 0; i < length; i++)
                {
                    var htmlFile = CreateHtmlFile(sheetList[i].HtmlContent);

                    if (string.IsNullOrEmpty(htmlFile)) return string.Empty;

                    Microsoft.Office.Interop.Excel.Workbook workBookTemp =  workBooks.Open(
                        htmlFile,
                        0,
                        true,
                        5,
                        "",
                        "",
                        true,
                        Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                        false,
                        false,
                        0,
                        true,
                        1, 
                        0);

                    //save workbook to new real Excel file
                    Worksheet sheetTemp = (Worksheet)workBookTemp.Worksheets[1];

                    // sheet to copy after
                    Worksheet defaultWorksheet = workBook.Worksheets[i + 1] as Worksheet;
                    sheetTemp.Name = "Sheet" + i;
                    sheetTemp.Copy(Type.Missing, defaultWorksheet);

                    // sheet is copied and setting page
                    Worksheet sheetCopy = workBook.Worksheets[i + 2] as Worksheet;
                    sheetCopy.PageSetup.Zoom = false;
                    sheetCopy.PageSetup.FitToPagesWide = 1;
                    sheetCopy.PageSetup.CenterHorizontally = true;

                    //print option
                    switch (sheetList[i].PageOrientation)
                    {
                        case PageOrientation.Landscape:
                            sheetCopy.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                            break;
                        case PageOrientation.Portrait:
                            sheetCopy.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                            break;
                    }

                    if (sheetList[i].IsAutoFit)
                    {
                        sheetCopy.Columns.AutoFit();
                        var usedRange = sheetCopy.UsedRange;
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
                        Microsoft.Office.Interop.Excel.Range offset = sheetCopy.UsedRange.Offset;
                    }

                    switch (format)
                    {
                        case DocumentType.Xlsx:
                            marginLeftRight = app.InchesToPoints(0.5);
                            marginTopBottom = app.InchesToPoints(0.75);
                            sheetCopy.PageSetup.TopMargin = marginTopBottom;
                            sheetCopy.PageSetup.BottomMargin = marginTopBottom;
                            sheetCopy.PageSetup.LeftMargin = marginLeftRight;
                            sheetCopy.PageSetup.RightMargin = marginLeftRight;
                            break;

                        case DocumentType.Pdf:
                            marginTopBottom = app.InchesToPoints(0.5);
                            sheetCopy.PageSetup.TopMargin = marginTopBottom;
                            sheetCopy.PageSetup.BottomMargin = marginTopBottom;
                            sheetCopy.PageSetup.LeftMargin = 0;
                            sheetCopy.PageSetup.RightMargin = 0;
                            break;
                    }

                    workBookTemp.Close();

                    //Delete temp html file
                    if (File.Exists(htmlFile))
                    {
                        File.Delete(htmlFile);
                    }
                }

                // Delete sheet 1 what is empty
                Worksheet emptySheet = workBook.Worksheets[1] as Worksheet;
                emptySheet.Delete();

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
                
                return docFile;
            }           
            catch (Exception ex)
            {
                LogExtensions.Log(ex);
                return string.Empty;
            }
            finally
            {                
                workBooks.Close();
                app.Workbooks.Close();

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
