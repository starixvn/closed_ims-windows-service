// -----------------------------------------------------------------------
// <copyright file="ExportFileInput.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

using System.Runtime.Serialization;
using Fanex.Workbook.Service.Business;
using System.Collections.Generic;

namespace Fanex.Workbook.Service.Entities
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    [DataContract]
    public class PageSetting
    {
        [DataMember]
        public string CssContent;

        [DataMember]
        public string HtmlContent;

        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation;

        [DataMember]
        public bool IsAutoFit;
    }

    public class DocumentFile
    {
        public string FileName;

        public WorkBookManager.DocumentType DocumentType;

        public PageSetting MainPage;

        public List<PageSetting> PageDetailList;
    }
}
