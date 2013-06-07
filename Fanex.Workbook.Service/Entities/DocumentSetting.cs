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
    public class DocumentSetting
    {
        //[DataMember]
        //public List<PageSetting> PageList;
        [DataMember]
        public string CssContent;

        [DataMember]
        public string HtmlContent;

        [DataMember]
        public WorkBookManager.DocumentType DocumentType;

        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation;
        //public DocumentSetting()
        //{
        //    PageList = new List<PageSetting>();
        //}
    }

    [DataContract]
    public class PageSetting
    {
        [DataMember]
        public string HtmlContent;

        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation;
    }

}
