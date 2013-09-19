// -----------------------------------------------------------------------
// <copyright file="DocumentSetting.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace Fanex.Workbook.Service.Entities
{
    using System.Runtime.Serialization;
    using Fanex.Workbook.Service.Business;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    [DataContract]
    public class DocumentSetting
    {
        [DataMember]
        public string CssContent;

        [DataMember]
        public string HtmlContent;

        [DataMember]
        public WorkBookManager.DocumentType DocumentType;

        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation;

        [DataMember]
        public bool IsAutoFit;
    }
}
