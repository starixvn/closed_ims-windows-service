// -----------------------------------------------------------------------
// <copyright file="DocumentSetting.cs" company="Microsoft">
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
