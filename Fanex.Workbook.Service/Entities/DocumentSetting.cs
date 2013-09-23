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
        public string CssContent
        {
            get;
            set;
        }

        [DataMember]
        public string HtmlContent
        {
            get;
            set;
        }

        [DataMember]
        public WorkBookManager.DocumentType DocumentType
        {
            get;
            set;
        }

        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation
        {
            get;
            set;
        }

        [DataMember]
        public bool IsAutoFit
        {
            get;
            set;
        }
    }
}
