// --------------------------------------------------------------------------
// <copyright file="DocumentSetting.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Entities
{
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Runtime.Serialization;
    using Fanex.Workbook.Service.Business;

    /// <summary>
    /// Class DocumentSetting.
    /// </summary>    
    public class DocumentSetting
    {
    }

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    [DataContract]
    [SuppressMessage(
        "StyleCop.CSharp.MaintainabilityRules",
        "SA1402:FileMayOnlyContainASingleClass",
        Justification = "Reviewed. Suppression is OK here.")]
    public class PageSetting
    {
        /// <summary>
        /// Gets or sets the content of the CSS.
        /// </summary>
        /// <value>The content of the CSS.</value>
        [DataMember]
        public string CssContent { get; set; }

        /// <summary>
        /// Gets or sets The HTML content.
        /// </summary>
        /// <value>The content of the HTML.</value>
        [DataMember]
        public string HtmlContent { get; set; }

        /// <summary>
        /// Gets or sets The page orientation.
        /// </summary>
        /// <value>The page orientation.</value>
        [DataMember]
        public WorkBookManager.PageOrientation PageOrientation { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is automatic fit.
        /// </summary>
        /// <value><c>true</c> if this instance is automatic fit; otherwise, <c>false</c>.</value>
        [DataMember]
        public bool IsAutoFit { get; set; }
    }

    /// <summary>
    /// Class DocumentFile.
    /// </summary>
    [SuppressMessage(
        "StyleCop.CSharp.MaintainabilityRules",
        "SA1402:FileMayOnlyContainASingleClass",
        Justification = "Reviewed. Suppression is OK here.")]
    public class DocumentFile
    {
        /// <summary>
        /// Gets or sets The file name.
        /// </summary>
        /// <value>The name of the file.</value>
        public string FileName { get; set; }

        /// <summary>
        /// Gets or sets The document type.
        /// </summary>
        /// <value>The type of the document.</value>
        public WorkBookManager.DocumentType DocumentType { get; set; }

        /// <summary>
        /// Gets or sets The main page.
        /// </summary>
        /// <value>The main page.</value>
        public PageSetting MainPage { get; set; }

        /// <summary>
        /// Gets or sets The page detail list.
        /// </summary>
        /// <value>The page detail list.</value>
        public List<PageSetting> PageDetailList { get; set; }
    }
}
