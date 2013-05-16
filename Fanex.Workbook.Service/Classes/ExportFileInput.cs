// -----------------------------------------------------------------------
// <copyright file="ExportFileInput.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExportService.Classes
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Runtime.Serialization;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    [DataContract]
    public class ExportFileInput
    {
        [DataMember]
        public string htmlContent;
        [DataMember]
        public List<string> listCssFilePath;
        [DataMember]
        public ExportService.WorkBookExtension.DocumentType format;
        [DataMember]
        public ExportService.WorkBookExtension.PageOrientation pageOrientation;
    }
}
