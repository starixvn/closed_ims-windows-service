// --------------------------------------------------------------------------
// <copyright file="MailExtension.cs" company="Nexcel Solutions Vietnam">
//     Copyright (c) Nexcel Solutions Vietnam. All rights reserved.
// </copyright>
// --------------------------------------------------------------------------

namespace Fanex.Workbook.Service.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Net.Mail;
    using Fanex.MailDeliver.Model;

    /// <summary>
    /// Class MailExtension.
    /// </summary>
    public class MailExtension
    {
        /// <summary>
        /// Sends the mail.
        /// </summary>
        /// <param name="to">Mail - To address.</param>
        /// <param name="title">The title.</param>
        /// <param name="subject">The subject.</param>
        /// <param name="body">The mail body.</param>
        public static void SendMail(string to, string title, string subject, string body)
        {
            try
            {
                var mailSetting = new MailSetting()
                    {
                        To = to,
                        Title = title,
                        Subject = subject,
                        Body = body
                    };

                DoSendMail(mailSetting);
            }
            catch (Exception ex)
            {
                FileExtension.WriteFile(ex.StackTrace);
            }
        }

        /// <summary>
        /// Does the send mail.
        /// </summary>
        /// <param name="mailSend">The mail send.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        [SuppressMessage(
            "StyleCop.CSharp.Nexcel.NexcelCustomRules",
            "SP2101:MethodMustNotContainMoreLinesThan",
            Justification = "Reviewed. Suppression is OK here.")]
        private static bool DoSendMail(MailSetting mailSend)
        {
            try
            {
                using (var client = new Fanex.MailDeliver.MailClient.API.MailDeliver())
                {
                    var mailItem = new Fanex.MailDeliver.Model.MailRecord
                    {
                        MailTo = mailSend.To,
                        MailCc = mailSend.Cc,
                        MailBcc = mailSend.Bcc,
                        Title = mailSend.Title,
                        Subject = mailSend.Subject,
                        Body = mailSend.Body
                    };

                    if (mailSend.Attachment != null && mailSend.Attachment.FileName != null)
                    {
                        mailItem.Attachments = new List<Fanex.MailDeliver.Model.MailAttachment>();
                        var attItem = new Fanex.MailDeliver.Model.MailAttachment
                        {
                            FileName = mailSend.Attachment.FileName,
                            Content = mailSend.Attachment.Content
                        };
                        if (mailSend.Attachment.ContentType.ToLower().IndexOf("excel") != -1)
                        {
                            attItem.ContentType = FileExtentions.MS_EXCEL;
                        }
                        else if (mailSend.Attachment.ContentType.ToLower().IndexOf("doc") != -1)
                        {
                            attItem.ContentType = FileExtentions.MS_DOC;
                        }
                        else
                        {
                            attItem.ContentType = FileExtentions.IMG_GIF;
                        }

                        mailItem.Attachments.Add(attItem);
                    }

                    if (mailSend.AltView != null)
                    {
                        mailItem.AlternateView = new List<MailLinkedResource>();
                        var linkedResource = mailSend.AltView.LinkedResources;
                        foreach (var linked in linkedResource)
                        {
                            var stream = linked.ContentStream;
                            byte[] buffer = new byte[stream.Length];
                            stream.Read(buffer, 0, buffer.Length);
                            var resource = new MailLinkedResource
                            {
                                FileName = string.Empty,
                                ImageType = FileExtentions.IMG_GIF,
                                ContentCID = linked.ContentId,
                                StreamContent = new MemoryStream(buffer)
                            };

                            mailItem.AlternateView.Add(resource);
                        }
                    }

                    client.SendMail(mailItem);
                }

                return true;
            }
            catch (Exception ex)
            {
                FileExtension.WriteFile(ex.StackTrace);
            }

            return false;
        }
    }

    /// <summary>
    /// Class MailSetting.
    /// </summary>
    [SuppressMessage(
        "StyleCop.CSharp.MaintainabilityRules",
        "SA1402:FileMayOnlyContainASingleClass",
        Justification = "Reviewed. Suppression is OK here.")]
    public class MailSetting
    {
        /// <summary>
        /// Gets or sets to.
        /// </summary>
        /// <value>To address.</value>
        public string To { get; set; }

        /// <summary>
        /// Gets or sets the cc.
        /// </summary>
        /// <value>The cc address.</value>
        public string Cc { get; set; }

        /// <summary>
        /// Gets or sets the BCC.
        /// </summary>
        /// <value>The BCC address.</value>
        public string Bcc { get; set; }

        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        /// <value>The title.</value>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the subject.
        /// </summary>
        /// <value>The subject.</value>
        public string Subject { get; set; }

        /// <summary>
        /// Gets or sets the body.
        /// </summary>
        /// <value>The mail body.</value>
        public string Body { get; set; }

        /// <summary>
        /// Gets or sets the attachment.
        /// </summary>
        /// <value>The attachment.</value>
        public MailAttachment Attachment { get; set; }

        /// <summary>
        /// Gets or sets the alt view.
        /// </summary>
        /// <value>The alt view.</value>
        public AlternateView AltView { get; set; }
    }

    /// <summary>
    /// Class MailAttachment.
    /// </summary>
    public class MailAttachment
    {
        /// <summary>
        /// Gets or sets the content.
        /// </summary>
        /// <value>The content.</value>
        public MemoryStream Content { get; set; }

        /// <summary>
        /// Gets or sets the type of the content.
        /// </summary>
        /// <value>The type of the content.</value>
        public string ContentType { get; set; }

        /// <summary>
        /// Gets or sets the name of the file.
        /// </summary>
        /// <value>The name of the file.</value>
        public string FileName { get; set; }
    }
}