// -----------------------------------------------------------------------
// <copyright file="MailExtension.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace Fanex.Workbook.Service.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net.Mail;
    using Fanex.MailDeliver.Model;

    /// <summary>
    /// Mail extensions.
    /// </summary>
    public class MailExtension
    {
        /// <summary>
        /// Send mail.
        /// </summary>
        /// <param name="to"></param>
        /// <param name="title"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
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
        /// Send mail method.
        /// </summary>
        /// <param name="mailSend"></param>
        /// <returns></returns>
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
                    //// send mail
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
    /// Mail setting.
    /// </summary>
    public class MailSetting
    {
        public string To 
        {
            get; 
            set; 
        }

        public string Cc 
        { 
            get; 
            set; 
        }

        public string Bcc 
        { 
            get; 
            set; 
        }

        public string Title 
        { 
            get; 
            set; 
        }

        public string Subject 
        { 
            get; 
            set; 
        }

        public string Body 
        { 
            get; 
            set; 
        }

        public MailAttachment Attachment 
        { 
            get; 
            set; 
        }

        public AlternateView AltView 
        { 
            get; 
            set; 
        }
    }

    /// <summary>
    /// Mail attachment.
    /// </summary>
    public class MailAttachment
    {
        public MemoryStream Content 
        { 
            get;
            set; 
        }

        public string ContentType 
        { 
            get; 
            set; 
        }

        public string FileName 
        {
            get; 
            set; 
        }
    }
}
