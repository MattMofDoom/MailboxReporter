using System;
using Microsoft.Exchange.WebServices.Data;

// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace MailboxReporter.Classes
{
    public class EmailRecord
    {
        public EmailRecord()
        {
            Id = string.Empty;
            InternetMessageId = string.Empty;
            ConversationId = string.Empty;
            SentDate = DateTime.Now;
            ReceivedDate = DateTime.Now;
            CreatedDate = DateTime.Now;
            ModifiedDate = DateTime.Now;
            ModifiedName = string.Empty;
            FromName = string.Empty;
            FromAddress = string.Empty;
            ReplyToName = string.Empty;
            ReplyToAddress = string.Empty;
            ToName = string.Empty;
            ToAddress = string.Empty;
            CcName = string.Empty;
            CcAddress = string.Empty;
            Priority = Importance.Normal;
            Subject = string.Empty;
            Body = string.Empty;
            BodyType = BodyType.HTML;
            Size = 0;
            AttachmentCount = 0;
            Attachments = string.Empty;
            IsRead = false;
        }

        public string Id { get; set; }
        public string InternetMessageId { get; set; }
        public string ConversationId { get; set; }
        public DateTime SentDate { get; set; }
        public DateTime ReceivedDate { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public string ModifiedName { get; set; }
        public string FromName { get; set; }
        public string FromAddress { get; set; }
        public string ReplyToName { get; set; }
        public string ReplyToAddress { get; set; }
        public string ToName { get; set; }
        public string ToAddress { get; set; }
        public string CcName { get; set; }
        public string CcAddress { get; set; }
        public Importance Priority { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public BodyType BodyType { get; set; }
        public int Size { get; set; }
        public int AttachmentCount { get; set; }
        public string Attachments { get; set; }
        public bool IsRead { get; set; }
    }
}