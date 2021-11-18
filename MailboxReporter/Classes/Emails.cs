using System;
using System.Linq;
using System.Net;
using Lurgle.Logging;
using Microsoft.Exchange.WebServices.Data;
using Task = System.Threading.Tasks.Task;

namespace MailboxReporter.Classes
{
    public static class Emails
    {
        public static async Task ListEmails(string address)
        {
            Log.Information().Add("[{Address:l}] Connecting ...", address);
            var sqlClient = new SqlServerClient();
            await sqlClient.Connect();

            var exchangeServer = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                UseDefaultCredentials = true,
                AcceptGzipEncoding = true,
                WebProxy = new WebProxy(),
                Credentials = !string.IsNullOrEmpty(Config.UserName) && !string.IsNullOrEmpty(Config.Password)
                    ? new WebCredentials(Config.UserName, Config.Password)
                    : new WebCredentials()
            };

            if (!string.IsNullOrEmpty(Config.Url))
                exchangeServer.Url = new Uri(Config.Url);
            else
                exchangeServer.AutodiscoverUrl(address, RedirectionUrlValidationCallback);

            var folderId = new FolderId(WellKnownFolderName.Inbox, new Mailbox(address));
            var inboxFolder = Folder.Bind(exchangeServer, folderId);
            var lastDay = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                    Config.LastTick.Subtract(TimeSpan.FromDays(1))));
            var lastDayUnread = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                    Config.LastTick.Subtract(TimeSpan.FromDays(1))),
                new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

            //Make an initial connection to get the total emails
            var itemCount = inboxFolder.TotalCount;
            var unreadCount = inboxFolder.UnreadCount;

            var emailProps = new PropertySet(BasePropertySet.FirstClassProperties);
            FindItemsResults<Item> fiResults;
            var iv = itemCount > 1000 ? new ItemView(1000) : itemCount > 0 ? new ItemView(itemCount) : new ItemView(1);
            var pageCount = 1;
            var totalCount = 0;

            Log.Information().Add("[{Address:l}] Total emails found: {ItemCount}, unread: {UnreadCount} ...",
                address, itemCount, unreadCount);

            if (!Config.FirstRun)
            {
                itemCount = exchangeServer.FindItems(folderId, lastDay, new ItemView(1)).TotalCount;
                unreadCount = exchangeServer.FindItems(folderId, lastDayUnread, new ItemView(1)).TotalCount;
                iv = itemCount > 1000 ? new ItemView(1000) : itemCount > 0 ? new ItemView(itemCount) : new ItemView(1);

                Log.Information().Add(
                    "[{Address:l}] Retrieving last 24 hours from {LastReportTick:l} - Emails found: {ItemCount}, unread: {UnreadCount}  ...",
                    address, Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"), itemCount, unreadCount);
            }
            else
            {
                Log.Information().Add("[{Address:l}] First run flag is set, retrieving ALL emails ...", address);
            }

            do
            {
                fiResults = Config.FirstRun
                    ? inboxFolder.FindItems(iv)
                    : exchangeServer.FindItems(folderId, lastDay, iv);
                iv.Offset += fiResults.Items.Count;

                if (fiResults.Items.Count <= 0) continue;
                Log.Information().Add("[{Address:l}] Page {PageCount}, Emails found: {ItemCount} ...",
                    address, pageCount, fiResults.Items.Count);

                foreach (var emailRecord in fiResults
                    .Select(mailItem => EmailMessage.Bind(exchangeServer, mailItem.Id, emailProps))
                    .Select(email => new EmailRecord
                    {
                        Id = email.Id.UniqueId,
                        ConversationId = email.ConversationId.UniqueId,
                        InternetMessageId = email.InternetMessageId,
                        SentDate = email.DateTimeCreated,
                        ReceivedDate = email.DateTimeReceived,
                        CreatedDate = email.DateTimeCreated,
                        ModifiedDate = email.LastModifiedTime,
                        ModifiedName = email.LastModifiedName,
                        BodyType = email.Body.BodyType,
                        Body = Config.IncludePartialBody && !string.IsNullOrEmpty(email.Body)
                            ? email.Body.Text.Length < Config.PartialBodyLength ? email.Body.Text :
                            email.Body.Text.Substring(0, Config.PartialBodyLength)
                            : "",
                        Subject = email.Subject,
                        FromName = email.Sender.Name,
                        FromAddress = email.Sender.Address,
                        ReplyToName = string.Join(",", email.ReplyTo.Select(x => x.Name).ToArray()),
                        ReplyToAddress = string.Join(",", email.ReplyTo.Select(x => x.Address).ToArray()),
                        ToName = string.Join(",", email.ToRecipients.Select(x => x.Name).ToArray()),
                        ToAddress = string.Join(",", email.ToRecipients.Select(x => x.Address).ToArray()),
                        CcName = string.Join(",", email.CcRecipients.Select(x => x.Name).ToArray()),
                        CcAddress = string.Join(",", email.CcRecipients.Select(x => x.Address).ToArray()),
                        IsRead = email.IsRead,
                        Priority = email.Importance,
                        AttachmentCount = email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).Count(),
                        Attachments = string.Join(",",
                            email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).ToArray()),
                        Size = email.Size
                    }))
                {
                    await sqlClient.AddOrUpdate(address, emailRecord);
                    totalCount++;
                }

                Log.Information().Add("[{Address:l}] Result set tally, Page {Page}: {TotalCount} items", address,
                    pageCount, totalCount);
                pageCount++;
            } while (fiResults.MoreAvailable);

            sqlClient.Disconnect();
            sqlClient.Dispose();

            Log.Information().Add("[{Address:l}] Completed!", address);
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return new Uri(redirectionUrl).Scheme.Equals("https", StringComparison.OrdinalIgnoreCase);
        }
    }
}