using System;
using System.Collections.Generic;
using System.Linq;
using Lurgle.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MailboxReporter.Classes
{
    public static class Emails
    {
        public static EmailResult ListEmails(string address)
        {
            var lastHours = new SearchFilter.SearchFilterCollection(
                LogicalOperator.And,
                new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                    Config.LastTick.Subtract(TimeSpan.FromHours(Config.LastHours))));

            var lastHoursUnread =
                new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                        Config.LastTick.Subtract(TimeSpan.FromHours(Config.LastHours))),
                    new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

            var exchangeServer =
                new ExchangeService(ExchangeVersion.Exchange2013_SP1, TimeZoneInfo.Local)
                {
                    UseDefaultCredentials = true,
                    AcceptGzipEncoding = Config.UseGzip,
                    WebProxy = null,
                    Credentials = !string.IsNullOrEmpty(Config.UserName) && !string.IsNullOrEmpty(Config.Password)
                        ? new WebCredentials(Config.UserName, Config.Password)
                        : new WebCredentials(),
                    Timeout = Config.ServerTimeout,
                    HttpHeaders = {new KeyValuePair<string, string>("X-AnchorMailbox", address)},
                    UserAgent = $"MailboxReporter/v{Logging.Config.AppVersion}"
                };

            if (!string.IsNullOrEmpty(Config.Url))
                exchangeServer.Url = new Uri(Config.Url);
            else
                exchangeServer.AutodiscoverUrl(address, RedirectionUrlValidationCallback);

            Log.Information().Add("[{Address:l}] Connecting ...", address);
            var startTime = DateTime.Now;
            var sqlClient = new SqlServerClient();
            sqlClient.Connect();

            var result = new EmailResult {ExchangeUrl = exchangeServer.Url.AbsoluteUri};

            FindItemsResults<Item> fiResults;

            if (Config.IsDebug)
                Log.Debug().Add("[{Address:l}] Connect to mailbox ...", address);
            var folderId = new FolderId(WellKnownFolderName.Inbox, new Mailbox(address));
            var inboxFolder = Folder.Bind(exchangeServer, folderId);
            var itemCount = inboxFolder.TotalCount;
            var unreadCount = inboxFolder.UnreadCount;

            var iv = itemCount > 1000 ? new ItemView(1000) :
                itemCount > 0 ? new ItemView(itemCount) : new ItemView(1);
            var pageCount = 1;
            var totalCount = 0;

            result.HttpHeaders = exchangeServer.HttpHeaders;
            result.HttpResponseHeaders = exchangeServer.HttpResponseHeaders;
            result.ServerInfo = exchangeServer.ServerInfo;

            Log.Information().AddProperty("Result", result, true).Add(
                "[{Address:l}] Total emails found: {ItemCount}, unread: {UnreadCount} ...",
                address, itemCount, unreadCount);

            if (!Config.FirstRun)
            {
                if (Config.IsDebug)
                    Log.Debug().Add(
                        "[{Address:l}] Search for total items for past {Hours} hours from {LastReportTick:l} ...",
                        address, Config.LastHours, Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"));
                fiResults = inboxFolder.FindItems(lastHours, new ItemView(1));
                itemCount = fiResults.TotalCount;

                if (Config.IsDebug)
                    Log.Debug().Add(
                        "[{Address:l}] Search for unread items for past {Hours} hours from {LastReportTick:l} ...",
                        address, Config.LastHours, Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"));
                fiResults = inboxFolder.FindItems(lastHoursUnread, new ItemView(1));
                unreadCount = fiResults.TotalCount;

                iv = itemCount > 1000 ? new ItemView(1000) :
                    itemCount > 0 ? new ItemView(itemCount) : new ItemView(1);

                result.HttpHeaders = exchangeServer.HttpHeaders;
                result.HttpResponseHeaders = exchangeServer.HttpResponseHeaders;
                result.ServerInfo = exchangeServer.ServerInfo;

                Log.Information().AddProperty("Result", result, true).Add(
                    "[{Address:l}] Retrieving last {Hours} hours from {LastReportTick:l} - Emails found: {ItemCount}, unread: {UnreadCount}  ...",
                    address, Config.LastHours, Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"), itemCount,
                    unreadCount);
            }
            else
            {
                Log.Information().Add("[{Address:l}] First run flag is set, retrieving ALL emails ...", address);
            }

            var now = DateTime.Now;
            iv.PropertySet = BasePropertySet.FirstClassProperties;

            do
            {
                fiResults = Config.FirstRun
                    ? inboxFolder.FindItems(iv)
                    : inboxFolder.FindItems(lastHours, iv);
                iv.Offset += fiResults.Items.Count;

                result.HttpHeaders = exchangeServer.HttpHeaders;
                result.HttpResponseHeaders = exchangeServer.HttpResponseHeaders;
                result.ServerInfo = exchangeServer.ServerInfo;

                if (fiResults.Items.Count <= 0) continue;
                Log.Information().AddProperty("Result", result).Add(
                    "[{Address:l}] Page {PageCount}, Emails found: {ItemCount} ...",
                    address, pageCount, fiResults.Items.Count);

                foreach (var emailRecord in fiResults.Items
                             .Select(mailItem => EmailMessage.Bind(exchangeServer, mailItem.Id,
                                 new PropertySet(BasePropertySet.FirstClassProperties)))
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
                                 AttachmentCount =
                                     email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).Count(),
                                 Attachments = string.Join(",",
                                     email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).ToArray()),
                                 Size = email.Size
                             }))
                {
                    sqlClient.AddOrUpdate(address, emailRecord);
                    totalCount++;
                }

                result.HttpHeaders = exchangeServer.HttpHeaders;
                result.HttpResponseHeaders = exchangeServer.HttpResponseHeaders;
                result.ServerInfo = exchangeServer.ServerInfo;

                Log.Information().AddProperty("Result", result, true).Add(
                    "[{Address:l}] Result set tally, Page {Page}: {CurrentCount} items of {TotalCount}, Duration: {Duration:N0} seconds",
                    address,
                    pageCount, totalCount, itemCount, (DateTime.Now - now).TotalSeconds);

                now = DateTime.Now;
                pageCount++;
            } while (fiResults.MoreAvailable);

            Log.Information().Add(
                "[{Address:l}] Mailbox Completed, Emails: {TotalCount}, Duration: {Duration:N0} seconds!",
                address, totalCount, (DateTime.Now - startTime).TotalSeconds);
            fiResults.Items.Clear();

            sqlClient.Disconnect();
            sqlClient.Dispose();

            return result;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return new Uri(redirectionUrl).Scheme.Equals("https", StringComparison.OrdinalIgnoreCase);
        }
    }
}