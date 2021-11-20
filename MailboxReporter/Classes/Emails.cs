using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using Lurgle.Logging;
using Microsoft.Exchange.WebServices.Data;

namespace MailboxReporter.Classes
{
    public static class Emails
    {
        public static async Task<EmailResult> ListEmails(string address)
        {
            Log.Information().Add("[{Address:l}] Connecting ...", address);
            var startTime = DateTime.Now;
            var sqlClient = new SqlServerClient();
            await sqlClient.Connect();

            var exchangeServer = new ExchangeService(ExchangeVersion.Exchange2016)
            {
                UseDefaultCredentials = true,
                AcceptGzipEncoding = Config.UseGzip,
                WebProxy = new WebProxy(),
                Credentials = !string.IsNullOrEmpty(Config.UserName) && !string.IsNullOrEmpty(Config.Password)
                    ? new WebCredentials(Config.UserName, Config.Password)
                    : new WebCredentials(),
                HttpHeaders = {new KeyValuePair<string, string>("X-AnchorMailbox", address)},
                Timeout = Config.ServerTimeout,
                UserAgent = $"MailboxReporter/v{Logging.Config.AppVersion}"
            };

            if (!string.IsNullOrEmpty(Config.Url))
                exchangeServer.Url = new Uri(Config.Url);
            else
                exchangeServer.AutodiscoverUrl(address, RedirectionUrlValidationCallback);

            var result = new EmailResult {ExchangeUrl = exchangeServer.Url.AbsoluteUri};

            try
            {
                var folderId = new FolderId(WellKnownFolderName.Inbox, new Mailbox(address));
                var inboxFolder = Folder.Bind(exchangeServer, folderId);
                var lastHours = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                        Config.LastTick.Subtract(TimeSpan.FromHours(Config.LastHours))));
                var lastHoursUnread = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                    new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived,
                        Config.LastTick.Subtract(TimeSpan.FromHours(Config.LastHours))),
                    new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));

                var itemCount = inboxFolder.TotalCount;
                var unreadCount = inboxFolder.UnreadCount;

                FindItemsResults<Item> fiResults;
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
                    itemCount = inboxFolder.FindItems(lastHours, new ItemView(1)).TotalCount;
                    unreadCount = inboxFolder.FindItems(lastHoursUnread, new ItemView(1)).TotalCount;
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

                    foreach (var emailRecord in fiResults
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
                            AttachmentCount = email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).Count(),
                            Attachments = string.Join(",",
                                email.Attachments.Where(x => !x.IsInline).Select(x => x.Name).ToArray()),
                            Size = email.Size
                        }))
                    {
                        await sqlClient.AddOrUpdate(address, emailRecord);
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
                    address, totalCount,
                    (DateTime.Now - startTime).TotalSeconds);
                fiResults.Items.Clear();
            }
            catch (Exception ex)
            {
                result.HttpHeaders = exchangeServer.HttpHeaders;
                result.HttpResponseHeaders = exchangeServer.HttpResponseHeaders;
                result.ServerInfo = exchangeServer.ServerInfo;
                result.ErrorInfo = ex;

                Log.Warning(ex).AddProperty("Result", result, true)
                    .Add("[{Address:l}] Error retrieving emails: {Message:l}", address, ex.Message);
            }

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