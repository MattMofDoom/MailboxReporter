using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

// ReSharper disable UnusedAutoPropertyAccessor.Global

namespace MailboxReporter.Classes
{
    public class EmailResult
    {
        public string ExchangeUrl { get; set; }
        public ExchangeServerInfo ServerInfo { get; set; }
        public IDictionary<string, string> HttpHeaders { get; set; }
        public IDictionary<string, string> HttpResponseHeaders { get; set; }
        public Exception ErrorInfo { get; set; }
    }
}