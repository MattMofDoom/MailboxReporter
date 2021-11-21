using System;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Threading;
using Lurgle.Logging;
using MailboxReporter.Classes;
using MailboxReporter.Enums;
using Timer = System.Timers.Timer;

namespace MailboxReporter
{
    public partial class Reporter : ServiceBase
    {
        private const int Available = 0;
        private const int Locked = 1;
        private int _lockState;
        private Timer _reportTimer;

        public Reporter()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback += CertificateValidationCallBack;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            AppDomain.CurrentDomain.UnhandledException += UnhandledException;
            Logging.Init();
            if (!Config.IsDebug)
            {
                Log.Information().Add("{Service:l} v{Version:l} started", Logging.Config.AppName,
                    Logging.Config.AppVersion);
            }
            else
            {
                Log.Information().Add("{Service:l} v{Version:l} started (Debug mode)", Logging.Config.AppName,
                    Logging.Config.AppVersion);
                Log.Debug().AddProperty("Mailboxes", Config.Addresses.Select(x => x.Address).ToList(), true)
                    .Add("Configured mailboxes: {Mailboxes:l}");
                Log.Debug().Add("Exchange URL: {Url:l}",
                    string.IsNullOrEmpty(Config.Url) ? "Autodiscover" : Config.Url);
                Log.Debug().Add("Use Gzip: {UseGzip}", Config.UseGzip);
                Log.Debug().Add("Authentication: {AuthenticationType:l}",
                    string.IsNullOrEmpty(Config.UserName) && string.IsNullOrEmpty(Config.Password)
                        ? "Service identity"
                        : "Supplied credentials");
                Log.Debug().Add("First Run: {FirstRun}", Config.FirstRun);
                Log.Debug().Add("Include Partial Body: {IncludePartialBody}", Config.IncludePartialBody);
                Log.Debug().Add("Partial body length: {PartialBodyLength}", Config.PartialBodyLength);
                Log.Debug().Add("Last X Hours: {LastHours}", Config.LastHours);
                Log.Debug().Add("Server timeout: {Timeout} seconds", Config.ServerTimeout / 1000);
                Log.Debug().Add("Poll interval: {PollInterval}", Config.PollInterval);
                Log.Debug().Add("Backoff interval: {BackoffInterval}", Config.BackoffInterval);
                Log.Debug().Add("Last tick: {LastTick:l}", Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"));
            }

            _reportTimer = new Timer {Interval = 1000, AutoReset = false};
            _reportTimer.Elapsed += ReportTick;
            _reportTimer.Start();
        }

        protected override void OnStop()
        {
            _reportTimer.Stop();
            Log.Information().Add("{Service:l} v{Version:l} stopped (Exit Code: {ExitCode})",
                Logging.Config.AppName, Logging.Config.AppVersion, AppErrors.Success);
            Logging.Close();
        }

        private void ReportTick(object sender, EventArgs e)
        {
            var currentAddress = string.Empty;

            try
            {
                if (Interlocked.CompareExchange(ref _lockState, Locked, Available) != Available) return;

                if (_reportTimer != null && !_reportTimer.AutoReset)
                {
                    _reportTimer.Interval = 1000;
                    _reportTimer.AutoReset = true;
                    _reportTimer.Start();
                }

                var thisReportTick = DateTime.Now;
                var currentThreads = Config.Addresses.Where(i => thisReportTick >= i.NextInterval)
                    .ToList();

                if (currentThreads.Any())
                {
                    Log.Information().AddProperty("Addresses", currentThreads.Select(x => x.Address).ToList(), true)
                        .Add("Begin retrieving emails for {Addresses} ...");

                    var isError = false;

                    foreach (var thread in currentThreads)
                        try
                        {
                            currentAddress = thread.Address;
                            var result = Emails.ListEmails(thread.Address);
                            foreach (var mailbox in Config.Addresses.Where(mailbox =>
                                thread.Address == mailbox.Address))
                                mailbox.NextInterval = DateTime.Now.AddMilliseconds(Config.PollInterval);
                        }
                        catch (Exception ex)
                        {
                            isError = true;
                            foreach (var mailbox in Config.Addresses.Where(mailbox =>
                                currentAddress == mailbox.Address))
                                mailbox.NextInterval = DateTime.Now.AddMilliseconds(Config.BackoffInterval);

                            Log.Warning()
                                .Add(
                                    "[{Address:l}] Errors detected - setting backoff interval of {Interval} seconds for next poll",
                                    currentAddress, Config.BackoffInterval / 1000);

                            Log.Exception(ex).Add("[{Address:l}] Error retrieving emails: {Message:l}",
                                thread.Address, ex.Message);
                        }

                    Log.Information().AddProperty("Addresses", currentThreads.Select(x => x.Address).ToList(), true)
                        .AddProperty("TotalDuration", (DateTime.Now - thisReportTick).TotalSeconds)
                        .Add(
                            "Completed retrieving and updating emails for {Addresses}, Total duration {TotalDuration:N0} seconds!");

                    if (!isError)
                    {
                        Config.SetLastTick(thisReportTick);

                        if (Config.FirstRun)
                            Config.DisableFirstRun();
                    }
                }

                currentThreads.Clear();
            }
            catch (Exception ex)
            {
                Log.Exception(ex).Add("Error retrieving emails: {Message:l}",
                    ex.Message);
            }

            Interlocked.CompareExchange(ref _lockState, Available, Locked);
        }

        private void UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Log.Exception((Exception) e.ExceptionObject)
                .Add("An unhandled exception has occurred - service will stop. ({Message:l})",
                    e.ExceptionObject.GetType());

            StopError(AppErrors.UnhandledException);
        }

        private void StopError(AppErrors errorType)
        {
            _reportTimer.Stop();
            Log.Information().Add("{Service:l} v{Version:l} stopped (Exit Code: {ExitCode})",
                Logging.Config.AppName, Logging.Config.AppVersion, errorType);
            Logging.Close();
            Environment.Exit((int) errorType);
        }

        private static bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
    }
}