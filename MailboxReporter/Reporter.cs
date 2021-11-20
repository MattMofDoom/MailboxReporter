using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
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
                Log.Information().AddProperty("Mailboxes", Config.Addresses.Select(x => x.Address).ToList(), true)
                    .Add("Configured mailboxes: {Mailboxes:l}");
                Log.Information().Add("Exchange URL: {Url:l}",
                    string.IsNullOrEmpty(Config.Url) ? "Autodiscover" : Config.Url);
                Log.Information().Add("Use Gzip: {UseGzip}", Config.UseGzip);
                Log.Information().Add("Authentication: {AuthenticationType:l}",
                    string.IsNullOrEmpty(Config.UserName) && string.IsNullOrEmpty(Config.Password)
                        ? "Service identity"
                        : "Supplied credentials");
                Log.Information().Add("First Run: {FirstRun}", Config.FirstRun);
                Log.Information().Add("Include Partial Body: {IncludePartialBody}", Config.IncludePartialBody);
                Log.Information().Add("Partial body length: {PartialBodyLength}", Config.PartialBodyLength);
                Log.Information().Add("Last X Hours: {LastHours}", Config.LastHours);
                Log.Information().Add("Server timeout: {Timeout} seconds", Config.ServerTimeout / 1000);
                Log.Information().Add("Concurrent Threads: {ConcurrentThreads}", Config.ConcurrentThreads);
                Log.Information().Add("Poll interval: {PollInterval}", Config.PollInterval);
                Log.Information().Add("Backoff interval: {BackoffInterval}", Config.BackoffInterval);
                Log.Information().Add("Last tick: {LastTick:l}", Config.LastTick.ToString("dd MMM yyyy HH:mm:ss"));
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

        private async void ReportTick(object sender, EventArgs e)
        {
            try
            {
                if (Interlocked.CompareExchange(ref _lockState, Locked, Available) != Available) return;

                if (_reportTimer != null && !_reportTimer.AutoReset)
                {
                    _reportTimer.Interval = 60000;
                    _reportTimer.AutoReset = true;
                    _reportTimer.Start();
                }

                var thisReportTick = DateTime.Now;
                var isError = false;
                var currentThreads = Config.Addresses.Where(i => thisReportTick >= i.NextInterval)
                    .ToList();

                if (currentThreads.Any())
                {
                    Log.Information().AddProperty("Addresses", Config.Addresses, true)
                        .Add("Begin retrieving emails ...");

                    var throttler =
                        new SemaphoreSlim(
                            currentThreads.Count < Config.ConcurrentThreads
                                ? currentThreads.Count
                                : Config.ConcurrentThreads, Config.ConcurrentThreads);
                    var allTasks = new List<Task>();

                    foreach (var thread in currentThreads)
                    {
                        await throttler.WaitAsync();

                        allTasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                var result = await Emails.ListEmails(thread.Address);

                                if (result.ErrorInfo != null)
                                {
                                    foreach (var t in Config.Addresses.Where(t =>
                                        Config.Addresses[0].Address == thread.Address))
                                        Config.Addresses[0].NextInterval =
                                            DateTime.Now.AddMilliseconds(Config.BackoffInterval);

                                    Log.Information()
                                        .Add(
                                            "[{Address:l}] Errors detected - setting backoff interval of {Interval} seconds for next poll",
                                            thread.Address, Config.BackoffInterval / 1000);
                                    isError = true;
                                }
                                else
                                {
                                    foreach (var unused in Config.Addresses.Where(t =>
                                        Config.Addresses[0].Address == thread.Address))
                                        Config.Addresses[0].NextInterval =
                                            DateTime.Now.AddMilliseconds(Config.PollInterval);
                                }
                            }
                            catch (Exception ex)
                            {
                                Log.Exception(ex).Add("[{Address:l}] Error retrieving emails: {Message:l}",
                                    thread.Address, ex.Message);
                            }
                            finally
                            {
                                throttler.Release();
                            }
                        }));
                    }

                    await Task.WhenAll(allTasks);
                    Log.Information().AddProperty("Addresses", Config.Addresses, true)
                        .Add("Completed retrieving and updating emails, Total duration {TotalDuration:N0} seconds!",
                            (DateTime.Now - thisReportTick).TotalSeconds);

                    if (!isError)
                    {
                        Config.SetLastTick(thisReportTick);

                        if (Config.FirstRun)
                            Config.DisableFirstRun();
                    }
                }
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