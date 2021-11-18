using System;
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
            AppDomain.CurrentDomain.UnhandledException += UnhandledException;
            Logging.Init();
            Log.Information().Add("{Service:l} v{Version:l} started", Logging.Config.AppName,
                Logging.Config.AppVersion);

            Log.Information().AddProperty("Mailboxes", Config.Addresses, true).Add("Configured mailboxes: {Mailboxes:l}");
            Log.Information().Add("Exchange URL: {Url:l}", string.IsNullOrEmpty(Config.Url) ? "Autodiscover" : Config.Url);
            Log.Information().Add("Authentication: {AuthenticationType:l}", string.IsNullOrEmpty(Config.UserName) && string.IsNullOrEmpty(Config.Password) ? "Service identity" : "Supplied credentials");
            Log.Information().Add("First Run: {FirstRun}", Config.FirstRun);
            Log.Information().Add("Include Partial Body: {IncludePartialBody}", Config.IncludePartialBody);
            Log.Information().Add("Last tick: {LastTick:l}", Config.LastTick.ToString("dd MMM yyyy hh:mm:ss"));

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

                Log.Information().AddProperty("Addresses", Config.Addresses, true).Add("Begin retrieving emails ...");

                var tasks = Enumerable.Range(0, Config.Addresses.Count)
                    .Select(i => Emails.ListEmails(Config.Addresses[i]));
                await Task.WhenAll(tasks);

                Log.Information().AddProperty("Addresses", Config.Addresses, true)
                    .Add("Completed retrieving and updating emails!");

                Config.SetLastTick(thisReportTick);

                if (Config.FirstRun)
                    Config.DisableFirstRun();

                Interlocked.CompareExchange(ref _lockState, Available, Locked);
            }
            catch (Exception ex)
            {
                Log.Exception(ex).Add("Error retrieving emails: {Message:l}",
                    ex.Message);
            }

            //foreach (var address in Config.Addresses)
            //    try
            //    {
            //        _thisReportTick = DateTime.Now;
            //        Log.Information().Add("[{Address:l}] Begin retrieving emails ...", address);
            //        await Emails.ListEmails(address, Config.FirstRun, _lastReportTick);
            //        Log.Information().Add("[{Address:l}] Completed retrieving and updating emails!", address);
            //        _lastReportTick = _thisReportTick;
            //    }
            //    catch (Exception ex)
            //    {
            //        Log.Exception(ex).Add("[{Address:l}] Error retrieving emails: {Message:l}", address,
            //            ex.Message);
            //    }
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