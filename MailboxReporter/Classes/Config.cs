using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;

namespace MailboxReporter.Classes
{
    public static class Config
    {
        public static readonly bool IsDebug;
        public static readonly string UserName;
        public static readonly string Password;
        public static readonly string Url;

        // ReSharper disable once FieldCanBeMadeReadOnly.Global
        public static List<MailboxSetting> Addresses = new List<MailboxSetting>();
        public static readonly bool UseGzip;
        public static readonly bool IncludePartialBody;
        public static readonly int PartialBodyLength;
        public static readonly int LastHours;
        public static readonly int ServerTimeout;
        public static readonly int PollInterval;
        public static readonly int BackoffInterval;
        public static DateTime LastTick;


        static Config()
        {
            IsDebug = GetBool(ConfigurationManager.AppSettings["IsDebug"]);
            UserName = ConfigurationManager.AppSettings["UserName"];
            Password = ConfigurationManager.AppSettings["Password"];
            Url = ConfigurationManager.AppSettings["Url"];
            foreach (var mailbox in (ConfigurationManager.AppSettings["Addresses"] ?? "")
                .Split(new[] {','}, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).ToList())
                Addresses.Add(new MailboxSetting
                    {Address = mailbox, NextInterval = DateTime.Now.AddSeconds(5)});
            UseGzip = GetBool(ConfigurationManager.AppSettings["UseGzip"]);
            FirstRun = GetBool(ConfigurationManager.AppSettings["FirstRun"]);
            IncludePartialBody = GetBool(ConfigurationManager.AppSettings["IncludePartialBody"]);
            PartialBodyLength = GetInt(ConfigurationManager.AppSettings["PartialBodyLength"]);
            if (PartialBodyLength < 0)
                PartialBodyLength = 100;
            LastHours = GetInt(ConfigurationManager.AppSettings["LastHours"]);
            if (LastHours <= 0)
                LastHours = 24;
            ServerTimeout = GetInt(ConfigurationManager.AppSettings["ServerTimeout"]);
            if (ServerTimeout <= 0 || ServerTimeout > 3600)
                ServerTimeout = 200000;
            else
                ServerTimeout *= 1000;
            PollInterval = GetInt(ConfigurationManager.AppSettings["PollInterval"]);
            if (PollInterval <= 0 || PollInterval > 3600)
                PollInterval = 60000;
            else
                PollInterval *= 1000;
            BackoffInterval = GetInt(ConfigurationManager.AppSettings["BackoffInterval"]);
            if (BackoffInterval <= 0 || BackoffInterval <= PollInterval || BackoffInterval > 3600)
                BackoffInterval = PollInterval * 10;
            LastTick = GetLastTick(ConfigurationManager.AppSettings["LastTick"]);
        }

        public static bool FirstRun { get; private set; }

        private static int GetInt(object sourceObject)
        {
            var sourceString = string.Empty;

            if (!Convert.IsDBNull(sourceObject))
                sourceString = (string) sourceObject;

            if (int.TryParse(sourceString, out var destInt))
                return destInt;

            return -1;
        }

        private static bool GetBool(object sourceObject)
        {
            var sourceString = string.Empty;

            if (!Convert.IsDBNull(sourceObject))
                sourceString = (string) sourceObject;

            return bool.TryParse(sourceString, out var destBool) && destBool;
        }

        public static void DisableFirstRun()
        {
            FirstRun = false;
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = config.AppSettings.Settings;
            settings["FirstRun"].Value = "False";
            config.Save();
        }

        private static DateTime GetLastTick(object sourceObject)
        {
            var sourceString = string.Empty;

            if (!Convert.IsDBNull(sourceObject))
                sourceString = (string) sourceObject;

            if (DateTime.TryParse(sourceString, out var lastTick))
                return lastTick;

            var nowTick = DateTime.Now;
            SetLastTick(nowTick);
            return nowTick;
        }

        public static void SetLastTick(DateTime lastTick)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = config.AppSettings.Settings;
            if (!settings.AllKeys.Contains("LastTick"))
                settings.Add(
                    new KeyValueConfigurationElement("LastTick", lastTick.ToString(CultureInfo.CurrentCulture)));
            else
                settings["LastTick"].Value = lastTick.ToString(CultureInfo.CurrentCulture);

            config.Save();
            LastTick = lastTick;
        }
    }
}