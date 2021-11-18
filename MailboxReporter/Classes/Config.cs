using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;

namespace MailboxReporter.Classes
{
    public static class Config
    {
        public static readonly string UserName;
        public static readonly string Password;
        public static readonly string Url;
        public static readonly List<string> Addresses;
        public static bool FirstRun { get; private set; }
        public static readonly bool IncludePartialBody;
        public static readonly int PartialBodyLength;
        public static DateTime LastTick;

        static Config()
        {
            UserName = ConfigurationManager.AppSettings["UserName"];
            Password = ConfigurationManager.AppSettings["Password"];
            Url = ConfigurationManager.AppSettings["Url"];
            Addresses = (ConfigurationManager.AppSettings["Addresses"] ?? "")
                .Split(new[] {','}, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim()).ToList();
            FirstRun = GetBool(ConfigurationManager.AppSettings["FirstRun"]);
            IncludePartialBody = GetBool(ConfigurationManager.AppSettings["IncludePartialBody"]);
            PartialBodyLength = GetInt(ConfigurationManager.AppSettings["PartialBodyLength"]);
            if (PartialBodyLength < 0)
                PartialBodyLength = 100;
            LastTick = GetLastTick(ConfigurationManager.AppSettings["LastTick"]);
        }

        private static int GetInt(object sourceObject)
        {
            var sourceString = string.Empty;

            if (!Convert.IsDBNull(sourceObject))
                sourceString = (string)sourceObject;

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
            {
                settings.Add(
                    new KeyValueConfigurationElement("LastTick", lastTick.ToString(CultureInfo.CurrentCulture)));
            }
            else
            {
                settings["LastTick"].Value = lastTick.ToString(CultureInfo.CurrentCulture);
            }

            config.Save();
            LastTick = lastTick;
        }
    }
}