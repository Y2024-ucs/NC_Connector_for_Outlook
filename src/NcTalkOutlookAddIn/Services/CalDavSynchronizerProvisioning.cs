// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal enum CalDavSynchronizerProvisioningOutcome
    {
        // Nothing to do: profile present and current, or an own profile of the user covers the calendar.
        Unchanged,
        // A new profile is needed; the caller backs up the Outlook calendar, then calls CreateProfile.
        CreateRequired,
        // A new profile was written; CalDav Synchronizer reads it on the next Outlook start.
        Created,
        // The stored password of our profile was renewed; also effective after a restart.
        PasswordUpdated,
        Skipped
    }

    /// <summary>
    /// Result of <see cref="CalDavSynchronizerProvisioning.Prepare"/>; carries what CreateProfile needs.
    /// </summary>
    internal sealed class CalDavSynchronizerProvisioningPlan
    {
        internal CalDavSynchronizerProvisioningOutcome Outcome { get; set; }
        internal string CalendarUrl { get; set; }
        internal string OptionsPath { get; set; }
    }

    /// <summary>
    /// Outlook-side facts gathered on the UI thread, so the file work can run on a pool thread.
    /// </summary>
    internal sealed class CalDavSynchronizerOutlookContext
    {
        internal string OutlookProfileName { get; set; }
        internal string CalendarEntryId { get; set; }
        internal string CalendarStoreId { get; set; }
        internal string AccountName { get; set; }
        internal string EmailAddress { get; set; }
    }

    /// <summary>
    /// Sets up the separately installed "Outlook CalDav Synchronizer" add-in for the user's personal
    /// Nextcloud calendar, so users need no manual configuration. Profiles the user created are never
    /// changed; a profile is created only when none covers the calendar URL or the Outlook calendar.
    /// CalDav Synchronizer reads its profiles only at Outlook start.
    /// </summary>
    internal static class CalDavSynchronizerProvisioning
    {
        internal const string AddInProgId = "CalDavSynchronizer.1";
        internal const string PersonalCalendarCollection = "personal";
        private const string ProfileName = "Nextcloud Kalender";
        private const string ProfilesFileName = "profiles.xml";
        private const string DefaultConfigFileName = "options.xml";
        private const string RegistryKey = @"Software\CalDavSynchronizer";
        private const string StateFileName = "caldav-synchronizer.xml";
        private const int SaltLength = 17;

        private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";
        private static readonly XNamespace Xsd = "http://www.w3.org/2001/XMLSchema";

        /// <summary>
        /// Decides what to do. Renews the credentials of our own profile directly; for a missing profile it
        /// only reports CreateRequired, so the Outlook calendar can be backed up before anything is written.
        /// </summary>
        internal static CalDavSynchronizerProvisioningPlan Prepare(
            TalkServiceConfiguration configuration,
            CalDavSynchronizerOutlookContext outlook)
        {
            var plan = new CalDavSynchronizerProvisioningPlan { Outcome = CalDavSynchronizerProvisioningOutcome.Skipped };
            if (configuration == null || !configuration.IsComplete() || outlook == null
                || string.IsNullOrWhiteSpace(outlook.OutlookProfileName)
                || string.IsNullOrWhiteSpace(outlook.CalendarEntryId))
            {
                return plan;
            }

            string calendarUrl = new DavDiscoveryService(configuration).DiscoverCalendarUrl(PersonalCalendarCollection);
            if (string.IsNullOrWhiteSpace(calendarUrl))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "CalDav Synchronizer provisioning skipped: personal calendar not found.");
                return plan;
            }
            plan.CalendarUrl = calendarUrl;

            DisableUpdateNotifications();
            string optionsPath = GetOrCreateOptionsPath(outlook.OutlookProfileName);
            XDocument options = LoadOptions(optionsPath);
            XElement root = options.Root;
            string ownId = LoadOwnProfileId(outlook.OutlookProfileName);

            XElement own = root.Elements("Options").FirstOrDefault(o =>
                ownId.Length > 0 && string.Equals((string)o.Element("Id"), ownId, StringComparison.OrdinalIgnoreCase));
            if (own != null)
            {
                plan.Outcome = RenewPasswordIfChanged(options, own, optionsPath, configuration, calendarUrl);
                return plan;
            }

            // The user's own profiles win: same calendar or same Outlook folder means it is set up already.
            XElement existing = root.Elements("Options").FirstOrDefault(o =>
                SameUrl((string)o.Element("CalenderUrl"), calendarUrl)
                || string.Equals((string)o.Element("OutlookFolderEntryId"), outlook.CalendarEntryId, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDav Synchronizer provisioning left an existing profile untouched (name=" + (string)existing.Element("Name") + ").");
                plan.Outcome = CalDavSynchronizerProvisioningOutcome.Unchanged;
                return plan;
            }

            plan.OptionsPath = optionsPath;
            plan.Outcome = CalDavSynchronizerProvisioningOutcome.CreateRequired;
            return plan;
        }

        /// <summary>
        /// Writes the new profile after the caller backed up the Outlook calendar. The options file is read
        /// again, so a profile the user added meanwhile still prevents a second one.
        /// </summary>
        internal static CalDavSynchronizerProvisioningOutcome CreateProfile(
            CalDavSynchronizerProvisioningPlan plan,
            TalkServiceConfiguration configuration,
            CalDavSynchronizerOutlookContext outlook)
        {
            if (plan == null || plan.Outcome != CalDavSynchronizerProvisioningOutcome.CreateRequired)
            {
                return CalDavSynchronizerProvisioningOutcome.Skipped;
            }

            XDocument options = LoadOptions(plan.OptionsPath);
            if (options.Root.Elements("Options").Any(o =>
                SameUrl((string)o.Element("CalenderUrl"), plan.CalendarUrl)
                || string.Equals((string)o.Element("OutlookFolderEntryId"), outlook.CalendarEntryId, StringComparison.OrdinalIgnoreCase)))
            {
                return CalDavSynchronizerProvisioningOutcome.Unchanged;
            }

            string id = Guid.NewGuid().ToString("D");
            options.Root.Add(BuildProfile(id, plan.CalendarUrl, configuration, outlook));
            SaveOptions(options, plan.OptionsPath);
            SaveOwnProfileId(outlook.OutlookProfileName, id);
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDav Synchronizer profile created (calendar=" + plan.CalendarUrl + ", outlookProfile=" + outlook.OutlookProfileName + ").");
            return CalDavSynchronizerProvisioningOutcome.Created;
        }

        private static CalDavSynchronizerProvisioningOutcome RenewPasswordIfChanged(
            XDocument options,
            XElement profile,
            string optionsPath,
            TalkServiceConfiguration configuration,
            string calendarUrl)
        {
            string stored = TryUnprotect((string)profile.Element("ProtectedPassword"), (string)profile.Element("Salt"));
            bool urlChanged = !SameUrl((string)profile.Element("CalenderUrl"), calendarUrl);
            bool userChanged = !string.Equals((string)profile.Element("UserName"), configuration.Username, StringComparison.Ordinal);
            if (string.Equals(stored, configuration.AppPassword, StringComparison.Ordinal) && !urlChanged && !userChanged)
            {
                return CalDavSynchronizerProvisioningOutcome.Unchanged;
            }

            string salt;
            string protectedPassword = Protect(configuration.AppPassword, out salt);
            SetValue(profile, "Salt", salt);
            SetValue(profile, "ProtectedPassword", protectedPassword);
            SetValue(profile, "UserName", configuration.Username);
            SetValue(profile, "CalenderUrl", calendarUrl);
            SaveOptions(options, optionsPath);
            DiagnosticsLogger.Log(LogCategories.Core, "CalDav Synchronizer profile credentials renewed.");
            return CalDavSynchronizerProvisioningOutcome.PasswordUpdated;
        }

        /// <summary>
        /// The IBP template, taken from a working profile: two-way merge of the Outlook default calendar,
        /// 60 days back and 365 ahead, Europe/Berlin, attendees, reminders and body mapped.
        /// </summary>
        private static XElement BuildProfile(
            string id,
            string calendarUrl,
            TalkServiceConfiguration configuration,
            CalDavSynchronizerOutlookContext outlook)
        {
            string salt;
            string protectedPassword = Protect(configuration.AppPassword, out salt);
            return new XElement("Options",
                new XElement("Inactive", "false"),
                new XElement("Name", ProfileName),
                new XElement("Id", id),
                new XElement("OutlookFolderEntryId", outlook.CalendarEntryId),
                new XElement("OutlookFolderStoreId", outlook.CalendarStoreId ?? string.Empty),
                new XElement("OutlookFolderAccountName", outlook.AccountName ?? string.Empty),
                new XElement("IgnoreSynchronizationTimeRange", "false"),
                new XElement("DaysToSynchronizeInThePast", "60"),
                new XElement("DaysToSynchronizeInTheFuture", "365"),
                new XElement("SynchronizationMode", "MergeInBothDirections"),
                new XElement("ConflictResolution", "Automatic"),
                new XElement("CalenderUrl", calendarUrl),
                new XElement("EmailAddress", outlook.EmailAddress ?? string.Empty),
                new XElement("UserName", configuration.Username),
                new XElement("SynchronizationIntervalInMinutes", "5"),
                new XElement("UseWebDavCollectionSync", "false"),
                new XElement("Salt", salt),
                new XElement("ProtectedPassword", protectedPassword),
                new XElement("UseAccountPassword", "false"),
                new XElement("ServerAdapterType", "Default"),
                new XElement("CloseAfterEachRequest", "false"),
                new XElement("PreemptiveAuthentication", "true"),
                new XElement("ForceBasicAuthentication", "true"),
                new XElement("EnableChangeTriggeredSynchronization", "true"),
                new XElement("IsChunkedSynchronizationEnabled", "true"),
                new XElement("ChunkSize", "100"),
                new XElement("ProxyOptions",
                    new XElement("ProxyUseDefault", "true"),
                    new XElement("ProxyUseManual", "false")),
                new XElement("MappingConfiguration",
                    new XAttribute(Xsi + "type", "EventMappingConfiguration"),
                    new XElement("MapReminder", "JustUpcoming"),
                    new XElement("MapSensitivityPrivateToClassConfidential", "false"),
                    new XElement("MapClassConfidentialToSensitivityPrivate", "false"),
                    new XElement("MapClassPublicToSensitivityPrivate", "false"),
                    new XElement("MapSensitivityPublicToDefault", "false"),
                    new XElement("MapAttendees", "true"),
                    new XElement("ScheduleAgentClient", "true"),
                    new XElement("SendNoAppointmentNotifications", "false"),
                    new XElement("OrganizerAsDelegate", "false"),
                    new XElement("MapBody", "true"),
                    new XElement("MapRtfBodyToXAltDesc", "false"),
                    new XElement("MapXAltDescToRtfBody", "false"),
                    new XElement("CreateEventsInUTC", "false"),
                    new XElement("UseIanaTz", "false"),
                    new XElement("EventTz", "Europe/Berlin"),
                    new XElement("IncludeHistoricalData", "false"),
                    new XElement("UseGlobalAppointmentID", "false"),
                    new XElement("IncludeEmptyEventCategoryFilter", "false"),
                    new XElement("InvertEventCategoryFilter", "false"),
                    new XElement("IsCategoryFilterSticky", "false"),
                    new XElement("CleanupDuplicateEvents", "false"),
                    new XElement("MapCustomProperties", "false"),
                    new XElement("MapEventColorToCategory", "false"),
                    new XElement("UserDefinedCustomPropertyMappings")),
                new XElement("ProfileTypeOrNull", "Nextcloud"));
        }

        /// <summary>
        /// Same scheme as CalDav Synchronizer's Options.Password: random 17-byte salt as entropy,
        /// UTF-16 password bytes, DPAPI for the current user.
        /// </summary>
        internal static string Protect(string password, out string salt)
        {
            var saltBytes = new byte[SaltLength];
            using (var random = RandomNumberGenerator.Create())
            {
                random.GetBytes(saltBytes);
            }
            salt = Convert.ToBase64String(saltBytes);
            byte[] data = ProtectedData.Protect(
                Encoding.Unicode.GetBytes(password ?? string.Empty),
                saltBytes,
                DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(data);
        }

        internal static string TryUnprotect(string protectedPassword, string salt)
        {
            try
            {
                byte[] data = ProtectedData.Unprotect(
                    Convert.FromBase64String(protectedPassword ?? string.Empty),
                    Convert.FromBase64String(salt ?? string.Empty),
                    DataProtectionScope.CurrentUser);
                return Encoding.Unicode.GetString(data);
            }
            catch (CryptographicException)
            {
                return null;
            }
            catch (FormatException)
            {
                return null;
            }
        }

        internal static bool SameUrl(string left, string right)
        {
            return string.Equals(
                (left ?? string.Empty).Trim().TrimEnd('/'),
                (right ?? string.Empty).Trim().TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase);
        }

        private static void SetValue(XElement profile, string name, string value)
        {
            XElement element = profile.Element(name);
            if (element == null)
            {
                profile.Add(new XElement(name, value ?? string.Empty));
            }
            else
            {
                element.Value = value ?? string.Empty;
            }
        }

        /// <summary>
        /// Mirrors CalDav Synchronizer's GetOrCreateDataDirectory: the entry for the Outlook profile in
        /// profiles.xml names the data directory and the options file.
        /// </summary>
        private static string GetOrCreateOptionsPath(string outlookProfileName)
        {
            string baseDirectory = GetDataDirectoryBase();
            Directory.CreateDirectory(baseDirectory);
            string profilesPath = Path.Combine(baseDirectory, ProfilesFileName);

            XDocument profiles = File.Exists(profilesPath)
                ? XDocument.Parse(File.ReadAllText(profilesPath))
                : new XDocument(new XElement("ArrayOfProfileEntry",
                    new XAttribute(XNamespace.Xmlns + "xsd", Xsd.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi", Xsi.NamespaceName)));
            XElement entry = profiles.Root.Elements("ProfileEntry").FirstOrDefault(e =>
                string.Equals((string)e.Attribute("ProfileName"), outlookProfileName, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
            {
                entry = new XElement("ProfileEntry",
                    new XAttribute("ProfileName", outlookProfileName),
                    new XAttribute("ConfigFileName", DefaultConfigFileName),
                    new XAttribute("DataDirectoryName", Guid.NewGuid().ToString()));
                profiles.Root.Add(entry);
                WriteSerializerStyle(profiles, profilesPath);
            }

            string dataDirectoryName = (string)entry.Attribute("DataDirectoryName");
            string dataDirectory = string.IsNullOrEmpty(dataDirectoryName)
                ? baseDirectory
                : Path.Combine(baseDirectory, dataDirectoryName);
            Directory.CreateDirectory(dataDirectory);
            string configFileName = (string)entry.Attribute("ConfigFileName");
            return Path.Combine(dataDirectory, string.IsNullOrEmpty(configFileName) ? DefaultConfigFileName : configFileName);
        }

        private static string GetDataDirectoryBase()
        {
            bool roaming = false;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKey))
            {
                object value = key != null ? key.GetValue("StoreAppDataInRoamingFolder") : null;
                roaming = value is int && (int)value != 0;
            }
            return Path.Combine(
                Environment.GetFolderPath(roaming
                    ? Environment.SpecialFolder.ApplicationData
                    : Environment.SpecialFolder.LocalApplicationData),
                "CalDavSynchronizer");
        }

        private static XDocument LoadOptions(string path)
        {
            if (!File.Exists(path))
            {
                return new XDocument(new XElement("ArrayOfOptions",
                    new XAttribute(XNamespace.Xmlns + "xsd", Xsd.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi", Xsi.NamespaceName)));
            }
            return XDocument.Parse(File.ReadAllText(path));
        }

        private static void SaveOptions(XDocument options, string path)
        {
            // CalDav Synchronizer keeps one backup per migration; keep ours next to it before changing.
            if (File.Exists(path) && !File.Exists(path + ".nc-connector.bak"))
            {
                File.Copy(path, path + ".nc-connector.bak");
            }
            WriteSerializerStyle(options, path);
        }

        /// <summary>
        /// Writes like CalDav Synchronizer's FileDataAccess: XmlSerializer text (utf-16 declaration)
        /// stored with File.WriteAllText. Written to a temporary file first, then swapped in.
        /// </summary>
        private static void WriteSerializerStyle(XDocument document, string path)
        {
            var builder = new StringBuilder();
            var settings = new XmlWriterSettings { Indent = true, Encoding = Encoding.Unicode };
            using (var writer = new StringWriter(builder))
            using (XmlWriter xml = XmlWriter.Create(writer, settings))
            {
                document.Save(xml);
            }
            string temp = path + ".tmp";
            File.WriteAllText(temp, builder.ToString());
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
        }

        private static void DisableUpdateNotifications()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryKey))
                {
                    // Updates come with the IBP package; users should not see update prompts.
                    if (key != null && key.GetValue("CheckForNewerVersions") == null)
                    {
                        key.SetValue("CheckForNewerVersions", 0, RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to disable CalDav Synchronizer update checks.", ex);
            }
        }

        private static string GetStatePath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), StateFileName);
        }

        private static string LoadOwnProfileId(string outlookProfileName)
        {
            try
            {
                string path = GetStatePath();
                if (!File.Exists(path)) return string.Empty;
                XElement profile = XDocument.Load(path).Root.Elements("Profile").FirstOrDefault(e =>
                    string.Equals((string)e.Attribute("outlookProfile"), outlookProfileName, StringComparison.OrdinalIgnoreCase));
                return profile != null ? ((string)profile.Attribute("id") ?? string.Empty) : string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CalDav Synchronizer provisioning state.", ex);
                return string.Empty;
            }
        }

        private static void SaveOwnProfileId(string outlookProfileName, string id)
        {
            string path = GetStatePath();
            XDocument document = File.Exists(path)
                ? XDocument.Load(path)
                : new XDocument(new XElement("CalDavSynchronizerProvisioning"));
            List<XElement> stale = document.Root.Elements("Profile").Where(e =>
                string.Equals((string)e.Attribute("outlookProfile"), outlookProfileName, StringComparison.OrdinalIgnoreCase)).ToList();
            stale.ForEach(e => e.Remove());
            document.Root.Add(new XElement("Profile",
                new XAttribute("outlookProfile", outlookProfileName),
                new XAttribute("id", id)));
            document.Save(path);
        }
    }
}
