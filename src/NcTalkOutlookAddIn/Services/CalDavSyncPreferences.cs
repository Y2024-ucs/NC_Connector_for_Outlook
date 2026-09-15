// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal enum CalDavSyncDirection
    {
        TwoWay = 0,
        NextcloudToOutlook = 1,
        OutlookToNextcloud = 2
    }

    internal sealed class CalDavSyncPreferences
    {
        private const string FileName = "caldav-sync.xml";

        internal CalDavSyncPreferences()
        {
            Configured = false;
            Enabled = true;
            Direction = CalDavSyncDirection.TwoWay;
            UseDefaultCalendarFolder = false;
            SyncOnStartup = true;
            IntervalMinutes = 15;
            SelectedCalendarHrefs = new List<string>();
        }

        internal bool Configured { get; set; }
        internal bool Enabled { get; set; }
        internal CalDavSyncDirection Direction { get; set; }
        internal bool UseDefaultCalendarFolder { get; set; }
        internal bool SyncOnStartup { get; set; }
        internal int IntervalMinutes { get; set; }
        internal List<string> SelectedCalendarHrefs { get; private set; }

        internal bool IsCalendarSelected(string href)
        {
            return !string.IsNullOrWhiteSpace(href)
                && SelectedCalendarHrefs.Any(v => string.Equals(v, href, StringComparison.OrdinalIgnoreCase));
        }

        internal void SetSelectedCalendars(IEnumerable<string> hrefs)
        {
            SelectedCalendarHrefs.Clear();
            foreach (string href in hrefs ?? Enumerable.Empty<string>())
            {
                string value = (href ?? string.Empty).Trim();
                if (value.Length == 0 || IsCalendarSelected(value))
                {
                    continue;
                }
                SelectedCalendarHrefs.Add(value);
            }
        }

        internal static CalDavSyncPreferences Load()
        {
            var preferences = new CalDavSyncPreferences();
            string path = GetPath();
            if (!File.Exists(path))
            {
                return preferences;
            }

            try
            {
                var document = new XmlDocument();
                document.Load(path);
                XmlElement root = document.DocumentElement;
                if (root == null || !string.Equals(root.Name, "CalDavSync", StringComparison.OrdinalIgnoreCase))
                {
                    return preferences;
                }

                preferences.Configured = ReadBool(root, "Configured", false);
                preferences.Enabled = ReadBool(root, "Enabled", true);
                preferences.UseDefaultCalendarFolder = ReadBool(root, "UseDefaultCalendarFolder", false);
                preferences.SyncOnStartup = ReadBool(root, "SyncOnStartup", true);
                preferences.IntervalMinutes = Math.Max(1, ReadInt(root, "IntervalMinutes", 15));

                CalDavSyncDirection direction;
                string directionText = ReadText(root, "Direction");
                preferences.Direction = Enum.TryParse(directionText, true, out direction)
                    ? direction
                    : CalDavSyncDirection.TwoWay;

                XmlElement selected = root["SelectedCalendars"];
                if (selected != null)
                {
                    foreach (XmlNode node in selected.SelectNodes("Calendar"))
                    {
                        XmlElement element = node as XmlElement;
                        if (element == null)
                        {
                            continue;
                        }
                        string href = (element.GetAttribute("href") ?? string.Empty).Trim();
                        if (href.Length > 0 && !preferences.IsCalendarSelected(href))
                        {
                            preferences.SelectedCalendarHrefs.Add(href);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CalDAV sync preferences.", ex);
            }
            return preferences;
        }

        internal void Save()
        {
            string path = GetPath();
            try
            {
                var document = new XmlDocument();
                XmlElement root = document.CreateElement("CalDavSync");
                document.AppendChild(root);
                AppendBool(document, root, "Configured", Configured);
                AppendBool(document, root, "Enabled", Enabled);
                AppendText(document, root, "Direction", Direction.ToString());
                AppendBool(document, root, "UseDefaultCalendarFolder", UseDefaultCalendarFolder);
                AppendBool(document, root, "SyncOnStartup", SyncOnStartup);
                AppendText(document, root, "IntervalMinutes", Math.Max(1, IntervalMinutes).ToString(CultureInfo.InvariantCulture));

                XmlElement selected = document.CreateElement("SelectedCalendars");
                foreach (string href in SelectedCalendarHrefs)
                {
                    XmlElement calendar = document.CreateElement("Calendar");
                    calendar.SetAttribute("href", href ?? string.Empty);
                    selected.AppendChild(calendar);
                }
                root.AppendChild(selected);

                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    Encoding = new UTF8Encoding(false)
                };
                using (XmlWriter writer = XmlWriter.Create(path, settings))
                {
                    document.Save(writer);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CalDAV sync preferences.", ex);
                throw;
            }
        }

        private static string GetPath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
        }

        private static bool ReadBool(XmlElement root, string name, bool fallback)
        {
            bool value;
            string text = ReadText(root, name);
            return bool.TryParse(text, out value) ? value : fallback;
        }

        private static int ReadInt(XmlElement root, string name, int fallback)
        {
            int value;
            string text = ReadText(root, name);
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
                ? value
                : fallback;
        }

        private static string ReadText(XmlElement root, string name)
        {
            XmlElement element = root[name];
            return element != null ? element.InnerText ?? string.Empty : string.Empty;
        }

        private static void AppendBool(XmlDocument document, XmlElement root, string name, bool value)
        {
            AppendText(document, root, name, value.ToString(CultureInfo.InvariantCulture));
        }

        private static void AppendText(XmlDocument document, XmlElement root, string name, string value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value ?? string.Empty;
            root.AppendChild(element);
        }
    }

    internal sealed class CalDavEventRecord
    {
        internal string CalendarHref { get; set; }
        internal string CalendarName { get; set; }
        internal string Href { get; set; }
        internal string ETag { get; set; }
        internal string Uid { get; set; }
        internal string Subject { get; set; }
        internal string Location { get; set; }
        internal string Description { get; set; }
        internal DateTime Start { get; set; }
        internal DateTime End { get; set; }
        internal bool AllDay { get; set; }
        internal bool Recurring { get; set; }
        internal bool Cancelled { get; set; }
    }

    internal sealed class CalDavEventSyncService
    {
        private const string UidPropertyName = "NC-CalDAV-UID";
        private const string HrefPropertyName = "NC-CalDAV-HREF";
        private const string ETagPropertyName = "NC-CalDAV-ETAG";
        private const string CalendarPropertyName = "NC-CalDAV-CALENDAR";
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CalDav = "urn:ietf:params:xml:ns:caldav";

        private readonly TalkServiceConfiguration _configuration;

        internal CalDavEventSyncService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            _configuration = configuration;
        }

        internal IList<CalDavEventRecord> DownloadEvents(CalDavCalendar calendar)
        {
            if (calendar == null || string.IsNullOrWhiteSpace(calendar.Href))
            {
                throw new ArgumentException("A CalDAV calendar is required.", "calendar");
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<cal:calendar-query xmlns:d=\"DAV:\" xmlns:cal=\"urn:ietf:params:xml:ns:caldav\">"
                + "<d:prop><d:getetag/><cal:calendar-data/></d:prop>"
                + "<cal:filter><cal:comp-filter name=\"VCALENDAR\"><cal:comp-filter name=\"VEVENT\"/></cal:comp-filter></cal:filter>"
                + "</cal:calendar-query>";

            XDocument document = SendDavRequest(calendar.Href, "REPORT", "1", body);
            var result = new List<CalDavEventRecord>();
            foreach (XElement responseElement in document.Descendants(Dav + "response"))
            {
                XElement prop = responseElement
                    .Elements(Dav + "propstat")
                    .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                    .Select(p => p.Element(Dav + "prop"))
                    .FirstOrDefault(p => p != null);
                if (prop == null)
                {
                    continue;
                }

                string ical = (string)prop.Element(CalDav + "calendar-data");
                if (string.IsNullOrWhiteSpace(ical))
                {
                    continue;
                }

                CalDavEventRecord record = ParseEvent(ical);
                if (record == null)
                {
                    continue;
                }
                record.CalendarHref = calendar.Href;
                record.CalendarName = calendar.DisplayName;
                record.Href = ResolveUri(calendar.Href, (string)responseElement.Element(Dav + "href"));
                record.ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim();
                result.Add(record);
            }
            return result;
        }

        internal int ImportIntoOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IEnumerable<CalDavEventRecord> events,
            bool useDefaultCalendarFolder)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException("outlookApplication");
            }
            if (calendar == null)
            {
                throw new ArgumentNullException("calendar");
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultCalendar = null;
            Outlook.MAPIFolder targetFolder = null;
            try
            {
                session = outlookApplication.Session;
                defaultCalendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                targetFolder = useDefaultCalendarFolder
                    ? defaultCalendar
                    : EnsureCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
                }

                Dictionary<string, string> existing = LoadExistingEntryIds(targetFolder, calendar.Href);
                int imported = 0;
                int skippedRecurring = 0;
                foreach (CalDavEventRecord source in events ?? Enumerable.Empty<CalDavEventRecord>())
                {
                    if (source == null || string.IsNullOrWhiteSpace(source.Uid))
                    {
                        continue;
                    }
                    if (source.Recurring)
                    {
                        skippedRecurring++;
                        continue;
                    }

                    Outlook.AppointmentItem target = null;
                    Outlook.Items items = null;
                    try
                    {
                        string entryId;
                        string key = BuildKey(source.Uid, source.CalendarHref);
                        if (existing.TryGetValue(key, out entryId) && !string.IsNullOrWhiteSpace(entryId))
                        {
                            target = session.GetItemFromID(entryId, targetFolder.StoreID) as Outlook.AppointmentItem;
                        }
                        if (target == null)
                        {
                            items = targetFolder.Items;
                            target = items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
                        }
                        if (target == null)
                        {
                            continue;
                        }

                        ApplyEvent(target, source);
                        target.Save();
                        imported++;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to import CalDAV event (uid=" + Safe(source.Uid) + ", href=" + Safe(source.Href) + ").",
                            ex);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(target, LogCategories.Core, "Failed to release CalDAV AppointmentItem.");
                        ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV Items collection.");
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV download sync imported/updated "
                    + imported
                    + " appointments (calendar="
                    + Safe(calendar.DisplayName)
                    + ", recurringSkipped="
                    + skippedRecurring
                    + ").");
                return imported;
            }
            finally
            {
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV target calendar folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release Outlook default calendar folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV import.");
            }
        }

        private static void ApplyEvent(Outlook.AppointmentItem target, CalDavEventRecord source)
        {
            target.Subject = source.Subject ?? string.Empty;
            target.Location = source.Location ?? string.Empty;
            target.Body = source.Description ?? string.Empty;
            target.AllDayEvent = source.AllDay;
            target.Start = source.Start;
            target.End = source.End > source.Start ? source.End : source.Start.AddMinutes(30);
            target.BusyStatus = source.Cancelled
                ? Outlook.OlBusyStatus.olFree
                : Outlook.OlBusyStatus.olBusy;

            WriteUserProperty(target, UidPropertyName, source.Uid);
            WriteUserProperty(target, HrefPropertyName, source.Href);
            WriteUserProperty(target, ETagPropertyName, source.ETag);
            WriteUserProperty(target, CalendarPropertyName, source.CalendarHref);
        }

        private static Dictionary<string, string> LoadExistingEntryIds(Outlook.MAPIFolder folder, string calendarHref)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                int count = items.Count;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    Outlook.AppointmentItem appointment = null;
                    try
                    {
                        raw = items[index];
                        appointment = raw as Outlook.AppointmentItem;
                        if (appointment == null)
                        {
                            continue;
                        }
                        string uid = ReadUserProperty(appointment, UidPropertyName);
                        string storedCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        string key = BuildKey(uid, storedCalendar);
                        if (!string.IsNullOrWhiteSpace(uid)
                            && string.Equals(storedCalendar, calendarHref, StringComparison.OrdinalIgnoreCase)
                            && !result.ContainsKey(key)
                            && !string.IsNullOrWhiteSpace(appointment.EntryID))
                        {
                            result[key] = appointment.EntryID;
                        }
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release existing CalDAV AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-appointment calendar item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV calendar items collection.");
            }
            return result;
        }

        private static string ReadUserProperty(Outlook.AppointmentItem item, string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                property = properties != null ? properties.Find(name, true) : null;
                return property != null ? Convert.ToString(property.Value) ?? string.Empty : string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV user properties.");
            }
        }

        private static void WriteUserProperty(Outlook.AppointmentItem item, string name, string value)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                if (properties == null)
                {
                    return;
                }
                property = properties.Find(name, true);
                if (property == null)
                {
                    property = properties.Add(name, Outlook.OlUserPropertyType.olText, true, Type.Missing);
                }
                property.Value = value ?? string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV user properties.");
            }
        }

        private static Outlook.MAPIFolder EnsureCalendarFolder(Outlook.MAPIFolder parent, string name)
        {
            Outlook.Folders folders = null;
            try
            {
                folders = parent.Folders;
                int count = folders.Count;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null && string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder match = folder;
                            folder = null;
                            return match;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release Outlook calendar subfolder.");
                    }
                }
                return folders.Add(name, Outlook.OlDefaultFolders.olFolderCalendar);
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release Outlook calendar folders collection.");
            }
        }

        private static string BuildCalendarFolderName(CalDavCalendar calendar)
        {
            string name = calendar != null ? calendar.DisplayName : string.Empty;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Kalender";
            }
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                name = name.Replace(c, '-');
            }
            return "Nextcloud - " + name.Trim();
        }

        private static string BuildKey(string uid, string calendarHref)
        {
            return (uid ?? string.Empty).Trim() + "|" + (calendarHref ?? string.Empty).Trim();
        }

        private static CalDavEventRecord ParseEvent(string ical)
        {
            List<string> lines = UnfoldLines(ical);
            bool inEvent = false;
            bool seenMaster = false;
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (string line in lines)
            {
                if (string.Equals(line, "BEGIN:VEVENT", StringComparison.OrdinalIgnoreCase))
                {
                    if (seenMaster)
                    {
                        inEvent = false;
                        continue;
                    }
                    inEvent = true;
                    seenMaster = true;
                    continue;
                }
                if (!inEvent)
                {
                    continue;
                }
                if (string.Equals(line, "END:VEVENT", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }
                string left = line.Substring(0, colon);
                string value = line.Substring(colon + 1);
                string[] parts = left.Split(';');
                string name = parts[0].Trim();
                if (string.Equals(name, "RECURRENCE-ID", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                if (!values.ContainsKey(name))
                {
                    values[name] = value;
                    parameters[name] = parts.Length > 1 ? string.Join(";", parts.Skip(1).ToArray()) : string.Empty;
                }
            }

            string uid;
            string startText;
            if (!values.TryGetValue("UID", out uid) || !values.TryGetValue("DTSTART", out startText))
            {
                return null;
            }

            string startParameters;
            parameters.TryGetValue("DTSTART", out startParameters);
            bool allDay = HasParameter(startParameters, "VALUE=DATE") || IsDateOnly(startText);
            DateTime start;
            if (!TryParseIcalDate(startText, startParameters, allDay, out start))
            {
                return null;
            }

            DateTime end = allDay ? start.AddDays(1) : start.AddMinutes(30);
            string endText;
            if (values.TryGetValue("DTEND", out endText))
            {
                string endParameters;
                parameters.TryGetValue("DTEND", out endParameters);
                DateTime parsedEnd;
                if (TryParseIcalDate(endText, endParameters, allDay, out parsedEnd))
                {
                    end = parsedEnd;
                }
            }

            string subject;
            string location;
            string description;
            string status;
            values.TryGetValue("SUMMARY", out subject);
            values.TryGetValue("LOCATION", out location);
            values.TryGetValue("DESCRIPTION", out description);
            values.TryGetValue("STATUS", out status);

            return new CalDavEventRecord
            {
                Uid = UnescapeText(uid),
                Subject = UnescapeText(subject),
                Location = UnescapeText(location),
                Description = UnescapeText(description),
                Start = start,
                End = end,
                AllDay = allDay,
                Recurring = values.ContainsKey("RRULE"),
                Cancelled = string.Equals(status, "CANCELLED", StringComparison.OrdinalIgnoreCase)
            };
        }

        private static List<string> UnfoldLines(string text)
        {
            string normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            string[] source = normalized.Split('\n');
            var result = new List<string>();
            foreach (string raw in source)
            {
                string line = raw ?? string.Empty;
                if ((line.StartsWith(" ") || line.StartsWith("\t")) && result.Count > 0)
                {
                    result[result.Count - 1] += line.Substring(1);
                }
                else
                {
                    result.Add(line);
                }
            }
            return result;
        }

        private static bool TryParseIcalDate(string value, string parameterText, bool allDay, out DateTime result)
        {
            result = DateTime.MinValue;
            string text = (value ?? string.Empty).Trim();
            if (allDay)
            {
                return DateTime.TryParseExact(text, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result);
            }

            bool utc = text.EndsWith("Z", StringComparison.OrdinalIgnoreCase);
            string format = utc ? "yyyyMMdd'T'HHmmss'Z'" : "yyyyMMdd'T'HHmmss";
            DateTime parsed;
            if (!DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
            {
                return false;
            }
            if (utc)
            {
                result = DateTime.SpecifyKind(parsed, DateTimeKind.Utc).ToLocalTime();
                return true;
            }

            string tzid = ReadParameter(parameterText, "TZID");
            if (!string.IsNullOrWhiteSpace(tzid))
            {
                try
                {
                    TimeZoneInfo zone = TimeZoneInfo.FindSystemTimeZoneById(MapTimeZoneId(tzid));
                    result = TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified), zone, TimeZoneInfo.Local);
                    return true;
                }
                catch
                {
                }
            }

            result = DateTime.SpecifyKind(parsed, DateTimeKind.Local);
            return true;
        }

        private static string MapTimeZoneId(string tzid)
        {
            string value = (tzid ?? string.Empty).Trim();
            if (string.Equals(value, "Europe/Berlin", StringComparison.OrdinalIgnoreCase))
            {
                return "W. Europe Standard Time";
            }
            if (string.Equals(value, "Europe/Vienna", StringComparison.OrdinalIgnoreCase))
            {
                return "W. Europe Standard Time";
            }
            return value;
        }

        private static bool HasParameter(string parameterText, string expected)
        {
            return (parameterText ?? string.Empty).IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ReadParameter(string parameterText, string name)
        {
            foreach (string part in (parameterText ?? string.Empty).Split(';'))
            {
                int equals = part.IndexOf('=');
                if (equals <= 0)
                {
                    continue;
                }
                if (string.Equals(part.Substring(0, equals).Trim(), name, StringComparison.OrdinalIgnoreCase))
                {
                    return part.Substring(equals + 1).Trim().Trim('"');
                }
            }
            return string.Empty;
        }

        private static bool IsDateOnly(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Trim().Length == 8
                && value.Trim().IndexOf('T') < 0;
        }

        private static string UnescapeText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            return value
                .Replace("\\n", "\r\n")
                .Replace("\\N", "\r\n")
                .Replace("\\,", ",")
                .Replace("\\;", ";")
                .Replace("\\\\", "\\");
        }

        private XDocument SendDavRequest(string url, string method, string depth, string body)
        {
            var client = new NcHttpClient(_configuration);
            var headers = new Dictionary<string, string>();
            if (!string.IsNullOrWhiteSpace(depth))
            {
                headers["Depth"] = depth;
            }
            NcHttpResponse response = client.Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Accept = "application/xml, text/xml",
                ContentType = "application/xml; charset=utf-8",
                Body = body,
                Headers = headers,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 60000
            });
            if (!response.HasHttpResponse || (int)response.StatusCode < 200 || (int)response.StatusCode >= 300)
            {
                throw new InvalidOperationException("CalDAV request failed for " + url + " (HTTP " + (int)response.StatusCode + ").");
            }
            return XDocument.Parse(response.ResponseText ?? string.Empty);
        }

        private static bool IsSuccessfulStatus(string status)
        {
            return !string.IsNullOrWhiteSpace(status) && status.IndexOf(" 200 ", StringComparison.Ordinal) >= 0;
        }

        private static string ResolveUri(string baseUrl, string href)
        {
            if (string.IsNullOrWhiteSpace(href))
            {
                return baseUrl ?? string.Empty;
            }
            Uri absolute;
            if (Uri.TryCreate(href, UriKind.Absolute, out absolute))
            {
                return absolute.AbsoluteUri;
            }
            Uri baseUri;
            return Uri.TryCreate(baseUrl, UriKind.Absolute, out baseUri)
                ? new Uri(baseUri, href).AbsoluteUri
                : href;
        }

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<empty>" : value.Trim();
        }
    }
}
