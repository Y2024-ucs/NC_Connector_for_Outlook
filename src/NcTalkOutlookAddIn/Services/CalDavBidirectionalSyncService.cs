// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CalDavBidirectionalSyncResult
    {
        internal int RemoteToOutlook { get; set; }
        internal int OutlookToRemote { get; set; }
        internal int BaselinesEstablished { get; set; }
        internal int Conflicts { get; set; }
        internal int SkippedRecurring { get; set; }
        internal int SkippedMeetings { get; set; }
        internal int Failures { get; set; }
        internal int RecoveryRemoteRemoved { get; set; }
        internal int RecoveryOutlookRemoved { get; set; }
        internal int RecoveryOutlookRelinked { get; set; }
    }

    // Regular synchronization for already linked calendar items.
    // Safety rules:
    // - never interprets absence as deletion
    // - uses ETag/If-Match for remote writes
    // - detects local/remote concurrent changes and leaves conflicts untouched
    // - recurring series and Outlook meetings are deliberately not modified yet
    //   (the model is prepared for a later dedicated implementation)
    internal sealed class CalDavBidirectionalSyncService
    {
        private const string UidPropertyName = "NC-CalDAV-UID";
        private const string HrefPropertyName = "NC-CalDAV-HREF";
        private const string ETagPropertyName = "NC-CalDAV-ETAG";
        private const string CalendarPropertyName = "NC-CalDAV-CALENDAR";
        private const string LocalSnapshotPropertyName = "NC-CalDAV-LOCAL-SNAPSHOT";

        private readonly TalkServiceConfiguration _configuration;

        internal CalDavBidirectionalSyncService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            _configuration = configuration;
        }

        internal CalDavBidirectionalSyncResult SyncLinkedEvents(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            CalDavSyncDirection direction)
        {
            var result = new CalDavBidirectionalSyncResult();
            if (outlookApplication == null || calendar == null)
            {
                return result;
            }

            Dictionary<string, CalDavEventRecord> remoteByUid = BuildRemoteByUid(remoteEvents);
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultCalendar = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items items = null;
            try
            {
                session = outlookApplication.Session;
                defaultCalendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                targetFolder = useDefaultCalendarFolder
                    ? defaultCalendar
                    : FindCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
                }
                if (targetFolder == null)
                {
                    return result;
                }

                items = targetFolder.Items;
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
                        if (string.IsNullOrWhiteSpace(uid)
                            || !string.Equals(storedCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        CalDavEventRecord remote;
                        if (!remoteByUid.TryGetValue(uid.Trim(), out remote))
                        {
                            // Missing on one side is not a deletion signal.
                            continue;
                        }

                        if (appointment.IsRecurring || remote.Recurring)
                        {
                            result.SkippedRecurring++;
                            continue;
                        }
                        if (appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                        {
                            result.SkippedMeetings++;
                            continue;
                        }

                        string localSnapshot = BuildOutlookSnapshot(appointment);
                        string storedSnapshot = ReadUserProperty(appointment, LocalSnapshotPropertyName);
                        string storedEtag = ReadUserProperty(appointment, ETagPropertyName);
                        string remoteEtag = (remote.ETag ?? string.Empty).Trim();

                        if (string.IsNullOrWhiteSpace(storedSnapshot))
                        {
                            // The very first per-item baseline must never silently choose one side.
                            // Only establish it when Outlook and the actual CalDAV payload already
                            // represent the same visible event. A difference is an initial conflict
                            // and remains untouched until the user resolves both sides to the same
                            // content (or a later conflict UI explicitly chooses a winner).
                            string remoteSnapshot = BuildRemoteSnapshot(remote);
                            if (!string.Equals(localSnapshot, remoteSnapshot, StringComparison.Ordinal))
                            {
                                result.Conflicts++;
                                DiagnosticsLogger.Log(
                                    LogCategories.Core,
                                    "CalDAV initial baseline conflict kept unchanged (calendar=" + Safe(calendar.DisplayName)
                                    + ", uid=" + Safe(uid)
                                    + ", reason=initial-local-remote-difference).");
                                continue;
                            }

                            WriteUserProperty(appointment, LocalSnapshotPropertyName, localSnapshot);
                            WriteUserProperty(appointment, ETagPropertyName, remoteEtag);
                            appointment.Save();
                            result.BaselinesEstablished++;
                            continue;
                        }

                        bool localChanged = !string.Equals(storedSnapshot, localSnapshot, StringComparison.Ordinal);
                        bool remoteChanged = !string.Equals(
                            NormalizeEtag(storedEtag),
                            NormalizeEtag(remoteEtag),
                            StringComparison.Ordinal);

                        if (localChanged && remoteChanged)
                        {
                            result.Conflicts++;
                            DiagnosticsLogger.Log(
                                LogCategories.Core,
                                "CalDAV conflict kept unchanged (calendar=" + Safe(calendar.DisplayName)
                                + ", uid=" + Safe(uid)
                                + ", reason=local-and-remote-changed).");
                            continue;
                        }

                        if (remoteChanged
                            && direction != CalDavSyncDirection.OutlookToNextcloud)
                        {
                            ApplyRemoteToOutlook(appointment, remote);
                            string newSnapshot = BuildOutlookSnapshot(appointment);
                            WriteUserProperty(appointment, LocalSnapshotPropertyName, newSnapshot);
                            WriteUserProperty(appointment, ETagPropertyName, remoteEtag);
                            appointment.Save();
                            result.RemoteToOutlook++;
                            continue;
                        }

                        if (localChanged
                            && direction != CalDavSyncDirection.NextcloudToOutlook
                            && !calendar.ReadOnly)
                        {
                            string href = !string.IsNullOrWhiteSpace(remote.Href)
                                ? remote.Href
                                : ReadUserProperty(appointment, HrefPropertyName);
                            string newEtag;
                            if (TryUpdateRemoteEvent(appointment, uid, href, remoteEtag, out newEtag))
                            {
                                WriteUserProperty(appointment, ETagPropertyName, newEtag);
                                WriteUserProperty(appointment, LocalSnapshotPropertyName, localSnapshot);
                                appointment.Save();
                                result.OutlookToRemote++;
                            }
                            else
                            {
                                result.Failures++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Failures++;
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "CalDAV bidirectional sync failed for one linked appointment; item was left unchanged where possible.",
                            ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV bidirectional AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV bidirectional folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV bidirectional Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV bidirectional target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV bidirectional default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV bidirectional sync.");
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDAV bidirectional linked sync completed (calendar=" + Safe(calendar.DisplayName)
                + ", remoteToOutlook=" + result.RemoteToOutlook
                + ", outlookToRemote=" + result.OutlookToRemote
                + ", baselines=" + result.BaselinesEstablished
                + ", conflicts=" + result.Conflicts
                + ", recurringSkipped=" + result.SkippedRecurring
                + ", meetingsSkipped=" + result.SkippedMeetings
                + ", failures=" + result.Failures
                + ", deletions=0).");
            return result;
        }

        // One-time guarded repair for the known first-merge damage pattern:
        // exactly one @nc4ol remote object plus exactly one original remote object
        // with the same visible event fingerprint. The original object always wins.
        internal CalDavBidirectionalSyncResult RepairGeneratedDuplicates(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder)
        {
            var result = new CalDavBidirectionalSyncResult();
            if (outlookApplication == null || calendar == null || calendar.ReadOnly)
            {
                return result;
            }

            Dictionary<string, List<CalDavEventRecord>> groups = GroupRemote(remoteEvents);
            foreach (List<CalDavEventRecord> group in groups.Values)
            {
                List<CalDavEventRecord> generated = group.Where(IsConnectorGeneratedUid).ToList();
                List<CalDavEventRecord> original = group.Where(v => v != null && !IsConnectorGeneratedUid(v)).ToList();
                if (generated.Count != 1 || original.Count != 1)
                {
                    continue;
                }

                CalDavEventRecord generatedItem = generated[0];
                CalDavEventRecord originalItem = original[0];
                if (generatedItem.Recurring || originalItem.Recurring)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(generatedItem.Href)
                    || string.IsNullOrWhiteSpace(generatedItem.ETag)
                    || string.IsNullOrWhiteSpace(originalItem.Uid))
                {
                    continue;
                }

                int generatedLocal;
                int originalLocal;
                string generatedEntryId;
                string originalEntryId;
                FindLinkedCopies(
                    outlookApplication,
                    calendar,
                    useDefaultCalendarFolder,
                    generatedItem.Uid,
                    originalItem.Uid,
                    out generatedLocal,
                    out originalLocal,
                    out generatedEntryId,
                    out originalEntryId);

                if (generatedLocal != 1 || originalLocal > 1)
                {
                    continue;
                }

                bool localReady = false;
                if (originalLocal == 1)
                {
                    localReady = DeleteNonMeetingNonRecurringByEntryId(outlookApplication, generatedEntryId);
                    if (localReady)
                    {
                        result.RecoveryOutlookRemoved++;
                    }
                }
                else
                {
                    localReady = RelinkNonMeetingNonRecurringByEntryId(
                        outlookApplication,
                        generatedEntryId,
                        originalItem,
                        calendar.Href);
                    if (localReady)
                    {
                        result.RecoveryOutlookRelinked++;
                    }
                }

                if (!localReady)
                {
                    continue;
                }

                if (DeleteRemoteGeneratedItem(generatedItem))
                {
                    result.RecoveryRemoteRemoved++;
                }
                else
                {
                    result.Failures++;
                }
            }

            if (result.RecoveryRemoteRemoved > 0
                || result.RecoveryOutlookRemoved > 0
                || result.RecoveryOutlookRelinked > 0
                || result.Failures > 0)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV guarded @nc4ol repair completed (calendar=" + Safe(calendar.DisplayName)
                    + ", remoteRemoved=" + result.RecoveryRemoteRemoved
                    + ", outlookRemoved=" + result.RecoveryOutlookRemoved
                    + ", outlookRelinked=" + result.RecoveryOutlookRelinked
                    + ", failures=" + result.Failures + ").");
            }
            return result;
        }

        private bool TryUpdateRemoteEvent(
            Outlook.AppointmentItem appointment,
            string uid,
            string href,
            string remoteEtag,
            out string newEtag)
        {
            newEtag = remoteEtag ?? string.Empty;
            if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(remoteEtag))
            {
                return false;
            }

            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "If-Match", remoteEtag.Trim() }
            };
            NcHttpResponse response = new NcHttpClient(_configuration).Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = href,
                Accept = "text/calendar, */*",
                ContentType = "text/calendar; charset=utf-8",
                Payload = BuildICalendar(appointment, uid),
                Headers = headers,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 60000
            });

            if (!response.HasHttpResponse
                || (int)response.StatusCode < 200
                || (int)response.StatusCode >= 300)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV update PUT failed (href=" + Safe(href)
                    + ", http=" + (response.HasHttpResponse
                        ? ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture)
                        : "transport")
                    + ").");
                return false;
            }

            string headerEtag;
            if (response.Headers != null && response.Headers.TryGetValue("ETag", out headerEtag))
            {
                newEtag = (headerEtag ?? string.Empty).Trim();
            }
            if (string.IsNullOrWhiteSpace(newEtag))
            {
                newEtag = remoteEtag.Trim();
            }
            return true;
        }

        private bool DeleteRemoteGeneratedItem(CalDavEventRecord item)
        {
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "If-Match", item.ETag.Trim() }
            };
            NcHttpResponse response = new NcHttpClient(_configuration).Send(new NcHttpRequestOptions
            {
                Method = "DELETE",
                Url = item.Href,
                Accept = "*/*",
                Headers = headers,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 60000
            });
            return response.HasHttpResponse
                && (int)response.StatusCode >= 200
                && (int)response.StatusCode < 300;
        }

        private static void ApplyRemoteToOutlook(Outlook.AppointmentItem target, CalDavEventRecord source)
        {
            target.Subject = source.Subject ?? string.Empty;
            target.Location = source.Location ?? string.Empty;
            target.Body = source.Description ?? string.Empty;
            target.AllDayEvent = source.AllDay;
            target.Start = source.Start;
            target.End = source.End > source.Start ? source.End : source.Start.AddMinutes(30);
            target.BusyStatus = source.Cancelled ? Outlook.OlBusyStatus.olFree : Outlook.OlBusyStatus.olBusy;
            WriteUserProperty(target, UidPropertyName, source.Uid);
            WriteUserProperty(target, HrefPropertyName, source.Href);
            WriteUserProperty(target, ETagPropertyName, source.ETag);
            WriteUserProperty(target, CalendarPropertyName, source.CalendarHref);
        }

        private static string BuildICalendar(Outlook.AppointmentItem appointment, string uid)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("BEGIN:VCALENDAR\r\n");
            builder.Append("VERSION:2.0\r\n");
            builder.Append("PRODID:-//NC Connector for Outlook//CalDAV Bidirectional Sync//EN\r\n");
            builder.Append("CALSCALE:GREGORIAN\r\n");
            builder.Append("BEGIN:VEVENT\r\n");
            builder.Append("UID:").Append(EscapeText(uid)).Append("\r\n");
            builder.Append("DTSTAMP:").Append(DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
            if (appointment.AllDayEvent)
            {
                builder.Append("DTSTART;VALUE=DATE:").Append(appointment.Start.ToString("yyyyMMdd", CultureInfo.InvariantCulture)).Append("\r\n");
                builder.Append("DTEND;VALUE=DATE:").Append(appointment.End.ToString("yyyyMMdd", CultureInfo.InvariantCulture)).Append("\r\n");
            }
            else
            {
                builder.Append("DTSTART:").Append(appointment.Start.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
                builder.Append("DTEND:").Append(appointment.End.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture)).Append("\r\n");
            }
            builder.Append("SUMMARY:").Append(EscapeText(appointment.Subject)).Append("\r\n");
            if (!string.IsNullOrWhiteSpace(appointment.Location))
            {
                builder.Append("LOCATION:").Append(EscapeText(appointment.Location)).Append("\r\n");
            }
            if (!string.IsNullOrWhiteSpace(appointment.Body))
            {
                builder.Append("DESCRIPTION:").Append(EscapeText(appointment.Body)).Append("\r\n");
            }
            builder.Append("END:VEVENT\r\n");
            builder.Append("END:VCALENDAR\r\n");
            return builder.ToString();
        }

        private static string BuildOutlookSnapshot(Outlook.AppointmentItem appointment)
        {
            return Normalize(appointment.Subject) + "|"
                + Normalize(appointment.Location) + "|"
                + Normalize(appointment.Body) + "|"
                + NormalizeDateTime(appointment.Start, appointment.AllDayEvent) + "|"
                + NormalizeDateTime(appointment.End, appointment.AllDayEvent) + "|"
                + (appointment.AllDayEvent ? "1" : "0");
        }

        private static string BuildRemoteSnapshot(CalDavEventRecord item)
        {
            if (item == null)
            {
                return string.Empty;
            }

            DateTime end = item.End > item.Start ? item.End : item.Start.AddMinutes(30);
            return Normalize(item.Subject) + "|"
                + Normalize(item.Location) + "|"
                + Normalize(item.Description) + "|"
                + NormalizeDateTime(item.Start, item.AllDay) + "|"
                + NormalizeDateTime(end, item.AllDay) + "|"
                + (item.AllDay ? "1" : "0");
        }

        private static Dictionary<string, CalDavEventRecord> BuildRemoteByUid(IList<CalDavEventRecord> remoteEvents)
        {
            var result = new Dictionary<string, CalDavEventRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (CalDavEventRecord item in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Uid))
                {
                    continue;
                }
                if (!result.ContainsKey(item.Uid.Trim()))
                {
                    result[item.Uid.Trim()] = item;
                }
            }
            return result;
        }

        private static Dictionary<string, List<CalDavEventRecord>> GroupRemote(IList<CalDavEventRecord> remoteEvents)
        {
            var result = new Dictionary<string, List<CalDavEventRecord>>(StringComparer.OrdinalIgnoreCase);
            foreach (CalDavEventRecord item in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (item == null)
                {
                    continue;
                }
                string key = BuildFingerprint(item.Subject, item.Location, item.Start, item.End, item.AllDay);
                List<CalDavEventRecord> list;
                if (!result.TryGetValue(key, out list))
                {
                    list = new List<CalDavEventRecord>();
                    result[key] = list;
                }
                list.Add(item);
            }
            return result;
        }

        private static bool IsConnectorGeneratedUid(CalDavEventRecord item)
        {
            return item != null
                && !string.IsNullOrWhiteSpace(item.Uid)
                && item.Uid.Trim().EndsWith("@nc4ol", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildFingerprint(string subject, string location, DateTime start, DateTime end, bool allDay)
        {
            return Normalize(subject) + "|"
                + Normalize(location) + "|"
                + NormalizeDateTime(start, allDay) + "|"
                + NormalizeDateTime(end, allDay) + "|"
                + (allDay ? "1" : "0");
        }

        private static string NormalizeDateTime(DateTime value, bool allDay)
        {
            if (allDay)
            {
                return value.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            DateTime local = value.Kind == DateTimeKind.Utc ? value.ToLocalTime() : value;
            return new DateTime(local.Year, local.Month, local.Day, local.Hour, local.Minute, 0)
                .ToString("yyyy-MM-dd'T'HH:mm", CultureInfo.InvariantCulture);
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim().Replace("\r\n", "\n").ToUpperInvariant();
        }

        private static string NormalizeEtag(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static string EscapeText(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\", "\\\\")
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n")
                .Replace(",", "\\,")
                .Replace(";", "\\;");
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV bidirectional user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV bidirectional user properties.");
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV bidirectional user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV bidirectional user properties.");
            }
        }

        private static Outlook.MAPIFolder FindCalendarFolder(Outlook.MAPIFolder parent, string name)
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
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CalDAV bidirectional calendar folder.");
                    }
                }
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release CalDAV bidirectional folders collection.");
            }
        }

        private static string BuildCalendarFolderName(CalDavCalendar calendar)
        {
            string name = string.IsNullOrWhiteSpace(calendar.DisplayName) ? "Kalender" : calendar.DisplayName.Trim();
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                name = name.Replace(c, '-');
            }
            return "Nextcloud - " + name;
        }

        private static void FindLinkedCopies(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            bool useDefaultCalendarFolder,
            string generatedUid,
            string originalUid,
            out int generatedCount,
            out int originalCount,
            out string generatedEntryId,
            out string originalEntryId)
        {
            generatedCount = 0;
            originalCount = 0;
            generatedEntryId = string.Empty;
            originalEntryId = string.Empty;
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultCalendar = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items items = null;
            try
            {
                session = outlookApplication.Session;
                defaultCalendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                targetFolder = useDefaultCalendarFolder ? defaultCalendar : FindCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
                }
                if (targetFolder == null)
                {
                    return;
                }
                items = targetFolder.Items;
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
                        string storedCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        if (!string.Equals(storedCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }
                        string uid = ReadUserProperty(appointment, UidPropertyName);
                        if (string.Equals(uid, generatedUid, StringComparison.OrdinalIgnoreCase))
                        {
                            generatedCount++;
                            generatedEntryId = appointment.EntryID ?? string.Empty;
                        }
                        else if (string.Equals(uid, originalUid, StringComparison.OrdinalIgnoreCase))
                        {
                            originalCount++;
                            originalEntryId = appointment.EntryID ?? string.Empty;
                        }
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV repair AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV repair folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV repair Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV repair target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV repair default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV repair scan.");
            }
        }

        private static bool DeleteNonMeetingNonRecurringByEntryId(Outlook.Application outlookApplication, string entryId)
        {
            if (outlookApplication == null || string.IsNullOrWhiteSpace(entryId))
            {
                return false;
            }
            Outlook.NameSpace session = null;
            object raw = null;
            Outlook.AppointmentItem appointment = null;
            try
            {
                session = outlookApplication.Session;
                raw = session.GetItemFromID(entryId, Type.Missing);
                appointment = raw as Outlook.AppointmentItem;
                if (appointment == null
                    || appointment.IsRecurring
                    || appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                {
                    return false;
                }
                appointment.Delete();
                return true;
            }
            finally
            {
                if (appointment != null)
                {
                    ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV repaired deleted appointment.");
                }
                else
                {
                    ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV repair item.");
                }
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV repair delete.");
            }
        }

        private static bool RelinkNonMeetingNonRecurringByEntryId(
            Outlook.Application outlookApplication,
            string entryId,
            CalDavEventRecord original,
            string calendarHref)
        {
            if (outlookApplication == null || string.IsNullOrWhiteSpace(entryId) || original == null)
            {
                return false;
            }
            Outlook.NameSpace session = null;
            object raw = null;
            Outlook.AppointmentItem appointment = null;
            try
            {
                session = outlookApplication.Session;
                raw = session.GetItemFromID(entryId, Type.Missing);
                appointment = raw as Outlook.AppointmentItem;
                if (appointment == null
                    || appointment.IsRecurring
                    || appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                {
                    return false;
                }
                WriteUserProperty(appointment, UidPropertyName, original.Uid);
                WriteUserProperty(appointment, HrefPropertyName, original.Href);
                WriteUserProperty(appointment, ETagPropertyName, original.ETag);
                WriteUserProperty(appointment, CalendarPropertyName, calendarHref);
                WriteUserProperty(appointment, LocalSnapshotPropertyName, BuildOutlookSnapshot(appointment));
                appointment.Save();
                return true;
            }
            finally
            {
                if (appointment != null)
                {
                    ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV repaired relinked appointment.");
                }
                else
                {
                    ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV repair item.");
                }
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV repair relink.");
            }
        }

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<empty>" : value.Trim();
        }
    }
}
