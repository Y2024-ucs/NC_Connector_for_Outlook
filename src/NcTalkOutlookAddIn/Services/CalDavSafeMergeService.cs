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
    internal sealed class CalDavMergeResult
    {
        internal int ImportedToOutlook { get; set; }
        internal int UploadedToNextcloud { get; set; }
        internal int MatchedExisting { get; set; }
        internal int SkippedRecurring { get; set; }
        internal int SkippedMeetings { get; set; }
        internal int UploadFailures { get; set; }
    }

    // Safety rule: this service NEVER deletes anything. It is intentionally suitable
    // for the first synchronization of an existing Outlook calendar with an existing
    // CalDAV calendar. Missing items are copied to the other side unless a per-item
    // state proves that both sides previously contained the item; in that case the
    // missing side is preserved for the deletion phase instead of being resurrected.
    internal sealed class CalDavSafeMergeService
    {
        private const string UidPropertyName = "NC-CalDAV-UID";
        private const string HrefPropertyName = "NC-CalDAV-HREF";
        private const string ETagPropertyName = "NC-CalDAV-ETAG";
        private const string CalendarPropertyName = "NC-CalDAV-CALENDAR";

        private readonly TalkServiceConfiguration _configuration;

        internal CalDavSafeMergeService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            _configuration = configuration;
        }

        internal CalDavMergeResult Merge(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            CalDavSyncDirection direction)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException("outlookApplication");
            }
            if (calendar == null || string.IsNullOrWhiteSpace(calendar.Href))
            {
                throw new ArgumentException("A CalDAV calendar is required.", "calendar");
            }

            var result = new CalDavMergeResult();
            IList<CalDavEventRecord> remote = remoteEvents ?? new List<CalDavEventRecord>();
            CalDavSyncStateStore syncState = CalDavSyncStateStore.Load();

            HashSet<string> matchedRemoteKeys = MatchRemoteToExistingOutlook(
                outlookApplication,
                calendar,
                remote,
                useDefaultCalendarFolder,
                result);

            if (direction != CalDavSyncDirection.OutlookToNextcloud)
            {
                int deletionIntentPreserved = 0;
                List<CalDavEventRecord> missingInOutlook = remote
                    .Where(v =>
                    {
                        if (v == null || matchedRemoteKeys.Contains(BuildRemoteKey(v)))
                        {
                            return false;
                        }

                        // In TwoWay/Outlook->Nextcloud mode, a remote event that was
                        // previously confirmed on both sides but is now absent locally
                        // is a possible local deletion. Do not re-import it here; the
                        // guarded deletion phase decides whether the remote copy may be
                        // deleted. In one-way Nextcloud->Outlook mode it is deliberately
                        // re-imported because Outlook is not authoritative for deletes.
                        if (direction != CalDavSyncDirection.NextcloudToOutlook
                            && syncState.WasItemSeenOnBothSides(calendar.Href, v.Uid))
                        {
                            deletionIntentPreserved++;
                            return false;
                        }
                        return true;
                    })
                    .ToList();

                if (deletionIntentPreserved > 0)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CalDAV safe merge preserved "
                        + deletionIntentPreserved.ToString(CultureInfo.InvariantCulture)
                        + " possible Outlook-side deletions for guarded deletion processing (calendar="
                        + Safe(calendar.DisplayName) + ").");
                }

                result.ImportedToOutlook = new CalDavEventSyncService(_configuration)
                    .ImportIntoOutlook(outlookApplication, calendar, missingInOutlook, useDefaultCalendarFolder);
            }

            if (direction == CalDavSyncDirection.NextcloudToOutlook || calendar.ReadOnly)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV safe merge upload skipped (calendar=" + Safe(calendar.DisplayName)
                    + ", direction=" + direction
                    + ", readOnly=" + calendar.ReadOnly + ").");
                return result;
            }

            UploadMissingOutlookAppointments(
                outlookApplication,
                calendar,
                remote,
                useDefaultCalendarFolder,
                result);

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDAV safe merge completed (calendar=" + Safe(calendar.DisplayName)
                + ", imported=" + result.ImportedToOutlook
                + ", uploaded=" + result.UploadedToNextcloud
                + ", matched=" + result.MatchedExisting
                + ", recurringSkipped=" + result.SkippedRecurring
                + ", meetingsSkipped=" + result.SkippedMeetings
                + ", uploadFailures=" + result.UploadFailures
                + ", deletions=0).");
            return result;
        }

        private HashSet<string> MatchRemoteToExistingOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            CalDavMergeResult result)
        {
            var matchedRemoteKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var remoteByUid = new Dictionary<string, CalDavEventRecord>(StringComparer.OrdinalIgnoreCase);
            var remoteByFingerprint = new Dictionary<string, Queue<CalDavEventRecord>>(StringComparer.OrdinalIgnoreCase);

            foreach (CalDavEventRecord remote in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (remote == null)
                {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(remote.Uid) && !remoteByUid.ContainsKey(remote.Uid.Trim()))
                {
                    remoteByUid[remote.Uid.Trim()] = remote;
                }

                string fingerprint = BuildFingerprint(remote.Subject, remote.Location, remote.Start, remote.End, remote.AllDay);
                Queue<CalDavEventRecord> queue;
                if (!remoteByFingerprint.TryGetValue(fingerprint, out queue))
                {
                    queue = new Queue<CalDavEventRecord>();
                    remoteByFingerprint[fingerprint] = queue;
                }
                queue.Enqueue(remote);
            }

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
                    : EnsureCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
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

                        string existingUid = ReadUserProperty(appointment, UidPropertyName);
                        string existingCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        if (!string.IsNullOrWhiteSpace(existingUid)
                            && string.Equals(existingCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            CalDavEventRecord alreadyLinked;
                            if (remoteByUid.TryGetValue(existingUid.Trim(), out alreadyLinked))
                            {
                                matchedRemoteKeys.Add(BuildRemoteKey(alreadyLinked));
                            }
                            continue;
                        }

                        if (appointment.IsRecurring
                            || appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                        {
                            continue;
                        }

                        string fingerprint = BuildFingerprint(
                            appointment.Subject,
                            appointment.Location,
                            appointment.Start,
                            appointment.End,
                            appointment.AllDayEvent);

                        Queue<CalDavEventRecord> candidates;
                        if (!remoteByFingerprint.TryGetValue(fingerprint, out candidates))
                        {
                            continue;
                        }

                        CalDavEventRecord matched = null;
                        while (candidates.Count > 0)
                        {
                            CalDavEventRecord candidate = candidates.Dequeue();
                            if (!matchedRemoteKeys.Contains(BuildRemoteKey(candidate)))
                            {
                                matched = candidate;
                                break;
                            }
                        }
                        if (matched == null)
                        {
                            continue;
                        }

                        WriteSyncProperties(appointment, matched.Uid, matched.Href, matched.ETag, calendar.Href);
                        appointment.Save();
                        matchedRemoteKeys.Add(BuildRemoteKey(matched));
                        result.MatchedExisting++;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "CalDAV first-sync pre-match failed for one Outlook appointment; item was left unchanged.",
                            ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV pre-match AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV pre-match folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV pre-match Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV pre-match target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV pre-match default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV pre-match.");
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDAV first-sync pre-match completed (calendar=" + Safe(calendar.DisplayName)
                + ", matched=" + matchedRemoteKeys.Count
                + ", remote=" + (remoteEvents != null ? remoteEvents.Count : 0)
                + ").");
            return matchedRemoteKeys;
        }

        private void UploadMissingOutlookAppointments(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            CalDavMergeResult result)
        {
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
                    : EnsureCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
                }

                var remoteByFingerprint = new Dictionary<string, CalDavEventRecord>(StringComparer.OrdinalIgnoreCase);
                foreach (CalDavEventRecord remote in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
                {
                    if (remote == null)
                    {
                        continue;
                    }
                    string fingerprint = BuildFingerprint(remote.Subject, remote.Location, remote.Start, remote.End, remote.AllDay);
                    if (!remoteByFingerprint.ContainsKey(fingerprint))
                    {
                        remoteByFingerprint[fingerprint] = remote;
                    }
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

                        if (appointment.IsRecurring)
                        {
                            result.SkippedRecurring++;
                            continue;
                        }
                        if (appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting)
                        {
                            result.SkippedMeetings++;
                            continue;
                        }

                        string existingUid = ReadUserProperty(appointment, UidPropertyName);
                        string existingCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        if (!string.IsNullOrWhiteSpace(existingUid)
                            && string.Equals(existingCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string fingerprint = BuildFingerprint(
                            appointment.Subject,
                            appointment.Location,
                            appointment.Start,
                            appointment.End,
                            appointment.AllDayEvent);
                        CalDavEventRecord matched;
                        if (remoteByFingerprint.TryGetValue(fingerprint, out matched))
                        {
                            WriteSyncProperties(appointment, matched.Uid, matched.Href, matched.ETag, calendar.Href);
                            appointment.Save();
                            result.MatchedExisting++;
                            continue;
                        }

                        string uid = Guid.NewGuid().ToString("D") + "@nc4ol";
                        string href = calendar.Href.TrimEnd('/') + "/" + Uri.EscapeDataString(uid) + ".ics";
                        string payload = BuildICalendar(appointment, uid);
                        string etag;
                        if (!TryCreateRemoteEvent(href, payload, out etag))
                        {
                            result.UploadFailures++;
                            continue;
                        }

                        WriteSyncProperties(appointment, uid, href, etag, calendar.Href);
                        appointment.Save();
                        result.UploadedToNextcloud++;
                    }
                    catch (Exception ex)
                    {
                        result.UploadFailures++;
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "CalDAV safe merge failed for one Outlook appointment; item was kept locally.",
                            ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV merge AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV merge folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV merge Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV merge target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV merge default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV safe merge.");
            }
        }

        private bool TryCreateRemoteEvent(string href, string payload, out string etag)
        {
            etag = string.Empty;
            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "If-None-Match", "*" }
            };
            NcHttpResponse response = new NcHttpClient(_configuration).Send(new NcHttpRequestOptions
            {
                Method = "PUT",
                Url = href,
                Accept = "text/calendar, */*",
                ContentType = "text/calendar; charset=utf-8",
                Payload = payload,
                Headers = headers,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 60000
            });

            if (!response.HasHttpResponse
                || ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300))
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV safe merge PUT failed (href=" + Safe(href)
                    + ", http=" + (response.HasHttpResponse ? ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture) : "transport")
                    + "). Local appointment was not modified.");
                return false;
            }

            if (response.Headers != null)
            {
                response.Headers.TryGetValue("ETag", out etag);
            }
            etag = (etag ?? string.Empty).Trim();
            return true;
        }

        private static string BuildICalendar(Outlook.AppointmentItem appointment, string uid)
        {
            var builder = new StringBuilder();
            builder.Append("BEGIN:VCALENDAR\r\n");
            builder.Append("VERSION:2.0\r\n");
            builder.Append("PRODID:-//NC Connector for Outlook//CalDAV Safe Merge//EN\r\n");
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

        private static string BuildRemoteKey(CalDavEventRecord value)
        {
            if (value == null)
            {
                return string.Empty;
            }
            if (!string.IsNullOrWhiteSpace(value.Href))
            {
                return "href:" + value.Href.Trim();
            }
            if (!string.IsNullOrWhiteSpace(value.Uid))
            {
                return "uid:" + value.Uid.Trim();
            }
            return "fp:" + BuildFingerprint(value.Subject, value.Location, value.Start, value.End, value.AllDay);
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
            return (value ?? string.Empty).Trim().ToUpperInvariant();
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV merge user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV merge user properties.");
            }
        }

        private static void WriteSyncProperties(
            Outlook.AppointmentItem item,
            string uid,
            string href,
            string etag,
            string calendarHref)
        {
            WriteUserProperty(item, UidPropertyName, uid);
            WriteUserProperty(item, HrefPropertyName, href);
            WriteUserProperty(item, ETagPropertyName, etag);
            WriteUserProperty(item, CalendarPropertyName, calendarHref);
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV merge user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV merge user properties.");
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
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CalDAV merge calendar folder.");
                    }
                }
                return folders.Add(name, Outlook.OlDefaultFolders.olFolderCalendar);
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release CalDAV merge folders collection.");
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

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<empty>" : value.Trim();
        }
    }
}
