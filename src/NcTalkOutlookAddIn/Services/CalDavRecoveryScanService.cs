// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CalDavRecoveryScanResult
    {
        internal int OutlookItems { get; set; }
        internal int OutlookDuplicateGroups { get; set; }
        internal int OutlookSafeRemoveCandidates { get; set; }
        internal int OutlookAmbiguousGroups { get; set; }
        internal int NextcloudItems { get; set; }
        internal int NextcloudDuplicateGroups { get; set; }
        internal int NextcloudSafeRemoveCandidates { get; set; }
        internal int NextcloudAmbiguousGroups { get; set; }

        internal bool HasFindings
        {
            get
            {
                return OutlookDuplicateGroups > 0 || NextcloudDuplicateGroups > 0;
            }
        }
    }

    // Analysis only. This service never deletes or changes Outlook/Nextcloud items.
    internal sealed class CalDavRecoveryScanService
    {
        private const string UidPropertyName = "NC-CalDAV-UID";
        private const string CalendarPropertyName = "NC-CalDAV-CALENDAR";

        internal CalDavRecoveryScanResult Scan(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
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

            var result = new CalDavRecoveryScanResult();
            ScanNextcloud(remoteEvents, result);
            ScanOutlook(outlookApplication, calendar, useDefaultCalendarFolder, result);

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDAV recovery scan (calendar=" + Safe(calendar.DisplayName)
                + ", outlookItems=" + result.OutlookItems
                + ", outlookDuplicateGroups=" + result.OutlookDuplicateGroups
                + ", outlookSafeToRemove=" + result.OutlookSafeRemoveCandidates
                + ", outlookAmbiguous=" + result.OutlookAmbiguousGroups
                + ", nextcloudItems=" + result.NextcloudItems
                + ", nextcloudDuplicateGroups=" + result.NextcloudDuplicateGroups
                + ", nextcloudSafeToRemove=" + result.NextcloudSafeRemoveCandidates
                + ", nextcloudAmbiguous=" + result.NextcloudAmbiguousGroups
                + ", deleted=0, modified=0).");

            return result;
        }

        private static void ScanNextcloud(
            IList<CalDavEventRecord> remoteEvents,
            CalDavRecoveryScanResult result)
        {
            var groups = new Dictionary<string, List<CalDavEventRecord>>(StringComparer.OrdinalIgnoreCase);
            foreach (CalDavEventRecord item in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (item == null)
                {
                    continue;
                }
                result.NextcloudItems++;
                string fingerprint = BuildFingerprint(item.Subject, item.Location, item.Start, item.End, item.AllDay);
                List<CalDavEventRecord> list;
                if (!groups.TryGetValue(fingerprint, out list))
                {
                    list = new List<CalDavEventRecord>();
                    groups[fingerprint] = list;
                }
                list.Add(item);
            }

            foreach (List<CalDavEventRecord> group in groups.Values)
            {
                if (group.Count < 2)
                {
                    continue;
                }

                int generated = group.Count(IsConnectorGeneratedUid);
                int original = group.Count - generated;
                if (generated == 0)
                {
                    continue;
                }

                result.NextcloudDuplicateGroups++;
                if (generated == 1 && original == 1 && group.Count == 2)
                {
                    // Exact pair: one original cloud item plus one item uploaded by our
                    // faulty first merge. Only the @nc4ol object would be removable.
                    result.NextcloudSafeRemoveCandidates++;
                }
                else
                {
                    result.NextcloudAmbiguousGroups++;
                }
            }
        }

        private static void ScanOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            bool useDefaultCalendarFolder,
            CalDavRecoveryScanResult result)
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
                    : FindCalendarFolder(defaultCalendar, BuildCalendarFolderName(calendar));
                if (useDefaultCalendarFolder)
                {
                    defaultCalendar = null;
                }
                if (targetFolder == null)
                {
                    return;
                }

                var groups = new Dictionary<string, List<OutlookScanItem>>(StringComparer.OrdinalIgnoreCase);
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

                        result.OutlookItems++;
                        string uid = ReadUserProperty(appointment, UidPropertyName);
                        string storedCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        bool linked = !string.IsNullOrWhiteSpace(uid)
                            && string.Equals(storedCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase);

                        string fingerprint = BuildFingerprint(
                            appointment.Subject,
                            appointment.Location,
                            appointment.Start,
                            appointment.End,
                            appointment.AllDayEvent);
                        List<OutlookScanItem> list;
                        if (!groups.TryGetValue(fingerprint, out list))
                        {
                            list = new List<OutlookScanItem>();
                            groups[fingerprint] = list;
                        }
                        list.Add(new OutlookScanItem { Uid = uid ?? string.Empty, Linked = linked });
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "CalDAV recovery scan could not inspect one Outlook appointment.",
                            ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV recovery AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV recovery folder item.");
                        }
                    }
                }

                foreach (List<OutlookScanItem> group in groups.Values)
                {
                    if (group.Count < 2 || !group.Any(v => v.Linked))
                    {
                        continue;
                    }

                    result.OutlookDuplicateGroups++;
                    int linkedCount = group.Count(v => v.Linked);
                    int unlinkedCount = group.Count - linkedCount;

                    bool exactLinkedPair = group.Count == 2
                        && linkedCount == 2
                        && !string.IsNullOrWhiteSpace(group[0].Uid)
                        && string.Equals(group[0].Uid, group[1].Uid, StringComparison.OrdinalIgnoreCase);
                    bool exactLinkedAndOriginalPair = group.Count == 2
                        && linkedCount == 1
                        && unlinkedCount == 1;

                    if (exactLinkedPair || exactLinkedAndOriginalPair)
                    {
                        result.OutlookSafeRemoveCandidates++;
                    }
                    else
                    {
                        result.OutlookAmbiguousGroups++;
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV recovery Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV recovery target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV recovery default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV recovery scan.");
            }
        }

        private static bool IsConnectorGeneratedUid(CalDavEventRecord item)
        {
            return item != null
                && !string.IsNullOrWhiteSpace(item.Uid)
                && item.Uid.Trim().EndsWith("@nc4ol", StringComparison.OrdinalIgnoreCase);
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CalDAV recovery user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CalDAV recovery user properties.");
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
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CalDAV recovery calendar folder.");
                    }
                }
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release CalDAV recovery folders collection.");
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

        private static string Safe(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<empty>" : value.Trim();
        }

        private sealed class OutlookScanItem
        {
            internal string Uid { get; set; }
            internal bool Linked { get; set; }
        }
    }
}
