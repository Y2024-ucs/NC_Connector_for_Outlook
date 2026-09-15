// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
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
            get { return OutlookDuplicateGroups > 0 || NextcloudDuplicateGroups > 0; }
        }
    }

    internal sealed class CalDavRecoveryCleanupResult
    {
        internal int OutlookRemoved { get; set; }
        internal int OutlookRelinked { get; set; }
        internal int OutlookSkipped { get; set; }
        internal int NextcloudRemoved { get; set; }
        internal int NextcloudRelinked { get; set; }
        internal int NextcloudSkipped { get; set; }
        internal int Failures { get; set; }
    }

    internal sealed class CalDavRecoveryScanService
    {
        private const string UidPropertyName = "NC-CalDAV-UID";
        private const string HrefPropertyName = "NC-CalDAV-HREF";
        private const string ETagPropertyName = "NC-CalDAV-ETAG";
        private const string CalendarPropertyName = "NC-CalDAV-CALENDAR";
        private const string ArmFileName = "caldav-recovery-cleanup.enable";
        private const string DoneFileName = "caldav-recovery-cleanup.done";

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

            CalDavRecoveryScanResult result = new CalDavRecoveryScanResult();
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

            if (result.HasFindings && IsCleanupArmed())
            {
                CalDavRecoveryCleanupResult cleanupResult = RunCleanup(
                    outlookApplication,
                    calendar,
                    remoteEvents,
                    useDefaultCalendarFolder);
                MarkCleanupCompleted(cleanupResult);
            }
            else if (result.HasFindings)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV recovery cleanup is not armed. Create " + GetArmFilePath()
                    + " to remove only safe duplicate candidates on the next Outlook start.");
            }

            return result;
        }

        private static void ScanNextcloud(IList<CalDavEventRecord> remoteEvents, CalDavRecoveryScanResult result)
        {
            Dictionary<string, List<CalDavEventRecord>> groups = GroupRemote(remoteEvents);
            foreach (CalDavEventRecord item in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (item != null)
                {
                    result.NextcloudItems++;
                }
            }

            int ambiguousIndex = 0;
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

                // Recovery stage 2: one connector-generated object is safe to remove
                // whenever at least one non-generated object with the same fingerprint exists.
                // Existing non-generated objects are never merged or deleted here.
                if (generated == 1 && original >= 1)
                {
                    result.NextcloudSafeRemoveCandidates++;
                }
                else
                {
                    result.NextcloudAmbiguousGroups++;
                    ambiguousIndex++;
                    LogNextcloudAmbiguousGroup(ambiguousIndex, group, generated, original);
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

                Dictionary<string, List<OutlookScanItem>> groups =
                    new Dictionary<string, List<OutlookScanItem>>(StringComparer.OrdinalIgnoreCase);
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

                        list.Add(new OutlookScanItem
                        {
                            Subject = appointment.Subject ?? string.Empty,
                            Location = appointment.Location ?? string.Empty,
                            Start = appointment.Start,
                            End = appointment.End,
                            AllDay = appointment.AllDayEvent,
                            Recurring = appointment.IsRecurring,
                            Meeting = appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting,
                            Uid = uid ?? string.Empty,
                            Href = ReadUserProperty(appointment, HrefPropertyName),
                            ETag = ReadUserProperty(appointment, ETagPropertyName),
                            Linked = linked
                        });
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery scan could not inspect one Outlook appointment.", ex);
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

                int ambiguousIndex = 0;
                foreach (KeyValuePair<string, List<OutlookScanItem>> pair in groups)
                {
                    List<OutlookScanItem> group = pair.Value;
                    if (group.Count < 2 || !group.Any(v => v.Linked))
                    {
                        continue;
                    }

                    result.OutlookDuplicateGroups++;
                    List<OutlookScanItem> linked = group.Where(v => v.Linked).ToList();
                    int unlinkedCount = group.Count - linked.Count;

                    bool exactLinkedPair = group.Count == 2
                        && linked.Count == 2
                        && !string.IsNullOrWhiteSpace(linked[0].Uid)
                        && string.Equals(linked[0].Uid, linked[1].Uid, StringComparison.OrdinalIgnoreCase);
                    bool exactLinkedAndOriginalPair = group.Count == 2
                        && linked.Count == 1
                        && unlinkedCount == 1;
                    bool generatedWithExistingOriginal = linked.Any(v => IsConnectorGeneratedUid(v.Uid) && !v.Meeting)
                        && linked.Any(v => !IsConnectorGeneratedUid(v.Uid));
                    bool exactLocalClone = HasExactNonMeetingClone(linked);

                    if (exactLinkedPair || exactLinkedAndOriginalPair || generatedWithExistingOriginal || exactLocalClone)
                    {
                        result.OutlookSafeRemoveCandidates++;
                    }
                    else
                    {
                        result.OutlookAmbiguousGroups++;
                        ambiguousIndex++;
                        LogOutlookAmbiguousGroup(ambiguousIndex, pair.Key, group, linked.Count, unlinkedCount);
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

        private static bool HasExactNonMeetingClone(IList<OutlookScanItem> linked)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (OutlookScanItem item in linked)
            {
                if (item == null || item.Meeting)
                {
                    continue;
                }
                string identity = BuildSyncIdentity(item.Uid, item.Href, item.ETag);
                if (string.IsNullOrWhiteSpace(identity))
                {
                    continue;
                }
                int count;
                counts.TryGetValue(identity, out count);
                count++;
                counts[identity] = count;
                if (count > 1)
                {
                    return true;
                }
            }
            return false;
        }

        private static void LogOutlookAmbiguousGroup(
            int index,
            string fingerprint,
            IList<OutlookScanItem> group,
            int linkedCount,
            int unlinkedCount)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("CalDAV recovery detail OUTLOOK ambiguous #")
                .Append(index.ToString(CultureInfo.InvariantCulture))
                .Append(" (count=").Append(group.Count.ToString(CultureInfo.InvariantCulture))
                .Append(", linked=").Append(linkedCount.ToString(CultureInfo.InvariantCulture))
                .Append(", unlinked=").Append(unlinkedCount.ToString(CultureInfo.InvariantCulture))
                .Append(", fingerprint=").Append(Safe(fingerprint)).Append(", items=");

            for (int i = 0; i < group.Count; i++)
            {
                OutlookScanItem item = group[i];
                if (i > 0)
                {
                    builder.Append(" || ");
                }
                builder.Append("[").Append(i + 1)
                    .Append(": subject=").Append(Safe(item.Subject))
                    .Append(", location=").Append(Safe(item.Location))
                    .Append(", start=").Append(item.Start.ToString("o", CultureInfo.InvariantCulture))
                    .Append(", end=").Append(item.End.ToString("o", CultureInfo.InvariantCulture))
                    .Append(", allDay=").Append(item.AllDay)
                    .Append(", recurring=").Append(item.Recurring)
                    .Append(", meeting=").Append(item.Meeting)
                    .Append(", linked=").Append(item.Linked)
                    .Append(", uid=").Append(Safe(item.Uid))
                    .Append(", href=").Append(Safe(item.Href))
                    .Append(", etag=").Append(Safe(item.ETag))
                    .Append("]");
            }
            builder.Append(").");
            DiagnosticsLogger.Log(LogCategories.Core, builder.ToString());
        }

        private static void LogNextcloudAmbiguousGroup(
            int index,
            IList<CalDavEventRecord> group,
            int generated,
            int original)
        {
            StringBuilder builder = new StringBuilder();
            CalDavEventRecord first = group.FirstOrDefault(v => v != null);
            builder.Append("CalDAV recovery detail NEXTCLOUD ambiguous #")
                .Append(index.ToString(CultureInfo.InvariantCulture))
                .Append(" (count=").Append(group.Count.ToString(CultureInfo.InvariantCulture))
                .Append(", generated=").Append(generated.ToString(CultureInfo.InvariantCulture))
                .Append(", original=").Append(original.ToString(CultureInfo.InvariantCulture));
            if (first != null)
            {
                builder.Append(", subject=").Append(Safe(first.Subject))
                    .Append(", location=").Append(Safe(first.Location))
                    .Append(", start=").Append(first.Start.ToString("o", CultureInfo.InvariantCulture))
                    .Append(", end=").Append(first.End.ToString("o", CultureInfo.InvariantCulture))
                    .Append(", allDay=").Append(first.AllDay);
            }
            builder.Append(", items=");
            for (int i = 0; i < group.Count; i++)
            {
                CalDavEventRecord item = group[i];
                if (i > 0)
                {
                    builder.Append(" || ");
                }
                builder.Append("[").Append(i + 1)
                    .Append(": generated=").Append(IsConnectorGeneratedUid(item))
                    .Append(", uid=").Append(Safe(item != null ? item.Uid : string.Empty))
                    .Append(", href=").Append(Safe(item != null ? item.Href : string.Empty))
                    .Append(", etag=").Append(Safe(item != null ? item.ETag : string.Empty))
                    .Append(", recurring=").Append(item != null && item.Recurring)
                    .Append("]");
            }
            builder.Append(").");
            DiagnosticsLogger.Log(LogCategories.Core, builder.ToString());
        }

        private static CalDavRecoveryCleanupResult RunCleanup(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder)
        {
            CalDavRecoveryCleanupResult result = new CalDavRecoveryCleanupResult();
            TalkServiceConfiguration configuration = BuildConfiguration(outlookApplication);
            if (configuration == null || !configuration.IsComplete())
            {
                result.Failures++;
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV recovery cleanup was armed but Nextcloud credentials could not be loaded. Nothing was deleted.");
                return result;
            }

            CleanupOutlook(outlookApplication, calendar, remoteEvents, useDefaultCalendarFolder, result);
            CleanupNextcloud(outlookApplication, calendar, remoteEvents, useDefaultCalendarFolder, configuration, result);

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CalDAV recovery cleanup completed (calendar=" + Safe(calendar.DisplayName)
                + ", outlookRemoved=" + result.OutlookRemoved
                + ", outlookRelinked=" + result.OutlookRelinked
                + ", outlookSkipped=" + result.OutlookSkipped
                + ", nextcloudRemoved=" + result.NextcloudRemoved
                + ", nextcloudRelinked=" + result.NextcloudRelinked
                + ", nextcloudSkipped=" + result.NextcloudSkipped
                + ", failures=" + result.Failures
                + ").");
            return result;
        }

        private static void CleanupOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            CalDavRecoveryCleanupResult result)
        {
            Dictionary<string, List<CalDavEventRecord>> remoteByFingerprint = GroupRemote(remoteEvents);
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

                Dictionary<string, List<OutlookCleanupItem>> groups =
                    new Dictionary<string, List<OutlookCleanupItem>>(StringComparer.OrdinalIgnoreCase);
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

                        string fingerprint = BuildFingerprint(
                            appointment.Subject,
                            appointment.Location,
                            appointment.Start,
                            appointment.End,
                            appointment.AllDayEvent);
                        List<OutlookCleanupItem> list;
                        if (!groups.TryGetValue(fingerprint, out list))
                        {
                            list = new List<OutlookCleanupItem>();
                            groups[fingerprint] = list;
                        }

                        list.Add(new OutlookCleanupItem
                        {
                            EntryId = appointment.EntryID ?? string.Empty,
                            Uid = ReadUserProperty(appointment, UidPropertyName),
                            Href = ReadUserProperty(appointment, HrefPropertyName),
                            ETag = ReadUserProperty(appointment, ETagPropertyName),
                            CalendarHref = ReadUserProperty(appointment, CalendarPropertyName),
                            Meeting = appointment.MeetingStatus != Outlook.OlMeetingStatus.olNonMeeting
                        });
                    }
                    catch (Exception ex)
                    {
                        result.Failures++;
                        DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery cleanup could not inspect one Outlook appointment.", ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV cleanup AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV cleanup folder item.");
                        }
                    }
                }

                foreach (KeyValuePair<string, List<OutlookCleanupItem>> pair in groups)
                {
                    List<OutlookCleanupItem> group = pair.Value;
                    if (group.Count < 2)
                    {
                        continue;
                    }

                    List<OutlookCleanupItem> linked = group.Where(v => v.IsLinkedTo(calendar.Href)).ToList();
                    List<OutlookCleanupItem> unlinked = group.Where(v => !v.IsLinkedTo(calendar.Href)).ToList();
                    HashSet<string> removedEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    bool changed = false;

                    // Stage 2a: remove only connector-generated Outlook copies when at least
                    // one non-generated linked copy with the same fingerprint exists.
                    List<OutlookCleanupItem> generatedLinked = linked
                        .Where(v => IsConnectorGeneratedUid(v.Uid) && !v.Meeting)
                        .ToList();
                    List<OutlookCleanupItem> originalLinked = linked
                        .Where(v => !IsConnectorGeneratedUid(v.Uid))
                        .ToList();
                    if (generatedLinked.Count > 0 && originalLinked.Count > 0)
                    {
                        foreach (OutlookCleanupItem generated in generatedLinked)
                        {
                            if (DeleteOutlookByEntryId(session, generated.EntryId))
                            {
                                result.OutlookRemoved++;
                                removedEntryIds.Add(generated.EntryId);
                                changed = true;
                            }
                            else
                            {
                                result.Failures++;
                            }
                        }
                    }

                    // Stage 2b: collapse only exact non-meeting local clones that carry
                    // identical UID + HREF + ETag. Different UIDs are never merged here.
                    Dictionary<string, List<OutlookCleanupItem>> exactGroups =
                        new Dictionary<string, List<OutlookCleanupItem>>(StringComparer.OrdinalIgnoreCase);
                    foreach (OutlookCleanupItem item in linked)
                    {
                        if (item.Meeting || removedEntryIds.Contains(item.EntryId))
                        {
                            continue;
                        }
                        string identity = BuildSyncIdentity(item.Uid, item.Href, item.ETag);
                        if (string.IsNullOrWhiteSpace(identity))
                        {
                            continue;
                        }
                        List<OutlookCleanupItem> same;
                        if (!exactGroups.TryGetValue(identity, out same))
                        {
                            same = new List<OutlookCleanupItem>();
                            exactGroups[identity] = same;
                        }
                        same.Add(item);
                    }
                    foreach (List<OutlookCleanupItem> exactGroup in exactGroups.Values)
                    {
                        for (int i = 1; i < exactGroup.Count; i++)
                        {
                            OutlookCleanupItem duplicate = exactGroup[i];
                            if (DeleteOutlookByEntryId(session, duplicate.EntryId))
                            {
                                result.OutlookRemoved++;
                                removedEntryIds.Add(duplicate.EntryId);
                                changed = true;
                            }
                            else
                            {
                                result.Failures++;
                            }
                        }
                    }

                    if (changed)
                    {
                        continue;
                    }

                    // Keep the original conservative two-item recovery rules.
                    if (group.Count == 2
                        && linked.Count == 2
                        && !linked[0].Meeting
                        && !linked[1].Meeting
                        && !string.IsNullOrWhiteSpace(linked[0].Uid)
                        && string.Equals(linked[0].Uid, linked[1].Uid, StringComparison.OrdinalIgnoreCase))
                    {
                        if (DeleteOutlookByEntryId(session, linked[1].EntryId))
                        {
                            result.OutlookRemoved++;
                        }
                        else
                        {
                            result.Failures++;
                        }
                        continue;
                    }

                    if (group.Count == 2 && linked.Count == 1 && unlinked.Count == 1
                        && !linked[0].Meeting && !unlinked[0].Meeting)
                    {
                        OutlookCleanupItem remove = linked[0];
                        OutlookCleanupItem keep = unlinked[0];
                        CalDavEventRecord preferredRemote = FindPreferredOriginalRemote(pair.Key, remoteByFingerprint);
                        string uid = preferredRemote != null ? preferredRemote.Uid : remove.Uid;
                        string href = preferredRemote != null ? preferredRemote.Href : remove.Href;
                        string etag = preferredRemote != null ? preferredRemote.ETag : remove.ETag;

                        if (RelinkOutlookByEntryId(session, keep.EntryId, uid, href, etag, calendar.Href))
                        {
                            result.OutlookRelinked++;
                            if (DeleteOutlookByEntryId(session, remove.EntryId))
                            {
                                result.OutlookRemoved++;
                            }
                            else
                            {
                                result.Failures++;
                            }
                        }
                        else
                        {
                            result.Failures++;
                        }
                        continue;
                    }

                    result.OutlookSkipped++;
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV cleanup Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV cleanup target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV cleanup default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV cleanup.");
            }
        }

        private static void CleanupNextcloud(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            IList<CalDavEventRecord> remoteEvents,
            bool useDefaultCalendarFolder,
            TalkServiceConfiguration configuration,
            CalDavRecoveryCleanupResult result)
        {
            Dictionary<string, List<CalDavEventRecord>> groups = GroupRemote(remoteEvents);
            foreach (List<CalDavEventRecord> group in groups.Values)
            {
                if (group.Count < 2)
                {
                    continue;
                }

                List<CalDavEventRecord> generated = group.Where(IsConnectorGeneratedUid).ToList();
                List<CalDavEventRecord> original = group.Where(v => !IsConnectorGeneratedUid(v)).ToList();
                if (generated.Count != 1 || original.Count < 1)
                {
                    continue;
                }

                CalDavEventRecord generatedItem = generated[0];
                if (string.IsNullOrWhiteSpace(generatedItem.Href) || string.IsNullOrWhiteSpace(generatedItem.ETag))
                {
                    result.NextcloudSkipped++;
                    continue;
                }

                if (original.Count == 1)
                {
                    result.NextcloudRelinked += RelinkGeneratedUidInOutlook(
                        outlookApplication,
                        calendar,
                        generatedItem,
                        original[0],
                        useDefaultCalendarFolder);
                }
                else
                {
                    // With multiple existing originals there is no safe target to relink to.
                    // The generated Outlook copy must already have been removed in stage 2a.
                    int remainingGenerated = CountLinkedUidInOutlook(
                        outlookApplication,
                        calendar,
                        generatedItem.Uid,
                        useDefaultCalendarFolder);
                    if (remainingGenerated > 0)
                    {
                        result.NextcloudSkipped++;
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CalDAV recovery kept @nc4ol remote object because a linked Outlook copy still exists (uid="
                            + Safe(generatedItem.Uid) + ", count="
                            + remainingGenerated.ToString(CultureInfo.InvariantCulture) + ").");
                        continue;
                    }
                }

                if (DeleteRemoteGeneratedItem(configuration, generatedItem))
                {
                    result.NextcloudRemoved++;
                }
                else
                {
                    result.Failures++;
                }
            }
        }

        private static int CountLinkedUidInOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            string uid,
            bool useDefaultCalendarFolder)
        {
            if (string.IsNullOrWhiteSpace(uid))
            {
                return 0;
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultCalendar = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items items = null;
            int found = 0;
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
                    return 0;
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
                        string currentUid = ReadUserProperty(appointment, UidPropertyName);
                        string storedCalendar = ReadUserProperty(appointment, CalendarPropertyName);
                        if (string.Equals(currentUid, uid, StringComparison.OrdinalIgnoreCase)
                            && string.Equals(storedCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            found++;
                        }
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV verification AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV verification folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV verification Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV verification target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV verification default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV verification.");
            }
            return found;
        }

        private static bool DeleteRemoteGeneratedItem(TalkServiceConfiguration configuration, CalDavEventRecord item)
        {
            Dictionary<string, string> headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "If-Match", item.ETag.Trim() }
            };
            NcHttpResponse response = new NcHttpClient(configuration).Send(new NcHttpRequestOptions
            {
                Method = "DELETE",
                Url = item.Href,
                Accept = "*/*",
                Headers = headers,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                TimeoutMs = 60000
            });

            bool ok = response.HasHttpResponse
                && (int)response.StatusCode >= 200
                && (int)response.StatusCode < 300;
            if (!ok)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV recovery DELETE failed/skipped (href=" + Safe(item.Href)
                    + ", uid=" + Safe(item.Uid)
                    + ", http=" + (response.HasHttpResponse
                        ? ((int)response.StatusCode).ToString(CultureInfo.InvariantCulture)
                        : "transport")
                    + ").");
            }
            return ok;
        }

        private static int RelinkGeneratedUidInOutlook(
            Outlook.Application outlookApplication,
            CalDavCalendar calendar,
            CalDavEventRecord generated,
            CalDavEventRecord original,
            bool useDefaultCalendarFolder)
        {
            if (string.IsNullOrWhiteSpace(generated.Uid))
            {
                return 0;
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultCalendar = null;
            Outlook.MAPIFolder targetFolder = null;
            Outlook.Items items = null;
            int updated = 0;
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
                    return 0;
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
                        if (!string.Equals(uid, generated.Uid, StringComparison.OrdinalIgnoreCase)
                            || !string.Equals(storedCalendar, calendar.Href, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        WriteSyncProperties(appointment, original.Uid, original.Href, original.ETag, calendar.Href);
                        appointment.Save();
                        updated++;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery could not relink one @nc4ol Outlook appointment.", ex);
                    }
                    finally
                    {
                        if (appointment != null)
                        {
                            ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release CalDAV relink AppointmentItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV relink folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CalDAV relink Items collection.");
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CalDAV relink target folder.");
                ComInteropScope.TryRelease(defaultCalendar, LogCategories.Core, "Failed to release CalDAV relink default folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDAV relink.");
            }
            return updated;
        }

        private static TalkServiceConfiguration BuildConfiguration(Outlook.Application outlookApplication)
        {
            object session = null;
            try
            {
                session = outlookApplication.Session;
                object rawProfileName = session.GetType().InvokeMember(
                    "CurrentProfileName",
                    BindingFlags.GetProperty,
                    null,
                    session,
                    null,
                    CultureInfo.InvariantCulture);
                string profileName = rawProfileName as string;
                SettingsStorage storage = new SettingsStorage(profileName);
                AddinSettings settings = storage.Load();
                return new TalkServiceConfiguration(settings.ServerUrl, settings.Username, settings.AppPassword);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery could not load current profile credentials.", ex);
                return null;
            }
            finally
            {
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after recovery configuration load.");
            }
        }

        private static bool DeleteOutlookByEntryId(Outlook.NameSpace session, string entryId)
        {
            if (session == null || string.IsNullOrWhiteSpace(entryId))
            {
                return false;
            }
            object raw = null;
            Outlook.AppointmentItem appointment = null;
            try
            {
                raw = session.GetItemFromID(entryId, Type.Missing);
                appointment = raw as Outlook.AppointmentItem;
                if (appointment == null)
                {
                    return false;
                }
                appointment.Delete();
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery could not delete one Outlook duplicate.", ex);
                return false;
            }
            finally
            {
                if (appointment != null)
                {
                    ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release deleted CalDAV AppointmentItem.");
                }
                else
                {
                    ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV cleanup item.");
                }
            }
        }

        private static bool RelinkOutlookByEntryId(
            Outlook.NameSpace session,
            string entryId,
            string uid,
            string href,
            string etag,
            string calendarHref)
        {
            if (session == null || string.IsNullOrWhiteSpace(entryId))
            {
                return false;
            }
            object raw = null;
            Outlook.AppointmentItem appointment = null;
            try
            {
                raw = session.GetItemFromID(entryId, Type.Missing);
                appointment = raw as Outlook.AppointmentItem;
                if (appointment == null)
                {
                    return false;
                }
                WriteSyncProperties(appointment, uid, href, etag, calendarHref);
                appointment.Save();
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery could not relink one Outlook appointment.", ex);
                return false;
            }
            finally
            {
                if (appointment != null)
                {
                    ComInteropScope.TryRelease(appointment, LogCategories.Core, "Failed to release relinked CalDAV AppointmentItem.");
                }
                else
                {
                    ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release CalDAV relink item.");
                }
            }
        }

        private static string GetArmFilePath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), ArmFileName);
        }

        private static string GetDoneFilePath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), DoneFileName);
        }

        private static bool IsCleanupArmed()
        {
            return File.Exists(GetArmFilePath());
        }

        private static void MarkCleanupCompleted(CalDavRecoveryCleanupResult result)
        {
            try
            {
                string donePath = GetDoneFilePath();
                string text = "completedUtc=" + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
                    + Environment.NewLine
                    + "outlookRemoved=" + result.OutlookRemoved.ToString(CultureInfo.InvariantCulture)
                    + Environment.NewLine
                    + "outlookRelinked=" + result.OutlookRelinked.ToString(CultureInfo.InvariantCulture)
                    + Environment.NewLine
                    + "nextcloudRemoved=" + result.NextcloudRemoved.ToString(CultureInfo.InvariantCulture)
                    + Environment.NewLine
                    + "nextcloudRelinked=" + result.NextcloudRelinked.ToString(CultureInfo.InvariantCulture)
                    + Environment.NewLine
                    + "failures=" + result.Failures.ToString(CultureInfo.InvariantCulture)
                    + Environment.NewLine;
                File.WriteAllText(donePath, text);
                string armPath = GetArmFilePath();
                if (File.Exists(armPath))
                {
                    File.Delete(armPath);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV recovery could not update one-shot marker files.", ex);
            }
        }

        private static Dictionary<string, List<CalDavEventRecord>> GroupRemote(IList<CalDavEventRecord> remoteEvents)
        {
            Dictionary<string, List<CalDavEventRecord>> groups =
                new Dictionary<string, List<CalDavEventRecord>>(StringComparer.OrdinalIgnoreCase);
            foreach (CalDavEventRecord item in remoteEvents ?? Enumerable.Empty<CalDavEventRecord>())
            {
                if (item == null)
                {
                    continue;
                }
                string fingerprint = BuildFingerprint(item.Subject, item.Location, item.Start, item.End, item.AllDay);
                List<CalDavEventRecord> list;
                if (!groups.TryGetValue(fingerprint, out list))
                {
                    list = new List<CalDavEventRecord>();
                    groups[fingerprint] = list;
                }
                list.Add(item);
            }
            return groups;
        }

        private static CalDavEventRecord FindPreferredOriginalRemote(
            string fingerprint,
            Dictionary<string, List<CalDavEventRecord>> remoteByFingerprint)
        {
            List<CalDavEventRecord> group;
            if (!remoteByFingerprint.TryGetValue(fingerprint, out group))
            {
                return null;
            }
            List<CalDavEventRecord> originals = group.Where(v => !IsConnectorGeneratedUid(v)).ToList();
            return originals.Count == 1 ? originals[0] : null;
        }

        private static bool IsConnectorGeneratedUid(CalDavEventRecord item)
        {
            return item != null && IsConnectorGeneratedUid(item.Uid);
        }

        private static bool IsConnectorGeneratedUid(string uid)
        {
            return !string.IsNullOrWhiteSpace(uid)
                && uid.Trim().EndsWith("@nc4ol", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSyncIdentity(string uid, string href, string etag)
        {
            if (string.IsNullOrWhiteSpace(uid)
                || string.IsNullOrWhiteSpace(href)
                || string.IsNullOrWhiteSpace(etag))
            {
                return string.Empty;
            }
            return uid.Trim() + "|" + href.Trim() + "|" + etag.Trim();
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
            internal string Subject { get; set; }
            internal string Location { get; set; }
            internal DateTime Start { get; set; }
            internal DateTime End { get; set; }
            internal bool AllDay { get; set; }
            internal bool Recurring { get; set; }
            internal bool Meeting { get; set; }
            internal string Uid { get; set; }
            internal string Href { get; set; }
            internal string ETag { get; set; }
            internal bool Linked { get; set; }
        }

        private sealed class OutlookCleanupItem
        {
            internal string EntryId { get; set; }
            internal string Uid { get; set; }
            internal string Href { get; set; }
            internal string ETag { get; set; }
            internal string CalendarHref { get; set; }
            internal bool Meeting { get; set; }

            internal bool IsLinkedTo(string calendarHref)
            {
                return !string.IsNullOrWhiteSpace(Uid)
                    && string.Equals(CalendarHref, calendarHref, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
