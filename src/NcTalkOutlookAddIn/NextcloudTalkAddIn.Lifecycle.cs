// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Add-in lifecycle and bootstrap/teardown flow.
    public sealed partial class NextcloudTalkAddIn
    {
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _outlookApplication = (Outlook.Application)application;
            string outlookProfileName = ResolveCurrentOutlookProfileName();
            _settingsStorage = new NcTalkOutlookAddIn.Settings.SettingsStorage(outlookProfileName);
            _currentSettings = _settingsStorage.Load();
            ConfigureDiagnosticsLogger(_currentSettings);
            InitializeOutlookUiSynchronizationContext();
            TryApplyTransportSecurityFromSettings("startup", false);
            TryApplyOfficeUiLanguage();
            LogCore("Add-in connected (Outlook version=" + (_outlookApplication != null ? _outlookApplication.Version : "unknown") + ").");
            if (!string.IsNullOrWhiteSpace(outlookProfileName))
            {
                LogCore("Using Outlook profile settings: " + outlookProfileName + ".");
            }
            if (_currentSettings != null)
            {
                LogSettings("Settings loaded (AuthMode=" + _currentSettings.AuthMode + ", IFB=" + _currentSettings.IfbEnabled + ", IfbPort=" + _currentSettings.IfbPort + ", Debug=" + _currentSettings.DebugLoggingEnabled + ", LogAnonymize=" + _currentSettings.LogAnonymizationEnabled + ").");
            }

            _freeBusyManager = new FreeBusyManager(
                _settingsStorage.DataDirectory,
                outlookProfileName);
            _freeBusyManager.Initialize(_outlookApplication);
            InitializeTalkAppointmentSync(outlookProfileName);
            InitializeTalkRoomLifecycle(
                _settingsStorage.DataDirectory,
                outlookProfileName);
            EnsureApplicationHook();
            EnsureInspectorHook();
            ApplyIfbSettings();
            StartUpdateCheckIfDue();
            StartCardDavContactSync();
            StartCalDavCalendarBootstrap();
        }

        private void InitializeOutlookUiSynchronizationContext()
        {
            try
            {
                _uiSynchronizationContext = new OutlookUiSynchronizationContext();
                LogCore(
                    "Outlook UI synchronization context initialized (threadId="
                    + _uiSynchronizationContext.ThreadId.ToString(CultureInfo.InvariantCulture)
                    + ", apartment="
                    + Thread.CurrentThread.GetApartmentState()
                    + ").");
            }
            catch (Exception ex)
            {
                _uiSynchronizationContext = null;
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to initialize Outlook UI synchronization context.", ex);
            }
        }

        private void StartUpdateCheckIfDue()
        {
            if (_currentSettings == null || _settingsStorage == null)
            {
                return;
            }

            bool hadInstallId = !string.IsNullOrWhiteSpace(_currentSettings.UpdateInstallId);
            UpdateCheckService.EnsureInstallId(_currentSettings);
            if (!hadInstallId)
            {
                _settingsStorage.Save(_currentSettings);
            }

            AddinSettings updateSettings = _currentSettings.Clone();
            Task.Run(async () =>
            {
                try
                {
                    UpdateCheckResult result = await _updateCheckService.CheckAsync(updateSettings, false).ConfigureAwait(false);
                    StoreUpdateCheckSettings(updateSettings);
                    if (UpdateCheckService.ShouldNotify(updateSettings, result))
                    {
                        PostUpdateNotification(result);
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Update check failed.", ex);
                }
            });
        }

        private void StartCardDavContactSync()
        {
            if (_currentSettings == null
                || _outlookApplication == null
                || _uiSynchronizationContext == null)
            {
                return;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV read-only sync skipped because Nextcloud credentials are incomplete.");
                return;
            }

            CardDavSyncPreferences preferences = CardDavSyncPreferences.Load();
            if (!preferences.Configured)
            {
                try
                {
                    using (var selectionForm = new CardDavFirstRunForm(preferences))
                    {
                        DialogResult result = selectionForm.ShowDialog();
                        selectionForm.ApplyTo(preferences, result == DialogResult.OK);
                    }
                    preferences.Save();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "CardDAV first-run selection failed.",
                        ex);
                    return;
                }
            }

            if (!preferences.Enabled
                || (!preferences.SyncCompanyDirectory && !preferences.SyncPersonalContacts))
            {
                DiagnosticsLogger.Log(LogCategories.Core, "CardDAV read-only sync disabled by user settings.");
                return;
            }

            bool syncCompanyDirectory = preferences.SyncCompanyDirectory;
            bool syncPersonalContacts = preferences.SyncPersonalContacts;

            Task.Run(async () =>
            {
                try
                {
                    var sync = new CardDavReadOnlySync(configuration);
                    IList<CardDavAddressBook> addressBooks =
                        new DavDiscoveryService(configuration).DiscoverAddressBooks();
                    var contacts = new List<CardDavContactRecord>();
                    int selectedAddressBooks = 0;
                    foreach (CardDavAddressBook addressBook in addressBooks)
                    {
                        bool systemAddressBook = IsSystemCardDavAddressBook(addressBook);
                        if ((systemAddressBook && !syncCompanyDirectory)
                            || (!systemAddressBook && !syncPersonalContacts))
                        {
                            continue;
                        }

                        selectedAddressBooks++;
                        contacts.AddRange(sync.DownloadContacts(addressBook));
                    }

                    int count = await RunOnOutlookUiThreadAsync(
                        () => _outlookApplication != null
                            ? sync.ImportIntoOutlook(_outlookApplication, contacts)
                            : 0).ConfigureAwait(false);
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV read-only startup sync completed (addressBooks="
                        + selectedAddressBooks
                        + ", contacts="
                        + count
                        + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "CardDAV read-only startup sync failed.",
                        ex);
                }
            });
        }

        private void StartCalDavCalendarBootstrap()
        {
            if (_currentSettings == null
                || _outlookApplication == null
                || _uiSynchronizationContext == null)
            {
                return;
            }

            var configuration = new TalkServiceConfiguration(
                _currentSettings.ServerUrl,
                _currentSettings.Username,
                _currentSettings.AppPassword);
            if (!configuration.IsComplete())
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CalDAV calendar setup skipped because Nextcloud credentials are incomplete.");
                return;
            }

            Task.Run(async () =>
            {
                try
                {
                    IList<CalDavCalendar> calendars =
                        new DavDiscoveryService(configuration).DiscoverCalendars();
                    CalDavSyncPreferences preferences = CalDavSyncPreferences.Load();

                    if (!preferences.Configured)
                    {
                        await RunOnOutlookUiThreadAsync(
                            () =>
                            {
                                using (var selectionForm = new CalDavFirstRunForm(calendars, preferences))
                                {
                                    DialogResult result = selectionForm.ShowDialog();
                                    selectionForm.ApplyTo(preferences, result == DialogResult.OK);
                                }
                                preferences.Save();
                            }).ConfigureAwait(false);
                    }

                    if (!preferences.Enabled)
                    {
                        DiagnosticsLogger.Log(LogCategories.Core, "CalDAV calendar sync disabled by user settings.");
                        return;
                    }

                    var selectedCalendars = new List<CalDavCalendar>();
                    int readOnlySelectedCount = 0;
                    foreach (CalDavCalendar calendar in calendars)
                    {
                        if (!preferences.IsCalendarSelected(calendar.Href))
                        {
                            continue;
                        }
                        selectedCalendars.Add(calendar);
                        if (calendar.ReadOnly)
                        {
                            readOnlySelectedCount++;
                        }
                    }

                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CalDAV calendar configuration ready (available="
                        + calendars.Count
                        + ", selected="
                        + selectedCalendars.Count
                        + ", readOnlySelected="
                        + readOnlySelectedCount
                        + ", direction="
                        + preferences.Direction
                        + ", target="
                        + (preferences.UseDefaultCalendarFolder ? "default" : "separate")
                        + ", startup="
                        + preferences.SyncOnStartup
                        + ", intervalMinutes="
                        + preferences.IntervalMinutes
                        + ").");

                    if (!preferences.SyncOnStartup || selectedCalendars.Count == 0)
                    {
                        return;
                    }

                    var downloader = new CalDavEventSyncService(configuration);
                    var downloaded = new List<KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>>>();
                    foreach (CalDavCalendar calendar in selectedCalendars)
                    {
                        IList<CalDavEventRecord> remoteEvents =
                            preferences.Direction == CalDavSyncDirection.OutlookToNextcloud
                                ? new List<CalDavEventRecord>()
                                : downloader.DownloadEvents(calendar);
                        downloaded.Add(new KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>>(
                            calendar,
                            remoteEvents));
                    }

                    int repaired = await RunOnOutlookUiThreadAsync(
                        () =>
                        {
                            if (_outlookApplication == null)
                            {
                                return 0;
                            }
                            int changed = 0;
                            var guardedRepair = new CalDavBidirectionalSyncService(configuration);
                            foreach (KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>> pair in downloaded)
                            {
                                CalDavBidirectionalSyncResult one = guardedRepair.RepairGeneratedDuplicates(
                                    _outlookApplication,
                                    pair.Key,
                                    pair.Value,
                                    preferences.UseDefaultCalendarFolder);
                                changed += one.RecoveryRemoteRemoved;
                            }
                            return changed;
                        }).ConfigureAwait(false);

                    if (repaired > 0)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CalDAV recovery repair changed " + repaired.ToString(CultureInfo.InvariantCulture)
                            + " remote connector-generated items; refreshing calendar snapshot before sync.");
                        downloaded.Clear();
                        foreach (CalDavCalendar calendar in selectedCalendars)
                        {
                            IList<CalDavEventRecord> remoteEvents =
                                preferences.Direction == CalDavSyncDirection.OutlookToNextcloud
                                    ? new List<CalDavEventRecord>()
                                    : downloader.DownloadEvents(calendar);
                            downloaded.Add(new KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>>(
                                calendar,
                                remoteEvents));
                        }
                    }

                    bool recoveryFindings = await RunOnOutlookUiThreadAsync(
                        () =>
                        {
                            if (_outlookApplication == null)
                            {
                                return false;
                            }

                            bool found = false;
                            var recovery = new CalDavRecoveryScanService();
                            foreach (KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>> pair in downloaded)
                            {
                                CalDavRecoveryScanResult scan = recovery.Scan(
                                    _outlookApplication,
                                    pair.Key,
                                    pair.Value,
                                    preferences.UseDefaultCalendarFolder);

                                // Only actual recovery work blocks synchronization. Different legitimate
                                // DAV UIDs with an identical visible fingerprint are diagnostic findings,
                                // not a reason to stop normal synchronization.
                                if (scan.OutlookSafeRemoveCandidates > 0
                                    || scan.NextcloudSafeRemoveCandidates > 0
                                    || scan.NextcloudAmbiguousGroups > 0)
                                {
                                    found = true;
                                }
                            }
                            return found;
                        }).ConfigureAwait(false);

                    if (recoveryFindings)
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CalDAV startup merge paused because actionable recovery candidates remain. No normal calendar sync was started.");
                        return;
                    }

                    CalDavMergeResult mergeResult = await RunOnOutlookUiThreadAsync(
                        () =>
                        {
                            var total = new CalDavMergeResult();
                            if (_outlookApplication == null)
                            {
                                return total;
                            }

                            var merge = new CalDavSafeMergeService(configuration);
                            foreach (KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>> pair in downloaded)
                            {
                                CalDavMergeResult one = merge.Merge(
                                    _outlookApplication,
                                    pair.Key,
                                    pair.Value,
                                    preferences.UseDefaultCalendarFolder,
                                    preferences.Direction);
                                total.ImportedToOutlook += one.ImportedToOutlook;
                                total.UploadedToNextcloud += one.UploadedToNextcloud;
                                total.MatchedExisting += one.MatchedExisting;
                                total.SkippedRecurring += one.SkippedRecurring;
                                total.SkippedMeetings += one.SkippedMeetings;
                                total.UploadFailures += one.UploadFailures;
                            }
                            return total;
                        }).ConfigureAwait(false);

                    CalDavBidirectionalSyncResult syncResult = await RunOnOutlookUiThreadAsync(
                        () =>
                        {
                            var total = new CalDavBidirectionalSyncResult();
                            if (_outlookApplication == null)
                            {
                                return total;
                            }

                            var sync = new CalDavBidirectionalSyncService(configuration);
                            foreach (KeyValuePair<CalDavCalendar, IList<CalDavEventRecord>> pair in downloaded)
                            {
                                CalDavBidirectionalSyncResult one = sync.SyncLinkedEvents(
                                    _outlookApplication,
                                    pair.Key,
                                    pair.Value,
                                    preferences.UseDefaultCalendarFolder,
                                    preferences.Direction);
                                total.RemoteToOutlook += one.RemoteToOutlook;
                                total.OutlookToRemote += one.OutlookToRemote;
                                total.BaselinesEstablished += one.BaselinesEstablished;
                                total.Conflicts += one.Conflicts;
                                total.SkippedRecurring += one.SkippedRecurring;
                                total.SkippedMeetings += one.SkippedMeetings;
                                total.Failures += one.Failures;
                            }
                            return total;
                        }).ConfigureAwait(false);

                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CalDAV startup synchronization completed (calendars="
                        + selectedCalendars.Count
                        + ", imported="
                        + mergeResult.ImportedToOutlook
                        + ", uploaded="
                        + mergeResult.UploadedToNextcloud
                        + ", matched="
                        + mergeResult.MatchedExisting
                        + ", remoteUpdates="
                        + syncResult.RemoteToOutlook
                        + ", localUpdates="
                        + syncResult.OutlookToRemote
                        + ", baselines="
                        + syncResult.BaselinesEstablished
                        + ", conflicts="
                        + syncResult.Conflicts
                        + ", recurringSkipped="
                        + (mergeResult.SkippedRecurring + syncResult.SkippedRecurring)
                        + ", meetingsSkipped="
                        + (mergeResult.SkippedMeetings + syncResult.SkippedMeetings)
                        + ", failures="
                        + (mergeResult.UploadFailures + syncResult.Failures)
                        + ", deletions=0).");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "CalDAV calendar setup/startup synchronization failed.",
                        ex);
                }
            });
        }

        private static bool IsSystemCardDavAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(
                    "z-server-generated--system",
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void StoreUpdateCheckSettings(AddinSettings updateSettings)
        {
            if (updateSettings == null || _currentSettings == null || _settingsStorage == null)
            {
                return;
            }

            _currentSettings.UpdateInstallId = updateSettings.UpdateInstallId ?? string.Empty;
            _currentSettings.UpdateLastCheckedAtUtc = updateSettings.UpdateLastCheckedAtUtc ?? string.Empty;
            _currentSettings.UpdateLatestVersion = updateSettings.UpdateLatestVersion ?? string.Empty;
            _currentSettings.UpdateReleaseUrl = updateSettings.UpdateReleaseUrl ?? string.Empty;
            _currentSettings.UpdateDownloadUrl = updateSettings.UpdateDownloadUrl ?? string.Empty;
            _currentSettings.UpdatePublishedAt = updateSettings.UpdatePublishedAt ?? string.Empty;
            _currentSettings.UpdateChangelogTitle = updateSettings.UpdateChangelogTitle ?? string.Empty;
            _currentSettings.UpdateChangelogText = updateSettings.UpdateChangelogText ?? string.Empty;
            _settingsStorage.Save(_currentSettings);
        }

        private void PostUpdateNotification(UpdateCheckResult result)
        {
            SynchronizationContext context = _uiSynchronizationContext;
            if (context == null)
            {
                DiagnosticsLogger.Log(LogCategories.Core, "Update notification skipped because no UI context is available.");
                return;
            }

            context.Post(_ => ShowUpdateNotification(result), null);
        }

        private void ShowUpdateNotification(UpdateCheckResult result)
        {
            try
            {
                if (_currentSettings == null || !UpdateCheckService.ShouldNotify(_currentSettings, result))
                {
                    return;
                }

                UpdateCheckService.MarkNotified(_currentSettings, result);
                if (_settingsStorage != null)
                {
                    _settingsStorage.Save(_currentSettings);
                }

                DialogResult answer = MessageBox.Show(
                    UpdateCheckService.BuildNotificationMessage(result),
                    Strings.UpdateAvailableTitle,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (answer != DialogResult.Yes)
                {
                    return;
                }

                string url = UpdateCheckService.GetPreferredOpenUrl(result);
                if (!string.IsNullOrWhiteSpace(url))
                {
                    BrowserLauncher.OpenUrl(
                        url,
                        LogCategories.Core,
                        "Failed to open update download URL.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Update notification failed.", ex);
            }
        }

        private void TryApplyOfficeUiLanguage()
        {
            try
            {
                if (_outlookApplication == null)
                {
                    return;
                }

                LanguageSettings languageSettings = _outlookApplication.LanguageSettings;
                if (languageSettings == null)
                {
                    return;
                }
                int lcid = languageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
                CultureInfo culture = CultureInfo.GetCultureInfo(lcid);
                Strings.SetPreferredUiLanguage(culture.Name);
                LogCore("Office UI language detected: " + culture.Name + " (LCID=" + lcid + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to detect Office UI language.", ex);
            }
        }

        private string ResolveCurrentOutlookProfileName()
        {
            if (_outlookApplication == null)
            {
                return string.Empty;
            }

            object session = null;
            try
            {
                session = _outlookApplication.Session;
                if (session == null)
                {
                    return string.Empty;
                }

                object rawProfileName = session.GetType().InvokeMember(
                    "CurrentProfileName",
                    BindingFlags.GetProperty,
                    null,
                    session,
                    null,
                    CultureInfo.InvariantCulture);

                string profileName = rawProfileName as string;
                return string.IsNullOrWhiteSpace(profileName) ? string.Empty : profileName.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve current Outlook profile name.", ex);
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Core,
                    "Failed to release Outlook session COM object after profile resolution.");
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            TearDownAddInState("disconnect", true);
            LogCore("Add-in disconnected (removeMode=" + removeMode + ").");
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
            TearDownAddInState("shutdown", false);
        }

        private void TearDownAddInState(string origin, bool clearOutlookApplication)
        {
            UnhookApplication();
            UnhookInspector();
            UnhookMailComposeSubscriptions();
            DisposeTalkAppointmentSync();
            DisposeTalkAppointmentSubscriptions();
            DisposeTalkRoomLifecycle();
            if (_freeBusyManager != null && _currentSettings != null && _currentSettings.IfbEnabled)
            {
                try
                {
                    var clone = _currentSettings.Clone();
                    clone.IfbEnabled = false;
                    _freeBusyManager.ApplySettings(clone);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Ifb,
                        "Failed to disable IFB during add-in " + (origin ?? "teardown") + ".",
                        ex);
                }
            }
            if (_freeBusyManager != null)
            {
                _freeBusyManager.Dispose();
            }
            _freeBusyManager = null;

            if (clearOutlookApplication)
            {
                _outlookApplication = null;
            }

            _ribbonUi = null;
            OutlookUiSynchronizationContext uiSynchronizationContext = _uiSynchronizationContext;
            _uiSynchronizationContext = null;
            if (uiSynchronizationContext != null)
            {
                try
                {
                    uiSynchronizationContext.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to dispose Outlook UI synchronization context.", ex);
                }
            }
            _deferredAppointmentEnsureState.ClearPendingKeys();
        }
    }
}
