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
using Timer = System.Threading.Timer;

namespace NcTalkOutlookAddIn
{
    // Add-in lifecycle and bootstrap/teardown flow.
    public sealed partial class NextcloudTalkAddIn
    {
        private static readonly TimeSpan CardDavStartupDelay = TimeSpan.FromSeconds(15);
        private static readonly TimeSpan CardDavUiPhaseYieldDelay = TimeSpan.FromMilliseconds(75);
        private Timer _cardDavStartupDelayTimer;
        private CardDavSyncStatusForm _cardDavSyncStatusForm;

        private sealed class CardDavSyncStatusForm : Form
        {
            private readonly Label _statusLabel;
            private readonly ProgressBar _progressBar;
            private readonly System.Windows.Forms.Timer _hideTimer;

            internal CardDavSyncStatusForm()
            {
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                TopMost = false;
                MaximizeBox = false;
                MinimizeBox = false;
                ControlBox = false;
                Width = 390;
                Height = 92;
                Text = "NC Connector";

                _statusLabel = new Label
                {
                    AutoSize = false,
                    Left = 12,
                    Top = 10,
                    Width = 356,
                    Height = 34,
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };
                Controls.Add(_statusLabel);

                _progressBar = new ProgressBar
                {
                    Left = 12,
                    Top = 51,
                    Width = 356,
                    Height = 15,
                    Style = ProgressBarStyle.Marquee,
                    MarqueeAnimationSpeed = 24
                };
                Controls.Add(_progressBar);

                _hideTimer = new System.Windows.Forms.Timer();
                _hideTimer.Tick += delegate
                {
                    _hideTimer.Stop();
                    Hide();
                };
            }

            protected override bool ShowWithoutActivation
            {
                get { return true; }
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    const int WS_EX_NOACTIVATE = 0x08000000;
                    CreateParams parameters = base.CreateParams;
                    parameters.ExStyle |= WS_EX_NOACTIVATE;
                    return parameters;
                }
            }

            internal void SetStatus(string text, bool busy)
            {
                _hideTimer.Stop();
                _statusLabel.Text = text ?? string.Empty;
                _progressBar.Style = busy ? ProgressBarStyle.Marquee : ProgressBarStyle.Continuous;
                if (!busy)
                {
                    _progressBar.Minimum = 0;
                    _progressBar.Maximum = 100;
                    _progressBar.Value = 100;
                }
                PositionBottomRight();
                if (!Visible)
                {
                    Show();
                }
                Refresh();
            }

            internal void HideAfter(int milliseconds)
            {
                _hideTimer.Stop();
                _hideTimer.Interval = Math.Max(250, milliseconds);
                _hideTimer.Start();
            }

            private void PositionBottomRight()
            {
                System.Drawing.Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
                Left = workingArea.Right - Width - 18;
                Top = workingArea.Bottom - Height - 18;
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _hideTimer.Stop();
                    _hideTimer.Dispose();
                }
                base.Dispose(disposing);
            }
        }

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
            // Startup resumes durable deletion jobs. It does not enumerate calendars
            // to rebuild subscriptions.
            InitializeTalkRoomLifecycle(
                _settingsStorage.DataDirectory,
                outlookProfileName);
            EnsureApplicationHook();
            EnsureInspectorHook();
            ApplyIfbSettings();
            StartUpdateCheckIfDue();
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
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Upstream update check disabled for IBP fork.");
        }

        private void ScheduleCardDavStartupSync()
        {
            OutlookUiSynchronizationContext uiContext = _uiSynchronizationContext;
            if (uiContext == null)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV startup sync not scheduled because the Outlook UI context is unavailable.");
                return;
            }

            Timer previousTimer = _cardDavStartupDelayTimer;
            _cardDavStartupDelayTimer = null;
            if (previousTimer != null)
            {
                try
                {
                    previousTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to dispose previous CardDAV startup delay timer.",
                        ex);
                }
            }

            _cardDavStartupDelayTimer = new Timer(
                state =>
                {
                    try
                    {
                        uiContext.Post(
                            ignored =>
                            {
                                try
                                {
                                    if (_outlookApplication == null || _currentSettings == null)
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
                                            "CardDAV delayed startup sync skipped because Nextcloud credentials are incomplete.");
                                        return;
                                    }

                                    // Start the recurring timer before the first network attempt.
                                    // If the server is unavailable now, the next 15-minute run retries automatically.
                                    CardDavBackgroundSyncManager.EnsureStarted(
                                        configuration,
                                        _outlookApplication);
                                    DiagnosticsLogger.Log(
                                        LogCategories.Core,
                                        "CardDAV delayed startup sync starting after Outlook startup.");
                                    StartCardDavContactSync();
                                }
                                catch (Exception ex)
                                {
                                    DiagnosticsLogger.LogException(
                                        LogCategories.Core,
                                        "CardDAV delayed startup sync launch failed.",
                                        ex);
                                }
                            },
                            null);
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to post delayed CardDAV startup sync to the Outlook UI thread.",
                            ex);
                    }
                },
                null,
                CardDavStartupDelay,
                Timeout.InfiniteTimeSpan);

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV startup sync scheduled (delaySeconds=15).");
        }

        private void ShowCardDavSyncStatus(string text)
        {
            if (_cardDavSyncStatusForm == null || _cardDavSyncStatusForm.IsDisposed)
            {
                _cardDavSyncStatusForm = new CardDavSyncStatusForm();
            }
            _cardDavSyncStatusForm.SetStatus(text, true);
        }

        private void CompleteCardDavSyncStatus(string text, int hideAfterMilliseconds)
        {
            if (_cardDavSyncStatusForm == null || _cardDavSyncStatusForm.IsDisposed)
            {
                _cardDavSyncStatusForm = new CardDavSyncStatusForm();
            }
            _cardDavSyncStatusForm.SetStatus(text, false);
            _cardDavSyncStatusForm.HideAfter(hideAfterMilliseconds);
        }

        private void DisposeCardDavSyncStatus()
        {
            CardDavSyncStatusForm statusForm = _cardDavSyncStatusForm;
            _cardDavSyncStatusForm = null;
            if (statusForm == null)
            {
                return;
            }
            try
            {
                statusForm.Close();
                statusForm.Dispose();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to dispose CardDAV sync status window.", ex);
            }
        }

        private void StartCardDavContactSync(bool userInitiated = false)
        {
            RunCardDavContactSyncAsync(userInitiated).ContinueWith(
                task => DiagnosticsLogger.LogException(LogCategories.Core,
                    "CardDAV sync launch failed.", task.Exception),
                CancellationToken.None, TaskContinuationOptions.OnlyOnFaulted, TaskScheduler.Default);
        }

        private async Task RunCardDavContactSyncAsync(bool userInitiated)
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
                if (userInitiated) MessageBox.Show("Bitte zuerst die Nextcloud-Zugangsdaten in den Einstellungen ergänzen.", "Kontakte synchronisieren");
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
                if (userInitiated) MessageBox.Show("Bitte die Kontaktsynchronisation und einen Kontaktbereich in den Einstellungen aktivieren.", "Kontakte synchronisieren");
                return;
            }

            if (!CardDavBackgroundSyncManager.TryBeginSync())
            {
                if (userInitiated) MessageBox.Show("Die Kontaktsynchronisation läuft bereits.", "Kontakte synchronisieren");
                return;
            }
            try
            {
                ShowCardDavSyncStatus("NC Connector: Kontakte werden geprüft ...");

                Dictionary<string, string> knownEtags =
                    CardDavReadOnlySync.LoadKnownEtagsFromOutlook(_outlookApplication);
                var sync = new CardDavReadOnlySync(configuration);
                IList<CardDavAddressBook> syncedAddressBooks = null;

                ShowCardDavSyncStatus("NC Connector: Serverdaten werden geladen ...");
                List<CardDavContactRecord> contacts = await Task.Run(() =>
                {
                    IList<CardDavAddressBook> addressBooks =
                        new DavDiscoveryService(configuration).DiscoverAddressBooks();
                    syncedAddressBooks = addressBooks;
                    var result = new List<CardDavContactRecord>();
                    foreach (CardDavAddressBook addressBook in addressBooks)
                    {
                        bool systemAddressBook = IsSystemCardDavAddressBook(addressBook);
                        if ((systemAddressBook && !preferences.SyncCompanyDirectory)
                            || (!systemAddressBook && !preferences.SyncPersonalContacts)) continue;
                        result.AddRange(sync.DownloadContacts(addressBook, knownEtags));
                    }
                    return result;
                }).ConfigureAwait(false);

                int count = 0;
                int groups = 0;
                int groupFolders = 0;

                await RunOnOutlookUiThreadAsync(() =>
                {
                    if (_outlookApplication == null) return;
                    ShowCardDavSyncStatus("NC Connector: Kontakte werden aktualisiert ...");
                    count = sync.ImportIntoOutlook(_outlookApplication, contacts);
                }).ConfigureAwait(false);

                await Task.Delay(CardDavUiPhaseYieldDelay).ConfigureAwait(false);

                await RunOnOutlookUiThreadAsync(() =>
                {
                    if (_outlookApplication == null) return;
                    ShowCardDavSyncStatus("NC Connector: Verteilerlisten werden abgeglichen ...");
                    groups = CardDavContactGroupSync.Reconcile(
                        configuration,
                        _outlookApplication,
                        syncedAddressBooks,
                        preferences);
                }).ConfigureAwait(false);

                await Task.Delay(CardDavUiPhaseYieldDelay).ConfigureAwait(false);

                await RunOnOutlookUiThreadAsync(() =>
                {
                    if (_outlookApplication == null) return;
                    ShowCardDavSyncStatus("NC Connector: Gruppenordner werden abgeglichen ...");
                    groupFolders = CardDavContactFolderSync.Reconcile(
                        configuration,
                        _outlookApplication,
                        syncedAddressBooks,
                        preferences);
                }).ConfigureAwait(false);

                await RunOnOutlookUiThreadAsync(() =>
                {
                    DiagnosticsLogger.Log(LogCategories.Core, "CardDAV sync completed (changedContacts="
                        + count + ", deleted=" + sync.DeletedCount + ", groups=" + groups
                        + ", groupFolders=" + groupFolders + ").");
                    CompleteCardDavSyncStatus(
                        "NC Connector: Synchronisierung abgeschlossen.",
                        userInitiated ? 900 : 2200);
                    if (userInitiated) MessageBox.Show("Nextcloud-Kontakte synchronisiert: " + count
                        + " geändert/neu, " + sync.DeletedCount + " gelöscht, " + groups
                        + " Gruppen und " + groupFolders + " Gruppenordner aktualisiert.", "Kontakte synchronisieren");
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    userInitiated
                        ? "CardDAV sync failed."
                        : "CardDAV automatic sync failed; existing Outlook contacts were kept and the background timer will retry.",
                    ex);
                if (_uiSynchronizationContext != null)
                {
                    _uiSynchronizationContext.Post(_ =>
                    {
                        CompleteCardDavSyncStatus(
                            userInitiated
                                ? "NC Connector: Synchronisierung fehlgeschlagen."
                                : "NC Connector: Server nicht erreichbar - neuer Versuch später.",
                            3500);
                        if (userInitiated)
                        {
                            MessageBox.Show("Kontaktsynchronisation fehlgeschlagen: "
                                + ex.Message, "Kontakte synchronisieren");
                        }
                    }, null);
                }
            }
            finally
            {
                CardDavBackgroundSyncManager.EndSync();
            }
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

        // IDTExtensibility2 requires this callback; runtime wiring happens in OnConnection
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        // Delay network and contact synchronization until Outlook reports startup complete.
        public void OnStartupComplete(ref Array custom)
        {
            ScheduleCardDavStartupSync();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            TearDownAddInState("shutdown", false);
        }

        // Outlook can call both shutdown callbacks, so teardown must stay idempotent
        private void TearDownAddInState(string origin, bool clearOutlookApplication)
        {
            Timer cardDavStartupDelayTimer = _cardDavStartupDelayTimer;
            _cardDavStartupDelayTimer = null;
            if (cardDavStartupDelayTimer != null)
            {
                try
                {
                    cardDavStartupDelayTimer.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to dispose CardDAV startup delay timer during add-in " + (origin ?? "teardown") + ".",
                        ex);
                }
            }

            DisposeCardDavSyncStatus();
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
            _ribbonUis.Clear();
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
