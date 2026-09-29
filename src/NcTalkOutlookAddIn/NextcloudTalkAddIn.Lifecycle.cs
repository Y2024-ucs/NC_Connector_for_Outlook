// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
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
        private const int CardDavUiSliceBudgetMilliseconds = 40;
        private const int CardDavRecentInputWindowMilliseconds = 900;
        private const int CardDavActiveInputBackoffMilliseconds = 300;
        private const int CardDavIdleYieldMilliseconds = 20;
        [StructLayout(LayoutKind.Sequential)]
        private struct LastInputInfo
        {
            internal uint cbSize;
            internal uint dwTime;
        }

        [DllImport("user32.dll")]
        private static extern bool GetLastInputInfo(ref LastInputInfo plii);

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

                                    // Before any sync starts, so the release notes do not overlap other dialogs.
                                    ShowWhatsNewIfUpdated();

                                    // Start the recurring timer before the first network attempt.
                                    // If the server is unavailable now, the next 15-minute run retries automatically.
                                    CardDavBackgroundSyncManager.EnsureStarted(
                                        configuration,
                                        _outlookApplication);
                                    DiagnosticsLogger.Log(
                                        LogCategories.Core,
                                        "CardDAV delayed startup sync starting after Outlook startup.");
                                    StartCardDavContactSync();
                                    StartCalDavSynchronizerProvisioning(configuration);
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

        /// <summary>
        /// Shows the release notes once after an update to a version that has notes.
        /// </summary>
        private void ShowWhatsNewIfUpdated()
        {
            try
            {
                ReleaseNote note = ReleaseNotes.Pending(AddinVersionInfo.GetVersion());
                if (note == null)
                {
                    return;
                }
                // Mark first: a crash inside the dialog must not show it on every start.
                ReleaseNotes.MarkShown(note);
                using (var form = new WhatsNewForm(note))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to show release notes.", ex);
            }
        }

        private void StartCalDavSynchronizerProvisioning(TalkServiceConfiguration configuration)
        {
            RunCalDavSynchronizerProvisioningAsync(configuration).ContinueWith(
                task => DiagnosticsLogger.LogException(LogCategories.Core,
                    "CalDav Synchronizer provisioning failed.", task.Exception),
                CancellationToken.None, TaskContinuationOptions.OnlyOnFaulted, TaskScheduler.Default);
        }

        /// <summary>
        /// Sets up CalDav Synchronizer for the personal Nextcloud calendar when that add-in is installed.
        /// It reads profiles only at Outlook start, so a new profile needs one restart, which the user is asked for.
        /// </summary>
        private async Task RunCalDavSynchronizerProvisioningAsync(TalkServiceConfiguration configuration)
        {
            CalDavSynchronizerOutlookContext context = await RunOnOutlookUiThreadAsync(
                () => ReadCalDavSynchronizerContext()).ConfigureAwait(false);
            if (context == null)
            {
                return;
            }

            CalDavSynchronizerProvisioningPlan plan = await Task.Run(
                () => CalDavSynchronizerProvisioning.Prepare(configuration, context)).ConfigureAwait(false);
            CalDavSynchronizerProvisioningOutcome outcome = plan.Outcome;
            string backupFolderName = null;
            if (outcome == CalDavSynchronizerProvisioningOutcome.CreateRequired)
            {
                // The first sync merges both calendars; keep a copy of the Outlook side before it can run.
                bool backedUp = await RunOnOutlookUiThreadAsync(
                    () => TryBackUpDefaultCalendar(out backupFolderName)).ConfigureAwait(false);
                if (!backedUp)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CalDav Synchronizer profile not created because the calendar backup failed; retrying at next start.");
                    return;
                }
                outcome = await Task.Run(
                    () => CalDavSynchronizerProvisioning.CreateProfile(plan, configuration, context)).ConfigureAwait(false);
            }
            if (outcome != CalDavSynchronizerProvisioningOutcome.Created
                && outcome != CalDavSynchronizerProvisioningOutcome.PasswordUpdated)
            {
                return;
            }

            string text = outcome == CalDavSynchronizerProvisioningOutcome.Created
                ? "Ihr Nextcloud-Kalender wurde für Outlook eingerichtet."
                : "Die Zugangsdaten Ihres Nextcloud-Kalenders wurden aktualisiert.";
            if (!string.IsNullOrWhiteSpace(backupFolderName))
            {
                text += Environment.NewLine + Environment.NewLine
                    + "Eine Sicherungskopie Ihres bisherigen Outlook-Kalenders liegt im Ordner \""
                    + backupFolderName + "\". Sie können sie löschen, sobald alle Termine wie gewünscht angezeigt werden.";
            }
            await RunOnOutlookUiThreadAsync(() =>
            {
                MessageBox.Show(
                    text + Environment.NewLine + Environment.NewLine
                    + "Damit die Einrichtung abgeschlossen werden kann, starten Sie bitte Outlook neu. "
                    + "Erst dann wird das von NC Connector erzeugte Kalender-Profil eingelesen.",
                    "NC Connector – Kalender",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// Copies the default calendar next to it as "Kalender - Sicherung vor Nextcloud-Sync (date)".
        /// An empty calendar needs no copy. Returns false when the copy failed.
        /// </summary>
        private bool TryBackUpDefaultCalendar(out string backupFolderName)
        {
            backupFolderName = null;
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder calendar = null;
            Outlook.MAPIFolder parent = null;
            Outlook.MAPIFolder copy = null;
            Outlook.Items items = null;
            try
            {
                session = _outlookApplication.Session;
                calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                items = calendar.Items;
                if (items == null || items.Count == 0)
                {
                    return true;
                }

                parent = calendar.Parent as Outlook.MAPIFolder;
                if (parent == null)
                {
                    return false;
                }

                // A copy outside the default calendar raises no reminders and is not part of the sync.
                copy = calendar.CopyTo(parent);
                string baseName = "Kalender - Sicherung vor Nextcloud-Sync ("
                    + DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + ")";
                string name = baseName;
                for (int attempt = 2; attempt < 100; attempt++)
                {
                    try
                    {
                        copy.Name = name;
                        break;
                    }
                    catch (COMException)
                    {
                        // A folder with that name exists already.
                        name = baseName + " " + attempt;
                    }
                }
                backupFolderName = copy.Name;
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "Outlook calendar backed up before CalDav Synchronizer setup (folder=" + backupFolderName
                    + ", items=" + items.Count + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Outlook calendar backup failed.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release calendar items after backup.");
                ComInteropScope.TryRelease(copy, LogCategories.Core, "Failed to release calendar backup folder.");
                ComInteropScope.TryRelease(parent, LogCategories.Core, "Failed to release calendar parent folder.");
                ComInteropScope.TryRelease(calendar, LogCategories.Core, "Failed to release default calendar after backup.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after calendar backup.");
            }
        }

        private CalDavSynchronizerOutlookContext ReadCalDavSynchronizerContext()
        {
            if (_outlookApplication == null || !IsComAddInInstalled(CalDavSynchronizerProvisioning.AddInProgId))
            {
                return null;
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder calendar = null;
            Outlook.Store store = null;
            try
            {
                session = _outlookApplication.Session;
                calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                store = calendar.Store;
                return new CalDavSynchronizerOutlookContext
                {
                    OutlookProfileName = session.CurrentProfileName,
                    CalendarEntryId = calendar.EntryID,
                    CalendarStoreId = calendar.StoreID,
                    AccountName = store != null ? store.DisplayName : string.Empty,
                    EmailAddress = FindAccountSmtpAddress(session, calendar.StoreID)
                };
            }
            finally
            {
                ComInteropScope.TryRelease(store, LogCategories.Core, "Failed to release calendar store after CalDav Synchronizer scan.");
                ComInteropScope.TryRelease(calendar, LogCategories.Core, "Failed to release default calendar after CalDav Synchronizer scan.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CalDav Synchronizer scan.");
            }
        }

        private bool IsComAddInInstalled(string progId)
        {
            COMAddIns addIns = null;
            try
            {
                addIns = _outlookApplication.COMAddIns;
                int count = addIns != null ? addIns.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    object key = index;
                    COMAddIn addIn = null;
                    try
                    {
                        addIn = addIns.Item(ref key);
                        if (addIn != null && string.Equals(addIn.ProgId, progId, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(addIn, LogCategories.Core, "Failed to release COM add-in entry.");
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to enumerate Outlook COM add-ins.", ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(addIns, LogCategories.Core, "Failed to release COM add-ins collection.");
            }
        }

        private static string FindAccountSmtpAddress(Outlook.NameSpace session, string storeId)
        {
            Outlook.Accounts accounts = null;
            string fallback = string.Empty;
            try
            {
                accounts = session.Accounts;
                int count = accounts != null ? accounts.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.Account account = null;
                    Outlook.Store delivery = null;
                    try
                    {
                        account = accounts[index];
                        string address = account.SmtpAddress ?? string.Empty;
                        if (fallback.Length == 0) fallback = address;
                        delivery = account.DeliveryStore;
                        if (delivery != null && string.Equals(delivery.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                        {
                            return address;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(delivery, LogCategories.Core, "Failed to release account delivery store.");
                        ComInteropScope.TryRelease(account, LogCategories.Core, "Failed to release Outlook account.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(accounts, LogCategories.Core, "Failed to release Outlook accounts.");
            }
            return fallback;
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
                || (!preferences.SyncPersonalContacts && !preferences.SyncCustomers))
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
                    CardDavReadOnlySync.LoadKnownEtagsFromOutlook(
                        _outlookApplication,
                        preferences.UseDefaultContactsFolder);
                var sync = new CardDavReadOnlySync(configuration);
                IList<CardDavAddressBook> syncedAddressBooks = null;

                ShowCardDavSyncStatus("NC Connector: Serverdaten werden geladen ...");
                List<CardDavContactRecord> contacts = await Task.Run(() =>
                {
                    IList<CardDavAddressBook> addressBooks = CardDavCustomerSync.WithoutCustomerAddressBook(
                        new DavDiscoveryService(configuration).DiscoverAddressBooks());
                    syncedAddressBooks = addressBooks;
                    var result = new List<CardDavContactRecord>();
                    foreach (CardDavAddressBook addressBook in addressBooks)
                    {
                        if (!preferences.IncludesAddressBook(addressBook, addressBooks)) continue;
                        result.AddRange(sync.DownloadContacts(addressBook, knownEtags));
                    }
                    return result;
                }).ConfigureAwait(false);

                int count = 0;
                int groups = 0;
                int groupFolders = 0;
                bool hasOutlookChanges = contacts.Count > 0 || sync.HasPendingDeletions;

                if (contacts.Count > 0)
                {
                    await RunOnOutlookUiThreadAsync(() =>
                    {
                        ShowCardDavSyncStatus("NC Connector: Kontakte werden aktualisiert ...");
                    }).ConfigureAwait(false);

                    count = await ImportCardDavContactsResponsivelyAsync(
                        sync,
                        contacts,
                        preferences.UseDefaultContactsFolder).ConfigureAwait(false);
                }

                if (sync.HasPendingDeletions)
                {
                    await WaitForCardDavUiWindowAsync().ConfigureAwait(false);
                    await RunOnOutlookUiThreadAsync(() =>
                    {
                        if (_outlookApplication == null) return;
                        ShowCardDavSyncStatus("NC Connector: Gelöschte Kontakte werden abgeglichen ...");
                        sync.ReconcileDeletedContacts(_outlookApplication);
                    }).ConfigureAwait(false);
                }

                if (hasOutlookChanges)
                {
                    await WaitForCardDavUiWindowAsync().ConfigureAwait(false);
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

                    await WaitForCardDavUiWindowAsync().ConfigureAwait(false);
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
                }
                else
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV Outlook write phases skipped because no remote contact changes were detected.");
                }

                string customerSummary = string.Empty;
                if (preferences.SyncCustomers)
                {
                    await RunOnOutlookUiThreadAsync(() =>
                    {
                        ShowCardDavSyncStatus("NC Connector: Kunden werden abgeglichen ...");
                    }).ConfigureAwait(false);
                    try
                    {
                        CardDavCustomerSyncResult customers = await new CardDavCustomerSync(
                            configuration,
                            _outlookApplication,
                            callback => RunOnOutlookUiThreadAsync(callback),
                            WaitForCardDavUiWindowAsync).RunAsync().ConfigureAwait(false);
                        customerSummary = Environment.NewLine + customers.Summary();
                    }
                    catch (Exception ex)
                    {
                        // The read-only contact sync above stays valid; report the customer part separately.
                        DiagnosticsLogger.LogException(LogCategories.Core, "CardDAV customer sync failed.", ex);
                        customerSummary = Environment.NewLine + "Kunden-Abgleich fehlgeschlagen: " + ex.Message;
                    }
                }

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
                        + " Gruppen und " + groupFolders + " Gruppenordner aktualisiert." + customerSummary, "Kontakte synchronisieren");
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

        private async Task<int> ImportCardDavContactsResponsivelyAsync(
            CardDavReadOnlySync sync,
            IList<CardDavContactRecord> contacts,
            bool useDefaultContactsFolder)
        {
            if (sync == null || contacts == null || contacts.Count == 0)
            {
                return 0;
            }

            int index = 0;
            int imported = 0;
            var importCache = new CardDavImportCache();
            while (index < contacts.Count)
            {
                int sliceImported = 0;
                await RunOnOutlookUiThreadAsync(() =>
                {
                    if (_outlookApplication == null)
                    {
                        return;
                    }

                    var stopwatch = Stopwatch.StartNew();
                    do
                    {
                        CardDavContactRecord contact = contacts[index];
                        sliceImported += sync.ImportIntoOutlook(
                            _outlookApplication,
                            new[] { contact },
                            useDefaultContactsFolder,
                            false,
                            false,
                            importCache);
                        index++;
                    }
                    while (index < contacts.Count
                        && stopwatch.ElapsedMilliseconds < CardDavUiSliceBudgetMilliseconds);
                }).ConfigureAwait(false);

                imported += sliceImported;

                if (index < contacts.Count)
                {
                    await WaitForCardDavUiWindowAsync().ConfigureAwait(false);
                }
            }

            return imported;
        }

        private static async Task WaitForCardDavUiWindowAsync()
        {
            int delay = WasUserRecentlyActive()
                ? CardDavActiveInputBackoffMilliseconds
                : CardDavIdleYieldMilliseconds;
            await Task.Delay(delay).ConfigureAwait(false);
        }

        private static bool WasUserRecentlyActive()
        {
            try
            {
                var info = new LastInputInfo();
                info.cbSize = (uint)Marshal.SizeOf(typeof(LastInputInfo));
                if (!GetLastInputInfo(ref info))
                {
                    return false;
                }

                uint now = unchecked((uint)Environment.TickCount);
                uint idleMilliseconds = unchecked(now - info.dwTime);
                return idleMilliseconds <= CardDavRecentInputWindowMilliseconds;
            }
            catch
            {
                return false;
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
